"""Small Bright Data asynchronous API client using the Python standard library."""

from __future__ import annotations

import json
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


DEFAULT_DATASET_ID = "gd_l7q7dkf244hwjntr0"
DEFAULT_BASE_URL = "https://api.brightdata.com/datasets/v3"
IN_PROGRESS_STATUSES = {"starting", "running", "collecting", "digesting"}
TERMINAL_FAILURE_STATUSES = {"failed", "canceled", "cancelled"}
TRANSIENT_HTTP_STATUSES = {429, 500, 502, 503, 504}


class BrightDataError(RuntimeError):
    """Raised for a Bright Data API or collection failure."""


@dataclass(frozen=True)
class SnapshotProgress:
    snapshot_id: str
    status: str
    response: dict[str, Any]


class BrightDataClient:
    def __init__(
        self,
        api_key: str,
        *,
        dataset_id: str = DEFAULT_DATASET_ID,
        base_url: str = DEFAULT_BASE_URL,
        request_timeout: float = 60.0,
        retry_attempts: int = 3,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        if not api_key or not api_key.strip():
            raise BrightDataError("A Bright Data API key is required.")
        self._api_key = api_key.strip()
        self.dataset_id = dataset_id
        self.base_url = base_url.rstrip("/")
        self.request_timeout = request_timeout
        self.retry_attempts = max(1, retry_attempts)
        self._sleep = sleep

    @property
    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "amazon-data-extractor/0.1.0",
        }

    def _request_json(
        self,
        method: str,
        path: str,
        *,
        query: dict[str, object] | None = None,
        payload: object | None = None,
    ) -> Any:
        url = f"{self.base_url}/{path.lstrip('/')}"
        if query:
            url = f"{url}?{urlencode(query)}"

        body = None if payload is None else json.dumps(payload).encode("utf-8")
        request = Request(url, data=body, headers=self._headers, method=method)

        for attempt in range(1, self.retry_attempts + 1):
            try:
                with urlopen(request, timeout=self.request_timeout) as response:
                    response_body = response.read().decode("utf-8-sig")
                    if not response_body:
                        return None
                    return json.loads(response_body)
            except HTTPError as exc:
                error_body = exc.read().decode("utf-8", errors="replace")
                if exc.code in TRANSIENT_HTTP_STATUSES and attempt < self.retry_attempts:
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    delay = float(retry_after) if retry_after and retry_after.isdigit() else 2**attempt
                    self._sleep(delay)
                    continue
                message = _extract_error_message(error_body) or exc.reason or "HTTP error"
                raise BrightDataError(f"Bright Data returned HTTP {exc.code}: {message}") from exc
            except URLError as exc:
                if attempt < self.retry_attempts:
                    self._sleep(2**attempt)
                    continue
                raise BrightDataError(f"Could not reach Bright Data: {exc.reason}") from exc
            except json.JSONDecodeError as exc:
                raise BrightDataError("Bright Data returned a response that was not valid JSON.") from exc

        raise BrightDataError("Bright Data request failed after all retry attempts.")

    def trigger_collection(
        self,
        asins: Sequence[str],
        *,
        language: str = "EN",
        zipcode: str | None = None,
    ) -> str:
        if not asins:
            raise BrightDataError("At least one ASIN is required.")

        payload: list[dict[str, object]] = []
        for asin in asins:
            item: dict[str, object] = {
                "url": f"https://www.amazon.com/dp/{asin}?th=1&psc=1&language=en_US",
                "language": language.upper(),
                "all_variations": False,
            }
            if zipcode:
                item["zipcode"] = zipcode
            payload.append(item)

        response = self._request_json(
            "POST",
            "trigger",
            query={
                "dataset_id": self.dataset_id,
                "format": "json",
                "include_errors": "true",
            },
            payload=payload,
        )
        if not isinstance(response, dict) or not response.get("snapshot_id"):
            raise BrightDataError("Bright Data did not return a snapshot ID.")
        return str(response["snapshot_id"])

    def get_progress(self, snapshot_id: str) -> SnapshotProgress:
        response = self._request_json("GET", f"progress/{snapshot_id}")
        if not isinstance(response, dict):
            raise BrightDataError("Bright Data returned an invalid progress response.")
        status = str(response.get("status", "")).strip().lower()
        if not status:
            raise BrightDataError("Bright Data progress response did not contain a status.")
        return SnapshotProgress(snapshot_id=snapshot_id, status=status, response=response)

    def wait_until_ready(
        self,
        snapshot_id: str,
        *,
        poll_interval: float = 10.0,
        timeout_seconds: float = 3600.0,
        on_progress: Callable[[SnapshotProgress], None] | None = None,
    ) -> SnapshotProgress:
        started = time.monotonic()
        last_status: str | None = None

        while True:
            progress = self.get_progress(snapshot_id)
            if on_progress and progress.status != last_status:
                on_progress(progress)
            last_status = progress.status

            if progress.status == "ready":
                return progress
            if progress.status in TERMINAL_FAILURE_STATUSES:
                detail = progress.response.get("error") or progress.response.get("message")
                suffix = f": {detail}" if detail else ""
                raise BrightDataError(f"Snapshot {snapshot_id} {progress.status}{suffix}")
            if progress.status not in IN_PROGRESS_STATUSES:
                raise BrightDataError(
                    f"Snapshot {snapshot_id} returned unknown status '{progress.status}'."
                )
            if time.monotonic() - started >= timeout_seconds:
                raise BrightDataError(
                    f"Timed out waiting for snapshot {snapshot_id}. "
                    "The snapshot ID was saved and can be checked in Bright Data."
                )
            self._sleep(max(1.0, poll_interval))

    def download_results(self, snapshot_id: str) -> list[dict[str, Any]]:
        response = self._request_json(
            "GET",
            f"snapshot/{snapshot_id}",
            query={"format": "json"},
        )
        if not isinstance(response, list):
            raise BrightDataError("Bright Data snapshot was not returned as a JSON record list.")
        if not all(isinstance(record, dict) for record in response):
            raise BrightDataError("Bright Data snapshot contained an invalid product record.")
        return response


def _extract_error_message(body: str) -> str:
    try:
        parsed = json.loads(body)
    except json.JSONDecodeError:
        return body.strip()[:500]
    if isinstance(parsed, dict):
        for key in ("error", "message", "detail"):
            if parsed.get(key):
                return str(parsed[key])[:500]
    return str(parsed)[:500]

