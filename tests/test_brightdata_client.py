import unittest
from unittest.mock import Mock

from amazon_data_extractor.brightdata_client import (
    BrightDataClient,
    BrightDataError,
    SnapshotProgress,
)


class StubClient(BrightDataClient):
    def __init__(self, responses):
        super().__init__("test-key", sleep=lambda _: None)
        self.responses = iter(responses)
        self.calls = []

    def _request_json(self, method, path, *, query=None, payload=None):
        self.calls.append((method, path, query, payload))
        return next(self.responses)


class BrightDataClientTests(unittest.TestCase):
    def test_trigger_uses_one_child_url_per_asin_and_disables_variations(self):
        client = StubClient([{"snapshot_id": "s_test"}])

        snapshot_id = client.trigger_collection(["B0G38V6MRN", "B0CHHSFMRL"])

        self.assertEqual(snapshot_id, "s_test")
        method, path, query, payload = client.calls[0]
        self.assertEqual((method, path), ("POST", "trigger"))
        self.assertEqual(query["dataset_id"], client.dataset_id)
        self.assertEqual(len(payload), 2)
        self.assertFalse(payload[0]["all_variations"])
        self.assertIn("B0G38V6MRN", payload[0]["url"])

    def test_wait_accepts_documented_and_observed_progress_statuses(self):
        client = BrightDataClient("test-key", sleep=lambda _: None)
        client.get_progress = Mock(
            side_effect=[
                SnapshotProgress("s_test", "starting", {}),
                SnapshotProgress("s_test", "collecting", {}),
                SnapshotProgress("s_test", "digesting", {}),
                SnapshotProgress("s_test", "ready", {}),
            ]
        )

        result = client.wait_until_ready("s_test", poll_interval=1, timeout_seconds=30)

        self.assertEqual(result.status, "ready")

    def test_wait_raises_on_failed_snapshot(self):
        client = BrightDataClient("test-key", sleep=lambda _: None)
        client.get_progress = Mock(
            return_value=SnapshotProgress("s_test", "failed", {"error": "bad URL"})
        )

        with self.assertRaisesRegex(BrightDataError, "bad URL"):
            client.wait_until_ready("s_test")

    def test_download_requires_record_list(self):
        client = StubClient([{"unexpected": "object"}])

        with self.assertRaisesRegex(BrightDataError, "JSON record list"):
            client.download_results("s_test")


if __name__ == "__main__":
    unittest.main()

