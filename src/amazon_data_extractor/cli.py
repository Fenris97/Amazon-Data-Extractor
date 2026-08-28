"""Command-line entry point for the first Amazon extraction workflow."""

from __future__ import annotations

import argparse
import getpass
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from . import __version__
from .asin_reader import InputWorkbookError, read_asins
from .brightdata_client import BrightDataClient, BrightDataError, SnapshotProgress
from .exporter import write_outputs


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="amazon-data-extractor",
        description="Read ASINs from Excel and collect current Amazon data through Bright Data.",
    )
    parser.add_argument("input_file", type=Path, help="Path to the .xlsx workbook")
    parser.add_argument("--sheet", help="Worksheet to read (default: active worksheet)")
    parser.add_argument("--column", help="ASIN column name (default: auto-detect)")
    parser.add_argument("--header-row", type=int, default=1, help="Header row number (default: 1)")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Parent output directory (default: output)",
    )
    parser.add_argument("--language", default="EN", help="Amazon language (default: EN)")
    parser.add_argument("--zipcode", help="Optional ZIP code for location-specific results")
    parser.add_argument(
        "--poll-seconds",
        type=float,
        default=10.0,
        help="Seconds between progress checks (default: 10)",
    )
    parser.add_argument(
        "--timeout-minutes",
        type=float,
        default=60.0,
        help="Maximum wait time in minutes (default: 60)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Validate Excel without calling Bright Data")
    parser.add_argument("--yes", action="store_true", help="Skip the submission confirmation")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _confirm_submission(count: int) -> bool:
    print()
    print(f"ASINs to submit: {count}")
    print(f"Estimated successfully delivered records: about {count}")
    answer = input("Submit this batch to Bright Data? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def _save_initial_metadata(run_dir: Path, metadata: dict[str, object]) -> Path:
    run_dir.mkdir(parents=True, exist_ok=True)
    metadata_path = run_dir / "run_metadata.json"
    metadata_path.write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return metadata_path


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    try:
        read_result = read_asins(
            args.input_file,
            sheet_name=args.sheet,
            column_name=args.column,
            header_row=args.header_row,
        )
    except InputWorkbookError as exc:
        print(f"Input error: {exc}", file=sys.stderr)
        return 2

    print(f"Worksheet: {read_result.sheet_name}")
    print(f"ASIN column: {read_result.column_name}")
    print(f"Valid unique ASINs: {len(read_result.asins)}")
    print(f"Duplicates skipped: {read_result.duplicate_count}")
    print(f"Invalid rows skipped: {read_result.invalid_count}")

    if not read_result.asins:
        print("No valid ASINs were found; nothing was submitted.", file=sys.stderr)
        return 2

    if args.dry_run:
        print("Dry run complete. Bright Data was not called and no records were used.")
        return 0

    if not args.yes and not _confirm_submission(len(read_result.asins)):
        print("Canceled. Bright Data was not called.")
        return 0

    api_key = os.environ.get("BRIGHTDATA_API_KEY", "").strip()
    if not api_key:
        api_key = getpass.getpass("Bright Data API key (hidden): ").strip()
    if not api_key:
        print("A Bright Data API key is required.", file=sys.stderr)
        return 2

    started_at = datetime.now(timezone.utc)
    run_name = f"amazon_extraction_{started_at.strftime('%Y%m%d_%H%M%S')}"
    run_dir = args.output_dir / run_name
    client = BrightDataClient(api_key)

    try:
        print("Submitting asynchronous batch...")
        snapshot_id = client.trigger_collection(
            read_result.asins,
            language=args.language,
            zipcode=args.zipcode,
        )
        metadata: dict[str, object] = {
            "schema_version": "1.0",
            "application_version": __version__,
            "started_at_utc": started_at.isoformat(),
            "source_workbook": args.input_file.name,
            "source_worksheet": read_result.sheet_name,
            "source_asin_column": read_result.column_name,
            "submitted_asin_count": len(read_result.asins),
            "invalid_row_count": read_result.invalid_count,
            "duplicate_row_count": read_result.duplicate_count,
            "dataset_id": client.dataset_id,
            "snapshot_id": snapshot_id,
            "all_variations": False,
            "language": args.language.upper(),
            "zipcode": args.zipcode or "",
            "status": "submitted",
        }
        metadata_path = _save_initial_metadata(run_dir, metadata)
        print(f"Snapshot ID: {snapshot_id}")
        print(f"Recovery information saved to: {metadata_path.resolve()}")

        def show_progress(progress: SnapshotProgress) -> None:
            print(f"Bright Data status: {progress.status}")

        client.wait_until_ready(
            snapshot_id,
            poll_interval=args.poll_seconds,
            timeout_seconds=args.timeout_minutes * 60,
            on_progress=show_progress,
        )
        print("Downloading completed snapshot...")
        records = client.download_results(snapshot_id)

        completed_at = datetime.now(timezone.utc)
        metadata.update(
            {
                "completed_at_utc": completed_at.isoformat(),
                "duration_seconds": round((completed_at - started_at).total_seconds(), 3),
                "returned_record_count": len(records),
                "error_record_count": sum(bool(record.get("error")) for record in records),
                "warning_record_count": sum(bool(record.get("warning")) for record in records),
                "status": "completed",
            }
        )
        raw_path, workbook_path, _ = write_outputs(
            run_dir,
            records=records,
            audit_rows=read_result.audit_rows,
            metadata=metadata,
        )
        print(f"Returned records: {len(records)}")
        print(f"Excel workbook: {workbook_path.resolve()}")
        print(f"Untouched JSON: {raw_path.resolve()}")
        return 0
    except BrightDataError as exc:
        print(f"Bright Data error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nStopped by user. Any displayed snapshot ID remains available in Bright Data.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())

