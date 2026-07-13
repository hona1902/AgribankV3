from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from agribank_v3.file_merge import merge_same_structure_csv_to_csv
from agribank_v3.settings import BranchProfile
from agribank_v3.settlement import SETTLEMENT_SPECS, SettlementOptions, SettlementRequest
from agribank_v3.settlement.processors import Mau1314Processor


def _log(message: str) -> None:
    print(message, flush=True)


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python -m agribank_v3.settlement.consolidation13_worker <request.json>", file=sys.stderr)
        return 2
    request_path = Path(sys.argv[1])
    payload = json.loads(request_path.read_text(encoding="utf-8"))
    source_paths = tuple(Path(path) for path in payload["source_paths"])
    output_path = Path(payload["output_path"])
    merged_csv_path = Path(payload["merged_csv_path"])
    profile = BranchProfile(**payload["profile"])
    options = SettlementOptions(**payload["options"])
    spec = SETTLEMENT_SPECS["consolidation.13"]
    started = time.perf_counter()
    merge_result = None
    try:
        _log(f"[TongHop13] start: {len(source_paths)} source files -> {output_path}")
        step = time.perf_counter()
        merge_result = merge_same_structure_csv_to_csv(source_paths, merged_csv_path)
        _log(
            f"[TongHop13] merge done in {time.perf_counter() - step:.1f}s, "
            f"rows={merge_result.row_count}"
        )
        processor = Mau1314Processor()
        request = SettlementRequest(
            spec=spec,
            profile=profile,
            options=options,
            source_paths=(merged_csv_path,),
        )
        step = time.perf_counter()
        records, report_date, currency_order = processor.read_source(
            request,
            merged_csv_path,
        )
        _log(
            f"[TongHop13] read/sort done in {time.perf_counter() - step:.1f}s, "
            f"records={len(records)}"
        )
        step = time.perf_counter()
        processor.save_mau13_streaming_workbook(
            request,
            records,
            report_date,
            currency_order,
            output_path,
        )
        _log(
            f"[TongHop13] save done in {time.perf_counter() - step:.1f}s; "
            f"total={time.perf_counter() - started:.1f}s"
        )
        _log(
            "[TongHop13] completed: "
            f"output={output_path}; "
            f"sources={merge_result.source_count if merge_result else 0}; "
            f"merged_rows={merge_result.row_count if merge_result else 0}; "
            f"processed_rows={len(records)}"
        )
        return 0
    finally:
        if merged_csv_path.exists():
            try:
                merged_csv_path.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
