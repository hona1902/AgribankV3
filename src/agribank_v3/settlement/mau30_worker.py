from __future__ import annotations

from dataclasses import asdict
import json
import sys
import time
from pathlib import Path

from agribank_v3.settings import BranchProfile
from agribank_v3.settlement import SETTLEMENT_SPECS, SettlementOptions, SettlementRequest
from agribank_v3.settlement.processors import Mau30Processor


def _log(message: str) -> None:
    print(message, flush=True)


def main() -> int:
    if len(sys.argv) != 2:
        print(
            "Usage: python -m agribank_v3.settlement.mau30_worker <request.json>",
            file=sys.stderr,
        )
        return 2
    request_path = Path(sys.argv[1])
    payload = json.loads(request_path.read_text(encoding="utf-8"))
    spec = SETTLEMENT_SPECS[payload["spec_key"]]
    profile = BranchProfile(**payload["profile"])
    options = SettlementOptions(**payload["options"])
    source_paths = tuple(Path(path) for path in payload["source_paths"])
    started = time.perf_counter()
    try:
        _log(
            f"[Mau30] start: spec={spec.key}; model={options.source_report_code}; "
            f"file={source_paths[0]}"
        )
        result = Mau30Processor().execute(
            SettlementRequest(
                spec=spec,
                profile=profile,
                options=options,
                source_paths=source_paths,
            )
        )
        _log(
            f"[Mau30] completed in {time.perf_counter() - started:.1f}s; "
            f"sheet={result.worksheet_name}; rows={result.processed_rows}; "
            f"output={result.output_path}"
        )
        print(
            json.dumps(
                {
                    "output_path": str(result.output_path),
                    "worksheet_name": result.worksheet_name,
                    "processed_rows": result.processed_rows,
                    "profile": asdict(profile),
                },
                ensure_ascii=True,
            ),
            flush=True,
        )
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
