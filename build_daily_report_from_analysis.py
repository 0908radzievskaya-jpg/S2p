from __future__ import annotations

import argparse
import datetime as dt
import json
import re
from pathlib import Path
from typing import List, Sequence

from openpyxl import Workbook

import mail_automation as ma


ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def item_from_analysis(path: Path) -> ma.ProcessedMessage:
    payload = json.loads(path.read_text(encoding="utf-8"))
    fields = {str(k): str(v) for k, v in (payload.get("fields") or {}).items()}
    attachments = [str(x) for x in payload.get("attachments", [])]
    links = [str(x) for x in payload.get("links", [])]
    downloaded_files = [str(x) for x in payload.get("downloaded_files", [])]
    subject = str(payload.get("subject", ""))
    sender = str(payload.get("sender", ""))
    fallback_field_text = "\n".join(
        str(fields.get(key, ""))
        for key in ("requirements", "constraints", "tech_params", "volumes")
    )
    fallback_text = "\n".join([
        subject,
        sender,
        fallback_field_text,
    ])
    is_relevant, relevance_reason = ma.assess_design_request_relevance(
        subject,
        sender,
        fallback_text,
        fallback_text,
        attachments,
        links,
        downloaded_files,
    )
    fields["is_relevant"] = "yes" if is_relevant else "no"
    fields["relevance_reason"] = relevance_reason
    return ma.ProcessedMessage(
        mailbox=str(payload.get("mailbox", "")),
        uid=str(payload.get("uid", "")),
        message_id=str(payload.get("message_id", "")),
        subject=str(payload.get("subject", "")),
        sender=str(payload.get("sender", "")),
        date_folder=path.parts[-5] if len(path.parts) >= 5 else dt.date.today().isoformat(),
        output_dir=str(path.parent.resolve()),
        attachments=attachments,
        links=links,
        downloaded_files=downloaded_files,
        sections=[str(x) for x in payload.get("sections", [])],
        extracted_fields=fields,
        notes=[str(x) for x in payload.get("notes", [])],
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build full daily report from analysis.json files")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--date", default=dt.date.today().isoformat())
    args = parser.parse_args(argv)

    config = ma.load_config(Path(args.config))
    day_root = ma.get_relevant_output_root(config) / args.date
    if not day_root.exists():
        day_root = Path(str(config.get("output_root", "archive"))) / args.date
    report_path = Path(str(config.get("report_root", "reports"))) / f"report-{args.date}.xlsx"

    items: List[ma.ProcessedMessage] = [item_from_analysis(path) for path in sorted(day_root.rglob("analysis.json"))]
    relevant_items = [item for item in items if item.extracted_fields.get("is_relevant") == "yes"]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Исходные данные"
    sheet.append(ma.REPORT_HEADERS)
    pir_sheet = workbook.create_sheet("ТЭП и ПИР")
    pir_sheet.append(ma.PIR_ESTIMATE_HEADERS)
    processed_at = dt.datetime.now().isoformat(timespec="seconds")
    for item in relevant_items:
        row = [ILLEGAL_XLSX_RE.sub("", str(value)) if value is not None else "" for value in ma.build_report_row(processed_at, item)]
        sheet.append(row)
        pir_row = [
            ILLEGAL_XLSX_RE.sub("", str(value)) if value is not None else ""
            for value in ma.build_pir_estimate_row(processed_at, item)
        ]
        pir_sheet.append(pir_row)
    ma.ensure_dir(report_path.parent)
    saved_report_path = report_path
    try:
        workbook.save(saved_report_path)
    except PermissionError:
        timestamp = dt.datetime.now().strftime("%H-%M-%S")
        saved_report_path = report_path.with_name(f"{report_path.stem}_filtered_{timestamp}{report_path.suffix}")
        workbook.save(saved_report_path)
    print(json.dumps({
        "processed_messages": len(relevant_items),
        "scanned_messages": len(items),
        "report_path": str(saved_report_path.resolve()),
        "requested_report_path": str(report_path.resolve()),
        "date": args.date,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
