from __future__ import annotations

import argparse
import datetime as dt
import email
import json
import re
from pathlib import Path
from typing import List, Sequence

from openpyxl import Workbook

import mail_automation as ma


ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def load_message(base_dir: Path):
    eml_path = base_dir / "message.eml"
    if not eml_path.exists():
        return None
    return email.message_from_bytes(eml_path.read_bytes())


def list_existing_attachments(base_dir: Path) -> List[str]:
    attachment_dir = base_dir / "attachments"
    if not attachment_dir.exists():
        return []
    return [str(path.resolve()) for path in sorted(attachment_dir.rglob("*")) if path.is_file()]


def rebuild_message_folder(base_dir: Path, config: dict[str, object]) -> ma.ProcessedMessage | None:
    analysis_path = base_dir / "analysis.json"
    if not analysis_path.exists():
        return None

    payload = json.loads(analysis_path.read_text(encoding="utf-8"))
    message = load_message(base_dir)

    subject = str(payload.get("subject", ""))
    sender = str(payload.get("sender", ""))
    message_id = str(payload.get("message_id", ""))
    uid = str(payload.get("uid", ""))
    mailbox = str(payload.get("mailbox", ""))

    body_text = ""
    html_body = ""
    if message is not None:
        subject = ma.decode_mime(message.get("Subject")) or subject
        sender = ma.sanitize_sender(message.get("From", "")) or sender
        body_text, html_body = ma.extract_message_bodies(message)

    message_txt_path = base_dir / "message.txt"
    saved_body_text = message_txt_path.read_text(encoding="utf-8", errors="ignore") if message_txt_path.exists() else body_text
    links = ma.extract_links(body_text or saved_body_text, html_body)

    refreshed_download_dir = base_dir / "downloads_refreshed"
    ma.ensure_dir(refreshed_download_dir)
    downloaded_files: List[str] = []
    for link in links:
        resources = ma.download_link_graph(
            link,
            refreshed_download_dir,
            int(config.get("download_timeout_seconds", 30)),
            str(config.get("user_agent", "mail-intake-automation/1.0")),
            int(config.get("link_follow_depth", 2)),
            int(config.get("max_downloaded_files_per_link", 20)),
        )
        if resources:
            downloaded_files.extend(resource.local_path for resource in resources)

    attachments = list_existing_attachments(base_dir)
    extracted_texts = [subject, saved_body_text, ma.html_to_text(html_body)]
    notes: List[str] = []
    for path_string in attachments + downloaded_files:
        text, item_notes = ma.extract_text_from_path(Path(path_string))
        if text:
            extracted_texts.append(text)
            lowered = text.lower()
            if "ошибка 404" in lowered or "ссылка устарела" in lowered:
                notes.append(f"Недоступная страница по ссылке: {Path(path_string).name}")
            if "авториз" in lowered and "тендер" in lowered:
                notes.append(f"Вероятно требуется авторизация для доступа к материалам: {Path(path_string).name}")
        notes.extend(item_notes)
    if links and not downloaded_files:
        notes.append("Ссылки найдены, но полезные материалы по ним повторно скачать не удалось")

    combined_text = "\n".join(text for text in extracted_texts if text)
    fields = ma.extract_field_bundle(subject, saved_body_text, combined_text, sender)
    sections = ma.classify_sections(extracted_texts)
    is_relevant, relevance_reason = ma.assess_design_request_relevance(
        subject,
        sender,
        saved_body_text,
        combined_text,
        attachments,
        links,
        downloaded_files,
    )
    fields["is_relevant"] = "yes" if is_relevant else "no"
    fields["relevance_reason"] = relevance_reason

    item = ma.ProcessedMessage(
        mailbox=mailbox,
        uid=uid,
        message_id=message_id,
        subject=subject,
        sender=sender,
        date_folder=base_dir.parts[-4] if len(base_dir.parts) >= 4 else dt.date.today().isoformat(),
        output_dir=str(base_dir.resolve()),
        attachments=attachments,
        links=links,
        downloaded_files=[str(Path(path).resolve()) for path in downloaded_files],
        sections=sections,
        extracted_fields=fields,
        notes=notes,
    )

    summary_payload = {
        "mailbox": item.mailbox,
        "uid": item.uid,
        "message_id": item.message_id,
        "subject": item.subject,
        "sender": item.sender,
        "sections": item.sections,
        "fields": item.extracted_fields,
        "attachments": item.attachments,
        "links": item.links,
        "downloaded_files": item.downloaded_files,
        "notes": item.notes,
        "is_relevant": is_relevant,
        "relevance_reason": relevance_reason,
    }
    analysis_path.write_text(json.dumps(summary_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return item


def write_full_report(report_path: Path, items: Sequence[ma.ProcessedMessage]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Исходные данные"
    sheet.append(ma.REPORT_HEADERS)
    processed_at = dt.datetime.now().isoformat(timespec="seconds")
    for item in items:
        if item.extracted_fields.get("is_relevant") != "yes":
            continue
        cleaned_row = [
            ILLEGAL_XLSX_RE.sub("", str(value)) if value is not None else ""
            for value in ma.build_report_row(processed_at, item)
        ]
        sheet.append(cleaned_row)
    ma.ensure_dir(report_path.parent)
    workbook.save(report_path)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Refresh today's archived messages and rebuild report")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--date", default=dt.date.today().isoformat())
    args = parser.parse_args(argv)

    config = ma.load_config(Path(args.config))
    ma.RUNTIME_CONFIG = config

    day_root = Path(str(config.get("output_root", "archive"))) / args.date
    if not day_root.exists():
        print(json.dumps({"processed_messages": 0, "report_path": "", "date": args.date}, ensure_ascii=False))
        return 0

    items: List[ma.ProcessedMessage] = []
    for analysis_path in sorted(day_root.rglob("analysis.json")):
        item = rebuild_message_folder(analysis_path.parent, config)
        if item is not None:
            items.append(item)

    report_path = Path(str(config.get("report_root", "reports"))) / f"report-{args.date}.xlsx"
    write_full_report(report_path, items)
    relevant_count = sum(1 for item in items if item.extracted_fields.get("is_relevant") == "yes")
    print(
        json.dumps(
            {
                "processed_messages": relevant_count,
                "scanned_messages": len(items),
                "report_path": str(report_path.resolve()),
                "date": args.date,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
