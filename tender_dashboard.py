from __future__ import annotations

import argparse
import base64
import datetime as dt
import getpass
import hmac
import hashlib
import html
import io
import json
import mimetypes
import os
import posixpath
import re
import shutil
import shlex
import subprocess
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Iterable, Sequence
from zipfile import ZipFile

try:
    import mail_automation as ma
except BaseException:  # pragma: no cover - dashboard has a local fallback for read-only display.
    ma = None


ROOT = Path(__file__).resolve().parent
STATIC_ROOT = ROOT / "dashboard"
STATE_SCHEMA_VERSION = 1

REPORT_HEADERS_FALLBACK = [
    "Дата обработки",
    "Почтовый ящик",
    "UID письма",
    "Тема",
    "Объект",
    "Заказчик/отправитель",
    "Контакты",
    "Адрес/местоположение",
    "Вид работ",
    "Стадия",
    "Сроки",
    "Состав документации",
    "Классификация разделов",
    "Исходные файлы",
    "Ссылки",
    "Ключевые технические параметры",
    "Объемы",
    "Требования",
    "Ограничения",
    "Недостающие сведения",
    "Риски",
    "Рекомендация следующего действия",
    "Статус достаточности",
    "Папка материалов",
]

STATUS_DEFS = {
    "unchecked": {"label": "Непроверена", "sort": 10},
    "relevant": {"label": "Релевантна", "sort": 20},
    "submitting": {"label": "В работе", "actionLabel": "Подаемся", "sort": 40},
    "not_relevant": {"label": "Не подходит", "sort": 90},
}
LEGACY_STATUS_MAP = {"analyzing": "unchecked"}
DASHBOARD_DISPLAY_HEADERS = [
    "Источник закупки",
    "Объект",
    "Заказчик",
    "НМЦК",
    "Окончание подачи предложений",
    "ТЗ",
    "Аналитическая записка",
    "Ссылки на закупку",
    "ТЭП",
    "Рекомендация следующего действия",
    "Папка материалов",
]
DEFAULT_MIN_NMC_RUB = 1_500_000
INTEGRATION_ENV_FIELDS = {
    "TENDER_DASHBOARD_USER": {"label": "Логин dashboard", "secret": False},
    "TENDER_DASHBOARD_PASSWORD": {"label": "Пароль dashboard", "secret": True},
    "TENDER360_BASE_URL": {"label": "Tender360 URL", "secret": False},
    "TENDER360_USERNAME": {"label": "Логин Tender360", "secret": False},
    "TENDER360_PASSWORD": {"label": "Пароль Tender360", "secret": True},
    "TENDER360_API_TOKEN": {"label": "API token Tender360", "secret": True},
    "TENDER_DASHBOARD_BITRIX_PORTAL_URL": {"label": "Портал Bitrix24", "secret": False},
    "TENDER_DASHBOARD_BITRIX_WEBHOOK_URL": {"label": "Webhook Bitrix24", "secret": True},
    "TENDER_DASHBOARD_BITRIX_RESPONSIBLE_ID": {"label": "ID ответственного CRM", "secret": False},
    "TENDER_DASHBOARD_BITRIX_GIP_USER_ID": {"label": "ID ГИПа в Bitrix24", "secret": False},
    "TENDER_DASHBOARD_BITRIX_TASK_CREATED_BY_ID": {"label": "ID автора задачи", "secret": False},
    "TENDER_DASHBOARD_BITRIX_TASK_GROUP_ID": {"label": "ID группы задач", "secret": False},
}

DOWNLOAD_DIR_NAMES = {"downloads", "downloads_refreshed", "01_downloaded_docs"}
METADATA_FILENAMES = {
    "03_document_index.json",
    "04_project_card.json",
    "05_tep.json",
    "06_scope.json",
    "07_requirements.json",
    "08_red_flags.json",
    "09_missing_data.json",
    "12_analysis_report.json",
    "14_pir_normative_estimate.json",
    "batch_index.json",
    "links.txt",
    "request_meta.json",
    "run_summary.json",
    "source_message_path.txt",
}
DEFAULT_RELEVANT_EXCEL_GLOBS = ["reports/Заявки_*.xlsx", "reports/report-*.xlsx"]
ANALYTIC_NOTE_FILENAME = "Аналитическая записка.md"
DEFAULT_DAILY_REPORT_EXCLUDE_PATTERNS = [
    r"поможем\s+оформить\s+банковск\w*\s+гарант",
    r"банковск\w*\s+гарант\w*\s+без\s+лишн\w*\s+сложност",
    r"включени\w*\s+в\s+реестр\s+минпромторг",
    r"как\s+забирать\s+прибыльн\w*\s+контракт\w*\s+на\s+этп",
    r"фас\s+начнет\s+проверять\s+закупк\w*\s+с\s+помощью\s+ии",
    r"другие\s+новости",
    r"пс\s+и\s+оборудовани\w*,?\s+используем\w*\s+на\s+объект\w*\s+опо",
    r"протокол\s+подведения\s+итогов\s+закупк",
    r"\bвебинар\w*\b",
    r"\bобучени\w*\b",
]
ANALYTIC_NOTE_HEADER = "Краткая аналитическая записка"
DEADLINE_HEADER = "Окончание подачи предложений"
DEADLINE_SOURCE_HEADERS = (
    DEADLINE_HEADER,
    "Дата окончания приема предложений",
    "Дата окончания приема заявок",
    "Крайний срок подачи предложений",
    "Крайний срок подачи заявок",
    "Срок подачи заявок",
    "Сроки",
)
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XLSX_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


@dataclass
class DashboardItem:
    id: str
    source_kind: str
    entry_date: str
    title: str
    columns: dict[str, str]
    material_dir: str
    source_path: str
    headers: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)
    downloaded_files: list[str] = field(default_factory=list)
    links: list[str] = field(default_factory=list)
    deep_analysis_dir: str = ""
    analysis: dict[str, Any] = field(default_factory=dict)
    status: str = "unchecked"
    status_label: str = STATUS_DEFS["unchecked"]["label"]
    first_seen_at: str = ""
    decision_updated_at: str = ""
    files_deleted_at: str = ""
    bitrix_lead_id: str = ""
    bitrix_task_id: str = ""
    bitrix_error: str = ""
    bitrix_task_error: str = ""
    is_new: bool = False
    compact: bool = False


def report_headers() -> list[str]:
    if ma is not None:
        return list(ma.REPORT_HEADERS)
    return list(REPORT_HEADERS_FALLBACK)


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def today_iso() -> str:
    return dt.date.today().isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple, set)):
        return "; ".join(clean_text(item) for item in value if clean_text(item))
    if isinstance(value, dict):
        for key in ("value", "description", "question_to_customer", "item", "title", "name"):
            if key in value:
                nested = clean_text(value.get(key))
                if nested:
                    return nested
        return json.dumps(value, ensure_ascii=False)
    return str(value).replace("\x00", "").strip()


def extract_value(value: Any) -> str:
    if isinstance(value, dict):
        if "value" in value:
            return clean_text(value.get("value"))
        if "description" in value:
            return clean_text(value.get("description"))
        if "question_to_customer" in value:
            return clean_text(value.get("question_to_customer"))
        if "item" in value:
            return clean_text(value.get("item"))
        return ""
    return clean_text(value)


def read_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return default


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore").strip()
    except OSError:
        return ""


def resolve_path(value: Any, base_dir: Path = ROOT) -> Path:
    path = Path(str(value))
    if path.is_absolute():
        return path
    return base_dir / path


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return read_json(path, {})


def dashboard_config(config: dict[str, Any]) -> dict[str, Any]:
    payload = config.get("dashboard", {})
    return payload if isinstance(payload, dict) else {}


def configured_state_path(config: dict[str, Any], config_path: Path) -> Path:
    dash = dashboard_config(config)
    value = dash.get("state_path") or ".state/dashboard_statuses.json"
    return resolve_path(value, config_path.parent)


def configured_state_dir(config: dict[str, Any], config_path: Path) -> Path:
    return resolve_path(config.get("state_dir", ".state"), config_path.parent)


def configured_source_counts_path(config: dict[str, Any], config_path: Path) -> Path:
    dash = dashboard_config(config)
    value = dash.get("source_counts_path") or ".state/dashboard_source_counts.json"
    return resolve_path(value, config_path.parent)


def configured_processed_messages_path(config: dict[str, Any], config_path: Path) -> Path:
    return configured_state_dir(config, config_path) / "processed_messages.json"


def load_state(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    path = configured_state_path(config, config_path)
    payload = read_json(path, {})
    if not isinstance(payload, dict):
        payload = {}
    payload.setdefault("version", STATE_SCHEMA_VERSION)
    payload.setdefault("items", {})
    return payload


def save_state(config: dict[str, Any], config_path: Path, state: dict[str, Any]) -> None:
    path = configured_state_path(config, config_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def dashboard_unprocessed_retention_days(config: dict[str, Any]) -> int:
    dash = dashboard_config(config)
    raw = dash.get("unprocessed_retention_days", config.get("unprocessed_download_retention_days", 10))
    try:
        return max(0, int(raw))
    except (TypeError, ValueError):
        return 10


def parse_state_datetime(value: Any) -> dt.datetime | None:
    text = clean_text(value)
    if not text:
        return None
    normalized = text.replace("Z", "+00:00")
    try:
        parsed = dt.datetime.fromisoformat(normalized)
    except ValueError:
        try:
            parsed = dt.datetime.combine(dt.date.fromisoformat(text[:10]), dt.time.min)
        except ValueError:
            return None
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed


def state_record_status(record: dict[str, Any]) -> str:
    status = clean_text(record.get("status")) or "unchecked"
    normalized = LEGACY_STATUS_MAP.get(status, status)
    return normalized if normalized in STATUS_DEFS else "unchecked"


def retain_unprocessed_state_record(record: dict[str, Any], retention_days: int) -> bool:
    if retention_days <= 0 or state_record_status(record) != "unchecked":
        return False
    seen_at = parse_state_datetime(record.get("first_seen_at")) or parse_state_datetime(record.get("entry_date"))
    if seen_at is None:
        return False
    return seen_at >= dt.datetime.now() - dt.timedelta(days=retention_days)


def state_list(value: Any) -> list[str]:
    if isinstance(value, (list, tuple, set)):
        return [clean_text(item) for item in value if clean_text(item)]
    text = clean_text(value)
    return [text] if text else []


def state_columns(value: Any) -> dict[str, str]:
    if not isinstance(value, dict):
        return {}
    return {clean_text(key): clean_text(val) for key, val in value.items() if clean_text(key)}


def snapshot_state_record(record: dict[str, Any], item: DashboardItem) -> bool:
    snapshot = {
        "source_kind": item.source_kind,
        "entry_date": item.entry_date,
        "title": item.title,
        "columns": state_columns(item.columns),
        "material_dir": item.material_dir,
        "source_path": item.source_path,
        "headers": state_list(item.headers),
        "attachments": state_list(item.attachments),
        "downloaded_files": state_list(item.downloaded_files),
        "links": state_list(item.links),
        "deep_analysis_dir": item.deep_analysis_dir,
    }
    changed = False
    for key, value in snapshot.items():
        if value in ("", [], {}):
            continue
        if record.get(key) != value:
            record[key] = value
            changed = True
    return changed


def item_from_state_record(item_id: str, record: dict[str, Any]) -> DashboardItem | None:
    title = clean_text(record.get("title"))
    columns = state_columns(record.get("columns"))
    if not title:
        title = first_column(columns, ("Объект", "Название объекта", "Тема"))
    if not title:
        return None
    if not columns:
        columns = {"Объект": title}
    elif not first_column(columns, ("Объект", "Название объекта", "Тема")):
        columns.setdefault("Объект", title)
    return DashboardItem(
        id=item_id,
        source_kind=clean_text(record.get("source_kind")) or "retained_unprocessed",
        entry_date=clean_text(record.get("entry_date")),
        title=title,
        columns=columns,
        material_dir=clean_text(record.get("material_dir")),
        source_path=clean_text(record.get("source_path")),
        headers=state_list(record.get("headers")),
        attachments=state_list(record.get("attachments")),
        downloaded_files=state_list(record.get("downloaded_files")),
        links=state_list(record.get("links")),
        deep_analysis_dir=clean_text(record.get("deep_analysis_dir")),
    )


def append_retained_unprocessed_items(
    items: list[DashboardItem],
    state: dict[str, Any],
    config: dict[str, Any],
) -> list[DashboardItem]:
    records = state.get("items")
    if not isinstance(records, dict):
        return items
    retention_days = dashboard_unprocessed_retention_days(config)
    present_ids = {item.id for item in items}
    retained: list[DashboardItem] = []
    for item_id, record in records.items():
        if item_id in present_ids or not isinstance(record, dict):
            continue
        if not retain_unprocessed_state_record(record, retention_days):
            continue
        item = item_from_state_record(clean_text(item_id), record)
        if item is not None and daily_report_row_excluded(item.columns, config):
            continue
        if item is not None:
            item = apply_daily_report_title_normalization(item)
            retained.append(item)
    if not retained:
        return items
    return [*items, *retained]


def integration_env_path(config: dict[str, Any], config_path: Path, explicit_path: str = "") -> Path:
    if explicit_path:
        return resolve_path(explicit_path, config_path.parent)
    dash = dashboard_config(config)
    configured = clean_text(dash.get("secrets_env_path") or os.environ.get("TENDER_DASHBOARD_ENV_FILE"))
    if configured:
        return resolve_path(configured, config_path.parent)
    if os.name == "nt":
        return config_path.parent / ".state" / "tender-dashboard.env"
    return Path("/etc/tender-dashboard/tender-dashboard.env")


def parse_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        try:
            tokens = shlex.split(stripped, comments=True, posix=True)
        except ValueError:
            tokens = [stripped]
        for token in tokens:
            if "=" not in token:
                continue
            key, value = token.split("=", 1)
            if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", key):
                values[key] = value
            break
    return values


def quote_env_value(value: str) -> str:
    text = clean_text(value)
    if "\n" in text or "\r" in text or "\x00" in text:
        raise ValueError("Значение переменной не может содержать перенос строки")
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def write_env_file(path: Path, values: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ordered_keys = sorted(values)
    body = "\n".join(f"{key}={quote_env_value(values[key])}" for key in ordered_keys) + "\n"
    path.write_text(body, encoding="utf-8")
    try:
        os.chmod(path, 0o660)
    except OSError:
        pass


def env_write_status(path: Path) -> dict[str, Any]:
    parent = path.parent
    exists = path.exists()
    parent_exists = parent.exists()
    writable = (exists and os.access(path, os.W_OK)) or (not exists and parent_exists and os.access(parent, os.W_OK))
    return {
        "path": str(path),
        "exists": exists,
        "writable": writable,
    }


def integration_config_value(key: str, config: dict[str, Any]) -> str:
    dash = dashboard_config(config)
    mapping = {
        "TENDER_DASHBOARD_USER": dash.get("auth_username"),
        "TENDER_DASHBOARD_PASSWORD": dash.get("auth_password"),
        "TENDER_DASHBOARD_BITRIX_PORTAL_URL": dash.get("bitrix_portal_url"),
        "TENDER_DASHBOARD_BITRIX_WEBHOOK_URL": dash.get("bitrix_webhook_url") or config.get("bitrix_webhook_url"),
        "TENDER_DASHBOARD_BITRIX_RESPONSIBLE_ID": dash.get("bitrix_responsible_id"),
        "TENDER_DASHBOARD_BITRIX_GIP_USER_ID": dash.get("bitrix_gip_user_id") or dash.get("bitrix_task_responsible_id"),
        "TENDER_DASHBOARD_BITRIX_TASK_CREATED_BY_ID": dash.get("bitrix_task_created_by_id"),
        "TENDER_DASHBOARD_BITRIX_TASK_GROUP_ID": dash.get("bitrix_task_group_id"),
    }
    return clean_text(mapping.get(key))


def integrations_status(config: dict[str, Any], config_path: Path, explicit_path: str = "") -> dict[str, Any]:
    path = integration_env_path(config, config_path, explicit_path)
    values = parse_env_file(path)
    fields = {}
    for key, meta in INTEGRATION_ENV_FIELDS.items():
        value = values.get(key) or os.environ.get(key, "") or integration_config_value(key, config)
        fields[key] = {
            "label": meta["label"],
            "secret": bool(meta["secret"]),
            "configured": bool(clean_text(value)),
        }
    return {"envFile": env_write_status(path), "fields": fields}


def save_integrations(
    payload: dict[str, Any],
    config: dict[str, Any],
    config_path: Path,
    explicit_path: str = "",
) -> dict[str, Any]:
    path = integration_env_path(config, config_path, explicit_path)
    values = parse_env_file(path)
    fields_payload = payload.get("fields", payload)
    if not isinstance(fields_payload, dict):
        raise ValueError("Некорректный payload")
    updated: list[str] = []
    for key in INTEGRATION_ENV_FIELDS:
        if key not in fields_payload:
            continue
        value = clean_text(fields_payload.get(key))
        if not value:
            continue
        values[key] = value
        os.environ[key] = value
        updated.append(key)
    if not updated:
        raise ValueError("Нет новых значений для сохранения")
    write_env_file(path, values)
    return {"saved": True, "updated": updated, "envFile": env_write_status(path)}


def configure_integrations_interactive(config: dict[str, Any], config_path: Path, explicit_path: str = "") -> dict[str, Any]:
    path = integration_env_path(config, config_path, explicit_path)
    current = integrations_status(config, config_path, explicit_path)
    print(f"Env file: {path}")
    print("Пустой ввод оставляет текущее значение без изменений.")
    fields: dict[str, str] = {}
    for key, meta in INTEGRATION_ENV_FIELDS.items():
        status = "задано" if current["fields"][key]["configured"] else "не задано"
        prompt = f"{meta['label']} ({key}, сейчас: {status}): "
        if meta["secret"]:
            value = getpass.getpass(prompt)
        else:
            value = input(prompt)
        value = clean_text(value)
        if value:
            fields[key] = value
    if not fields:
        return {"saved": False, "message": "Новые значения не введены", "envFile": env_write_status(path)}
    result = save_integrations({"fields": fields}, config, config_path, explicit_path)
    result["status"] = integrations_status(config, config_path, explicit_path)
    return result


def stable_id(kind: str, basis: str) -> str:
    digest = hashlib.sha1(f"{kind}:{basis}".encode("utf-8", errors="ignore")).hexdigest()
    return digest[:20]


def path_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except (OSError, ValueError):
        return False


def configured_roots(config: dict[str, Any], config_path: Path) -> dict[str, Path]:
    base = config_path.parent
    roots = {
        "output": resolve_path(config.get("output_root", "archive"), base),
        "relevant": resolve_path(config.get("relevant_root", config.get("output_root", "релевантные")), base),
        "sorted": resolve_path(config.get("sorted_output_root", "sorted_mail"), base),
        "reports": resolve_path(config.get("report_root", "reports"), base),
    }
    return roots


def analysis_scan_roots(config: dict[str, Any], config_path: Path) -> list[Path]:
    dash = dashboard_config(config)
    explicit_roots = dash.get("analysis_roots")
    if isinstance(explicit_roots, list) and explicit_roots:
        return [resolve_path(root, config_path.parent) for root in explicit_roots]

    roots = configured_roots(config, config_path)
    scan_roots = [roots["relevant"]]
    if bool(dash.get("include_archive_root", False)):
        scan_roots.append(roots["output"])
    return [root for root in scan_roots if root.exists()]


def extract_date_from_path(path: Path, roots: Iterable[Path]) -> str:
    for root in roots:
        try:
            rel = path.resolve().relative_to(root.resolve())
        except (OSError, ValueError):
            continue
        parts = rel.parts
        if parts and looks_like_date(parts[0]):
            return parts[0]
    for part in path.parts:
        if looks_like_date(part):
            return part
    try:
        return dt.date.fromtimestamp(path.stat().st_mtime).isoformat()
    except OSError:
        return today_iso()


def looks_like_date(value: str) -> bool:
    try:
        dt.date.fromisoformat(value)
        return True
    except ValueError:
        return False


def normalize_path_string(value: Any) -> str:
    if not value:
        return ""
    try:
        return str(Path(str(value)).resolve())
    except OSError:
        return str(value)


def split_semicolon_paths(value: str) -> list[str]:
    return [part.strip() for part in value.split(";") if part.strip()]


def split_links(value: str) -> list[str]:
    urls = re.findall(r"https?://[^\s;]+", value)
    if urls:
        return list(dict.fromkeys(url.rstrip(".,)") for url in urls))
    parts = re.split(r"[\n;]+", value)
    return [part.strip() for part in parts if part.strip()]


def normalize_match_key(value: str) -> str:
    return re.sub(r"[^0-9a-zа-яё]+", "", value.lower())


def first_column(columns: dict[str, str], names: Sequence[str]) -> str:
    for name in names:
        value = clean_text(columns.get(name))
        if value:
            return value
    return ""


def offer_deadline(columns: dict[str, str]) -> str:
    return first_column(columns, DEADLINE_SOURCE_HEADERS)


def display_headers_with_deadline(headers: Sequence[str]) -> list[str]:
    ordered = [DEADLINE_HEADER]
    for header in headers:
        if not header or header in DEADLINE_SOURCE_HEADERS:
            continue
        ordered.append(header)
    return ordered


def split_note_parts(value: str) -> list[str]:
    parts = re.split(r"[;\n]+", value)
    cleaned: list[str] = []
    service_tokens = {"high", "medium", "low", "critical", "info", "высокий", "средний", "низкий"}
    for part in parts:
        text = clean_text(part).strip(" .;")
        if text and text.lower() not in service_tokens and text not in cleaned:
            cleaned.append(text)
    return cleaned


def limited_join(parts: Sequence[str], limit: int = 4) -> str:
    unique: list[str] = []
    seen: set[str] = set()
    for part in parts:
        text = clean_text(part)
        key = normalize_match_key(text)[:140]
        if text and key not in seen:
            unique.append(text)
            seen.add(key)
    selected = unique[:limit]
    if not selected:
        return ""
    suffix = f"; еще {len(unique) - limit}" if len(unique) > limit else ""
    return "; ".join(selected) + suffix


def is_placeholder_text(value: str) -> bool:
    lowered = clean_text(value).lower()
    return not lowered or lowered in {"не определено", "не найдено", "нет", "-", "—"}


def build_analytic_note(columns: dict[str, str], deep_match: DashboardItem | None = None) -> str:
    title = first_column(columns, ("Название объекта", "Объект", "Тема"))
    customer = first_column(columns, ("Заказчик", "Заказчик/отправитель"))
    deadline = offer_deadline(columns)
    docs = first_column(columns, ("ТЗ", "Документация", "Исходные файлы"))
    problems = split_note_parts(first_column(columns, ("Проблемы", "Риски")))
    missing = split_note_parts(first_column(columns, ("Чего не хватает", "Недостающие сведения", "Недостающие сведения/ИРД")))
    recommendation = first_column(columns, ("Рекомендация", "Рекомендация следующего действия"))
    tep_found: list[str] = []
    tep_missing: list[str] = []
    scope: list[str] = []
    requirements: list[str] = []
    primary_doc = ""
    project_card: dict[str, str] = {}

    if deep_match is not None:
        analysis = deep_match.analysis or {}
        primary_doc = clean_text(analysis.get("primary_technical_document"))
        project_card = analysis.get("project_card", {}) if isinstance(analysis.get("project_card"), dict) else {}
        tep_found = [clean_text(value) for value in analysis.get("tep_found", []) if clean_text(value)]
        tep_missing = [clean_text(value) for value in analysis.get("tep_missing", []) if clean_text(value)]
        scope = [clean_text(value) for value in analysis.get("scope", []) if clean_text(value)]
        requirements = [clean_text(value) for value in analysis.get("requirements", []) if clean_text(value)]
        deep_problems = [clean_text(value) for value in analysis.get("red_flags", []) if clean_text(value)]
        deep_missing = [clean_text(value) for value in analysis.get("missing_items", []) if clean_text(value)]
        problems.extend(part for part in deep_problems if part not in problems)
        missing.extend(part for part in deep_missing if part not in missing)
        if primary_doc:
            problems = [part for part in problems if "не найдено техническое задание" not in part.lower()]
        if not recommendation:
            recommendation = limited_join(
                [clean_text(value) for value in analysis.get("missing_questions", []) if clean_text(value)],
                3,
            ) or first_column(deep_match.columns, ("Рекомендация следующего действия",))
        if is_placeholder_text(docs):
            docs = primary_doc or first_column(deep_match.columns, ("Состав документации", "Исходные файлы"))

    intro_parts = []
    if title:
        intro_parts.append(f"Объект: {title}")
    if customer:
        intro_parts.append(f"заказчик: {customer}")
    if deadline:
        intro_parts.append(f"срок подачи: {deadline}")
    if project_card.get("address"):
        intro_parts.append(f"адрес/регион: {project_card['address']}")
    if project_card.get("stage"):
        intro_parts.append(f"стадия/состав: {project_card['stage']}")
    elif scope:
        intro_parts.append(f"состав работ: {limited_join(scope, 4)}")
    if project_card.get("work_type"):
        intro_parts.append(f"вид/тип работ: {project_card['work_type']}")
    intro = ". ".join(intro_parts)

    problem_text = limited_join(problems)
    missing_text = limited_join(missing)
    tep_text = limited_join(tep_found, 6)
    tep_missing_text = limited_join(tep_missing, 6)
    requirements_text = limited_join(requirements, 6)
    note_parts: list[str] = []
    if intro:
        note_parts.append(intro + ".")
    if deep_match is not None:
        note_parts.append(
            f"Основание анализа: разобран текст ТЗ/задания на проектирование/ИРД и исходных файлов"
            f"{' по основному техдокументу ' + primary_doc if primary_doc else ''}."
        )
    else:
        note_parts.append(
            "Основание анализа: в dashboard найдена только Excel-сводка; папка глубокого разбора ТЗ/ИРД для этой строки не сопоставлена, поэтому нужен прогон исходных файлов через анализатор."
        )
    if tep_text:
        note_parts.append(f"ТЭП и основные параметры объекта: {tep_text}.")
        if tep_missing_text:
            note_parts.append(f"Не извлечены или не подтверждены количественные ТЭПы: {tep_missing_text}.")
    elif deep_match is not None:
        note_parts.append(
            f"ТЭП и основные параметры объекта: количественные параметры в ТЗ/исходниках не извлечены"
            f"{' (не хватает: ' + tep_missing_text + ')' if tep_missing_text else ''}."
        )
    else:
        note_parts.append("ТЭП и основные параметры объекта: отсутствуют в доступной Excel-сводке; требуется извлечь из ТЗ/ИРД.")
    if requirements_text:
        note_parts.append(f"Ключевые требования из задания: {requirements_text}.")
    if problem_text:
        note_parts.append(f"Анализ ТЗ/ИРД выявил проблемы, неточности или риски: {problem_text}.")
    else:
        note_parts.append("Анализ ТЗ/ИРД: явные критические проблемы в доступных данных не выделены, но полноту исходников нужно подтвердить.")
    if missing_text:
        note_parts.append(f"Недостатки ТЗ/ИРД и исходных данных: требуется уточнить {missing_text}.")
    else:
        note_parts.append("Недостатки ТЗ/ИРД: перечень обязательных уточнений не сформирован; проверить полноту задания на проектирование и ИРД вручную.")
    if docs:
        note_parts.append(f"Проверяемые материалы: {docs}.")
    if recommendation:
        note_parts.append(f"Рекомендуемое действие: {recommendation}.")
    elif missing_text or problem_text:
        note_parts.append("Рекомендуемое действие: до подачи запросить недостающие исходные данные и зафиксировать границы ответственности.")
    else:
        note_parts.append("Рекомендуемое действие: выполнить финальную проверку ТЗ/ИРД перед подачей.")
    return " ".join(note_parts)


def unique_headers(headers: Sequence[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for header in headers:
        base = clean_text(header)
        if not base:
            continue
        count = seen.get(base, 0)
        seen[base] = count + 1
        result.append(base if count == 0 else f"{base} {count + 1}")
    return result


def column_index_from_ref(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return max(index - 1, 0)


def xlsx_shared_strings(zip_file: ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for node in root.findall(f"{XLSX_NS}si"):
        values.append("".join(text.text or "" for text in node.iter(f"{XLSX_NS}t")))
    return values


def xlsx_first_sheet_path(zip_file: ZipFile) -> str:
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    sheet = workbook.find(f"{XLSX_NS}sheets/{XLSX_NS}sheet")
    if sheet is None:
        return "xl/worksheets/sheet1.xml"
    relation_id = sheet.attrib.get(f"{XLSX_REL_NS}id", "")
    rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.attrib.get("Id") == relation_id:
            target = rel.attrib.get("Target", "").lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    return "xl/worksheets/sheet1.xml"


def xlsx_cell_text(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{XLSX_NS}t")).strip()
    value = cell.find(f"{XLSX_NS}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        index = int(value.text)
        return shared_strings[index] if index < len(shared_strings) else ""
    return value.text.strip()


def read_xlsx_rows(path: Path) -> list[list[str]]:
    with ZipFile(path) as zip_file:
        shared_strings = xlsx_shared_strings(zip_file)
        sheet_path = xlsx_first_sheet_path(zip_file)
        root = ET.fromstring(zip_file.read(sheet_path))
        rows: list[list[str]] = []
        for row in root.findall(f"{XLSX_NS}sheetData/{XLSX_NS}row"):
            values: list[str] = []
            for cell in row.findall(f"{XLSX_NS}c"):
                index = column_index_from_ref(cell.attrib.get("r", "A"))
                while len(values) <= index:
                    values.append("")
                values[index] = xlsx_cell_text(cell, shared_strings)
            rows.append(values)
        return rows


def relevant_excel_globs(config: dict[str, Any]) -> list[str]:
    dash = dashboard_config(config)
    configured = dash.get("relevant_excel_globs")
    patterns: list[str] = []
    if isinstance(configured, list) and configured:
        patterns = [clean_text(value) for value in configured if clean_text(value)]
    if not patterns:
        patterns = list(DEFAULT_RELEVANT_EXCEL_GLOBS)
    for pattern in DEFAULT_RELEVANT_EXCEL_GLOBS:
        if pattern not in patterns:
            patterns.append(pattern)
    return patterns


def daily_report_exclude_patterns(config: dict[str, Any]) -> list[str]:
    patterns = list(DEFAULT_DAILY_REPORT_EXCLUDE_PATTERNS)
    configured = dashboard_config(config).get("daily_report_exclude_patterns")
    if isinstance(configured, list):
        for value in configured:
            pattern = clean_text(value)
            if pattern and pattern not in patterns:
                patterns.append(pattern)
    return patterns


def is_daily_report_path(path: Path | str) -> bool:
    name = Path(clean_text(path)).name.lower()
    return bool(re.fullmatch(r"report-\d{4}-\d{2}-\d{2}\.xlsx", name))


def daily_report_search_text(columns: dict[str, str]) -> str:
    priority_headers = (
        "Название объекта",
        "Объект",
        "Тема",
        "Вид работ",
        "Заказчик",
        "Заказчик/отправитель",
        "Рекомендация следующего действия",
        "Ссылки",
    )
    parts = [first_column(columns, (header,)) for header in priority_headers]
    parts.extend(clean_text(value) for value in columns.values())
    return " ".join(part for part in parts if part).lower()


def daily_report_row_excluded(columns: dict[str, str], config: dict[str, Any]) -> bool:
    text = daily_report_search_text(columns)
    if not text:
        return False
    for pattern in daily_report_exclude_patterns(config):
        try:
            if re.search(pattern, text, flags=re.IGNORECASE):
                return True
        except re.error:
            continue
    return False


def normalize_daily_report_object_title(value: str) -> str:
    title = clean_text(value)
    if not title:
        return ""
    title = re.sub(r"^[^\wа-яё]+", "", title, flags=re.IGNORECASE)
    title = re.sub(r"^новые\s+закупки\s*//\s*", "", title, flags=re.IGNORECASE)
    title = re.sub(r"^\(?\s*лот\s*\d+\s*\)?\s*", "", title, flags=re.IGNORECASE)
    title = re.sub(r"^уточнени\w*\s+по\s+тендеру\s*[:\-]\s*", "", title, flags=re.IGNORECASE)
    title = re.sub(r"^приглашени\w*\s+на\s+участие\s+в\s+тендере\s*", "", title, flags=re.IGNORECASE)
    title = re.sub(
        r"^приглашени\w*\s+на\s+закупку\s+[^:]{1,80}:\s*[a-zа-я0-9_.\-]+\s*",
        "",
        title,
        flags=re.IGNORECASE,
    )
    title = re.sub(r"^[\s:;.,/\\\-–—]+", "", title)
    return title or clean_text(value)


def has_daily_report_title_wrapper(value: str) -> bool:
    title = clean_text(value)
    if not title:
        return False
    return bool(
        re.search(r"новые\s+закупки\s*//", title, flags=re.IGNORECASE)
        or re.search(r"уточнени\w*\s+по\s+тендеру\s*[:\-]", title, flags=re.IGNORECASE)
        or re.search(r"приглашени\w*\s+на\s+участие\s+в\s+тендере", title, flags=re.IGNORECASE)
        or re.search(r"приглашени\w*\s+на\s+закупку\s+[^:]{1,80}:", title, flags=re.IGNORECASE)
        or re.match(r"^[^\wа-яё]*\(?\s*лот\s*\d+\s*\)?", title, flags=re.IGNORECASE)
    )


def apply_daily_report_title_normalization(item: DashboardItem) -> DashboardItem:
    raw_title = first_column(item.columns, ("Название объекта", "Объект", "Тема")) or item.title
    if not is_daily_report_path(item.source_path) and not has_daily_report_title_wrapper(raw_title):
        return item
    normalized = normalize_daily_report_object_title(raw_title)
    if not normalized:
        return item
    item.title = normalized
    if "Название объекта" in item.columns:
        item.columns["Название объекта"] = normalized
    if "Объект" in item.columns or "Название объекта" not in item.columns:
        item.columns["Объект"] = normalized
    return item


def iter_excel_paths(config: dict[str, Any], config_path: Path) -> list[Path]:
    paths: dict[str, Path] = {}
    for pattern in relevant_excel_globs(config):
        pattern_path = Path(pattern)
        if pattern_path.is_absolute():
            candidates = pattern_path.parent.glob(pattern_path.name)
        else:
            candidates = config_path.parent.glob(pattern)
        for path in candidates:
            if path.is_file():
                paths[str(path.resolve())] = path
    return sorted(paths.values(), key=lambda path: (path.stat().st_mtime, str(path)))


def date_from_excel_path(path: Path) -> str:
    dates = re.findall(r"\d{4}-\d{2}-\d{2}", path.name)
    if dates:
        return dates[-1]
    try:
        return dt.date.fromtimestamp(path.stat().st_mtime).isoformat()
    except OSError:
        return today_iso()


def date_from_excel_row(columns: dict[str, str], excel_path: Path) -> str:
    for header in ("Дата входа", "Дата обработки"):
        value = clean_text(columns.get(header))
        if not value:
            continue
        match = re.search(r"\d{4}-\d{2}-\d{2}", value)
        if match:
            return match.group(0)
    return date_from_excel_path(excel_path)


def extract_iso_date(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    return match.group(0) if match else ""


def normalize_source_counts(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        return {"total": 0, "byDate": {}}
    raw_by_date = payload.get("byDate") or payload.get("by_date") or {}
    by_date: dict[str, int] = {}
    if isinstance(raw_by_date, dict):
        for key, value in raw_by_date.items():
            date = extract_iso_date(key)
            if not date:
                continue
            try:
                count = int(value)
            except (TypeError, ValueError):
                continue
            if count > 0:
                by_date[date] = by_date.get(date, 0) + count
    try:
        total = int(payload.get("total", 0))
    except (TypeError, ValueError):
        total = 0
    if total <= 0:
        total = sum(by_date.values())
    return {
        "total": total,
        "byDate": dict(sorted(by_date.items(), reverse=True)),
        "generatedAt": clean_text(payload.get("generatedAt") or payload.get("generated_at")),
    }


def source_counts_from_processed_messages(path: Path) -> dict[str, Any]:
    payload = read_json(path, {})
    processed = payload.get("processed") if isinstance(payload, dict) else payload
    if isinstance(processed, dict):
        records = list(processed.values())
    elif isinstance(processed, list):
        records = processed
    else:
        return {"total": 0, "byDate": {}}

    by_date: dict[str, int] = {}
    total = 0
    for record in records:
        total += 1
        date = ""
        if isinstance(record, dict):
            for key in ("processed_at", "date", "entry_date", "created_at"):
                date = extract_iso_date(record.get(key))
                if date:
                    break
            if not date:
                date = extract_iso_date(record.get("path"))
        else:
            date = extract_iso_date(record)
        if date:
            by_date[date] = by_date.get(date, 0) + 1
    return {
        "total": total,
        "byDate": dict(sorted(by_date.items(), reverse=True)),
        "generatedAt": now_iso(),
    }


def source_analysis_counts(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    counts_path = configured_source_counts_path(config, config_path)
    payload = normalize_source_counts(read_json(counts_path, {}))
    if payload["total"] or payload["byDate"]:
        return payload
    return source_counts_from_processed_messages(configured_processed_messages_path(config, config_path))


def summarize_file_names(paths: Sequence[str], limit: int = 8) -> str:
    names = [Path(path).name for path in paths if path]
    if len(names) > limit:
        return "; ".join(names[:limit]) + f"; еще {len(names) - limit}"
    return "; ".join(names)


def processed_message_from_payload(path: Path, payload: dict[str, Any], date_folder: str) -> Any:
    fields = {str(key): clean_text(value) for key, value in (payload.get("fields") or {}).items()}
    attachments = [normalize_path_string(value) for value in payload.get("attachments", [])]
    links = [clean_text(value) for value in payload.get("links", [])]
    downloaded_files = [normalize_path_string(value) for value in payload.get("downloaded_files", [])]
    output_dir = normalize_path_string(payload.get("output_dir") or path.parent)
    return ma.ProcessedMessage(
        mailbox=clean_text(payload.get("mailbox")),
        uid=clean_text(payload.get("uid")),
        message_id=clean_text(payload.get("message_id")),
        subject=clean_text(payload.get("subject")),
        sender=clean_text(payload.get("sender")),
        date_folder=date_folder,
        output_dir=output_dir,
        attachments=attachments,
        links=links,
        downloaded_files=downloaded_files,
        sections=[clean_text(value) for value in payload.get("sections", [])],
        extracted_fields=fields,
        notes=[clean_text(value) for value in payload.get("notes", [])],
    )


def fallback_mail_columns(path: Path, payload: dict[str, Any], date_folder: str) -> dict[str, str]:
    fields = {str(key): clean_text(value) for key, value in (payload.get("fields") or {}).items()}
    attachments = [normalize_path_string(value) for value in payload.get("attachments", [])]
    downloaded_files = [normalize_path_string(value) for value in payload.get("downloaded_files", [])]
    links = [clean_text(value) for value in payload.get("links", [])]
    values = [
        now_iso(),
        clean_text(payload.get("mailbox")),
        clean_text(payload.get("uid")),
        clean_text(payload.get("subject")),
        fields.get("object", ""),
        fields.get("customer", "") or clean_text(payload.get("sender")),
        fields.get("contacts", ""),
        fields.get("address", ""),
        fields.get("work_type", ""),
        fields.get("stage", ""),
        fields.get("deadline", ""),
        summarize_file_names([*attachments, *downloaded_files]),
        "; ".join(clean_text(value) for value in payload.get("sections", [])),
        "; ".join([*attachments, *downloaded_files]),
        "; ".join(links),
        fields.get("tech_params", ""),
        fields.get("volumes", ""),
        fields.get("requirements", ""),
        fields.get("constraints", ""),
        fields.get("missing_info", ""),
        "",
        "",
        "",
        normalize_path_string(payload.get("output_dir") or path.parent),
    ]
    return dict(zip(report_headers(), values, strict=False))


def item_from_mail_analysis(path: Path, config: dict[str, Any], config_path: Path) -> DashboardItem | None:
    payload = read_json(path, {})
    if not isinstance(payload, dict):
        return None
    roots = analysis_scan_roots(config, config_path)
    date_folder = extract_date_from_path(path, roots)
    if ma is not None:
        try:
            processed = processed_message_from_payload(path, payload, date_folder)
            row = ma.build_report_row(now_iso(), processed)
            columns = dict(zip(report_headers(), [clean_text(value) for value in row], strict=False))
            attachments = list(processed.attachments)
            downloaded_files = list(processed.downloaded_files)
            links = list(processed.links)
            material_dir = processed.output_dir
            title = processed.subject or columns.get("Объект", "")
        except Exception:
            columns = fallback_mail_columns(path, payload, date_folder)
            attachments = [normalize_path_string(value) for value in payload.get("attachments", [])]
            downloaded_files = [normalize_path_string(value) for value in payload.get("downloaded_files", [])]
            links = [clean_text(value) for value in payload.get("links", [])]
            material_dir = columns.get("Папка материалов") or normalize_path_string(path.parent)
            title = clean_text(payload.get("subject")) or columns.get("Объект", "")
    else:
        columns = fallback_mail_columns(path, payload, date_folder)
        attachments = [normalize_path_string(value) for value in payload.get("attachments", [])]
        downloaded_files = [normalize_path_string(value) for value in payload.get("downloaded_files", [])]
        links = [clean_text(value) for value in payload.get("links", [])]
        material_dir = columns.get("Папка материалов") or normalize_path_string(path.parent)
        title = clean_text(payload.get("subject")) or columns.get("Объект", "")

    basis = material_dir or str(path.parent.resolve())
    return DashboardItem(
        id=stable_id("mail", basis),
        source_kind="mail_analysis",
        entry_date=date_folder,
        title=title or Path(material_dir).name,
        columns=columns,
        material_dir=material_dir,
        source_path=str(path.resolve()),
        attachments=attachments,
        downloaded_files=downloaded_files,
        links=links,
    )


def field_value(payload: dict[str, Any], key: str) -> str:
    return extract_value(payload.get(key, {}))


def list_values(payload: Any, key: str | None = None) -> list[str]:
    source = payload.get(key, []) if key and isinstance(payload, dict) else payload
    if not isinstance(source, list):
        return []
    return [extract_value(item) for item in source if extract_value(item)]


def flatten_values(payload: Any, prefix: str = "") -> list[str]:
    values: list[str] = []
    if isinstance(payload, dict):
        if "value" in payload:
            text = extract_value(payload)
            if text:
                values.append(f"{prefix}: {text}" if prefix else text)
            return values
        for key, value in payload.items():
            label = f"{prefix}.{key}" if prefix else str(key)
            values.extend(flatten_values(value, label))
    elif isinstance(payload, list):
        for item in payload:
            values.extend(flatten_values(item, prefix))
    else:
        text = clean_text(payload)
        if text:
            values.append(f"{prefix}: {text}" if prefix else text)
    return values


def compact_list(values: Sequence[str], limit: int = 12) -> str:
    cleaned = []
    for value in values:
        text = clean_text(value)
        if text and text not in cleaned:
            cleaned.append(text)
    if len(cleaned) > limit:
        return "; ".join(cleaned[:limit]) + f"; еще {len(cleaned) - limit}"
    return "; ".join(cleaned)


def read_links_from_file(path: Path) -> list[str]:
    text = read_text(path)
    return [line.strip() for line in text.splitlines() if line.strip()]


def extract_mailbox_from_source_path(path_text: str) -> str:
    if not path_text:
        return ""
    parts = Path(path_text).parts
    for index, part in enumerate(parts):
        if looks_like_date(part) and index + 1 < len(parts):
            return parts[index + 1]
    return ""


def source_message_dir(report_dir: Path) -> Path | None:
    source_path = read_text(report_dir / "_input" / "source_message_path.txt")
    if not source_path:
        return None
    candidate = Path(source_path)
    if candidate.name.lower() in {"message.eml", "message.txt"}:
        return candidate.parent
    return candidate


def useful_source_file(source_file: str) -> bool:
    lowered = source_file.lower()
    if not lowered:
        return True
    noisy_tokens = (
        "privacy",
        "contacts",
        "soglasie",
        "unsubscribe",
        "address_req",
        "request_meta",
        "source_message_path",
        "links.txt",
    )
    return not any(token in lowered for token in noisy_tokens)


def tep_label(section: str, key: str) -> str:
    labels = {
        "roads.length_m": "протяженность дорог, м",
        "roads.width_m": "ширина дорог, м",
        "roads.area_m2": "площадь дорог/покрытий, м2",
        "roads.road_category": "категория дороги",
        "roads.lanes": "количество полос",
        "networks.network_type": "тип инженерной сети",
        "networks.length_m": "протяженность сети, м",
        "networks.diameter_mm": "диаметр сети, мм",
        "networks.connection_points": "точки подключения",
        "buildings.gfa_m2": "общая площадь здания, м2",
        "buildings.floors": "этажность",
        "buildings.site_area_m2": "площадь участка, м2",
        "buildings.parking_spaces": "машино-места",
        "landscaping.area_m2": "площадь благоустройства, м2",
        "landscaping.hardscape_area_m2": "площадь твердых покрытий, м2",
        "landscaping.green_area_m2": "площадь озеленения, м2",
    }
    return labels.get(f"{section}.{key}", f"{section}.{key}")


def summarize_tep_payload(tep: Any) -> tuple[list[str], list[str]]:
    found: list[str] = []
    missing: list[str] = []
    if not isinstance(tep, dict):
        return found, missing
    for section, fields in tep.items():
        if not isinstance(fields, dict):
            continue
        for key, payload in fields.items():
            label = tep_label(str(section), str(key))
            if not isinstance(payload, dict):
                value = clean_text(payload)
                if value:
                    found.append(f"{label}: {value}")
                continue
            value = payload.get("value")
            text = clean_text(value)
            source_file = clean_text(payload.get("source_file"))
            confidence = float(payload.get("confidence") or 0)
            source_confidence = clean_text(payload.get("source_confidence"))
            if text and useful_source_file(source_file) and source_confidence != "missing" and confidence >= 0.5:
                source = f" ({source_file})" if source_file else ""
                found.append(f"{label}: {text}{source}")
            elif source_confidence == "missing" or not text:
                missing.append(label)
    return found, missing


def yes_no_unclear(value: str) -> str:
    lowered = value.lower()
    if lowered in {"yes", "да", "true", "1"}:
        return "да"
    if lowered in {"no", "нет", "false", "0"}:
        return "нет"
    if lowered in {"unclear", "unknown", "неясно"}:
        return "неясно"
    return value


def summarize_requirements_payload(requirements: Any) -> list[str]:
    if not isinstance(requirements, dict):
        return []
    labels = {
        "expertise_required": "экспертиза",
        "bim_required": "BIM",
        "estimate_required": "сметная документация",
        "survey_required": "изыскания",
        "approvals_required": "согласования",
    }
    parts: list[str] = []
    for key, label in labels.items():
        value = extract_value(requirements.get(key, {}))
        if value:
            parts.append(f"{label}: {yes_no_unclear(value)}")
    formats = list_values(requirements, "output_formats")
    if formats:
        parts.append(f"форматы выдачи: {limited_join(formats, 5)}")
    return parts


def summarize_red_flags(red_flags: Any) -> list[str]:
    if not isinstance(red_flags, list):
        return []
    parts: list[str] = []
    for item in red_flags:
        if not isinstance(item, dict):
            text = clean_text(item)
            if text:
                parts.append(text)
            continue
        description = clean_text(item.get("description"))
        impact = clean_text(item.get("impact"))
        action = clean_text(item.get("recommended_action"))
        text = description
        if impact:
            text = f"{text} ({impact})" if text else impact
        if action:
            text = f"{text}; {action}" if text else action
        if text:
            parts.append(text)
    return parts


def summarize_missing_data(missing_data: Any) -> tuple[list[str], list[str]]:
    if not isinstance(missing_data, list):
        return [], []
    items: list[str] = []
    questions: list[str] = []
    for item in missing_data:
        if isinstance(item, dict):
            missing = clean_text(item.get("item"))
            why = clean_text(item.get("why_needed"))
            question = clean_text(item.get("question_to_customer"))
            if missing:
                items.append(f"{missing} ({why})" if why else missing)
            if question:
                questions.append(question)
        else:
            text = clean_text(item)
            if text:
                items.append(text)
    return items, questions


def item_from_deep_analysis(path: Path, config: dict[str, Any], config_path: Path) -> DashboardItem | None:
    report_dir = path.parent
    analysis_report = read_json(path, {})
    project_card = read_json(report_dir / "04_project_card.json", {})
    tep = read_json(report_dir / "05_tep.json", {})
    scope = read_json(report_dir / "06_scope.json", {})
    requirements = read_json(report_dir / "07_requirements.json", {})
    red_flags = read_json(report_dir / "08_red_flags.json", [])
    missing_data = read_json(report_dir / "09_missing_data.json", [])
    document_index = read_json(report_dir / "03_document_index.json", [])
    meta = read_json(report_dir / "_input" / "request_meta.json", {})
    pir = read_json(report_dir / "14_pir_normative_estimate.json", {})

    if not isinstance(project_card, dict):
        project_card = {}
    if not isinstance(analysis_report, dict):
        analysis_report = {}
    if not isinstance(scope, dict):
        scope = {}
    if not isinstance(requirements, dict):
        requirements = {}
    if not isinstance(document_index, list):
        document_index = []
    if not isinstance(meta, dict):
        meta = {}
    if not isinstance(pir, dict):
        pir = {}

    all_files = [normalize_path_string(doc.get("path")) for doc in document_index if isinstance(doc, dict) and doc.get("path")]
    downloaded_files = [
        normalize_path_string(doc.get("path"))
        for doc in document_index
        if isinstance(doc, dict) and doc.get("path") and clean_text(doc.get("source")) == "downloaded"
    ]
    input_files = [
        normalize_path_string(doc.get("path"))
        for doc in document_index
        if isinstance(doc, dict) and doc.get("path") and clean_text(doc.get("source")) == "input"
    ]
    links = []
    if isinstance(meta.get("links"), list):
        links.extend(clean_text(link) for link in meta.get("links", []) if clean_text(link))
    links.extend(read_links_from_file(report_dir / "_input" / "links.txt"))
    links = list(dict.fromkeys(links))

    included_scope = list_values(scope, "included_scope")
    output_formats = list_values(requirements, "output_formats")
    missing_items, missing_questions = summarize_missing_data(missing_data)
    risk_texts = summarize_red_flags(red_flags)
    tep_found, tep_missing = summarize_tep_payload(tep)
    requirement_parts = summarize_requirements_payload(requirements)
    primary_doc_payload = analysis_report.get("primary_technical_document", {})
    primary_doc = ""
    if isinstance(primary_doc_payload, dict):
        primary_doc_name = clean_text(primary_doc_payload.get("original_filename"))
        primary_doc_type = clean_text(primary_doc_payload.get("document_type"))
        primary_doc_confidence = clean_text(primary_doc_payload.get("confidence"))
        primary_doc = primary_doc_name
        if primary_doc_type:
            primary_doc += f" ({primary_doc_type}"
            if primary_doc_confidence:
                primary_doc += f", confidence={primary_doc_confidence}"
            primary_doc += ")"

    object_name = field_value(project_card, "object_name") or clean_text(meta.get("object")) or report_dir.name
    customer = field_value(project_card, "customer") or clean_text(meta.get("customer"))
    address = field_value(project_card, "address") or field_value(project_card, "region")
    deadline = (
        field_value(project_card, "deadline_for_offer")
        or field_value(project_card, "deadline_for_works")
        or clean_text(meta.get("deadline"))
    )
    work_type = field_value(project_card, "request_type") or field_value(project_card, "object_type")
    stage = field_value(project_card, "project_stage") or compact_list(included_scope, limit=5)
    requirements_text = compact_list(requirement_parts or flatten_values(requirements), limit=10)
    tep_text = compact_list(tep_found, limit=10)
    missing_text = compact_list(missing_items, limit=10)
    risk_summary = compact_list(risk_texts, limit=8)
    recommendation = compact_list(missing_questions, limit=8)
    source_dir = source_message_dir(report_dir)
    source_path_text = str(source_dir) if source_dir is not None else ""
    mailbox = extract_mailbox_from_source_path(source_path_text)
    entry_date = clean_text(meta.get("date")) or extract_date_from_path(report_dir, [configured_roots(config, config_path)["reports"]])
    sufficiency = "частично достаточно" if missing_items else "достаточно"
    if clean_text(pir.get("status")) in {"needs_normative_data", "blocked"}:
        sufficiency = "частично достаточно"

    analysis_details = {
        "primary_technical_document": primary_doc,
        "project_card": {
            "object_name": object_name,
            "customer": customer,
            "address": address,
            "work_type": work_type,
            "stage": stage,
            "deadline": deadline,
        },
        "tep_found": tep_found,
        "tep_missing": tep_missing,
        "scope": included_scope,
        "requirements": requirement_parts,
        "red_flags": risk_texts,
        "missing_items": missing_items,
        "missing_questions": missing_questions,
        "documents": [Path(path_text).name for path_text in all_files],
        "report_dir": str(report_dir.resolve()),
        "source_dir": source_path_text,
    }

    values = [
        dt.datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
        mailbox,
        "",
        object_name,
        object_name,
        customer,
        "",
        address,
        work_type,
        stage,
        deadline,
        summarize_file_names(all_files),
        compact_list([*included_scope, *output_formats], limit=10),
        "; ".join(all_files),
        "; ".join(links),
        tep_text,
        tep_text,
        requirements_text,
        risk_summary,
        missing_text,
        risk_summary,
        recommendation,
        sufficiency,
        str(report_dir.resolve()),
    ]
    columns = dict(zip(report_headers(), values, strict=False))
    basis = str(source_dir.resolve()) if source_dir is not None else str(report_dir.resolve())
    return DashboardItem(
        id=stable_id("mail", basis) if source_dir is not None else stable_id("deep", basis),
        source_kind="deep_analysis",
        entry_date=entry_date,
        title=object_name,
        columns=columns,
        material_dir=str(report_dir.resolve()),
        source_path=str(path.resolve()),
        attachments=input_files,
        downloaded_files=downloaded_files,
        links=links,
        deep_analysis_dir=str(report_dir.resolve()),
        analysis=analysis_details,
    )


def merge_items(existing: DashboardItem, incoming: DashboardItem) -> DashboardItem:
    # Prefer deep analysis values, but keep mail fields when the deep adapter has blanks.
    if incoming.source_kind == "deep_analysis":
        primary, secondary = incoming, existing
    else:
        primary, secondary = existing, incoming
    headers = report_headers()
    merged_columns = {}
    for header in headers:
        merged_columns[header] = primary.columns.get(header) or secondary.columns.get(header, "")
    primary.columns = merged_columns
    primary.attachments = sorted(set([*primary.attachments, *secondary.attachments]))
    primary.downloaded_files = sorted(set([*primary.downloaded_files, *secondary.downloaded_files]))
    primary.links = sorted(set([*primary.links, *secondary.links]))
    if secondary.deep_analysis_dir and not primary.deep_analysis_dir:
        primary.deep_analysis_dir = secondary.deep_analysis_dir
    if secondary.analysis and not primary.analysis:
        primary.analysis = dict(secondary.analysis)
    if secondary.source_path and secondary.source_path not in primary.source_path:
        primary.source_path = f"{primary.source_path}; {secondary.source_path}"
    if secondary.entry_date and (not primary.entry_date or secondary.entry_date > primary.entry_date):
        primary.entry_date = secondary.entry_date
    return primary


def collect_deep_analysis_items(config: dict[str, Any], config_path: Path) -> list[DashboardItem]:
    items: list[DashboardItem] = []
    reports_root = configured_roots(config, config_path)["reports"]
    if reports_root.exists():
        for report_path in sorted(reports_root.rglob("12_analysis_report.json")):
            item = item_from_deep_analysis(report_path, config, config_path)
            if item is None:
                continue
            items.append(item)
    return items


def find_deep_match(title: str, deep_items: Sequence[DashboardItem]) -> DashboardItem | None:
    key = normalize_match_key(title)
    if not key:
        return None
    lookup = {normalize_match_key(item.title): item for item in deep_items if normalize_match_key(item.title)}
    if key in lookup:
        return lookup[key]
    for candidate_key, item in lookup.items():
        if len(key) > 24 and (key in candidate_key or candidate_key in key):
            return item
    return None


def item_from_excel_row(
    excel_path: Path,
    row_number: int,
    headers: Sequence[str],
    row: Sequence[str],
    deep_items: Sequence[DashboardItem],
    config: dict[str, Any],
) -> DashboardItem | None:
    columns = {
        header: clean_text(row[index]) if index < len(row) else ""
        for index, header in enumerate(headers)
        if header
    }
    if is_daily_report_path(excel_path) and daily_report_row_excluded(columns, config):
        return None
    raw_title = first_column(columns, ("Название объекта", "Объект", "Тема"))
    if not raw_title:
        return None
    title = normalize_daily_report_object_title(raw_title) if is_daily_report_path(excel_path) else raw_title
    if is_daily_report_path(excel_path):
        if "Название объекта" in columns:
            columns["Название объекта"] = title
        if "Объект" in columns or "Название объекта" not in columns:
            columns["Объект"] = title
    customer = first_column(columns, ("Заказчик", "Заказчик/отправитель"))
    entry_date = date_from_excel_row(columns, excel_path)
    deep_match = find_deep_match(title, deep_items)
    if deep_match is not None and deep_match.entry_date:
        entry_date = deep_match.entry_date

    links = split_links(first_column(columns, ("Ссылки в интернете", "Ссылки")))
    if deep_match is not None:
        links = list(dict.fromkeys([*links, *deep_match.links]))
    attachments = list(deep_match.attachments) if deep_match is not None else []
    downloaded_files = list(deep_match.downloaded_files) if deep_match is not None else []
    material_dir = deep_match.material_dir if deep_match is not None else str(excel_path.parent.resolve())
    item_id = deep_match.id if deep_match is not None else stable_id(
        "excel",
        f"{excel_path.resolve()}:{row_number}:{raw_title}:{customer}",
    )

    columns[DEADLINE_HEADER] = offer_deadline(columns)
    display_headers = display_headers_with_deadline(headers)
    if ANALYTIC_NOTE_HEADER not in display_headers:
        insert_after = "Чего не хватает" if "Чего не хватает" in display_headers else "Проблемы"
        if insert_after in display_headers:
            display_headers.insert(display_headers.index(insert_after) + 1, ANALYTIC_NOTE_HEADER)
        else:
            display_headers.append(ANALYTIC_NOTE_HEADER)
    if "Исходные файлы" not in display_headers:
        display_headers.append("Исходные файлы")
    if "Папка материалов" not in display_headers:
        display_headers.append("Папка материалов")
    columns[ANALYTIC_NOTE_HEADER] = build_analytic_note(columns, deep_match)
    columns.setdefault("Исходные файлы", "; ".join([*attachments, *downloaded_files]))
    columns.setdefault("Папка материалов", material_dir)

    return DashboardItem(
        id=item_id,
        source_kind="excel_relevant",
        entry_date=entry_date,
        title=title,
        columns=columns,
        material_dir=material_dir,
        source_path=str(excel_path.resolve()),
        headers=display_headers,
        attachments=attachments,
        downloaded_files=downloaded_files,
        links=links,
        deep_analysis_dir=deep_match.deep_analysis_dir if deep_match is not None else "",
        analysis=dict(deep_match.analysis) if deep_match is not None else {},
    )


def collect_excel_relevant_items(config: dict[str, Any], config_path: Path) -> list[DashboardItem]:
    excel_paths = iter_excel_paths(config, config_path)
    if not excel_paths:
        return []
    deep_items = collect_deep_analysis_items(config, config_path)
    items: list[DashboardItem] = []
    for excel_path in excel_paths:
        try:
            rows = read_xlsx_rows(excel_path)
        except (OSError, KeyError, ET.ParseError):
            continue
        if not rows:
            continue
        headers = unique_headers(rows[0])
        if not headers:
            continue
        for row_number, row in enumerate(rows[1:], start=2):
            if not any(clean_text(value) for value in row):
                continue
            item = item_from_excel_row(excel_path, row_number, headers, row, deep_items, config)
            if item is not None:
                items.append(item)
    return items


def collect_analysis_items(config: dict[str, Any], config_path: Path) -> list[DashboardItem]:
    items_by_id: dict[str, DashboardItem] = {}
    for root in analysis_scan_roots(config, config_path):
        for analysis_path in sorted(root.rglob("analysis.json")):
            item = item_from_mail_analysis(analysis_path, config, config_path)
            if item is None:
                continue
            if item.id in items_by_id:
                items_by_id[item.id] = merge_items(items_by_id[item.id], item)
            else:
                items_by_id[item.id] = item
    for item in collect_deep_analysis_items(config, config_path):
        if item.id in items_by_id:
            items_by_id[item.id] = merge_items(items_by_id[item.id], item)
        else:
            items_by_id[item.id] = item
    return list(items_by_id.values())


def collect_items(config: dict[str, Any], config_path: Path) -> list[DashboardItem]:
    excel_items = collect_excel_relevant_items(config, config_path)
    if excel_items:
        return excel_items
    return collect_analysis_items(config, config_path)


def apply_state(items: list[DashboardItem], state: dict[str, Any], mutate: bool = True) -> bool:
    changed = False
    records = state.setdefault("items", {})
    current = now_iso()
    today = today_iso()
    for item in items:
        record = records.get(item.id)
        if not isinstance(record, dict):
            if not mutate:
                record = {}
            else:
                record = {
                    "status": "unchecked",
                    "first_seen_at": current,
                    "entry_date": item.entry_date,
                    "title": item.title,
                }
                records[item.id] = record
                changed = True
        elif mutate:
            if not record.get("title") and item.title:
                record["title"] = item.title
                changed = True
            if not record.get("entry_date") and item.entry_date:
                record["entry_date"] = item.entry_date
                changed = True
            if not record.get("first_seen_at"):
                record["first_seen_at"] = current
                changed = True
        if mutate and snapshot_state_record(record, item):
            changed = True

        status = clean_text(record.get("status")) or "unchecked"
        normalized_status = LEGACY_STATUS_MAP.get(status, status)
        if normalized_status not in STATUS_DEFS:
            normalized_status = "unchecked"
        if mutate and normalized_status != status:
            record["status"] = normalized_status
            changed = True
        status = normalized_status
        item.status = status
        item.status_label = STATUS_DEFS[status]["label"]
        item.first_seen_at = clean_text(record.get("first_seen_at"))
        item.decision_updated_at = clean_text(record.get("decision_updated_at"))
        item.files_deleted_at = clean_text(record.get("files_deleted_at"))
        item.bitrix_lead_id = clean_text(record.get("bitrix_lead_id"))
        item.bitrix_task_id = clean_text(record.get("bitrix_task_id"))
        item.bitrix_error = clean_text(record.get("bitrix_error"))
        item.bitrix_task_error = clean_text(record.get("bitrix_task_error"))
        item.is_new = status == "unchecked" and item.first_seen_at[:10] == today
        item.compact = status == "not_relevant" and bool(item.files_deleted_at)
    return changed


def item_sort_key(item: DashboardItem) -> tuple[int, int, str]:
    status_sort = STATUS_DEFS.get(item.status, STATUS_DEFS["unchecked"])["sort"]
    if item.is_new:
        status_sort = 0
    try:
        date_sort = -dt.date.fromisoformat(item.entry_date).toordinal()
    except ValueError:
        date_sort = 0
    return (status_sort, date_sort, item.title.lower())


def dashboard_file_link_mode(config: dict[str, Any]) -> str:
    mode = clean_text(dashboard_config(config).get("file_link_mode")).lower()
    if mode in {"download", "http"}:
        return "download"
    return "local"


def dashboard_download_url(path_text: str) -> str:
    return "/api/download?path=" + urllib.parse.quote(path_text, safe="")


def dashboard_browse_url(path_text: str) -> str:
    return "/api/browse?path=" + urllib.parse.quote(path_text, safe="")


def map_path_for_dashboard(path_text: str, config: dict[str, Any]) -> str:
    if not path_text:
        return ""
    dash = dashboard_config(config)
    local_prefix = clean_text(dash.get("local_path_prefix"))
    network_prefix = clean_text(dash.get("network_path_prefix"))
    if not network_prefix:
        return path_text
    if local_prefix:
        norm_path = os.path.normcase(os.path.normpath(path_text))
        norm_local = os.path.normcase(os.path.normpath(local_prefix))
        if norm_path == norm_local:
            return network_prefix
        if norm_path.startswith(norm_local + os.sep):
            remainder = os.path.normpath(path_text)[len(os.path.normpath(local_prefix)):].lstrip("\\/")
            return str(Path(network_prefix) / remainder)
    return path_text


def file_url(path_text: str) -> str:
    if not path_text:
        return ""
    normalized = path_text.replace("\\", "/")
    if path_text.startswith("\\\\"):
        return "file://///" + normalized.lstrip("/")
    if re.match(r"^[a-zA-Z]:[\\/]", path_text):
        return "file:///" + urllib.parse.quote(normalized, safe="/:")
    parsed = urllib.parse.urlparse(path_text)
    if parsed.scheme:
        return path_text
    try:
        return Path(path_text).absolute().as_uri()
    except ValueError:
        return "file:///" + urllib.parse.quote(normalized)


def looks_like_local_path(path_text: str) -> bool:
    text = clean_text(path_text)
    return bool(re.match(r"^[a-zA-Z]:[\\/]", text) or text.startswith("\\\\") or text.startswith("//"))


def linked_path(path_text: str, config: dict[str, Any]) -> dict[str, Any]:
    source = normalize_path_string(path_text)
    mapped = map_path_for_dashboard(source, config)
    if dashboard_file_link_mode(config) == "download":
        is_dir = False
        try:
            is_dir = Path(source).resolve().is_dir()
        except OSError:
            is_dir = False
        return {
            "path": mapped,
            "url": dashboard_browse_url(source) if is_dir else dashboard_download_url(source),
            "name": Path(mapped).name or mapped,
            "local": False,
            "download": not is_dir,
        }
    return {
        "path": mapped,
        "url": file_url(mapped),
        "name": Path(mapped).name or mapped,
        "local": looks_like_local_path(mapped),
    }


def linked_material_folder(item: DashboardItem, config: dict[str, Any]) -> dict[str, Any]:
    source = clean_text(item.analysis.get("source_dir") if isinstance(item.analysis, dict) else "")
    if not source:
        specific_dirs = item_specific_material_dirs(item)
        source = str(specific_dirs[0]) if specific_dirs else ""
    if not source:
        return {"path": "", "url": "", "name": ""}
    mapped = map_path_for_dashboard(source, config)
    if mapped and mapped != source and looks_like_local_path(mapped):
        return {
            "path": mapped,
            "url": file_url(mapped),
            "name": "папка на диске Z" if re.match(r"^z:[\\/]", mapped, flags=re.IGNORECASE) else (Path(mapped).name or mapped),
            "local": False,
            "browserOnly": True,
        }
    return linked_path(item.material_dir, config) if item.material_dir else {"path": "", "url": "", "name": ""}


def extract_urls_from_text(value: str) -> list[str]:
    matches = re.findall(r"https?://[^\s;,)]+", value or "")
    return list(dict.fromkeys(match.rstrip(".,)") for match in matches))


def platform_label_from_text(value: str) -> str:
    lowered = value.lower()
    platform_tokens = [
        ("zakupki360", "Закупки 360"),
        ("zakupki.gov", "ЕИС zakupki.gov.ru"),
        ("b2b-center", "B2B-Center"),
        ("b2b-rts", "B2B-РТС"),
        ("rts-tender", "РТС-тендер"),
        ("sberbank-ast", "Сбербанк-АСТ"),
        ("etp-ets", "ЭТП ЭТС"),
        ("etp.gpb", "ЭТП ГПБ"),
        ("etpgpb", "ЭТП ГПБ"),
        ("тэк-торг", "ТЭК-Торг"),
        ("tektorg", "ТЭК-Торг"),
        ("roseltorg", "Росэлторг"),
        ("fabrikant", "Фабрикант"),
        ("tender.pro", "Tender.Pro"),
        ("lot-online", "Lot-online"),
    ]
    for token, label in platform_tokens:
        if token in lowered:
            return label
    return ""


def procurement_source(item: DashboardItem) -> str:
    columns = item.columns
    link_text = " ".join([*item.links, first_column(columns, ("Ссылки в интернете", "Ссылки"))])
    platform = platform_label_from_text(link_text)
    if platform:
        return platform
    for path_text in [item.source_path, item.material_dir, clean_text(item.analysis.get("source_dir") if isinstance(item.analysis, dict) else "")]:
        platform = platform_label_from_text(path_text)
        if platform:
            return platform
    return first_column(columns, ("Почтовый ящик", "Источник закупки", "Источник")) or "—"


def read_item_request_text(item: DashboardItem) -> str:
    candidates = []
    for base_text in [item.deep_analysis_dir, item.material_dir]:
        if not base_text:
            continue
        base = Path(base_text)
        candidates.extend([base / "_input" / "request.txt", base / "request.txt"])
    for path in candidates:
        try:
            if path.exists() and path.is_file():
                return path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
    return ""


def extract_nmc_from_text(value: str) -> str:
    if not value:
        return ""
    patterns = [
        r"(?:Цена/НМЦК|НМЦК)\s*[:\-]?\s*([0-9][0-9\s.,]*(?:₽|руб\.?|р\.?)?)",
        r"(?:Начальная\s*\(?максимальная\)?\s*цена(?:\s*(?:контракта|договора))?|Начальная цена|Общая стоимость закупки)\s*[:\-]?\s*([0-9][0-9\s.,]*(?:₽|руб\.?|р\.?)?)",
        r"\(на сумму\s*([0-9][0-9\s.,]*(?:₽|руб\.?|р\.?)?)\s*\)",
    ]
    for pattern in patterns:
        match = re.search(pattern, value, flags=re.IGNORECASE)
        if match:
            amount = re.sub(r"\s+", " ", match.group(1)).strip()
            if re.search(r"\d", amount):
                return amount if re.search(r"₽|руб|р\.", amount, flags=re.IGNORECASE) else f"{amount} ₽"
    if re.search(r"без указания цены", value, flags=re.IGNORECASE):
        return "Без указания цены"
    return ""


def procurement_nmc(item: DashboardItem) -> str:
    direct = first_column(
        item.columns,
        (
            "НМЦК",
            "Цена/НМЦК",
            "Начальная цена",
            "Начальная максимальная цена контракта",
            "Начальная (максимальная) цена контракта",
            "Начальная (максимальная) цена договора",
            "Стоимость",
        ),
    )
    if direct:
        return direct
    from_columns = extract_nmc_from_text(" ".join(clean_text(value) for value in item.columns.values()))
    if from_columns:
        return from_columns
    return extract_nmc_from_text(read_item_request_text(item))


def rub_amount_from_text(value: str) -> float | None:
    text = clean_text(value)
    if not text or re.search(r"без указания цены", text, flags=re.IGNORECASE):
        return None
    match = re.search(r"\d[\d\s.,]*", text)
    if not match:
        return None
    raw = re.sub(r"\s+", "", match.group(0))
    if "," in raw:
        whole, fraction = raw.rsplit(",", 1)
        if 0 < len(fraction) <= 2:
            normalized = re.sub(r"\D", "", whole) + "." + re.sub(r"\D", "", fraction)
        else:
            normalized = re.sub(r"\D", "", raw)
    elif "." in raw:
        whole, fraction = raw.rsplit(".", 1)
        if 0 < len(fraction) <= 2:
            normalized = re.sub(r"\D", "", whole) + "." + re.sub(r"\D", "", fraction)
        else:
            normalized = re.sub(r"\D", "", raw)
    else:
        normalized = re.sub(r"\D", "", raw)
    try:
        return float(normalized) if normalized else None
    except ValueError:
        return None


def dashboard_min_nmc_rub(config: dict[str, Any]) -> float:
    raw = dashboard_config(config).get("min_nmc_rub", DEFAULT_MIN_NMC_RUB)
    try:
        return float(raw)
    except (TypeError, ValueError):
        return float(DEFAULT_MIN_NMC_RUB)


def dashboard_item_visible_by_nmc(item: DashboardItem, config: dict[str, Any]) -> bool:
    minimum = dashboard_min_nmc_rub(config)
    if minimum <= 0:
        return True
    amount = rub_amount_from_text(procurement_nmc(item))
    return amount is None or amount >= minimum


def procurement_customer(item: DashboardItem) -> str:
    value = first_column(item.columns, ("Заказчик", "Заказчик/отправитель"))
    if not value:
        return ""
    value = re.split(r"\s+(?:Регион/площадка|Цена/НМЦК|НМЦК|Начальная цена)\s*:", value, maxsplit=1, flags=re.IGNORECASE)[0]
    return value.strip(" ;")


def tep_summary(item: DashboardItem) -> str:
    if isinstance(item.analysis, dict):
        tep_found = item.analysis.get("tep_found")
        if isinstance(tep_found, list):
            text = limited_join([clean_text(value) for value in tep_found], limit=8)
            if text:
                return text
    return first_column(item.columns, ("ТЭП", "Ключевые технические параметры", "Объемы"))


def technical_assignment_score(path: Path) -> int:
    name = path.name.lower()
    if any(token in name for token in ("нмцк", "обоснование", "контракт", "договор", "заявк", "протокол", "разъяснен")):
        return 0
    score = 0
    if "описание объекта закупки" in name:
        score += 80
    if "техническое задание" in name or "техническое_задание" in name:
        score += 80
    if re.search(r"(^|[^а-яa-z])тз([^а-яa-z]|$)", name):
        score += 60
    if "задани" in name:
        score += 35
    if "техническ" in name:
        score += 30
    if "объект" in name and "закуп" in name:
        score += 30
    if "описание" in name:
        score += 20
    if path.suffix.lower() in {".pdf", ".doc", ".docx", ".xls", ".xlsx"}:
        score += 10
    return score


def item_specific_material_dirs(item: DashboardItem) -> list[Path]:
    dirs: list[Path] = []
    for value in [item.deep_analysis_dir, item.material_dir]:
        text = clean_text(value)
        if not text:
            continue
        path = Path(text)
        if path.name.lower() in {"reports", "релевантные", "archive", "sorted_mail"}:
            continue
        try:
            if path.exists() and path.is_dir():
                dirs.append(path.resolve())
        except OSError:
            continue
    result: list[Path] = []
    seen: set[str] = set()
    for path in dirs:
        key = str(path)
        if key not in seen:
            seen.add(key)
            result.append(path)
    return result


def technical_assignment_paths(item: DashboardItem) -> list[str]:
    candidates: dict[str, Path] = {}
    direct_values = [
        first_column(item.columns, ("ТЗ", "Техническое задание", "Описание объекта закупки")),
        *item.attachments,
        *item.downloaded_files,
        *split_semicolon_paths(item.columns.get("Исходные файлы", "")),
    ]
    for value in direct_values:
        text = clean_text(value)
        if not text or re.match(r"^[a-z][a-z0-9+.-]*://", text, flags=re.IGNORECASE):
            continue
        if os.name != "nt" and re.match(r"^[A-Za-z]:[\\/]", text):
            continue
        path = Path(text)
        try:
            if path.exists() and path.is_file():
                candidates[str(path.resolve())] = path.resolve()
        except OSError:
            continue

    for base in item_specific_material_dirs(item):
        try:
            for path in base.rglob("*"):
                if path.is_file():
                    candidates[str(path.resolve())] = path.resolve()
        except OSError:
            continue

    scored = [(technical_assignment_score(path), path) for path in candidates.values()]
    selected = [path for score, path in sorted(scored, key=lambda pair: (-pair[0], pair[1].name.lower())) if score > 0]
    return [str(path) for path in selected[:3]]


def analytic_note_slug(item: DashboardItem) -> str:
    base = normalize_match_key(item.title or first_column(item.columns, ("Объект", "Название объекта", "Тема")))[:60]
    return f"{item.id}_{base or 'zakupka'}.md"


def analytic_note_target_path(item: DashboardItem, config: dict[str, Any], config_path: Path) -> Path:
    for base in item_specific_material_dirs(item):
        if base.exists() and base.is_dir():
            return base / ANALYTIC_NOTE_FILENAME
    reports_root = configured_roots(config, config_path)["reports"]
    date_part = item.entry_date if looks_like_date(item.entry_date) else today_iso()
    return reports_root / "analytic_notes" / date_part / analytic_note_slug(item)


def markdown_link_lines(values: Sequence[str]) -> list[str]:
    lines: list[str] = []
    for value in values:
        text = clean_text(value)
        if not text:
            continue
        if re.match(r"^https?://", text, flags=re.IGNORECASE):
            lines.append(f"- <{text}>")
        else:
            lines.append(f"- `{text}`")
    return lines


def analytic_note_markdown(item: DashboardItem) -> str:
    columns = item.columns
    title = first_column(columns, ("Объект", "Название объекта", "Тема")) or item.title
    customer = procurement_customer(item) or first_column(columns, ("Заказчик", "Заказчик/отправитель"))
    links = list(dict.fromkeys([*item.links, *extract_urls_from_text(first_column(columns, ("Ссылки в интернете", "Ссылки")))]))
    source_files = split_semicolon_paths(columns.get("Исходные файлы", "")) or [*item.attachments, *item.downloaded_files]
    technical_files = technical_assignment_paths(item)
    note = first_column(columns, (ANALYTIC_NOTE_HEADER,)) or build_analytic_note(columns)
    recommendation = first_column(columns, ("Рекомендация следующего действия", "Рекомендация"))
    parts = [
        "# Аналитическая записка",
        "",
        f"Дата формирования: {now_iso()}",
        f"Дата входа: {item.entry_date or 'не указана'}",
        f"Источник закупки: {procurement_source(item)}",
        "",
        "## Закупка",
        "",
        f"Объект: {title or 'не указан'}",
        f"Заказчик: {customer or 'не указан'}",
        f"НМЦК: {procurement_nmc(item) or 'не указана'}",
        f"Окончание подачи предложений: {offer_deadline(columns) or 'не указано'}",
        f"ТЭП: {tep_summary(item) or 'не извлечены'}",
        "",
        "## Ссылки на закупку",
        "",
        *(markdown_link_lines(links) or ["- не найдены"]),
        "",
        "## ТЗ и материалы",
        "",
        *(markdown_link_lines(technical_files) or markdown_link_lines(source_files) or ["- не найдены"]),
        "",
        "## Анализ",
        "",
        note or "Аналитическая часть не сформирована.",
        "",
        "## Рекомендация",
        "",
        recommendation or "Выполнить ручную проверку ТЗ, сроков и состава документации перед решением о подаче.",
        "",
    ]
    return "\n".join(parts)


def ensure_analytic_note_file(item: DashboardItem, config: dict[str, Any], config_path: Path) -> str:
    if item.compact:
        return ""
    try:
        target = analytic_note_target_path(item, config, config_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        content = analytic_note_markdown(item)
        if not target.exists() or target.read_text(encoding="utf-8", errors="ignore") != content:
            target.write_text(content, encoding="utf-8")
        return str(target.resolve())
    except OSError:
        return ""


def dashboard_display_columns(item: DashboardItem) -> dict[str, str]:
    columns = item.columns
    return {
        "Источник закупки": procurement_source(item),
        "Объект": first_column(columns, ("Объект", "Название объекта", "Тема")) or item.title,
        "Заказчик": procurement_customer(item),
        "НМЦК": procurement_nmc(item),
        "Окончание подачи предложений": offer_deadline(columns),
        "ТЗ": "Файлы удалены" if item.compact else "",
        "Аналитическая записка": "Файлы удалены" if item.compact else "",
        "Ссылки на закупку": first_column(columns, ("Ссылки в интернете", "Ссылки")),
        "ТЭП": tep_summary(item),
        "Рекомендация следующего действия": first_column(columns, ("Рекомендация следующего действия", "Рекомендация")),
        "Папка материалов": "",
    }


def path_from_file_url(value: str) -> str:
    parsed = urllib.parse.urlparse(value)
    if parsed.scheme != "file":
        return value
    path = urllib.request.url2pathname(parsed.path)
    if parsed.netloc:
        return "\\\\" + parsed.netloc + path.replace("/", "\\")
    return path


def open_dashboard_path(path_text: str) -> dict[str, str]:
    target = path_from_file_url(clean_text(path_text))
    if not target:
        raise ValueError("Путь не передан")
    if not looks_like_local_path(target):
        raise ValueError("Можно открыть только локальный или сетевой путь")
    if not os.path.exists(target):
        raise FileNotFoundError(f"Файл или папка не найдены: {target}")
    if os.name == "nt":
        os.startfile(target)  # type: ignore[attr-defined]
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([opener, target])
    return {"status": "opened", "path": target}


def dashboard_allowed_roots(config: dict[str, Any], config_path: Path) -> list[Path]:
    roots = configured_roots(config, config_path)
    allowed = [roots["relevant"], roots["reports"]]
    dash = dashboard_config(config)
    if bool(dash.get("include_archive_root", False)):
        allowed.append(roots["output"])
    extra_roots = dash.get("download_roots")
    if isinstance(extra_roots, list):
        allowed.extend(resolve_path(root, config_path.parent) for root in extra_roots)
    return [root for root in allowed if root.exists()]


def safe_dashboard_download_path(path_text: str, config: dict[str, Any], config_path: Path) -> Path:
    target = path_from_file_url(clean_text(path_text))
    if not target:
        raise ValueError("Путь не передан")
    try:
        resolved = Path(target).resolve()
    except OSError as exc:
        raise FileNotFoundError(target) from exc
    if not resolved.exists():
        raise FileNotFoundError(f"Файл или папка не найдены: {target}")
    allowed_roots = dashboard_allowed_roots(config, config_path)
    if not any(path_within(resolved, root) or resolved == root.resolve() for root in allowed_roots):
        raise PermissionError("Путь вне разрешенных папок dashboard")
    return resolved


def content_disposition(filename: str, disposition: str = "inline") -> str:
    fallback = re.sub(r"[^A-Za-z0-9._ -]+", "_", filename) or "download"
    quoted = urllib.parse.quote(filename)
    return f'{disposition}; filename="{fallback}"; filename*=UTF-8\'\'{quoted}'


def send_stream(
    handler: BaseHTTPRequestHandler,
    stream: Any,
    content_type: str,
    content_length: int,
    filename: str,
    disposition: str = "inline",
) -> None:
    handler.send_response(200)
    handler.send_header("Content-Type", content_type)
    handler.send_header("Content-Length", str(content_length))
    handler.send_header("Content-Disposition", content_disposition(filename, disposition))
    handler.end_headers()
    while True:
        chunk = stream.read(1024 * 256)
        if not chunk:
            break
        try:
            handler.wfile.write(chunk)
        except (BrokenPipeError, ConnectionResetError):
            break


def send_dashboard_download(
    handler: BaseHTTPRequestHandler,
    path_text: str,
    config: dict[str, Any],
    config_path: Path,
) -> None:
    target = safe_dashboard_download_path(path_text, config, config_path)
    if target.is_dir():
        archive_name = f"{target.name or 'materials'}.zip"
        buffer = io.BytesIO()
        with ZipFile(buffer, "w") as zip_file:
            for file_path in sorted(path for path in target.rglob("*") if path.is_file()):
                try:
                    arcname = file_path.relative_to(target)
                except ValueError:
                    continue
                zip_file.write(file_path, arcname.as_posix())
        buffer.seek(0)
        send_stream(
            handler,
            buffer,
            "application/zip",
            buffer.getbuffer().nbytes,
            archive_name,
            disposition="attachment",
        )
        return
    content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
    with target.open("rb") as stream:
        send_stream(handler, stream, content_type, target.stat().st_size, target.name)


def send_html(handler: BaseHTTPRequestHandler, body: str, status: int = 200) -> None:
    data = body.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def send_dashboard_browse(
    handler: BaseHTTPRequestHandler,
    path_text: str,
    config: dict[str, Any],
    config_path: Path,
) -> None:
    target = safe_dashboard_download_path(path_text, config, config_path)
    if target.is_file():
        send_dashboard_download(handler, str(target), config, config_path)
        return
    rows: list[str] = []
    for child in sorted(target.iterdir(), key=lambda path: (path.is_file(), path.name.lower())):
        icon = "Папка" if child.is_dir() else "Файл"
        url = dashboard_browse_url(str(child)) if child.is_dir() else dashboard_download_url(str(child))
        try:
            size = "" if child.is_dir() else f"{child.stat().st_size / 1024:,.1f} КБ".replace(",", " ")
        except OSError:
            size = ""
        rows.append(
            "<tr>"
            f"<td>{html.escape(icon)}</td>"
            f'<td><a href="{html.escape(url, quote=True)}">{html.escape(child.name)}</a></td>'
            f"<td>{html.escape(size)}</td>"
            "</tr>"
        )
    zip_url = dashboard_download_url(str(target))
    title = html.escape(target.name or str(target))
    body = f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <style>
    body {{ margin: 0; padding: 24px; font-family: "Segoe UI", Arial, sans-serif; color: #17202a; background: #f5f7fb; }}
    main {{ max-width: 1180px; margin: 0 auto; }}
    h1 {{ margin: 0 0 8px; font-size: 24px; }}
    .path {{ color: #647084; margin-bottom: 18px; word-break: break-all; }}
    a {{ color: #1f6fd1; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .action {{ display: inline-block; margin-bottom: 16px; padding: 8px 12px; border: 1px solid #d9e0ea; border-radius: 8px; background: #fff; }}
    table {{ width: 100%; border-collapse: collapse; background: #fff; border: 1px solid #d9e0ea; }}
    th, td {{ padding: 8px 10px; border-bottom: 1px solid #d9e0ea; text-align: left; font-size: 14px; }}
    th {{ background: #eef3f8; }}
  </style>
</head>
<body>
<main>
  <h1>{title}</h1>
  <div class="path">{html.escape(str(target))}</div>
  <a class="action" href="{html.escape(zip_url, quote=True)}">Скачать всю папку ZIP</a>
  <table>
    <thead><tr><th>Тип</th><th>Имя</th><th>Размер</th></tr></thead>
    <tbody>{''.join(rows) if rows else '<tr><td colspan="3">Папка пустая</td></tr>'}</tbody>
  </table>
</main>
</body>
</html>"""
    send_html(handler, body)


def headers_for_items(items: Sequence[DashboardItem]) -> list[str]:
    return list(DASHBOARD_DISPLAY_HEADERS)


def serialise_item(item: DashboardItem, config: dict[str, Any], config_path: Path, headers: Sequence[str]) -> dict[str, Any]:
    columns = dashboard_display_columns(item)

    files = [linked_path(path, config) for path in sorted(set([*item.attachments, *item.downloaded_files])) if path]
    downloaded = [linked_path(path, config) for path in sorted(set(item.downloaded_files)) if path]
    folder = linked_material_folder(item, config)
    technical_assignment_files = [linked_path(path, config) for path in technical_assignment_paths(item)]
    analytic_note_path = ensure_analytic_note_file(item, config, config_path)
    analytic_note_file = linked_path(analytic_note_path, config) if analytic_note_path else {}
    return {
        "id": item.id,
        "sourceKind": item.source_kind,
        "entryDate": item.entry_date,
        "title": item.title,
        "status": item.status,
        "statusLabel": item.status_label,
        "isNew": item.is_new,
        "compact": item.compact,
        "firstSeenAt": item.first_seen_at,
        "decisionUpdatedAt": item.decision_updated_at,
        "filesDeletedAt": item.files_deleted_at,
        "bitrixLeadId": item.bitrix_lead_id,
        "bitrixTaskId": item.bitrix_task_id,
        "bitrixError": item.bitrix_error,
        "bitrixTaskError": item.bitrix_task_error,
        "columns": columns,
        "folder": folder,
        "files": files,
        "downloadedFiles": downloaded,
        "technicalAssignmentFiles": technical_assignment_files,
        "analyticNoteFile": analytic_note_file,
        "links": item.links,
        "sourcePath": item.source_path,
        "deepAnalysisDir": item.deep_analysis_dir,
    }


def dashboard_payload(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    state = load_state(config, config_path)
    items = collect_items(config, config_path)
    items = append_retained_unprocessed_items(items, state, config)
    source_counts = source_analysis_counts(config, config_path)
    changed = apply_state(items, state, mutate=True)
    if changed:
        save_state(config, config_path, state)
    all_items = items
    items = [item for item in all_items if dashboard_item_visible_by_nmc(item, config)]
    items.sort(key=item_sort_key)
    headers = headers_for_items(items)
    serialised = [serialise_item(item, config, config_path, headers) for item in items]
    stats = {
        "total": len(items),
        "excludedByNmc": len(all_items) - len(items),
        "minNmcRub": dashboard_min_nmc_rub(config),
        "new": sum(1 for item in items if item.is_new),
        "unchecked": sum(1 for item in items if item.status == "unchecked"),
        "notRelevant": sum(1 for item in items if item.status == "not_relevant"),
        "inWork": sum(1 for item in items if item.status == "submitting"),
    }
    dates = sorted({item.entry_date for item in items if item.entry_date}, reverse=True)
    return {
        "generatedAt": now_iso(),
        "headers": headers,
        "statuses": STATUS_DEFS,
        "stats": stats,
        "dates": dates,
        "sourceCounts": source_counts,
        "fileLinkMode": dashboard_file_link_mode(config),
        "items": serialised,
    }


def record_list_append(record: dict[str, Any], key: str, values: Iterable[str]) -> None:
    previous = record.get(key, [])
    if not isinstance(previous, list):
        previous = []
    cleaned = [clean_text(value) for value in values if clean_text(value)]
    if cleaned:
        record[key] = sorted(set([*previous, *cleaned]))


def cleanup_roots(config: dict[str, Any], config_path: Path) -> list[Path]:
    roots = configured_roots(config, config_path)
    allowed = [roots["reports"], roots["relevant"]]
    if bool(dashboard_config(config).get("include_archive_root", False)):
        allowed.append(roots["output"])
    return [root.resolve() for root in allowed]


def cleanup_relative_path(path: Path, roots: Sequence[Path]) -> str:
    try:
        resolved = path.resolve()
    except OSError:
        return ""
    for root in roots:
        try:
            relative = resolved.relative_to(root.resolve())
        except (OSError, ValueError):
            continue
        if not relative.parts:
            return ""
        return "/".join([root.name, relative.as_posix()])
    return ""


def cleanup_candidate_strings(item: DashboardItem) -> list[str]:
    candidates = [
        item.material_dir,
        item.deep_analysis_dir,
        clean_text(item.analysis.get("report_dir") if isinstance(item.analysis, dict) else ""),
        clean_text(item.analysis.get("source_dir") if isinstance(item.analysis, dict) else ""),
        *item.attachments,
        *item.downloaded_files,
        *split_semicolon_paths(item.columns.get("Исходные файлы", "")),
    ]
    return list(dict.fromkeys(clean_text(candidate) for candidate in candidates if clean_text(candidate)))


def resolve_cleanup_candidate(path_text: str, config_path: Path) -> Path | None:
    text = clean_text(path_text).strip().strip('"')
    if not text or re.match(r"^[a-z][a-z0-9+.-]*://", text, flags=re.IGNORECASE):
        return None
    if re.match(r"^[A-Za-z]:[\\/]", text):
        return None
    path = Path(text)
    if not path.is_absolute():
        path = config_path.parent / path
    try:
        return path.resolve()
    except OSError:
        return None


def safe_cleanup_file(path: Path, roots: Sequence[Path]) -> bool:
    try:
        if not path.exists() or not path.is_file():
            return False
    except OSError:
        return False
    if path.name in METADATA_FILENAMES or path.suffix.lower() == ".json":
        return False
    if not any(path_within(path, root) for root in roots):
        return False
    for root in roots:
        try:
            relative = path.resolve().relative_to(root.resolve())
        except (OSError, ValueError):
            continue
        if len(relative.parts) == 1 and path.suffix.lower() in {".xlsx", ".xls"}:
            return False
    return True


def iter_cleanup_files(path: Path, roots: Sequence[Path], errors: list[str] | None = None) -> list[Path]:
    try:
        is_file = path.is_file()
        exists = path.exists()
        is_dir = path.is_dir()
    except OSError as exc:
        if errors is not None:
            errors.append(f"{path}: {exc}")
        return []
    if is_file:
        return [path] if safe_cleanup_file(path, roots) else []
    if not exists or not is_dir:
        return []
    if not any(path_within(path, root) for root in roots):
        return []
    try:
        if any(path.resolve() == root.resolve() for root in roots):
            return []
    except OSError as exc:
        if errors is not None:
            errors.append(f"{path}: {exc}")
        return []
    result: list[Path] = []
    try:
        for candidate in path.rglob("*"):
            try:
                if safe_cleanup_file(candidate, roots):
                    result.append(candidate)
            except OSError as exc:
                if errors is not None:
                    errors.append(f"{candidate}: {exc}")
    except OSError as exc:
        if errors is not None:
            errors.append(f"{path}: {exc}")
    return result


def prune_empty_dirs_to_roots(start: Path, roots: Sequence[Path]) -> list[str]:
    removed: list[str] = []
    current = start
    while True:
        try:
            if any(current.resolve() == root.resolve() for root in roots):
                break
        except OSError:
            break
        if not any(path_within(current, root) for root in roots):
            break
        try:
            current.rmdir()
            removed.append(str(current))
        except OSError:
            break
        current = current.parent
    return removed


def delete_not_relevant_files(item: DashboardItem, record: dict[str, Any], config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    roots = cleanup_roots(config, config_path)
    candidate_texts = cleanup_candidate_strings(item)
    files: dict[str, Path] = {}
    relative_candidates: list[str] = []
    errors: list[str] = []
    for text in candidate_texts:
        path = resolve_cleanup_candidate(text, config_path)
        if path is None:
            continue
        relative = cleanup_relative_path(path, roots)
        if relative:
            relative_candidates.append(relative)
        for candidate in iter_cleanup_files(path, roots, errors):
            files[str(candidate)] = candidate

    deleted: list[str] = []
    deleted_relative: list[str] = []
    deleted_dirs: list[str] = []
    for path in sorted(files.values(), key=lambda value: len(value.parts), reverse=True):
        relative = cleanup_relative_path(path, roots)
        try:
            path.unlink()
            deleted.append(str(path))
            if relative:
                deleted_relative.append(relative)
            deleted_dirs.extend(prune_empty_dirs_to_roots(path.parent, roots))
        except OSError as exc:
            errors.append(f"{path}: {exc}")

    record_list_append(record, "deleted_files", deleted)
    record_list_append(record, "deleted_relative_paths", [*deleted_relative, *relative_candidates])
    record_list_append(record, "local_delete_candidates", candidate_texts)
    record_list_append(record, "deleted_dirs", deleted_dirs)
    if errors:
        record["cleanup_error"] = "; ".join(errors)
    else:
        record.pop("cleanup_error", None)
    record["files_deleted_at"] = now_iso()
    return {
        "status": "cleaned" if deleted else "marked",
        "deletedCount": len(deleted),
        "deletedFiles": deleted,
        "deletedDirs": deleted_dirs,
        "errors": errors,
    }


def update_item_status(
    item_id: str,
    status: str,
    config: dict[str, Any],
    config_path: Path,
) -> dict[str, Any]:
    if status not in STATUS_DEFS:
        raise ValueError(f"Unknown status: {status}")
    state = load_state(config, config_path)
    records = state.setdefault("items", {})
    record = records.setdefault(item_id, {})
    record["status"] = status
    record["decision_updated_at"] = now_iso()
    record.pop("bitrix_error", None)
    record.pop("bitrix_task_error", None)

    bitrix_result: dict[str, Any] | None = None
    cleanup_result: dict[str, Any] | None = None
    if status in {"relevant", "not_relevant"}:
        items = collect_items(config, config_path)
        apply_state(items, state, mutate=False)
        item = next((candidate for candidate in items if candidate.id == item_id), None)
        if item is None:
            raise KeyError(item_id)
    else:
        item = None

    if status == "relevant" and item is not None:
        bitrix_result = push_item_to_bitrix(item, record, config)
    elif status == "not_relevant" and item is not None:
        try:
            cleanup_result = delete_not_relevant_files(item, record, config, config_path)
        except OSError as exc:
            record["cleanup_error"] = str(exc)
            record["files_deleted_at"] = now_iso()
            cleanup_result = {"status": "marked", "deletedCount": 0, "deletedFiles": [], "deletedDirs": [], "errors": [str(exc)]}

    save_state(config, config_path, state)
    payload = dashboard_payload(config, config_path)
    payload["bitrix"] = bitrix_result
    payload["cleanup"] = cleanup_result
    return payload


def bitrix_method_url(webhook: str, method: str) -> str:
    webhook = webhook.strip()
    if not webhook:
        return ""
    if method in webhook:
        return webhook
    parsed = urllib.parse.urlsplit(webhook)
    path = parsed.path.rstrip("/")
    parts = path.split("/") if path else []
    if parts and (parts[-1].endswith(".json") or "." in parts[-1]):
        path = "/".join(parts[:-1])
    path = path.rstrip("/") + f"/{method}.json"
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, path, "", ""))


def call_bitrix_method(webhook: str, method: str, fields: dict[str, Any], timeout: int) -> dict[str, Any]:
    url = bitrix_method_url(webhook, method)
    body = urllib.parse.urlencode(fields, doseq=True).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8", errors="ignore"))
    return payload if isinstance(payload, dict) else {"result": payload}


def extract_bitrix_task_id(payload: dict[str, Any]) -> str:
    result = payload.get("result")
    if isinstance(result, dict):
        task = result.get("task")
        if isinstance(task, dict):
            return clean_text(task.get("id"))
        return clean_text(result.get("id"))
    return clean_text(result)


def bitrix_task_deadline(dash: dict[str, Any]) -> str:
    hours = clean_text(dash.get("bitrix_task_deadline_hours"))
    days = clean_text(dash.get("bitrix_task_deadline_days"))
    try:
        if hours:
            delta = dt.timedelta(hours=int(hours))
        elif days:
            delta = dt.timedelta(days=int(days))
        else:
            return ""
    except ValueError:
        return ""
    return (dt.datetime.now() + delta).isoformat(timespec="seconds")


def push_item_to_bitrix(item: DashboardItem, record: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    dash = dashboard_config(config)
    webhook = clean_text(
        os.environ.get("TENDER_DASHBOARD_BITRIX_WEBHOOK_URL")
        or dash.get("bitrix_webhook_url")
        or config.get("bitrix_webhook_url")
    )
    if not webhook:
        portal = clean_text(os.environ.get("TENDER_DASHBOARD_BITRIX_PORTAL_URL") or dash.get("bitrix_portal_url"))
        record["bitrix_error"] = (
            "Не задан TENDER_DASHBOARD_BITRIX_WEBHOOK_URL"
            + (f" для портала {portal}" if portal else "")
        )
        return {"status": "not_configured", "message": record["bitrix_error"]}

    columns = item.columns
    title = first_column(columns, ("Название объекта", "Объект", "Тема")) or item.title
    mapped_source_files = "; ".join(
        map_path_for_dashboard(path, config)
        for path in (split_semicolon_paths(columns.get("Исходные файлы", "")) or [*item.attachments, *item.downloaded_files])
    )
    customer = first_column(columns, ("Заказчик", "Заказчик/отправитель"))
    deadline = offer_deadline(columns)
    links_text = first_column(columns, ("Ссылки в интернете", "Ссылки"))
    risks = first_column(columns, ("Проблемы", "Риски"))
    missing = first_column(columns, ("Чего не хватает", "Недостающие сведения"))
    analytic_note = first_column(columns, (ANALYTIC_NOTE_HEADER,))
    recommendation = first_column(columns, ("Рекомендация", "Рекомендация следующего действия"))
    comments = "\n".join(
        part
        for part in (
            f"Дата входа: {item.entry_date}",
            f"Объект: {title}",
            f"Заказчик/отправитель: {customer}",
            f"Сроки: {deadline}",
            f"Статус достаточности: {columns.get('Статус достаточности', '')}",
            f"Папка материалов: {map_path_for_dashboard(item.material_dir, config)}",
            f"Исходные файлы: {mapped_source_files}",
            f"Ссылки: {links_text}",
            f"Риски: {risks}",
            f"Чего не хватает: {missing}",
            f"Аналитическая записка: {analytic_note}",
            f"Рекомендация: {recommendation}",
        )
        if part and not part.endswith(": ")
    )
    fields: dict[str, str] = {
        "fields[TITLE]": f"Закупка: {title}",
        "fields[SOURCE_ID]": clean_text(dash.get("bitrix_source_id")) or "EMAIL",
        "fields[COMMENTS]": comments,
    }
    if customer:
        fields["fields[COMPANY_TITLE]"] = customer
    responsible_id = clean_text(os.environ.get("TENDER_DASHBOARD_BITRIX_RESPONSIBLE_ID") or dash.get("bitrix_responsible_id"))
    if responsible_id:
        fields["fields[ASSIGNED_BY_ID]"] = responsible_id

    timeout = int(dash.get("bitrix_timeout_seconds") or 20)

    lead_id = clean_text(record.get("bitrix_lead_id"))
    lead_created = False
    if not lead_id:
        try:
            payload = call_bitrix_method(webhook, "crm.lead.add", fields, timeout)
        except Exception as exc:
            record["bitrix_error"] = str(exc)
            return {"status": "error", "message": str(exc)}

        lead_id = clean_text(payload.get("result"))
        if not lead_id:
            error = clean_text(payload.get("error_description") or payload.get("error") or payload)
            record["bitrix_error"] = error
            return {"status": "error", "message": error}
        lead_created = True
        record["bitrix_lead_id"] = lead_id
        record["bitrix_uploaded_at"] = now_iso()
        record.pop("bitrix_error", None)

    task_result = push_bitrix_gip_task(item, record, config, webhook, lead_id, title, comments, timeout)
    if task_result.get("status") in {"sent", "already_sent"}:
        return {
            "status": "sent" if lead_created else "already_sent",
            "leadId": lead_id,
            "taskId": clean_text(task_result.get("taskId")),
        }
    return {
        "status": "partial" if lead_id else clean_text(task_result.get("status")) or "error",
        "leadId": lead_id,
        "message": clean_text(task_result.get("message")),
    }


def push_bitrix_gip_task(
    item: DashboardItem,
    record: dict[str, Any],
    config: dict[str, Any],
    webhook: str,
    lead_id: str,
    title: str,
    comments: str,
    timeout: int,
) -> dict[str, Any]:
    existing_task_id = clean_text(record.get("bitrix_task_id"))
    if existing_task_id:
        return {"status": "already_sent", "taskId": existing_task_id}

    dash = dashboard_config(config)
    gip_user_id = clean_text(
        os.environ.get("TENDER_DASHBOARD_BITRIX_GIP_USER_ID")
        or dash.get("bitrix_gip_user_id")
        or dash.get("bitrix_task_responsible_id")
        or os.environ.get("TENDER_DASHBOARD_BITRIX_RESPONSIBLE_ID")
        or dash.get("bitrix_responsible_id")
    )
    if not gip_user_id:
        record["bitrix_task_error"] = "Не задан dashboard.bitrix_gip_user_id"
        return {"status": "not_configured", "message": record["bitrix_task_error"]}

    description = "\n\n".join(
        part
        for part in (
            f"Лид CRM: #{lead_id}",
            "Задача: оценить закупку и подготовить решение по участию.",
            comments,
        )
        if part
    )
    fields: dict[str, Any] = {
        "fields[TITLE]": f"Оценить закупку: {title}",
        "fields[DESCRIPTION]": description,
        "fields[RESPONSIBLE_ID]": gip_user_id,
        "fields[UF_CRM_TASK][]": [f"L_{lead_id}"],
    }
    created_by = clean_text(os.environ.get("TENDER_DASHBOARD_BITRIX_TASK_CREATED_BY_ID") or dash.get("bitrix_task_created_by_id"))
    if created_by:
        fields["fields[CREATED_BY]"] = created_by
    group_id = clean_text(os.environ.get("TENDER_DASHBOARD_BITRIX_TASK_GROUP_ID") or dash.get("bitrix_task_group_id"))
    if group_id:
        fields["fields[GROUP_ID]"] = group_id
    deadline = bitrix_task_deadline(dash)
    if deadline:
        fields["fields[DEADLINE]"] = deadline

    try:
        payload = call_bitrix_method(webhook, "tasks.task.add", fields, timeout)
    except Exception as exc:
        record["bitrix_task_error"] = str(exc)
        return {"status": "error", "message": str(exc)}

    task_id = extract_bitrix_task_id(payload)
    if task_id:
        record["bitrix_task_id"] = task_id
        record["bitrix_task_created_at"] = now_iso()
        record.pop("bitrix_task_error", None)
        return {"status": "sent", "taskId": task_id}
    error = clean_text(payload.get("error_description") or payload.get("error") or payload)
    record["bitrix_task_error"] = error
    return {"status": "error", "message": error}


def safe_download_file(path_text: str, config: dict[str, Any], config_path: Path) -> Path | None:
    if not path_text:
        return None
    path = Path(path_text)
    try:
        resolved = path.resolve()
    except OSError:
        return None
    if not resolved.exists() or not resolved.is_file():
        return None
    roots = configured_roots(config, config_path)
    allowed_roots = [roots["relevant"], roots["reports"]]
    if bool(dashboard_config(config).get("include_archive_root", False)):
        allowed_roots.append(roots["output"])
    if not any(path_within(resolved, root) for root in allowed_roots):
        return None
    if not any(part in DOWNLOAD_DIR_NAMES for part in resolved.parts):
        return None
    return resolved


def file_is_older_than(path: Path, cutoff: dt.datetime) -> bool:
    try:
        return dt.datetime.fromtimestamp(path.stat().st_mtime) <= cutoff
    except OSError:
        return False


def deadline_datetimes_from_text(text: str) -> list[dt.datetime]:
    result: list[dt.datetime] = []
    value = clean_text(text)
    if not value:
        return result
    patterns = (
        r"(?P<day>\d{1,2})[./-](?P<month>\d{1,2})[./-](?P<year>20\d{2})(?:\s*[T ]\s*(?P<hour>\d{1,2})[:.](?P<minute>\d{2}))?",
        r"(?P<year>20\d{2})[./-](?P<month>\d{1,2})[./-](?P<day>\d{1,2})(?:\s*[T ]\s*(?P<hour>\d{1,2})[:.](?P<minute>\d{2}))?",
    )
    for pattern in patterns:
        for match in re.finditer(pattern, value):
            try:
                hour = int(match.group("hour") or 23)
                minute = int(match.group("minute") or 59)
                second = 59 if not match.group("hour") else 0
                result.append(
                    dt.datetime(
                        int(match.group("year")),
                        int(match.group("month")),
                        int(match.group("day")),
                        hour,
                        minute,
                        second,
                    )
                )
            except ValueError:
                continue
    return result


def unchecked_item_deadline_expired(item: DashboardItem, now: dt.datetime) -> bool:
    deadline = offer_deadline(item.columns)
    dates = deadline_datetimes_from_text(deadline)
    return bool(dates) and max(dates) < now


def prune_empty_download_dirs(start: Path) -> None:
    current = start
    stop: Path | None = None
    for candidate in [start, *start.parents]:
        if candidate.name in DOWNLOAD_DIR_NAMES:
            stop = candidate
            break
    if stop is None:
        return
    while path_within(current, stop) or current == stop:
        try:
            current.rmdir()
        except OSError:
            break
        if current == stop:
            break
        current = current.parent


def cleanup_downloaded_files(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    dash = dashboard_config(config)
    retention_days = int(dash.get("file_retention_days") or 3)
    unprocessed_retention_days = dashboard_unprocessed_retention_days(config)
    now = dt.datetime.now()
    cutoff = now - dt.timedelta(days=retention_days)
    unprocessed_cutoff = now - dt.timedelta(days=unprocessed_retention_days)
    state = load_state(config, config_path)
    items = collect_items(config, config_path)
    items = append_retained_unprocessed_items(items, state, config)
    apply_state(items, state, mutate=True)
    records = state.setdefault("items", {})

    deleted: list[str] = []
    skipped: list[str] = []
    for item in items:
        record = records.setdefault(item.id, {})
        status = state_record_status(record)
        if status in {"relevant", "submitting"}:
            skipped.append(item.id)
            continue
        if status == "unchecked" and not unchecked_item_deadline_expired(item, now):
            skipped.append(f"{item.id}: unchecked_deadline_not_expired")
            continue
        item_cutoff = unprocessed_cutoff if status == "unchecked" else cutoff
        deleted_for_item: list[str] = []
        for path_text in sorted(set(item.downloaded_files)):
            path = safe_download_file(path_text, config, config_path)
            if path is None:
                continue
            if not file_is_older_than(path, item_cutoff):
                continue
            try:
                path.unlink()
                deleted.append(str(path))
                deleted_for_item.append(str(path))
                prune_empty_download_dirs(path.parent)
            except OSError as exc:
                skipped.append(f"{path}: {exc}")
        if deleted_for_item:
            previous = record.get("deleted_files", [])
            if not isinstance(previous, list):
                previous = []
            record["deleted_files"] = sorted(set([*previous, *deleted_for_item]))
            record["files_deleted_at"] = now_iso()

    save_state(config, config_path, state)
    return {
        "retentionDays": retention_days,
        "unprocessedRetentionDays": unprocessed_retention_days,
        "cutoff": cutoff.isoformat(timespec="seconds"),
        "unprocessedCutoff": unprocessed_cutoff.isoformat(timespec="seconds"),
        "deletedCount": len(deleted),
        "deletedFiles": deleted,
        "skipped": skipped,
    }


def read_request_json(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length") or "0")
    if length <= 0:
        return {}
    raw = handler.rfile.read(length)
    try:
        payload = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def send_json(handler: BaseHTTPRequestHandler, payload: Any, status: int = 200) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def send_static(handler: BaseHTTPRequestHandler, path: Path) -> None:
    if not path.exists() or not path.is_file():
        handler.send_error(404)
        return
    data = path.read_bytes()
    content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
    if path.suffix == ".js":
        content_type = "application/javascript; charset=utf-8"
    elif path.suffix in {".html", ".css"}:
        content_type = f"text/{path.suffix.lstrip('.')}; charset=utf-8"
    handler.send_response(200)
    handler.send_header("Content-Type", content_type)
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def dashboard_auth_credentials(config: dict[str, Any]) -> tuple[str, str] | None:
    dash = dashboard_config(config)
    username = clean_text(os.environ.get("TENDER_DASHBOARD_USER") or dash.get("auth_username"))
    password = clean_text(os.environ.get("TENDER_DASHBOARD_PASSWORD") or dash.get("auth_password"))
    if not username and not password:
        return None
    if not username or not password:
        raise ValueError("Для Basic Auth нужно задать и логин, и пароль")
    return username, password


def request_authorized(handler: BaseHTTPRequestHandler, config: dict[str, Any]) -> bool:
    credentials = dashboard_auth_credentials(config)
    if credentials is None:
        return True
    header = handler.headers.get("Authorization", "")
    scheme, _, encoded = header.partition(" ")
    if scheme.lower() != "basic" or not encoded:
        return False
    try:
        decoded = base64.b64decode(encoded).decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        return False
    username, _, password = decoded.partition(":")
    expected_user, expected_password = credentials
    return hmac.compare_digest(username, expected_user) and hmac.compare_digest(password, expected_password)


def send_auth_required(handler: BaseHTTPRequestHandler) -> None:
    handler.send_response(401)
    handler.send_header("WWW-Authenticate", 'Basic realm="Tender dashboard", charset="UTF-8"')
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    payload = json.dumps({"error": "auth_required"}, ensure_ascii=False).encode("utf-8")
    handler.send_header("Content-Length", str(len(payload)))
    handler.end_headers()
    handler.wfile.write(payload)


def make_handler(config: dict[str, Any], config_path: Path) -> type[BaseHTTPRequestHandler]:
    class DashboardHandler(BaseHTTPRequestHandler):
        server_version = "TenderDashboard/1.0"

        def log_message(self, format: str, *args: Any) -> None:  # noqa: A003 - inherited API name.
            print(f"{self.address_string()} - {format % args}")

        def do_GET(self) -> None:  # noqa: N802 - inherited API name.
            if not request_authorized(self, config):
                send_auth_required(self)
                return
            parsed = urllib.parse.urlparse(self.path)
            route = parsed.path
            if route == "/api/items":
                send_json(self, dashboard_payload(config, config_path))
                return
            if route == "/api/integrations":
                send_json(self, integrations_status(config, config_path))
                return
            if route == "/api/download":
                query = urllib.parse.parse_qs(parsed.query)
                try:
                    send_dashboard_download(self, (query.get("path") or [""])[0], config, config_path)
                except PermissionError as exc:
                    send_json(self, {"error": str(exc)}, 403)
                except ValueError as exc:
                    send_json(self, {"error": str(exc)}, 400)
                except FileNotFoundError as exc:
                    send_json(self, {"error": str(exc)}, 404)
                return
            if route == "/api/browse":
                query = urllib.parse.parse_qs(parsed.query)
                try:
                    send_dashboard_browse(self, (query.get("path") or [""])[0], config, config_path)
                except PermissionError as exc:
                    send_json(self, {"error": str(exc)}, 403)
                except ValueError as exc:
                    send_json(self, {"error": str(exc)}, 400)
                except FileNotFoundError as exc:
                    send_json(self, {"error": str(exc)}, 404)
                return
            if route in {"/", "/index.html"}:
                send_static(self, STATIC_ROOT / "index.html")
                return
            safe_path = posixpath.normpath(urllib.parse.unquote(route)).lstrip("/")
            if safe_path.startswith("api/") or safe_path.startswith(".."):
                self.send_error(404)
                return
            send_static(self, STATIC_ROOT / safe_path)

        def do_POST(self) -> None:  # noqa: N802 - inherited API name.
            if not request_authorized(self, config):
                send_auth_required(self)
                return
            parsed = urllib.parse.urlparse(self.path)
            parts = [part for part in parsed.path.strip("/").split("/") if part]
            try:
                if parts == ["api", "cleanup"]:
                    send_json(self, cleanup_downloaded_files(config, config_path))
                    return
                if parts == ["api", "open-path"]:
                    payload = read_request_json(self)
                    send_json(self, open_dashboard_path(clean_text(payload.get("path"))))
                    return
                if parts == ["api", "integrations"]:
                    payload = read_request_json(self)
                    result = save_integrations(payload, config, config_path)
                    result["status"] = integrations_status(config, config_path)
                    send_json(self, result)
                    return
                if len(parts) == 4 and parts[0] == "api" and parts[1] == "items" and parts[3] == "status":
                    payload = read_request_json(self)
                    item_id = urllib.parse.unquote(parts[2])
                    result = update_item_status(item_id, clean_text(payload.get("status")), config, config_path)
                    send_json(self, result)
                    return
                if len(parts) == 4 and parts[0] == "api" and parts[1] == "items" and parts[3] == "bitrix":
                    state = load_state(config, config_path)
                    records = state.setdefault("items", {})
                    item_id = urllib.parse.unquote(parts[2])
                    record = records.setdefault(item_id, {})
                    items = collect_items(config, config_path)
                    apply_state(items, state, mutate=True)
                    item = next((candidate for candidate in items if candidate.id == item_id), None)
                    if item is None:
                        send_json(self, {"error": "not_found"}, 404)
                        return
                    result = push_item_to_bitrix(item, record, config)
                    save_state(config, config_path, state)
                    payload = dashboard_payload(config, config_path)
                    payload["bitrix"] = result
                    send_json(self, payload)
                    return
            except ValueError as exc:
                send_json(self, {"error": str(exc)}, 400)
                return
            except FileNotFoundError as exc:
                send_json(self, {"error": str(exc)}, 404)
                return
            except KeyError:
                send_json(self, {"error": "not_found"}, 404)
                return
            except Exception as exc:
                send_json(self, {"error": str(exc)}, 500)
                return
            self.send_error(404)

    return DashboardHandler


def export_payload(config: dict[str, Any], config_path: Path, output_path: Path) -> dict[str, Any]:
    payload = dashboard_payload(config, config_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"output": str(output_path.resolve()), "items": len(payload.get("items", []))}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Dashboard по входящим закупкам")
    parser.add_argument("command", nargs="?", choices=("serve", "cleanup", "export", "configure-secrets"), default="serve")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--host")
    parser.add_argument("--port", type=int)
    parser.add_argument("--output", default=str(STATIC_ROOT / "dashboard-data.json"))
    parser.add_argument("--env-file", default="")
    args = parser.parse_args(argv)

    config_path = resolve_path(args.config)
    config = load_config(config_path)
    dash = dashboard_config(config)

    if args.command == "cleanup":
        print(json.dumps(cleanup_downloaded_files(config, config_path), ensure_ascii=False, indent=2))
        return 0
    if args.command == "export":
        print(json.dumps(export_payload(config, config_path, resolve_path(args.output)), ensure_ascii=False, indent=2))
        return 0
    if args.command == "configure-secrets":
        print(json.dumps(configure_integrations_interactive(config, config_path, args.env_file), ensure_ascii=False, indent=2))
        return 0

    host = args.host or clean_text(dash.get("host")) or "127.0.0.1"
    port = args.port or int(dash.get("port") or 8765)
    handler = make_handler(config, config_path)
    server = ThreadingHTTPServer((host, port), handler)
    try:
        print(json.dumps({"dashboard_url": f"http://{host}:{port}/", "config": str(config_path.resolve())}, ensure_ascii=False))
    except (AttributeError, OSError):
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
