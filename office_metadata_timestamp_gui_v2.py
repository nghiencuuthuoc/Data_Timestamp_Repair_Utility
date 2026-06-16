# -*- coding: utf-8 -*-
"""
Office Metadata Timestamp GUI v2
Copyright 2026 // NGHIEN CUU THUOC // RnD PHARMA PLUS // WWW.NGHIENCUUTHUOC.COM

Purpose
-------
Paste files or folders, scan internal document metadata, and set Windows file
timestamps to the document's real Content created date.

This v2 fixes two practical issues found in real PharmSolu data:
1. Some corrupted filesystem timestamps can make datetime.fromtimestamp() raise
   [Errno 22] Invalid argument on Windows. v2 never lets bad filesystem time stop
   metadata scanning.
2. Legacy .doc/.xls/.ppt metadata is normalized more safely, and the selected
   target date is shown with a detailed decision/rejection note.

Supported metadata sources
--------------------------
- Modern Office OOXML: .docx .xlsx .pptx and macro/template variants
- Legacy Office OLE: .doc .xls .ppt, when olefile is installed
- PDF metadata: /CreationDate and /ModDate
- Optional filename/path date fallback: for names like 17.05.2016

Recommended dependency
----------------------
pip install olefile

Safety
------
- Scan first.
- Review the table and CSV log.
- Apply timestamps only after you confirm.
"""

from __future__ import annotations

import csv
import ctypes
import datetime as dt
import os
import platform
import queue
import re
import sys
import threading
import traceback
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_NAME = "Office_Metadata_Timestamp_GUI_v2"
APP_TITLE = "Office Metadata Timestamp GUI v2"
DEFAULT_LOG_DIR = "office_timestamp_logs"

SUPPORTED_EXTS = {
    ".docx", ".docm", ".dotx", ".dotm",
    ".xlsx", ".xlsm", ".xltx", ".xltm",
    ".pptx", ".pptm", ".potx", ".potm",
    ".doc", ".xls", ".ppt",
    ".pdf",
}

OOXML_EXTS = {
    ".docx", ".docm", ".dotx", ".dotm",
    ".xlsx", ".xlsm", ".xltx", ".xltm",
    ".pptx", ".pptm", ".potx", ".potm",
}
OLE_EXTS = {".doc", ".xls", ".ppt"}
PDF_EXTS = {".pdf"}


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------


def set_console_title(title: str) -> None:
    """Set a clear Windows console title. Safe no-op on non-Windows systems."""
    if os.name == "nt":
        try:
            ctypes.windll.kernel32.SetConsoleTitleW(title)
        except Exception:
            pass


def is_windows() -> bool:
    return platform.system().lower() == "windows"


def now_local() -> dt.datetime:
    return dt.datetime.now().astimezone()


def local_tz() -> dt.tzinfo:
    return now_local().tzinfo or dt.datetime.now().astimezone().tzinfo


def normalize_datetime(value: object) -> Optional[dt.datetime]:
    """
    Normalize Office/PDF/Python datetime values into a local timezone-aware datetime.

    Rules:
    - Naive datetimes are treated as local time. This matches how Office Properties
      are displayed in Windows Explorer for most user workflows.
    - Aware datetimes are converted to local time.
    - Dates outside Python's normal safe range are rejected.
    """
    if value is None:
        return None

    if isinstance(value, dt.datetime):
        try:
            if value.tzinfo is None:
                return value.replace(tzinfo=local_tz())
            return value.astimezone()
        except Exception:
            return None

    return None


def safe_fromtimestamp(timestamp: float) -> Tuple[Optional[dt.datetime], str]:
    """
    Convert filesystem timestamp safely.

    On Windows, badly corrupted file times can make datetime.fromtimestamp raise:
      OSError: [Errno 22] Invalid argument

    This function returns (None, error_message) instead of crashing the scan.
    """
    try:
        return dt.datetime.fromtimestamp(timestamp).astimezone(), ""
    except Exception as exc:
        return None, f"invalid_timestamp:{exc}"


def dt_to_ts(value: dt.datetime) -> float:
    normalized = normalize_datetime(value)
    if normalized is None:
        raise ValueError("Invalid datetime value.")
    return normalized.timestamp()


def safe_iso(value: Optional[dt.datetime]) -> str:
    return value.isoformat(timespec="seconds") if value else ""


def safe_display(value: Optional[dt.datetime]) -> str:
    if not value:
        return ""
    return value.strftime("%Y-%m-%d %H:%M:%S")


def safe_year(value: Optional[dt.datetime]) -> str:
    return str(value.year) if value else ""


# ---------------------------------------------------------------------------
# Metadata date parsing
# ---------------------------------------------------------------------------


def parse_iso_datetime(text: str) -> Optional[dt.datetime]:
    """Parse ISO-like Office dates such as 2016-05-11T12:43:00Z."""
    if not text:
        return None

    raw = str(text).strip()
    if not raw:
        return None

    try:
        if raw.endswith("Z"):
            raw = raw[:-1] + "+00:00"
        parsed = dt.datetime.fromisoformat(raw)
        return normalize_datetime(parsed)
    except Exception:
        return None


PDF_DATE_RE = re.compile(
    r"D:(?P<year>\d{4})(?P<month>\d{2})?(?P<day>\d{2})?"
    r"(?P<hour>\d{2})?(?P<minute>\d{2})?(?P<second>\d{2})?"
    r"(?P<tz>Z|[+-]\d{2}'?\d{2}'?)?"
)


def parse_pdf_datetime(text: str) -> Optional[dt.datetime]:
    if not text:
        return None

    match = PDF_DATE_RE.search(str(text))
    if not match:
        return None

    data = match.groupdict()
    try:
        year = int(data["year"])
        month = int(data["month"] or 1)
        day = int(data["day"] or 1)
        hour = int(data["hour"] or 0)
        minute = int(data["minute"] or 0)
        second = int(data["second"] or 0)
        parsed = dt.datetime(year, month, day, hour, minute, second)

        tz_text = data.get("tz")
        if tz_text:
            if tz_text == "Z":
                parsed = parsed.replace(tzinfo=dt.timezone.utc)
            else:
                clean = tz_text.replace("'", "")
                sign = 1 if clean[0] == "+" else -1
                tz_hour = int(clean[1:3])
                tz_min = int(clean[3:5])
                parsed = parsed.replace(
                    tzinfo=dt.timezone(sign * dt.timedelta(hours=tz_hour, minutes=tz_min))
                )

        return normalize_datetime(parsed)
    except Exception:
        return None


def parse_filename_or_path_date(path: Path) -> Tuple[Optional[dt.datetime], str]:
    """
    Extract useful dates from filename/path when document metadata is missing.

    Supported examples:
      GTCN ... 17.05.2016 .doc -> 2016-05-17
      report_2016-05-17.docx   -> 2016-05-17
      20160517_report.docx     -> 2016-05-17

    Time is set to 00:00:00 because filenames normally do not contain time.
    """
    text = str(path)

    # dd.mm.yyyy, dd-mm-yyyy, dd_mm_yyyy, dd mm yyyy
    for match in re.finditer(r"(?<!\d)(\d{1,2})[.\-_\s](\d{1,2})[.\-_\s]((?:19|20)\d{2})(?!\d)", text):
        day, month, year = map(int, match.groups())
        try:
            return dt.datetime(year, month, day, 0, 0, 0, tzinfo=local_tz()), "filename_date_dd_mm_yyyy"
        except Exception:
            pass

    # yyyy-mm-dd, yyyy.mm.dd, yyyy_mm_dd, yyyy mm dd
    for match in re.finditer(r"(?<!\d)((?:19|20)\d{2})[.\-_\s](\d{1,2})[.\-_\s](\d{1,2})(?!\d)", text):
        year, month, day = map(int, match.groups())
        try:
            return dt.datetime(year, month, day, 0, 0, 0, tzinfo=local_tz()), "filename_date_yyyy_mm_dd"
        except Exception:
            pass

    # yyyymmdd
    for match in re.finditer(r"(?<!\d)((?:19|20)\d{2})(\d{2})(\d{2})(?!\d)", text):
        year, month, day = map(int, match.groups())
        try:
            return dt.datetime(year, month, day, 0, 0, 0, tzinfo=local_tz()), "filename_date_yyyymmdd"
        except Exception:
            pass

    return None, ""


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class FileSystemTimes:
    created: Optional[dt.datetime] = None
    modified: Optional[dt.datetime] = None
    accessed: Optional[dt.datetime] = None
    error: str = ""


@dataclass
class MetadataResult:
    created: Optional[dt.datetime] = None
    modified: Optional[dt.datetime] = None
    author: str = ""
    last_saved_by: str = ""
    source: str = ""
    status: str = "not_scanned"
    message: str = ""


@dataclass
class ScanResult:
    path: Path
    exists: bool = False
    fs_times: FileSystemTimes = field(default_factory=FileSystemTimes)
    metadata: MetadataResult = field(default_factory=MetadataResult)
    filename_date: Optional[dt.datetime] = None
    filename_date_source: str = ""
    target_time: Optional[dt.datetime] = None
    target_source: str = ""
    status: str = "pending"
    message: str = ""
    applied: bool = False
    creation_set: bool = False
    modified_set: bool = False
    accessed_set: bool = False


# ---------------------------------------------------------------------------
# Filesystem time handling
# ---------------------------------------------------------------------------


def get_fs_times_safe(path: Path) -> FileSystemTimes:
    """Return filesystem timestamps without allowing bad times to stop scanning."""
    result = FileSystemTimes()

    try:
        st = path.stat()
    except Exception as exc:
        result.error = f"stat_error:{exc}"
        return result

    errors: List[str] = []

    if is_windows():
        result.created, err = safe_fromtimestamp(st.st_ctime)
        if err:
            errors.append(f"creation={err}")

    result.modified, err = safe_fromtimestamp(st.st_mtime)
    if err:
        errors.append(f"modified={err}")

    result.accessed, err = safe_fromtimestamp(st.st_atime)
    if err:
        errors.append(f"accessed={err}")

    result.error = "; ".join(errors)
    return result


# ---------------------------------------------------------------------------
# Metadata readers
# ---------------------------------------------------------------------------


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def read_ooxml_metadata(path: Path) -> MetadataResult:
    result = MetadataResult(source="ooxml_core_xml")

    try:
        with zipfile.ZipFile(path, "r") as z:
            try:
                raw = z.read("docProps/core.xml")
            except KeyError:
                result.status = "missing_metadata"
                result.message = "docProps/core.xml was not found."
                return result

        root = ET.fromstring(raw)
        fields: Dict[str, str] = {}

        for item in root.iter():
            name = local_name(item.tag)
            text = (item.text or "").strip()
            if text:
                fields[name] = text

        result.created = parse_iso_datetime(fields.get("created", ""))
        result.modified = parse_iso_datetime(fields.get("modified", ""))
        result.author = fields.get("creator", "")
        result.last_saved_by = fields.get("lastModifiedBy", "")

        if result.created or result.modified:
            result.status = "ok"
            result.message = "OOXML metadata found."
        else:
            result.status = "missing_date"
            result.message = "OOXML core.xml has no created/modified date."

        return result

    except zipfile.BadZipFile:
        result.status = "bad_zip"
        result.message = "Not a valid OOXML zip package. It may be a legacy Office file renamed to .docx/.xlsx/.pptx."
        return result
    except Exception as exc:
        result.status = "error"
        result.message = f"OOXML read error: {exc}"
        return result


def read_ole_metadata(path: Path, source_label: str = "ole_summary_information") -> MetadataResult:
    result = MetadataResult(source=source_label)

    try:
        import olefile  # type: ignore
    except Exception:
        result.status = "missing_dependency"
        result.message = "Install dependency first: pip install olefile"
        return result

    try:
        if not olefile.isOleFile(str(path)):
            result.status = "bad_file"
            result.message = "Not a valid OLE compound document."
            return result

        with olefile.OleFileIO(str(path)) as ole:
            meta = ole.get_metadata()

            result.created = normalize_datetime(getattr(meta, "create_time", None))
            result.modified = normalize_datetime(getattr(meta, "last_saved_time", None))
            result.author = str(getattr(meta, "author", "") or "")
            result.last_saved_by = str(getattr(meta, "last_saved_by", "") or "")

        if result.created or result.modified:
            result.status = "ok"
            result.message = "OLE metadata found."
        else:
            result.status = "missing_date"
            result.message = "OLE metadata has no create_time/last_saved_time."

        return result

    except Exception as exc:
        result.status = "error"
        result.message = f"OLE read error: {exc}"
        return result


def read_pdf_metadata(path: Path) -> MetadataResult:
    result = MetadataResult(source="pdf_raw_metadata")

    try:
        with path.open("rb") as f:
            head = f.read(1024 * 1024)
            try:
                f.seek(max(0, path.stat().st_size - 1024 * 1024))
                tail = f.read(1024 * 1024)
            except Exception:
                tail = b""

        text = (head + b"\n" + tail).decode("latin-1", errors="ignore")

        created_match = re.search(r"/CreationDate\s*\((.*?)\)", text)
        modified_match = re.search(r"/ModDate\s*\((.*?)\)", text)

        if created_match:
            result.created = parse_pdf_datetime(created_match.group(1))
        if modified_match:
            result.modified = parse_pdf_datetime(modified_match.group(1))

        if result.created or result.modified:
            result.status = "ok"
            result.message = "PDF metadata found."
        else:
            result.status = "missing_date"
            result.message = "No PDF CreationDate/ModDate was found."

        return result

    except Exception as exc:
        result.status = "error"
        result.message = f"PDF read error: {exc}"
        return result


def merge_metadata(primary: MetadataResult, fallback: MetadataResult) -> MetadataResult:
    """Use fallback only when primary did not provide usable dates."""
    if primary.created or primary.modified:
        return primary
    if fallback.created or fallback.modified:
        return fallback
    if primary.status in {"error", "bad_zip", "missing_metadata", "missing_date"}:
        # Keep both messages visible.
        return MetadataResult(
            source=f"{primary.source}; {fallback.source}",
            status=fallback.status,
            message=f"{primary.message} | {fallback.message}",
            created=fallback.created,
            modified=fallback.modified,
            author=fallback.author,
            last_saved_by=fallback.last_saved_by,
        )
    return primary


def read_document_metadata(path: Path) -> MetadataResult:
    """Auto-detect and read metadata. Allows renamed legacy Office files."""
    ext = path.suffix.lower()

    if ext in OOXML_EXTS:
        ooxml = read_ooxml_metadata(path)
        if ooxml.created or ooxml.modified:
            return ooxml
        # Some old .doc files are renamed .docx. Try OLE as fallback.
        ole = read_ole_metadata(path, source_label="ole_fallback_for_renamed_ooxml")
        return merge_metadata(ooxml, ole)

    if ext in OLE_EXTS:
        return read_ole_metadata(path)

    if ext in PDF_EXTS:
        return read_pdf_metadata(path)

    # Unknown extension: still try OOXML then OLE then PDF-like raw metadata.
    ooxml = read_ooxml_metadata(path)
    if ooxml.created or ooxml.modified:
        return ooxml

    ole = read_ole_metadata(path, source_label="ole_auto_detect")
    if ole.created or ole.modified:
        return ole

    pdf = read_pdf_metadata(path)
    if pdf.created or pdf.modified:
        return pdf

    return MetadataResult(
        source="unsupported_or_no_metadata",
        status="unsupported_or_no_metadata",
        message=f"No supported metadata was found for extension: {ext}",
    )


# ---------------------------------------------------------------------------
# Target timestamp selection
# ---------------------------------------------------------------------------


def validate_candidate_date(
    value: Optional[dt.datetime],
    label: str,
    min_year: int,
    allow_future_days: int,
    reject_future_dates: bool,
) -> Tuple[bool, str]:
    if value is None:
        return False, f"{label}:empty"

    normalized = normalize_datetime(value)
    if normalized is None:
        return False, f"{label}:not_datetime"

    if normalized.year < min_year:
        return False, f"{label}:year_before_{min_year}"

    if reject_future_dates:
        latest = now_local() + dt.timedelta(days=allow_future_days)
        if normalized > latest:
            return False, f"{label}:future_rejected_{safe_display(normalized)}"

    return True, f"{label}:ok"


def choose_target_time(
    metadata: MetadataResult,
    filename_date: Optional[dt.datetime],
    filename_date_source: str,
    prefer_created: bool,
    allow_modified_fallback: bool,
    allow_filename_fallback: bool,
    min_year: int,
    allow_future_days: int,
    reject_future_dates: bool,
) -> Tuple[Optional[dt.datetime], str, str]:
    candidates: List[Tuple[str, Optional[dt.datetime]]] = []

    if prefer_created:
        candidates.append(("metadata_created_content_created", metadata.created))
        if allow_modified_fallback:
            candidates.append(("metadata_modified_date_last_saved", metadata.modified))
    else:
        candidates.append(("metadata_modified_date_last_saved", metadata.modified))
        if allow_modified_fallback:
            candidates.append(("metadata_created_content_created", metadata.created))

    if allow_filename_fallback:
        candidates.append((filename_date_source or "filename_or_path_date", filename_date))

    notes: List[str] = []

    for label, value in candidates:
        valid, note = validate_candidate_date(
            value,
            label=label,
            min_year=min_year,
            allow_future_days=allow_future_days,
            reject_future_dates=reject_future_dates,
        )
        notes.append(note)
        if valid:
            return normalize_datetime(value), label, "; ".join(notes)

    return None, "", "; ".join(notes)


# ---------------------------------------------------------------------------
# Path parsing and file iteration
# ---------------------------------------------------------------------------


def parse_pasted_paths(text: str) -> List[Path]:
    items: List[Path] = []

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        quoted = re.findall(r'"([^"]+)"', line)
        if quoted:
            for value in quoted:
                cleaned = value.strip()
                if cleaned:
                    items.append(Path(cleaned))
            continue

        line = line.strip().strip("'").strip('"')
        if line:
            items.append(Path(line))

    seen = set()
    unique: List[Path] = []
    for item in items:
        key = str(item).lower()
        if key not in seen:
            unique.append(item)
            seen.add(key)
    return unique


def iter_input_files(paths: Sequence[Path], recursive: bool, supported_only: bool) -> Iterable[Path]:
    for item in paths:
        try:
            if item.is_file():
                if not supported_only or item.suffix.lower() in SUPPORTED_EXTS:
                    yield item
                continue

            if item.is_dir():
                if recursive:
                    for root, _dirnames, filenames in os.walk(item):
                        root_path = Path(root)
                        for filename in filenames:
                            file_path = root_path / filename
                            if not supported_only or file_path.suffix.lower() in SUPPORTED_EXTS:
                                yield file_path
                else:
                    for child in item.iterdir():
                        if child.is_file() and (not supported_only or child.suffix.lower() in SUPPORTED_EXTS):
                            yield child
        except Exception:
            continue


# ---------------------------------------------------------------------------
# Windows timestamp setter
# ---------------------------------------------------------------------------


def win32_long_path(path: Path) -> str:
    text = str(path.resolve())
    if text.startswith("\\\\?\\"):
        return text
    if text.startswith("\\\\"):
        return "\\\\?\\UNC\\" + text.lstrip("\\")
    return "\\\\?\\" + text


class FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", ctypes.c_uint32),
        ("dwHighDateTime", ctypes.c_uint32),
    ]


def datetime_to_filetime(value: dt.datetime) -> FILETIME:
    normalized = normalize_datetime(value)
    if normalized is None:
        raise ValueError("Cannot convert invalid datetime to FILETIME.")
    timestamp = normalized.timestamp()
    filetime_int = int((timestamp + 11644473600) * 10000000)
    return FILETIME(filetime_int & 0xFFFFFFFF, filetime_int >> 32)


def set_times_windows(
    path: Path,
    new_time: dt.datetime,
    set_created: bool,
    set_modified: bool,
    set_accessed: bool,
) -> Dict[str, bool]:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
    ]
    create_file.restype = ctypes.c_void_p

    set_file_time = kernel32.SetFileTime
    set_file_time.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
    ]
    set_file_time.restype = ctypes.c_int

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [ctypes.c_void_p]
    close_handle.restype = ctypes.c_int

    FILE_WRITE_ATTRIBUTES = 0x0100
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    flags = FILE_FLAG_BACKUP_SEMANTICS if path.is_dir() else 0

    handle = create_file(
        win32_long_path(path),
        FILE_WRITE_ATTRIBUTES,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        flags,
        None,
    )

    if handle == INVALID_HANDLE_VALUE:
        raise ctypes.WinError(ctypes.get_last_error())

    try:
        ft = datetime_to_filetime(new_time)
        created_ptr = ctypes.byref(ft) if set_created else None
        accessed_ptr = ctypes.byref(ft) if set_accessed else None
        modified_ptr = ctypes.byref(ft) if set_modified else None

        ok = set_file_time(handle, created_ptr, accessed_ptr, modified_ptr)
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())

        return {
            "creation_set": bool(set_created),
            "modified_set": bool(set_modified),
            "accessed_set": bool(set_accessed),
        }
    finally:
        close_handle(handle)


def set_times_cross_platform(
    path: Path,
    new_time: dt.datetime,
    set_created: bool,
    set_modified: bool,
    set_accessed: bool,
) -> Dict[str, bool]:
    if is_windows():
        return set_times_windows(path, new_time, set_created, set_modified, set_accessed)

    current = get_fs_times_safe(path)
    access_time = dt_to_ts(new_time if set_accessed else (current.accessed or now_local()))
    modified_time = dt_to_ts(new_time if set_modified else (current.modified or now_local()))
    os.utime(str(path), (access_time, modified_time), follow_symlinks=False)

    return {
        "creation_set": False,
        "modified_set": bool(set_modified),
        "accessed_set": bool(set_accessed),
    }


# ---------------------------------------------------------------------------
# CSV logging
# ---------------------------------------------------------------------------


def make_log_path(action: str) -> Path:
    log_dir = Path(DEFAULT_LOG_DIR)
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return log_dir / f"{APP_NAME}_{action}_{stamp}.csv"


def write_scan_log(log_path: Path, results: Sequence[ScanResult]) -> None:
    with log_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "path",
                "exists",
                "status",
                "message",
                "metadata_source",
                "target_source",
                "fs_creation_time",
                "fs_modified_time",
                "fs_accessed_time",
                "fs_time_error",
                "metadata_created",
                "metadata_modified",
                "filename_date",
                "author",
                "last_saved_by",
                "target_time",
                "applied",
                "creation_set",
                "modified_set",
                "accessed_set",
            ]
        )

        for item in results:
            writer.writerow(
                [
                    str(item.path),
                    item.exists,
                    item.status,
                    item.message,
                    item.metadata.source,
                    item.target_source,
                    safe_iso(item.fs_times.created),
                    safe_iso(item.fs_times.modified),
                    safe_iso(item.fs_times.accessed),
                    item.fs_times.error,
                    safe_iso(item.metadata.created),
                    safe_iso(item.metadata.modified),
                    safe_iso(item.filename_date),
                    item.metadata.author,
                    item.metadata.last_saved_by,
                    safe_iso(item.target_time),
                    item.applied,
                    item.creation_set,
                    item.modified_set,
                    item.accessed_set,
                ]
            )


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class OfficeTimestampGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        set_console_title(APP_NAME)

        self.title(APP_TITLE)
        self.geometry("1380x820")
        self.minsize(1050, 680)

        self.results: List[ScanResult] = []
        self.worker_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None

        self._build_vars()
        self._build_ui()
        self._poll_queue()

    def _build_vars(self) -> None:
        self.recursive_var = tk.BooleanVar(value=True)
        self.supported_only_var = tk.BooleanVar(value=True)
        self.prefer_created_var = tk.BooleanVar(value=True)
        self.modified_fallback_var = tk.BooleanVar(value=False)
        self.filename_fallback_var = tk.BooleanVar(value=True)
        self.reject_future_var = tk.BooleanVar(value=True)
        self.set_created_var = tk.BooleanVar(value=True)
        self.set_modified_var = tk.BooleanVar(value=True)
        self.set_accessed_var = tk.BooleanVar(value=True)

        self.min_year_var = tk.StringVar(value="1900")
        self.future_days_var = tk.StringVar(value="1")

        self.status_var = tk.StringVar(value="Ready.")
        self.summary_var = tk.StringVar(value="No scan result yet.")

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        input_frame = ttk.LabelFrame(root, text="Paste files or folders")
        input_frame.pack(fill="x", padx=4, pady=4)

        self.path_text = tk.Text(input_frame, height=7, wrap="none")
        self.path_text.pack(side="left", fill="both", expand=True, padx=(6, 2), pady=6)

        scroll_y = ttk.Scrollbar(input_frame, orient="vertical", command=self.path_text.yview)
        scroll_y.pack(side="left", fill="y", pady=6)
        self.path_text.configure(yscrollcommand=scroll_y.set)

        button_box = ttk.Frame(input_frame)
        button_box.pack(side="left", fill="y", padx=6, pady=6)
        ttk.Button(button_box, text="Paste clipboard", command=self.paste_clipboard).pack(fill="x", pady=2)
        ttk.Button(button_box, text="Add files", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(button_box, text="Add folder", command=self.add_folder).pack(fill="x", pady=2)
        ttk.Button(button_box, text="Clear", command=self.clear_paths).pack(fill="x", pady=2)

        options_frame = ttk.LabelFrame(root, text="Options")
        options_frame.pack(fill="x", padx=4, pady=4)

        left_opts = ttk.Frame(options_frame)
        left_opts.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        mid_opts = ttk.Frame(options_frame)
        mid_opts.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        right_opts = ttk.Frame(options_frame)
        right_opts.pack(side="left", fill="x", expand=True, padx=6, pady=6)

        ttk.Checkbutton(left_opts, text="Scan folders recursively", variable=self.recursive_var).pack(anchor="w")
        ttk.Checkbutton(left_opts, text="Only supported extensions", variable=self.supported_only_var).pack(anchor="w")
        ttk.Checkbutton(left_opts, text="Prefer Content created date", variable=self.prefer_created_var).pack(anchor="w")
        ttk.Checkbutton(left_opts, text="Use filename/path date fallback", variable=self.filename_fallback_var).pack(anchor="w")

        ttk.Checkbutton(mid_opts, text="Allow Date last saved fallback", variable=self.modified_fallback_var).pack(anchor="w")
        ttk.Checkbutton(mid_opts, text="Reject future metadata dates", variable=self.reject_future_var).pack(anchor="w")

        year_row = ttk.Frame(mid_opts)
        year_row.pack(anchor="w", fill="x", pady=2)
        ttk.Label(year_row, text="Minimum valid year:").pack(side="left")
        ttk.Entry(year_row, textvariable=self.min_year_var, width=8).pack(side="left", padx=4)

        future_row = ttk.Frame(mid_opts)
        future_row.pack(anchor="w", fill="x", pady=2)
        ttk.Label(future_row, text="Future allowance days:").pack(side="left")
        ttk.Entry(future_row, textvariable=self.future_days_var, width=8).pack(side="left", padx=4)

        ttk.Label(right_opts, text="Set these file timestamps:").pack(anchor="w")
        ttk.Checkbutton(right_opts, text="CreationTime", variable=self.set_created_var).pack(anchor="w")
        ttk.Checkbutton(right_opts, text="ModifiedTime / LastWriteTime", variable=self.set_modified_var).pack(anchor="w")
        ttk.Checkbutton(right_opts, text="AccessedTime", variable=self.set_accessed_var).pack(anchor="w")

        action_frame = ttk.Frame(root)
        action_frame.pack(fill="x", padx=4, pady=4)
        self.scan_button = ttk.Button(action_frame, text="Scan metadata", command=self.start_scan)
        self.scan_button.pack(side="left", padx=4)
        self.apply_button = ttk.Button(action_frame, text="Apply timestamps", command=self.start_apply, state="disabled")
        self.apply_button.pack(side="left", padx=4)
        ttk.Button(action_frame, text="Save CSV log", command=self.save_csv_log).pack(side="left", padx=4)
        ttk.Button(action_frame, text="Open log folder", command=self.open_log_folder).pack(side="left", padx=4)
        ttk.Label(action_frame, textvariable=self.status_var).pack(side="left", padx=12)

        result_frame = ttk.LabelFrame(root, text="Results")
        result_frame.pack(fill="both", expand=True, padx=4, pady=4)

        columns = (
            "status", "target_time", "target_source", "metadata_created", "metadata_modified",
            "filename_date", "source", "fs_created", "fs_modified", "fs_error", "message", "path",
        )
        headings = {
            "status": "Status",
            "target_time": "New target time",
            "target_source": "Target source",
            "metadata_created": "Metadata created",
            "metadata_modified": "Metadata modified",
            "filename_date": "Filename date",
            "source": "Metadata source",
            "fs_created": "FS created",
            "fs_modified": "FS modified",
            "fs_error": "FS time error",
            "message": "Decision / Message",
            "path": "Path",
        }
        widths = {
            "status": 110,
            "target_time": 150,
            "target_source": 220,
            "metadata_created": 150,
            "metadata_modified": 150,
            "filename_date": 140,
            "source": 180,
            "fs_created": 150,
            "fs_modified": 150,
            "fs_error": 220,
            "message": 420,
            "path": 620,
        }

        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w", stretch=True)

        tree_y = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        tree_x = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_y.grid(row=0, column=1, sticky="ns")
        tree_x.grid(row=1, column=0, sticky="ew")
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        summary_frame = ttk.Frame(root)
        summary_frame.pack(fill="x", padx=4, pady=4)
        ttk.Label(summary_frame, textvariable=self.summary_var).pack(side="left", anchor="w")
        help_text = (
            "v2 fix: bad filesystem times no longer stop scanning. "
            "For files with Date last saved = 2062, keep Date last saved fallback OFF."
        )
        ttk.Label(summary_frame, text=help_text, foreground="#444444").pack(side="right", anchor="e")

    def paste_clipboard(self) -> None:
        try:
            text = self.clipboard_get()
            if text:
                self.path_text.insert("end", text)
                if not text.endswith("\n"):
                    self.path_text.insert("end", "\n")
        except Exception as exc:
            messagebox.showerror("Clipboard error", str(exc))

    def add_files(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="Select files",
            filetypes=[
                ("Supported documents", "*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.doc *.xls *.ppt *.pdf"),
                ("All files", "*.*"),
            ],
        )
        for filename in filenames:
            self.path_text.insert("end", f'"{filename}"\n')

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            self.path_text.insert("end", f'"{folder}"\n')

    def clear_paths(self) -> None:
        self.path_text.delete("1.0", "end")

    def get_int_option(self, var: tk.StringVar, label: str, default: int) -> int:
        text = var.get().strip()
        if not text:
            return default
        try:
            return int(text)
        except Exception:
            raise ValueError(f"{label} must be a valid number.")

    def get_input_files(self) -> List[Path]:
        paths = parse_pasted_paths(self.path_text.get("1.0", "end"))
        if not paths:
            raise ValueError("Please paste at least one file or folder path.")

        files = list(iter_input_files(paths, self.recursive_var.get(), self.supported_only_var.get()))
        unique: List[Path] = []
        seen = set()
        for file_path in files:
            key = str(file_path).lower()
            if key not in seen:
                unique.append(file_path)
                seen.add(key)

        if not unique:
            raise ValueError("No matching files were found. Check paths and extension filter.")
        return unique

    def set_busy(self, busy: bool) -> None:
        self.scan_button.configure(state="disabled" if busy else "normal")
        self.apply_button.configure(state="disabled" if busy or not self.results else "normal")

    def start_scan(self) -> None:
        try:
            files = self.get_input_files()
            min_year = self.get_int_option(self.min_year_var, "Minimum valid year", 1900)
            future_days = self.get_int_option(self.future_days_var, "Future allowance days", 1)
        except Exception as exc:
            messagebox.showerror("Input error", str(exc))
            return

        self.tree.delete(*self.tree.get_children())
        self.results = []
        self.status_var.set("Scanning metadata...")
        self.summary_var.set(f"Scanning {len(files)} files...")
        self.set_busy(True)

        options = {
            "prefer_created": self.prefer_created_var.get(),
            "allow_modified_fallback": self.modified_fallback_var.get(),
            "allow_filename_fallback": self.filename_fallback_var.get(),
            "reject_future_dates": self.reject_future_var.get(),
            "min_year": min_year,
            "allow_future_days": future_days,
        }
        self.worker_thread = threading.Thread(target=self.scan_worker, args=(files, options), daemon=True)
        self.worker_thread.start()

    def scan_worker(self, files: Sequence[Path], options: Dict[str, object]) -> None:
        results: List[ScanResult] = []

        for index, file_path in enumerate(files, start=1):
            item = ScanResult(path=file_path, exists=file_path.exists())

            try:
                # v2: read metadata and filesystem times independently.
                item.metadata = read_document_metadata(file_path)
                item.fs_times = get_fs_times_safe(file_path)
                item.filename_date, item.filename_date_source = parse_filename_or_path_date(file_path)

                target_time, target_source, decision_note = choose_target_time(
                    item.metadata,
                    filename_date=item.filename_date,
                    filename_date_source=item.filename_date_source,
                    prefer_created=bool(options["prefer_created"]),
                    allow_modified_fallback=bool(options["allow_modified_fallback"]),
                    allow_filename_fallback=bool(options["allow_filename_fallback"]),
                    min_year=int(options["min_year"]),
                    allow_future_days=int(options["allow_future_days"]),
                    reject_future_dates=bool(options["reject_future_dates"]),
                )

                item.target_time = target_time
                item.target_source = target_source

                fs_note = f" | FS time issue: {item.fs_times.error}" if item.fs_times.error else ""
                if target_time:
                    item.status = "ready"
                    item.message = f"{item.metadata.message} | {decision_note}{fs_note}"
                else:
                    item.status = "no_valid_date"
                    item.message = f"{item.metadata.message} | {decision_note}{fs_note}"

            except Exception as exc:
                item.status = "error"
                item.message = f"Scan error: {exc}"

            results.append(item)
            self.worker_queue.put(("row", item))
            self.worker_queue.put(("status", f"Scanned {index}/{len(files)}"))

        self.worker_queue.put(("scan_done", results))

    def start_apply(self) -> None:
        ready = [item for item in self.results if item.target_time and item.status in {"ready", "applied"}]
        if not ready:
            messagebox.showinfo("No ready files", "No files have a valid target timestamp.")
            return

        if not (self.set_created_var.get() or self.set_modified_var.get() or self.set_accessed_var.get()):
            messagebox.showerror("Option error", "Select at least one timestamp type to set.")
            return

        answer = messagebox.askyesno(
            "Confirm timestamp update",
            f"Apply timestamps to {len(ready)} files?\n\nA CSV log will be saved automatically.",
        )
        if not answer:
            return

        self.status_var.set("Applying timestamps...")
        self.set_busy(True)
        options = {
            "set_created": self.set_created_var.get(),
            "set_modified": self.set_modified_var.get(),
            "set_accessed": self.set_accessed_var.get(),
        }
        self.worker_thread = threading.Thread(target=self.apply_worker, args=(ready, options), daemon=True)
        self.worker_thread.start()

    def apply_worker(self, items: Sequence[ScanResult], options: Dict[str, object]) -> None:
        updated_by_path: Dict[str, ScanResult] = {}

        for index, item in enumerate(items, start=1):
            try:
                if item.target_time is None:
                    raise ValueError("Missing target_time.")

                result = set_times_cross_platform(
                    item.path,
                    item.target_time,
                    set_created=bool(options["set_created"]),
                    set_modified=bool(options["set_modified"]),
                    set_accessed=bool(options["set_accessed"]),
                )
                item.applied = True
                item.creation_set = bool(result.get("creation_set"))
                item.modified_set = bool(result.get("modified_set"))
                item.accessed_set = bool(result.get("accessed_set"))
                item.status = "applied"
                item.message = "Timestamp applied."
            except Exception as exc:
                item.status = "apply_error"
                item.message = f"Apply error: {exc}"

            updated_by_path[str(item.path)] = item
            self.worker_queue.put(("status", f"Applied {index}/{len(items)}"))

        self.worker_queue.put(("apply_done", updated_by_path))

    def _poll_queue(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()

                if event == "row":
                    self.insert_result_row(payload)

                elif event == "status":
                    self.status_var.set(str(payload))

                elif event == "scan_done":
                    self.results = list(payload)
                    self.refresh_tree()
                    self.set_busy(False)

                    ready_count = sum(1 for item in self.results if item.target_time and item.status == "ready")
                    not_ready_count = len(self.results) - ready_count
                    self.summary_var.set(
                        f"Scan complete. Total: {len(self.results)} | Ready: {ready_count} | Not ready/errors: {not_ready_count}"
                    )
                    self.status_var.set("Scan complete.")
                    self.apply_button.configure(state="normal" if ready_count else "disabled")

                    log_path = make_log_path("scan")
                    write_scan_log(log_path, self.results)

                elif event == "apply_done":
                    updated_by_path = payload
                    for i, item in enumerate(self.results):
                        new_item = updated_by_path.get(str(item.path))
                        if new_item:
                            self.results[i] = new_item

                    self.refresh_tree()
                    self.set_busy(False)

                    applied_count = sum(1 for item in self.results if item.applied)
                    error_count = sum(1 for item in self.results if item.status == "apply_error")
                    self.summary_var.set(f"Apply complete. Applied: {applied_count} | Apply errors: {error_count}")
                    self.status_var.set("Apply complete.")

                    log_path = make_log_path("apply")
                    write_scan_log(log_path, self.results)
                    messagebox.showinfo("Done", f"Apply complete.\n\nCSV log:\n{log_path.resolve()}")
        except queue.Empty:
            pass

        self.after(200, self._poll_queue)

    def insert_result_row(self, item: ScanResult) -> None:
        self.tree.insert("", "end", values=self.row_values(item))

    def refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for item in self.results:
            self.insert_result_row(item)

    def row_values(self, item: ScanResult) -> Tuple[str, ...]:
        return (
            item.status,
            safe_display(item.target_time),
            item.target_source,
            safe_display(item.metadata.created),
            safe_display(item.metadata.modified),
            safe_display(item.filename_date),
            item.metadata.source,
            safe_display(item.fs_times.created),
            safe_display(item.fs_times.modified),
            item.fs_times.error,
            item.message,
            str(item.path),
        )

    def save_csv_log(self) -> None:
        if not self.results:
            messagebox.showinfo("No results", "There is no scan result to save.")
            return

        filename = filedialog.asksaveasfilename(
            title="Save CSV log",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        try:
            write_scan_log(Path(filename), self.results)
            messagebox.showinfo("Saved", f"CSV log saved:\n{filename}")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    def open_log_folder(self) -> None:
        folder = Path(DEFAULT_LOG_DIR)
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(str(folder.resolve()))
            elif sys.platform == "darwin":
                os.system(f'open "{folder.resolve()}"')
            else:
                os.system(f'xdg-open "{folder.resolve()}"')
        except Exception as exc:
            messagebox.showerror("Open folder error", str(exc))


def main() -> int:
    try:
        app = OfficeTimestampGUI()
        app.mainloop()
        return 0
    except Exception:
        log_dir = Path(DEFAULT_LOG_DIR)
        log_dir.mkdir(parents=True, exist_ok=True)
        error_log = log_dir / f"{APP_NAME}_fatal_error.txt"
        error_log.write_text(traceback.format_exc(), encoding="utf-8")
        try:
            messagebox.showerror("Fatal error", f"Fatal error. See:\n{error_log.resolve()}")
        except Exception:
            print(f"Fatal error. See: {error_log.resolve()}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
