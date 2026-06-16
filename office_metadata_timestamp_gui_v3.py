# -*- coding: utf-8 -*-
"""
Office Metadata Timestamp GUI v3
Copyright 2026 // NGHIEN CUU THUOC // RnD PHARMA PLUS // WWW.NGHIENCUUTHUOC.COM

New in v3
---------
- Future-date filter: scan only files whose Windows filesystem timestamp is
  greater than the current local date/time.
- Optional date-only mode: scan only files whose filesystem date is after today.
- The filter is applied before reading Office/PDF metadata, so large folders
  are faster and clean files are skipped.
- Robust filesystem timestamp handling: corrupted timestamps are logged without
  stopping metadata scanning.

Purpose
-------
Paste files or folders, scan internal document metadata, and set Windows file
timestamps to the document's "Content created" date.

Supported metadata sources
--------------------------
- OOXML Office: .docx, .xlsx, .pptx and macro/template variants
- Legacy Office OLE: .doc, .xls, .ppt when olefile is installed
- PDF: /CreationDate and /ModDate when present
- Filename/path fallback: dates such as 17.05.2016 or 2016-05-17

Recommended dependency
----------------------
pip install olefile
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


APP_NAME = "Office_Metadata_Timestamp_GUI_v3"
APP_TITLE = "Office Metadata Timestamp GUI v3"
DEFAULT_LOG_DIR = "office_timestamp_logs_v3"

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


def set_console_title(title: str) -> None:
    if os.name == "nt":
        try:
            ctypes.windll.kernel32.SetConsoleTitleW(title)
        except Exception:
            pass


def is_windows() -> bool:
    return platform.system().lower() == "windows"


def now_local() -> dt.datetime:
    return dt.datetime.now().astimezone()


def local_midnight(value: Optional[dt.datetime] = None) -> dt.datetime:
    value = value or now_local()
    return value.replace(hour=0, minute=0, second=0, microsecond=0)


def as_local_aware(value: dt.datetime) -> dt.datetime:
    if value.tzinfo is None:
        return value.astimezone()
    return value.astimezone()


def safe_fromtimestamp(timestamp: float) -> Tuple[Optional[dt.datetime], str]:
    try:
        return dt.datetime.fromtimestamp(timestamp).astimezone(), ""
    except Exception as exc:
        return None, str(exc)


def safe_iso(value: Optional[dt.datetime]) -> str:
    return value.isoformat(timespec="seconds") if value else ""


def safe_display(value: Optional[dt.datetime]) -> str:
    if not value:
        return ""
    return value.strftime("%Y-%m-%d %H:%M:%S")


def parse_iso_datetime(text: str) -> Optional[dt.datetime]:
    if not text:
        return None

    value = str(text).strip()
    if not value:
        return None

    try:
        if value.endswith("Z"):
            value = value[:-1] + "+00:00"
        parsed = dt.datetime.fromisoformat(value)
        return as_local_aware(parsed)
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
                offset = dt.timedelta(hours=tz_hour, minutes=tz_min) * sign
                parsed = parsed.replace(tzinfo=dt.timezone(offset))

        return as_local_aware(parsed)
    except Exception:
        return None


def parse_filename_or_path_date(path: Path) -> Optional[dt.datetime]:
    """
    Extract a likely date from file name or path.

    Supported examples:
      17.05.2016
      17-05-2016
      17_05_2016
      2016-05-17
      2016.05.17
    """
    text = str(path)

    patterns = [
        # dd.mm.yyyy, dd-mm-yyyy, dd_mm_yyyy
        r"(?<!\d)(?P<d>[0-3]?\d)[.\-_ ](?P<m>[0-1]?\d)[.\-_ ](?P<y>(?:19|20)\d{2})(?!\d)",
        # yyyy.mm.dd, yyyy-mm-dd, yyyy_mm_dd
        r"(?<!\d)(?P<y>(?:19|20)\d{2})[.\-_ ](?P<m>[0-1]?\d)[.\-_ ](?P<d>[0-3]?\d)(?!\d)",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text):
            try:
                year = int(match.group("y"))
                month = int(match.group("m"))
                day = int(match.group("d"))
                if 1900 <= year <= now_local().year:
                    return dt.datetime(year, month, day, 0, 0, 0).astimezone()
            except Exception:
                continue

    return None


@dataclass
class FileSystemTimes:
    created: Optional[dt.datetime]
    modified: Optional[dt.datetime]
    accessed: Optional[dt.datetime]
    raw_created: Optional[float]
    raw_modified: Optional[float]
    raw_accessed: Optional[float]
    future_fields: str = ""
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
    fs_times: Optional[FileSystemTimes] = None
    metadata: MetadataResult = field(default_factory=MetadataResult)
    target_time: Optional[dt.datetime] = None
    target_source: str = ""
    status: str = "pending"
    message: str = ""
    applied: bool = False
    creation_set: bool = False
    modified_set: bool = False
    accessed_set: bool = False


def future_threshold_epoch(date_only: bool, allowance_days: int) -> float:
    """
    Return epoch threshold for future filtering.

    date_only=True:
      Future means date is after today, so threshold is tomorrow 00:00 local
      plus allowance days.

    date_only=False:
      Future means timestamp is later than now plus allowance days.
    """
    if date_only:
        threshold_dt = local_midnight() + dt.timedelta(days=1 + allowance_days)
    else:
        threshold_dt = now_local() + dt.timedelta(days=allowance_days)

    return threshold_dt.timestamp()


def get_fs_times(path: Path, date_only: bool, allowance_days: int) -> FileSystemTimes:
    """
    Read filesystem timestamps safely.

    Raw timestamps are compared numerically to detect future dates. This avoids
    datetime conversion errors for corrupted future or out-of-range timestamps.
    """
    errors: List[str] = []
    st = path.stat()

    raw_created = st.st_ctime if is_windows() else None
    raw_modified = st.st_mtime
    raw_accessed = st.st_atime

    created, created_err = safe_fromtimestamp(raw_created) if raw_created is not None else (None, "")
    modified, modified_err = safe_fromtimestamp(raw_modified)
    accessed, accessed_err = safe_fromtimestamp(raw_accessed)

    if created_err:
        errors.append(f"creation:{created_err}")
    if modified_err:
        errors.append(f"modified:{modified_err}")
    if accessed_err:
        errors.append(f"accessed:{accessed_err}")

    threshold = future_threshold_epoch(date_only=date_only, allowance_days=allowance_days)
    future_fields: List[str] = []

    if raw_created is not None and raw_created >= threshold:
        future_fields.append("creation")
    if raw_modified >= threshold:
        future_fields.append("modified")
    if raw_accessed >= threshold:
        future_fields.append("accessed")

    return FileSystemTimes(
        created=created,
        modified=modified,
        accessed=accessed,
        raw_created=raw_created,
        raw_modified=raw_modified,
        raw_accessed=raw_accessed,
        future_fields="+".join(future_fields),
        error="; ".join(errors),
    )


def has_future_fs_time(fs_times: FileSystemTimes) -> bool:
    return bool(fs_times.future_fields)


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
            result.message = "No created/modified date was found in core.xml."

        return result

    except zipfile.BadZipFile:
        # Some files have a wrong extension, for example .docx that is really .doc.
        ole_result = read_ole_metadata(path)
        if ole_result.status in {"ok", "missing_date"}:
            ole_result.source = "ole_fallback_wrong_extension"
            return ole_result

        result.status = "bad_file"
        result.message = "File is not a valid OOXML zip package."
        return result

    except Exception as exc:
        result.status = "error"
        result.message = str(exc)
        return result


def read_ole_metadata(path: Path) -> MetadataResult:
    result = MetadataResult(source="ole_summary_information")

    try:
        import olefile
    except Exception:
        result.status = "missing_dependency"
        result.message = "Install dependency first: pip install olefile"
        return result

    try:
        if not olefile.isOleFile(str(path)):
            result.status = "bad_file"
            result.message = "File is not a valid OLE compound document."
            return result

        with olefile.OleFileIO(str(path)) as ole:
            meta = ole.get_metadata()

            created = getattr(meta, "create_time", None)
            modified = getattr(meta, "last_saved_time", None)

            if isinstance(created, dt.datetime):
                result.created = as_local_aware(created)

            if isinstance(modified, dt.datetime):
                result.modified = as_local_aware(modified)

            result.author = str(getattr(meta, "author", "") or "")
            result.last_saved_by = str(getattr(meta, "last_saved_by", "") or "")

        if result.created or result.modified:
            result.status = "ok"
            result.message = "OLE metadata found."
        else:
            result.status = "missing_date"
            result.message = "No create_time/last_saved_time was found."

        return result

    except Exception as exc:
        result.status = "error"
        result.message = str(exc)
        return result


def read_pdf_metadata(path: Path) -> MetadataResult:
    result = MetadataResult(source="pdf_raw_metadata")

    try:
        size = path.stat().st_size
        with path.open("rb") as f:
            head = f.read(1024 * 1024)
            try:
                f.seek(max(0, size - 1024 * 1024))
                tail = f.read(1024 * 1024)
            except Exception:
                tail = b""

        raw = head + b"\n" + tail
        text = raw.decode("latin-1", errors="ignore")

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
        result.message = str(exc)
        return result


def read_document_metadata(path: Path) -> MetadataResult:
    ext = path.suffix.lower()

    if ext in OOXML_EXTS:
        return read_ooxml_metadata(path)

    if ext in OLE_EXTS:
        return read_ole_metadata(path)

    if ext in PDF_EXTS:
        return read_pdf_metadata(path)

    return MetadataResult(
        source="unsupported",
        status="unsupported",
        message=f"Unsupported extension: {ext}",
    )


def is_valid_metadata_date(
    value: Optional[dt.datetime],
    min_year: int,
    reject_future_dates: bool,
    future_allowance_days: int,
) -> Tuple[bool, str]:
    if value is None:
        return False, "empty"

    local_value = as_local_aware(value)
    if local_value.year < min_year:
        return False, f"year_before_{min_year}"

    if reject_future_dates:
        latest = now_local() + dt.timedelta(days=future_allowance_days)
        if local_value > latest:
            return False, "future_date_rejected"

    return True, "ok"


def choose_target_time(
    path: Path,
    metadata: MetadataResult,
    prefer_created: bool,
    allow_modified_fallback: bool,
    use_path_date_fallback: bool,
    min_year: int,
    reject_future_dates: bool,
    future_allowance_days: int,
) -> Tuple[Optional[dt.datetime], str, str]:
    candidates: List[Tuple[str, Optional[dt.datetime]]] = []

    if prefer_created:
        candidates.append(("metadata_created_content_created", metadata.created))
        if allow_modified_fallback:
            candidates.append(("metadata_modified_last_saved", metadata.modified))
    else:
        candidates.append(("metadata_modified_last_saved", metadata.modified))
        if allow_modified_fallback:
            candidates.append(("metadata_created_content_created", metadata.created))

    if use_path_date_fallback:
        candidates.append(("filename_or_path_date", parse_filename_or_path_date(path)))

    rejection_notes: List[str] = []
    for label, value in candidates:
        valid, reason = is_valid_metadata_date(
            value,
            min_year=min_year,
            reject_future_dates=reject_future_dates,
            future_allowance_days=future_allowance_days,
        )
        if valid:
            return as_local_aware(value), label, "ok"
        rejection_notes.append(f"{label}:{reason}")

    return None, "", "; ".join(rejection_notes) if rejection_notes else "no_candidate"


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
    local_value = as_local_aware(value)
    timestamp = local_value.timestamp()
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

    CreateFileW = kernel32.CreateFileW
    CreateFileW.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
    ]
    CreateFileW.restype = ctypes.c_void_p

    SetFileTime = kernel32.SetFileTime
    SetFileTime.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
    ]
    SetFileTime.restype = ctypes.c_int

    CloseHandle = kernel32.CloseHandle
    CloseHandle.argtypes = [ctypes.c_void_p]
    CloseHandle.restype = ctypes.c_int

    FILE_WRITE_ATTRIBUTES = 0x0100
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    flags = FILE_FLAG_BACKUP_SEMANTICS if path.is_dir() else 0

    handle = CreateFileW(
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
        filetime = datetime_to_filetime(new_time)
        created_ptr = ctypes.byref(filetime) if set_created else None
        accessed_ptr = ctypes.byref(filetime) if set_accessed else None
        modified_ptr = ctypes.byref(filetime) if set_modified else None

        ok = SetFileTime(handle, created_ptr, accessed_ptr, modified_ptr)
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())

        return {
            "creation_set": bool(set_created),
            "modified_set": bool(set_modified),
            "accessed_set": bool(set_accessed),
        }
    finally:
        CloseHandle(handle)


def set_times_cross_platform(
    path: Path,
    new_time: dt.datetime,
    set_created: bool,
    set_modified: bool,
    set_accessed: bool,
) -> Dict[str, bool]:
    if is_windows():
        return set_times_windows(path, new_time, set_created, set_modified, set_accessed)

    current_stat = path.stat()
    access_ts = new_time.timestamp() if set_accessed else current_stat.st_atime
    modified_ts = new_time.timestamp() if set_modified else current_stat.st_mtime
    os.utime(str(path), (access_ts, modified_ts), follow_symlinks=False)

    return {
        "creation_set": False,
        "modified_set": bool(set_modified),
        "accessed_set": bool(set_accessed),
    }


def make_log_path(action: str) -> Path:
    log_dir = Path(DEFAULT_LOG_DIR)
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return log_dir / f"{APP_NAME}_{action}_{stamp}.csv"


def write_scan_log(log_path: Path, results: Sequence[ScanResult], skipped_not_future: int = 0) -> None:
    with log_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["summary_skipped_not_future", skipped_not_future])
        writer.writerow([])
        writer.writerow(
            [
                "path",
                "status",
                "message",
                "metadata_source",
                "target_source",
                "future_fs_fields",
                "fs_time_error",
                "fs_creation_time",
                "fs_modified_time",
                "fs_accessed_time",
                "metadata_created",
                "metadata_modified",
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
            fs = item.fs_times
            writer.writerow(
                [
                    str(item.path),
                    item.status,
                    item.message,
                    item.metadata.source,
                    item.target_source,
                    fs.future_fields if fs else "",
                    fs.error if fs else "",
                    safe_iso(fs.created if fs else None),
                    safe_iso(fs.modified if fs else None),
                    safe_iso(fs.accessed if fs else None),
                    safe_iso(item.metadata.created),
                    safe_iso(item.metadata.modified),
                    item.metadata.author,
                    item.metadata.last_saved_by,
                    safe_iso(item.target_time),
                    item.applied,
                    item.creation_set,
                    item.modified_set,
                    item.accessed_set,
                ]
            )


class OfficeTimestampGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        set_console_title(APP_NAME)

        self.title(APP_TITLE)
        self.geometry("1360x820")
        self.minsize(1080, 680)

        self.results: List[ScanResult] = []
        self.skipped_not_future = 0
        self.enumerated_count = 0
        self.worker_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()

        self._build_vars()
        self._build_ui()
        self._poll_queue()

    def _build_vars(self) -> None:
        self.recursive_var = tk.BooleanVar(value=True)
        self.supported_only_var = tk.BooleanVar(value=True)

        self.future_only_var = tk.BooleanVar(value=True)
        self.future_date_only_var = tk.BooleanVar(value=True)
        self.show_skipped_var = tk.BooleanVar(value=False)

        self.prefer_created_var = tk.BooleanVar(value=True)
        self.modified_fallback_var = tk.BooleanVar(value=False)
        self.reject_future_metadata_var = tk.BooleanVar(value=True)
        self.path_date_fallback_var = tk.BooleanVar(value=True)

        self.set_created_var = tk.BooleanVar(value=True)
        self.set_modified_var = tk.BooleanVar(value=True)
        self.set_accessed_var = tk.BooleanVar(value=True)

        self.min_year_var = tk.StringVar(value="1900")
        self.future_days_var = tk.StringVar(value="0")

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

        col1 = ttk.Frame(options_frame)
        col2 = ttk.Frame(options_frame)
        col3 = ttk.Frame(options_frame)
        col4 = ttk.Frame(options_frame)

        for col in (col1, col2, col3, col4):
            col.pack(side="left", fill="x", expand=True, padx=8, pady=6)

        ttk.Label(col1, text="Input scanning").pack(anchor="w")
        ttk.Checkbutton(col1, text="Scan folders recursively", variable=self.recursive_var).pack(anchor="w")
        ttk.Checkbutton(col1, text="Only supported extensions", variable=self.supported_only_var).pack(anchor="w")

        ttk.Label(col2, text="v3 future-date filter").pack(anchor="w")
        ttk.Checkbutton(
            col2,
            text="Scan only files with filesystem date > current date",
            variable=self.future_only_var,
        ).pack(anchor="w")
        ttk.Checkbutton(
            col2,
            text="Compare by date only, not exact time",
            variable=self.future_date_only_var,
        ).pack(anchor="w")
        ttk.Checkbutton(
            col2,
            text="Show skipped non-future files",
            variable=self.show_skipped_var,
        ).pack(anchor="w")

        future_row = ttk.Frame(col2)
        future_row.pack(anchor="w", fill="x", pady=2)
        ttk.Label(future_row, text="Future allowance days:").pack(side="left")
        ttk.Entry(future_row, textvariable=self.future_days_var, width=8).pack(side="left", padx=4)

        ttk.Label(col3, text="Metadata source").pack(anchor="w")
        ttk.Checkbutton(col3, text="Prefer Content created date", variable=self.prefer_created_var).pack(anchor="w")
        ttk.Checkbutton(col3, text="Allow Date last saved fallback", variable=self.modified_fallback_var).pack(anchor="w")
        ttk.Checkbutton(col3, text="Reject future metadata dates", variable=self.reject_future_metadata_var).pack(anchor="w")
        ttk.Checkbutton(col3, text="Use filename/path date fallback", variable=self.path_date_fallback_var).pack(anchor="w")

        year_row = ttk.Frame(col3)
        year_row.pack(anchor="w", fill="x", pady=2)
        ttk.Label(year_row, text="Minimum valid year:").pack(side="left")
        ttk.Entry(year_row, textvariable=self.min_year_var, width=8).pack(side="left", padx=4)

        ttk.Label(col4, text="Set these file timestamps").pack(anchor="w")
        ttk.Checkbutton(col4, text="CreationTime", variable=self.set_created_var).pack(anchor="w")
        ttk.Checkbutton(col4, text="ModifiedTime / LastWriteTime", variable=self.set_modified_var).pack(anchor="w")
        ttk.Checkbutton(col4, text="AccessedTime", variable=self.set_accessed_var).pack(anchor="w")

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
            "status",
            "future_fields",
            "target_time",
            "metadata_created",
            "metadata_modified",
            "source",
            "fs_created",
            "fs_modified",
            "fs_accessed",
            "message",
            "path",
        )

        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=19)

        headings = {
            "status": "Status",
            "future_fields": "Future FS fields",
            "target_time": "New target time",
            "metadata_created": "Metadata created",
            "metadata_modified": "Metadata modified",
            "source": "Source",
            "fs_created": "FS created",
            "fs_modified": "FS modified",
            "fs_accessed": "FS accessed",
            "message": "Message",
            "path": "Path",
        }

        widths = {
            "status": 110,
            "future_fields": 120,
            "target_time": 150,
            "metadata_created": 150,
            "metadata_modified": 150,
            "source": 160,
            "fs_created": 150,
            "fs_modified": 150,
            "fs_accessed": 150,
            "message": 260,
            "path": 560,
        }

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
        ttk.Label(
            summary_frame,
            text="Recommended: future filter ON, Content created ON, Date last saved fallback OFF.",
            foreground="#444444",
        ).pack(side="right", anchor="e")

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

    def get_input_paths(self) -> List[Path]:
        paths = parse_pasted_paths(self.path_text.get("1.0", "end"))
        if not paths:
            raise ValueError("Please paste at least one file or folder path.")
        return paths

    def set_busy(self, busy: bool) -> None:
        self.scan_button.configure(state="disabled" if busy else "normal")
        self.apply_button.configure(state="disabled" if busy or not self.results else "normal")

    def start_scan(self) -> None:
        try:
            input_paths = self.get_input_paths()
            min_year = self.get_int_option(self.min_year_var, "Minimum valid year", 1900)
            future_days = self.get_int_option(self.future_days_var, "Future allowance days", 0)
        except Exception as exc:
            messagebox.showerror("Input error", str(exc))
            return

        self.tree.delete(*self.tree.get_children())
        self.results = []
        self.skipped_not_future = 0
        self.enumerated_count = 0

        self.status_var.set("Scanning...")
        self.summary_var.set("Enumerating and filtering files...")
        self.set_busy(True)

        options = {
            "recursive": self.recursive_var.get(),
            "supported_only": self.supported_only_var.get(),
            "future_only": self.future_only_var.get(),
            "future_date_only": self.future_date_only_var.get(),
            "show_skipped": self.show_skipped_var.get(),
            "prefer_created": self.prefer_created_var.get(),
            "allow_modified_fallback": self.modified_fallback_var.get(),
            "reject_future_metadata": self.reject_future_metadata_var.get(),
            "use_path_date_fallback": self.path_date_fallback_var.get(),
            "min_year": min_year,
            "future_days": future_days,
        }

        thread = threading.Thread(target=self.scan_worker, args=(input_paths, options), daemon=True)
        thread.start()

    def scan_worker(self, input_paths: Sequence[Path], options: Dict[str, object]) -> None:
        results: List[ScanResult] = []
        skipped_not_future = 0
        enumerated = 0

        seen = set()

        for file_path in iter_input_files(
            input_paths,
            recursive=bool(options["recursive"]),
            supported_only=bool(options["supported_only"]),
        ):
            key = str(file_path).lower()
            if key in seen:
                continue
            seen.add(key)

            enumerated += 1

            try:
                fs_times = get_fs_times(
                    file_path,
                    date_only=bool(options["future_date_only"]),
                    allowance_days=int(options["future_days"]),
                )

                if bool(options["future_only"]) and not has_future_fs_time(fs_times):
                    skipped_not_future += 1

                    if bool(options["show_skipped"]):
                        skipped_item = ScanResult(
                            path=file_path,
                            fs_times=fs_times,
                            status="skipped_not_future",
                            message="Skipped because filesystem timestamp is not greater than current date.",
                        )
                        results.append(skipped_item)
                        self.worker_queue.put(("row", skipped_item))

                    if enumerated % 100 == 0:
                        self.worker_queue.put(
                            ("status", f"Checked {enumerated} files | Future-date files: {len(results)} | Skipped: {skipped_not_future}")
                        )
                    continue

                metadata = read_document_metadata(file_path)

                target_time, target_source, target_message = choose_target_time(
                    path=file_path,
                    metadata=metadata,
                    prefer_created=bool(options["prefer_created"]),
                    allow_modified_fallback=bool(options["allow_modified_fallback"]),
                    use_path_date_fallback=bool(options["use_path_date_fallback"]),
                    min_year=int(options["min_year"]),
                    reject_future_dates=bool(options["reject_future_metadata"]),
                    future_allowance_days=int(options["future_days"]),
                )

                if target_time:
                    status = "ready"
                    message = metadata.message
                    if fs_times.error:
                        message = f"{message} | FS time warning: {fs_times.error}"
                else:
                    status = "no_valid_date"
                    message = f"{metadata.message} | {target_message}"
                    if fs_times.error:
                        message = f"{message} | FS time warning: {fs_times.error}"

                item = ScanResult(
                    path=file_path,
                    fs_times=fs_times,
                    metadata=metadata,
                    target_time=target_time,
                    target_source=target_source,
                    status=status,
                    message=message,
                )

            except Exception as exc:
                item = ScanResult(
                    path=file_path,
                    status="error",
                    message=str(exc),
                )

            results.append(item)
            self.worker_queue.put(("row", item))

            if enumerated % 25 == 0:
                self.worker_queue.put(
                    ("status", f"Checked {enumerated} files | Displayed: {len(results)} | Skipped: {skipped_not_future}")
                )

        self.worker_queue.put(("scan_done", (results, skipped_not_future, enumerated)))

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
            (
                f"Apply timestamps to {len(ready)} files?\n\n"
                "Only files displayed as ready will be updated."
            ),
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

        thread = threading.Thread(target=self.apply_worker, args=(ready, options), daemon=True)
        thread.start()

    def apply_worker(self, items: Sequence[ScanResult], options: Dict[str, object]) -> None:
        updated_by_path: Dict[str, ScanResult] = {}

        for index, item in enumerate(items, start=1):
            try:
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
                item.message = str(exc)

            updated_by_path[str(item.path)] = item

            if index % 10 == 0 or index == len(items):
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
                    results, skipped_not_future, enumerated = payload
                    self.results = list(results)
                    self.skipped_not_future = int(skipped_not_future)
                    self.enumerated_count = int(enumerated)

                    self.refresh_tree()
                    self.set_busy(False)

                    ready_count = sum(1 for item in self.results if item.target_time and item.status == "ready")
                    error_count = sum(1 for item in self.results if item.status in {"error", "no_valid_date"})
                    displayed = len(self.results)

                    self.summary_var.set(
                        f"Scan complete. Enumerated: {self.enumerated_count} | "
                        f"Skipped non-future: {self.skipped_not_future} | "
                        f"Displayed: {displayed} | Ready: {ready_count} | Not ready/errors: {error_count}"
                    )
                    self.status_var.set("Scan complete.")
                    self.apply_button.configure(state="normal" if ready_count else "disabled")

                    log_path = make_log_path("scan")
                    write_scan_log(log_path, self.results, self.skipped_not_future)

                elif event == "apply_done":
                    updated_by_path = payload
                    for index, item in enumerate(self.results):
                        new_item = updated_by_path.get(str(item.path))
                        if new_item:
                            self.results[index] = new_item

                    self.refresh_tree()
                    self.set_busy(False)

                    applied_count = sum(1 for item in self.results if item.applied)
                    error_count = sum(1 for item in self.results if item.status == "apply_error")
                    self.summary_var.set(
                        f"Apply complete. Applied: {applied_count} | Apply errors: {error_count}"
                    )
                    self.status_var.set("Apply complete.")

                    log_path = make_log_path("apply")
                    write_scan_log(log_path, self.results, self.skipped_not_future)
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
        fs = item.fs_times
        return (
            item.status,
            fs.future_fields if fs else "",
            safe_display(item.target_time),
            safe_display(item.metadata.created),
            safe_display(item.metadata.modified),
            item.metadata.source,
            safe_display(fs.created if fs else None),
            safe_display(fs.modified if fs else None),
            safe_display(fs.accessed if fs else None),
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
            write_scan_log(Path(filename), self.results, self.skipped_not_future)
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
