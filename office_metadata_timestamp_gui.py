# -*- coding: utf-8 -*-
"""
Office Metadata Timestamp GUI
Copyright 2026 // NGHIEN CUU THUOC // RnD PHARMA PLUS // WWW.NGHIENCUUTHUOC.COM

Purpose
-------
Paste files or folders, scan internal document metadata such as:
- Word/Excel/PowerPoint OOXML: Content created / Date last saved
- Legacy Office OLE files: .doc, .xls, .ppt metadata when olefile is installed
- PDF: /CreationDate and /ModDate when present

Then set Windows file timestamps to the internal "Content created" date:
- CreationTime
- LastWriteTime / ModifiedTime
- LastAccessTime

Safety
------
- The GUI scans first.
- Nothing is changed until you click "Apply timestamps".
- A CSV log is written for every scan/apply operation.

Recommended dependency
----------------------
pip install olefile

olefile is needed for legacy .doc/.xls/.ppt metadata reading.
Modern .docx/.xlsx/.pptx files do not require external libraries.
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


APP_NAME = "Office_Metadata_Timestamp_GUI"
APP_TITLE = "Office Metadata Timestamp GUI"
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


def as_local_aware(value: dt.datetime) -> dt.datetime:
    """Return a timezone-aware local datetime."""
    if value.tzinfo is None:
        return value.astimezone()
    return value.astimezone()


def ts_to_dt(timestamp: float) -> dt.datetime:
    return dt.datetime.fromtimestamp(timestamp).astimezone()


def dt_to_ts(value: dt.datetime) -> float:
    return as_local_aware(value).timestamp()


def safe_iso(value: Optional[dt.datetime]) -> str:
    return value.isoformat(timespec="seconds") if value else ""


def safe_display(value: Optional[dt.datetime]) -> str:
    if not value:
        return ""
    return value.strftime("%Y-%m-%d %H:%M:%S")


def parse_iso_datetime(text: str) -> Optional[dt.datetime]:
    """Parse ISO-like Office dates into local timezone-aware datetime."""
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
    """
    Parse common PDF dates:
      D:20160511124300
      D:20160511124300+07'00'
      D:20160511124300Z
    """
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


@dataclass
class FileSystemTimes:
    created: Optional[dt.datetime]
    modified: dt.datetime
    accessed: dt.datetime


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


def get_creation_time(path: Path, stat_result: os.stat_result) -> Optional[dt.datetime]:
    """
    On Windows, st_ctime is CreationTime.
    On Unix-like systems, st_ctime is metadata-change time, not CreationTime.
    """
    if is_windows():
        try:
            return ts_to_dt(stat_result.st_ctime)
        except Exception:
            return None
    return None


def get_fs_times(path: Path) -> FileSystemTimes:
    st = path.stat()
    return FileSystemTimes(
        created=get_creation_time(path, st),
        modified=ts_to_dt(st.st_mtime),
        accessed=ts_to_dt(st.st_atime),
    )


def local_name(tag: str) -> str:
    """Extract local XML tag name without namespace."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def read_ooxml_metadata(path: Path) -> MetadataResult:
    """Read Office Open XML core metadata from docProps/core.xml."""
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
            result.message = "Metadata found."
        else:
            result.status = "missing_date"
            result.message = "No created/modified date was found in core.xml."

        return result

    except zipfile.BadZipFile:
        result.status = "bad_file"
        result.message = "File is not a valid OOXML zip package."
        return result
    except Exception as exc:
        result.status = "error"
        result.message = str(exc)
        return result


def read_ole_metadata(path: Path) -> MetadataResult:
    """
    Read legacy Office OLE metadata using olefile, if available.
    Supports .doc, .xls, .ppt.
    """
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
            result.message = "Metadata found."
        else:
            result.status = "missing_date"
            result.message = "No create_time/last_saved_time was found."

        return result

    except Exception as exc:
        result.status = "error"
        result.message = str(exc)
        return result


def read_pdf_metadata(path: Path) -> MetadataResult:
    """
    Lightweight PDF metadata reader.
    It scans the first and last part of the PDF for /CreationDate and /ModDate.
    """
    result = MetadataResult(source="pdf_raw_metadata")

    try:
        with path.open("rb") as f:
            head = f.read(1024 * 1024)
            try:
                f.seek(max(0, path.stat().st_size - 1024 * 1024))
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
    allow_future_days: int,
    reject_future_dates: bool,
) -> Tuple[bool, str]:
    if value is None:
        return False, "empty"

    local_value = as_local_aware(value)
    if local_value.year < min_year:
        return False, f"year_before_{min_year}"

    if reject_future_dates:
        latest = now_local() + dt.timedelta(days=allow_future_days)
        if local_value > latest:
            return False, "future_date_rejected"

    return True, "ok"


def choose_target_time(
    metadata: MetadataResult,
    prefer_created: bool,
    allow_modified_fallback: bool,
    min_year: int,
    allow_future_days: int,
    reject_future_dates: bool,
) -> Tuple[Optional[dt.datetime], str, str]:
    """
    Select target timestamp from metadata.

    Default:
      1. Content created
      2. Date last saved only if fallback is enabled and valid
    """
    candidates: List[Tuple[str, Optional[dt.datetime]]] = []

    if prefer_created:
        candidates.append(("metadata_created_content_created", metadata.created))
        if allow_modified_fallback:
            candidates.append(("metadata_modified_last_saved", metadata.modified))
    else:
        candidates.append(("metadata_modified_last_saved", metadata.modified))
        if allow_modified_fallback:
            candidates.append(("metadata_created_content_created", metadata.created))

    rejection_notes = []
    for label, value in candidates:
        valid, reason = is_valid_metadata_date(
            value,
            min_year=min_year,
            allow_future_days=allow_future_days,
            reject_future_dates=reject_future_dates,
        )
        if valid:
            return as_local_aware(value), label, "ok"
        rejection_notes.append(f"{label}:{reason}")

    return None, "", "; ".join(rejection_notes) if rejection_notes else "no_candidate"


def parse_pasted_paths(text: str) -> List[Path]:
    """
    Parse one path per line. Supports quoted paths pasted from Explorer or command line.
    """
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

    # Windows FILETIME: 100-ns intervals since 1601-01-01 UTC.
    filetime_int = int((timestamp + 11644473600) * 10000000)

    return FILETIME(
        filetime_int & 0xFFFFFFFF,
        filetime_int >> 32,
    )


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

    current = get_fs_times(path)
    access_time = dt_to_ts(new_time if set_accessed else current.accessed)
    modified_time = dt_to_ts(new_time if set_modified else current.modified)
    os.utime(str(path), (access_time, modified_time), follow_symlinks=False)

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


def write_scan_log(log_path: Path, results: Sequence[ScanResult]) -> None:
    with log_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "path",
                "status",
                "message",
                "metadata_source",
                "target_source",
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
            writer.writerow(
                [
                    str(item.path),
                    item.status,
                    item.message,
                    item.metadata.source,
                    item.target_source,
                    safe_iso(item.fs_times.created if item.fs_times else None),
                    safe_iso(item.fs_times.modified if item.fs_times else None),
                    safe_iso(item.fs_times.accessed if item.fs_times else None),
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
        self.geometry("1280x780")
        self.minsize(1000, 650)

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
            "status",
            "target_time",
            "metadata_created",
            "metadata_modified",
            "source",
            "fs_created",
            "fs_modified",
            "message",
            "path",
        )

        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=18)

        headings = {
            "status": "Status",
            "target_time": "New target time",
            "metadata_created": "Metadata created",
            "metadata_modified": "Metadata modified",
            "source": "Source",
            "fs_created": "FS created",
            "fs_modified": "FS modified",
            "message": "Message",
            "path": "Path",
        }

        widths = {
            "status": 110,
            "target_time": 150,
            "metadata_created": 150,
            "metadata_modified": 150,
            "source": 160,
            "fs_created": 150,
            "fs_modified": 150,
            "message": 240,
            "path": 520,
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

        help_text = (
            "Tip: For legacy .doc/.xls/.ppt files, install olefile first. "
            "The target date normally comes from Office 'Content created'. "
            "Keep 'Date last saved fallback' off when files have bogus saved dates such as year 2062."
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

        files = list(
            iter_input_files(
                paths,
                recursive=self.recursive_var.get(),
                supported_only=self.supported_only_var.get(),
            )
        )

        unique: List[Path] = []
        seen = set()
        for file_path in files:
            key = str(file_path).lower()
            if key not in seen:
                unique.append(file_path)
                seen.add(key)

        if not unique:
            raise ValueError("No matching files were found.")

        return unique

    def set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.scan_button.configure(state=state)
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
            "reject_future_dates": self.reject_future_var.get(),
            "min_year": min_year,
            "allow_future_days": future_days,
        }

        self.worker_thread = threading.Thread(
            target=self.scan_worker,
            args=(files, options),
            daemon=True,
        )
        self.worker_thread.start()

    def scan_worker(self, files: Sequence[Path], options: Dict[str, object]) -> None:
        results: List[ScanResult] = []

        for index, file_path in enumerate(files, start=1):
            try:
                fs_times = get_fs_times(file_path)
                metadata = read_document_metadata(file_path)

                target_time, target_source, target_message = choose_target_time(
                    metadata,
                    prefer_created=bool(options["prefer_created"]),
                    allow_modified_fallback=bool(options["allow_modified_fallback"]),
                    min_year=int(options["min_year"]),
                    allow_future_days=int(options["allow_future_days"]),
                    reject_future_dates=bool(options["reject_future_dates"]),
                )

                if target_time:
                    status = "ready"
                    message = metadata.message
                else:
                    status = "no_valid_date"
                    message = f"{metadata.message} | {target_message}"

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
            (
                f"Apply timestamps to {len(ready)} files?\n\n"
                "This will modify selected Windows file timestamps.\n"
                "A CSV log will be saved automatically."
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

        self.worker_thread = threading.Thread(
            target=self.apply_worker,
            args=(ready, options),
            daemon=True,
        )
        self.worker_thread.start()

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
                    error_count = sum(1 for item in self.results if item.status in {"error", "no_valid_date"})
                    self.summary_var.set(
                        f"Scan complete. Total: {len(self.results)} | Ready: {ready_count} | Not ready/errors: {error_count}"
                    )
                    self.status_var.set("Scan complete.")
                    self.apply_button.configure(state="normal" if ready_count else "disabled")

                    log_path = make_log_path("scan")
                    write_scan_log(log_path, self.results)

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
                    write_scan_log(log_path, self.results)
                    messagebox.showinfo("Done", f"Apply complete.\n\nCSV log:\n{log_path.resolve()}")

        except queue.Empty:
            pass

        self.after(200, self._poll_queue)

    def insert_result_row(self, item: ScanResult) -> None:
        self.tree.insert(
            "",
            "end",
            values=self.row_values(item),
        )

    def refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for item in self.results:
            self.insert_result_row(item)

    def row_values(self, item: ScanResult) -> Tuple[str, ...]:
        return (
            item.status,
            safe_display(item.target_time),
            safe_display(item.metadata.created),
            safe_display(item.metadata.modified),
            item.metadata.source,
            safe_display(item.fs_times.created if item.fs_times else None),
            safe_display(item.fs_times.modified if item.fs_times else None),
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
        messagebox.showerror("Fatal error", f"Fatal error. See:\n{error_log.resolve()}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
