# -*- coding: utf-8 -*-
"""
PharmSolu Timestamp Setter
Copyright 2026 // NGHIEN CUU THUOC // RnD PHARMA PLUS

This tool scans a folder such as D:\PharmSolu and resets suspicious timestamps.
It can run in interactive mode or from a .bat launcher.

Main use cases
--------------
1) Set only the year, preserving original month/day/time.
   Example: 2029-05-14 14:24:00 -> 2015-05-14 14:24:00

2) Set one exact date/time for all selected files.
   Example: 2015-01-01 00:00:00

3) Detect a year from the file/folder path.
   Example: "CHINESE PHARMACOPOEIA 2015 CMSP" -> year 2015

Safety
------
Default mode is DRY-RUN. No file timestamps are changed unless --apply is used.

Examples
--------
Preview suspicious files only:
    python pharmsolu_timestamp_setter.py --root "D:\PharmSolu" --mode year --year 2015

Apply to suspicious files only:
    python pharmsolu_timestamp_setter.py --root "D:\PharmSolu" --mode year --year 2015 --apply

Apply exact datetime:
    python pharmsolu_timestamp_setter.py --root "D:\PharmSolu" --mode exact --datetime "2015-01-01 00:00:00" --apply

Use year from path:
    python pharmsolu_timestamp_setter.py --root "D:\PharmSolu" --mode path-year --apply

Interactive:
    python pharmsolu_timestamp_setter.py --interactive
"""

from __future__ import annotations

import argparse
import csv
import ctypes
import datetime as dt
import os
import platform
import re
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple


APP_TITLE = "PharmSolu_Timestamp_Setter"
DEFAULT_ROOT = r"D:\PharmSolu"
DEFAULT_LOG_DIR = "timestamp_setter_logs"


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


def ts_to_dt(timestamp: float) -> dt.datetime:
    return dt.datetime.fromtimestamp(timestamp).astimezone()


def dt_to_ts(value: dt.datetime) -> float:
    if value.tzinfo is None:
        value = value.astimezone()
    return value.timestamp()


def safe_iso(value: Optional[dt.datetime]) -> str:
    return value.isoformat(timespec="seconds") if value else ""


def parse_datetime_text(text: str) -> dt.datetime:
    """
    Parse one of:
      YYYY
      YYYY-MM-DD
      YYYY-MM-DD HH:MM
      YYYY-MM-DD HH:MM:SS
    """
    text = text.strip()
    formats = [
        "%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
    ]

    last_error: Optional[Exception] = None
    for fmt in formats:
        try:
            parsed = dt.datetime.strptime(text, fmt)
            if fmt == "%Y":
                parsed = parsed.replace(month=1, day=1, hour=0, minute=0, second=0)
            return parsed.astimezone()
        except Exception as exc:
            last_error = exc

    raise ValueError(
        "Invalid datetime format. Use YYYY, YYYY-MM-DD, "
        "YYYY-MM-DD HH:MM, or YYYY-MM-DD HH:MM:SS."
    ) from last_error


def safe_replace_year(value: dt.datetime, target_year: int) -> dt.datetime:
    """Replace only the year. Handles February 29 for non-leap years."""
    try:
        return value.replace(year=target_year)
    except ValueError:
        if value.month == 2 and value.day == 29:
            return value.replace(year=target_year, day=28)
        return dt.datetime(
            target_year,
            1,
            1,
            value.hour,
            value.minute,
            value.second,
            value.microsecond,
            tzinfo=value.tzinfo,
        )


def extract_year_from_path(path: Path, max_year: Optional[int] = None) -> Optional[int]:
    """
    Extract a plausible year from the path.

    Examples:
      CHINESE PHARMACOPOEIA 2015 CMSP -> 2015
      BP 2024 -> 2024
    """
    if max_year is None:
        max_year = now_local().year

    years = [int(y) for y in re.findall(r"(?:19|20)\d{2}", str(path))]
    plausible = [y for y in years if 1900 <= y <= max_year]
    if not plausible:
        return None
    return max(plausible)


@dataclass
class TimestampInfo:
    created: Optional[dt.datetime]
    modified: dt.datetime
    accessed: dt.datetime


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


def get_timestamp_info(path: Path) -> TimestampInfo:
    st = path.stat()
    return TimestampInfo(
        created=get_creation_time(path, st),
        modified=ts_to_dt(st.st_mtime),
        accessed=ts_to_dt(st.st_atime),
    )


def is_suspicious_datetime(
    value: Optional[dt.datetime],
    min_year: int,
    future_margin_days: int,
) -> bool:
    if value is None:
        return False

    latest_allowed = now_local() + dt.timedelta(days=future_margin_days)

    if value > latest_allowed:
        return True

    if value.year < min_year:
        return True

    return False


def suspicious_reasons(
    info: TimestampInfo,
    min_year: int,
    future_margin_days: int,
) -> Tuple[bool, str]:
    reasons = []

    if is_suspicious_datetime(info.created, min_year, future_margin_days):
        reasons.append("creation")

    if is_suspicious_datetime(info.modified, min_year, future_margin_days):
        reasons.append("modified")

    if is_suspicious_datetime(info.accessed, min_year, future_margin_days):
        reasons.append("accessed")

    return bool(reasons), "+".join(reasons)


def win32_long_path(path: Path) -> str:
    """Return a Windows long-path-safe string."""
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
    """
    Convert local datetime to Windows FILETIME.

    FILETIME is the number of 100-nanosecond intervals since 1601-01-01 UTC.
    """
    if value.tzinfo is None:
        timestamp = value.astimezone().timestamp()
    else:
        timestamp = value.timestamp()

    filetime_int = int((timestamp + 11644473600) * 10000000)
    return FILETIME(
        filetime_int & 0xFFFFFFFF,
        filetime_int >> 32,
    )


def set_all_times_windows(path: Path, new_time: dt.datetime) -> Dict[str, object]:
    """
    Set CreationTime, LastAccessTime, and LastWriteTime on Windows using ctypes.
    This avoids external dependencies such as pywin32.
    """
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
        ft = datetime_to_filetime(new_time)
        ok = SetFileTime(handle, ctypes.byref(ft), ctypes.byref(ft), ctypes.byref(ft))
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())

        return {
            "creation_set": True,
            "modified_set": True,
            "accessed_set": True,
        }
    finally:
        CloseHandle(handle)


def set_all_times_cross_platform(path: Path, new_time: dt.datetime) -> Dict[str, object]:
    """
    Set timestamps. On Windows, CreationTime is also changed.
    On non-Windows, only modified/accessed can be changed reliably.
    """
    if is_windows():
        return set_all_times_windows(path, new_time)

    timestamp = dt_to_ts(new_time)
    os.utime(str(path), (timestamp, timestamp), follow_symlinks=False)
    return {
        "creation_set": False,
        "modified_set": True,
        "accessed_set": True,
    }


def iter_paths(root: Path, include_dirs: bool) -> Iterable[Path]:
    """
    Yield files first and directories later.
    Directory timestamps should be set after child files.
    """
    files = []
    dirs = []

    for current_root, dirnames, filenames in os.walk(root):
        current = Path(current_root)

        for filename in filenames:
            files.append(current / filename)

        if include_dirs:
            for dirname in dirnames:
                dirs.append(current / dirname)

    for item in files:
        yield item

    for item in sorted(dirs, key=lambda p: len(str(p)), reverse=True):
        yield item

    if include_dirs:
        yield root


def build_new_time_for_item(
    path: Path,
    info: TimestampInfo,
    mode: str,
    target_year: Optional[int],
    target_datetime: Optional[dt.datetime],
) -> Tuple[Optional[dt.datetime], str]:
    """
    Return the new timestamp and source/method description.
    """
    if mode == "exact":
        if target_datetime is None:
            raise ValueError("Mode exact requires target_datetime.")
        return target_datetime, "manual_exact_datetime"

    if mode == "year":
        if target_year is None:
            raise ValueError("Mode year requires target_year.")
        return safe_replace_year(info.modified, target_year), "manual_year_preserve_month_day_time"

    if mode == "path-year":
        detected_year = extract_year_from_path(path)
        if detected_year:
            return (
                safe_replace_year(info.modified, detected_year),
                f"path_year_{detected_year}_preserve_month_day_time",
            )

        if target_year is not None:
            return (
                safe_replace_year(info.modified, target_year),
                f"fallback_manual_year_{target_year}_preserve_month_day_time",
            )

        return None, "skipped_no_year_in_path"

    raise ValueError(f"Unsupported mode: {mode}")


def write_log_header(log_path: Path) -> None:
    if log_path.exists():
        return

    with log_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "path",
                "item_type",
                "status",
                "reason",
                "method",
                "old_creation_time",
                "old_modified_time",
                "old_accessed_time",
                "new_time",
                "creation_set",
                "modified_set",
                "accessed_set",
                "error",
            ]
        )


def append_log(log_path: Path, row: Dict[str, object]) -> None:
    with log_path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                row.get("path", ""),
                row.get("item_type", ""),
                row.get("status", ""),
                row.get("reason", ""),
                row.get("method", ""),
                row.get("old_creation_time", ""),
                row.get("old_modified_time", ""),
                row.get("old_accessed_time", ""),
                row.get("new_time", ""),
                row.get("creation_set", ""),
                row.get("modified_set", ""),
                row.get("accessed_set", ""),
                row.get("error", ""),
            ]
        )


def prompt_text(label: str, default: str = "") -> str:
    if default:
        value = input(f"{label} [{default}]: ").strip()
        return value or default
    return input(f"{label}: ").strip()


def prompt_yes_no(label: str, default: bool) -> bool:
    default_text = "Y/n" if default else "y/N"
    value = input(f"{label} [{default_text}]: ").strip().lower()

    if not value:
        return default

    return value in {"y", "yes", "1", "true"}


def prompt_int(label: str, default: Optional[int] = None) -> Optional[int]:
    while True:
        default_text = "" if default is None else f" [{default}]"
        value = input(f"{label}{default_text}: ").strip()

        if not value and default is not None:
            return default

        if not value and default is None:
            return None

        try:
            return int(value)
        except Exception:
            print("Please enter a valid number.")


def fill_interactive_args(args: argparse.Namespace) -> argparse.Namespace:
    print("=" * 80)
    print("PharmSolu Timestamp Setter - Interactive Mode")
    print("=" * 80)

    args.root = prompt_text("Enter root folder", args.root or DEFAULT_ROOT)

    print("")
    print("Choose timestamp mode:")
    print("  1 - Manual year; preserve existing month/day/time")
    print("      Example: 2029-05-14 14:24:00 -> 2015-05-14 14:24:00")
    print("  2 - Exact date/time for every selected item")
    print("      Example: 2015-01-01 00:00:00")
    print("  3 - Detect year from path; optional fallback year")
    print("      Example path contains: CHINESE PHARMACOPOEIA 2015 CMSP")
    choice = prompt_text("Enter choice", "1")

    if choice == "1":
        args.mode = "year"
        args.year = prompt_int("Enter target year", args.year)
    elif choice == "2":
        args.mode = "exact"
        while True:
            text = prompt_text("Enter target date/time", args.datetime or "2015-01-01 00:00:00")
            try:
                args.target_datetime_obj = parse_datetime_text(text)
                args.datetime = text
                break
            except Exception as exc:
                print(str(exc))
    elif choice == "3":
        args.mode = "path-year"
        fallback = prompt_int("Fallback year if no year is found in path; press Enter to skip files without path year", args.year)
        args.year = fallback
    else:
        print("Unknown choice. Using manual year mode.")
        args.mode = "year"
        args.year = prompt_int("Enter target year", args.year)

    args.suspicious_only = prompt_yes_no(
        "Process suspicious timestamps only",
        True if args.suspicious_only is None else args.suspicious_only,
    )

    args.include_dirs = prompt_yes_no("Also process folders", args.include_dirs)
    args.apply = prompt_yes_no("Apply changes now", args.apply)

    limit = prompt_int("Limit number of processed suspicious/selected items; 0 means no limit", args.limit)
    args.limit = 0 if limit is None else limit

    return args


def validate_args(args: argparse.Namespace) -> None:
    if not args.root:
        raise ValueError("Root folder is required.")

    root = Path(args.root)
    if not root.exists():
        raise ValueError(f"Root folder does not exist: {root}")

    if args.suspicious_only is None:
        args.suspicious_only = True

    current_year = now_local().year

    if args.mode == "year":
        if args.year is None:
            raise ValueError("--year is required when --mode year is used.")
        if not (1900 <= args.year <= current_year):
            raise ValueError(f"--year must be between 1900 and {current_year}.")

    if args.mode == "path-year":
        if args.year is not None and not (1900 <= args.year <= current_year):
            raise ValueError(f"Fallback --year must be between 1900 and {current_year}.")

    if args.mode == "exact":
        if not args.datetime and not getattr(args, "target_datetime_obj", None):
            raise ValueError("--datetime is required when --mode exact is used.")

        if not getattr(args, "target_datetime_obj", None):
            args.target_datetime_obj = parse_datetime_text(args.datetime)

        if args.target_datetime_obj > now_local() + dt.timedelta(days=args.future_margin_days):
            raise ValueError(
                "Target datetime is in the future. "
                "Use a past/current datetime for timestamp repair."
            )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Set suspicious file/folder timestamps under a selected path."
    )

    parser.add_argument("--interactive", action="store_true", help="Ask for input values interactively.")
    parser.add_argument("--root", default="", help="Root folder to scan. Example: D:\\PharmSolu")
    parser.add_argument(
        "--mode",
        choices=["year", "exact", "path-year"],
        default="year",
        help="How to calculate the new timestamp.",
    )
    parser.add_argument("--year", type=int, default=None, help="Target year or fallback year.")
    parser.add_argument(
        "--datetime",
        default="",
        help='Exact target datetime. Example: "2015-01-01 00:00:00"',
    )

    scope = parser.add_mutually_exclusive_group()
    scope.add_argument(
        "--suspicious-only",
        dest="suspicious_only",
        action="store_true",
        default=None,
        help="Only process files/folders with future or too-old timestamps.",
    )
    scope.add_argument(
        "--all-items",
        dest="suspicious_only",
        action="store_false",
        help="Process all files/folders under the root.",
    )

    parser.add_argument("--include-dirs", action="store_true", help="Also process folder timestamps.")
    parser.add_argument("--apply", action="store_true", help="Actually change timestamps. Default is dry-run.")
    parser.add_argument("--min-year", type=int, default=1990, help="Treat timestamps before this year as suspicious.")
    parser.add_argument(
        "--future-margin-days",
        type=int,
        default=1,
        help="Treat timestamps later than now plus this many days as suspicious.",
    )
    parser.add_argument("--log-dir", default=DEFAULT_LOG_DIR, help="Folder for CSV logs.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of selected items. 0 means no limit.")

    args = parser.parse_args()

    if len(sys.argv) == 1:
        args.interactive = True

    if args.interactive:
        args = fill_interactive_args(args)

    return args


def main() -> int:
    set_console_title(APP_TITLE)

    try:
        args = parse_args()
        validate_args(args)
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 2

    root = Path(args.root)
    target_datetime = getattr(args, "target_datetime_obj", None)
    if args.mode == "exact" and target_datetime is None:
        target_datetime = parse_datetime_text(args.datetime)

    log_dir = Path(args.log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    mode_name = "apply" if args.apply else "dry_run"
    log_path = log_dir / f"{APP_TITLE}_{mode_name}_{stamp}.csv"
    error_log_path = log_dir / f"{APP_TITLE}_errors_{stamp}.txt"
    write_log_header(log_path)

    print("=" * 80)
    print("PharmSolu Timestamp Setter")
    print("=" * 80)
    print(f"Root folder          : {root}")
    print(f"Mode                 : {args.mode}")
    print(f"Target year          : {args.year if args.year is not None else ''}")
    print(f"Target datetime      : {safe_iso(target_datetime) if target_datetime else ''}")
    print(f"Scope                : {'suspicious only' if args.suspicious_only else 'all items'}")
    print(f"Include directories  : {args.include_dirs}")
    print(f"Minimum year         : {args.min_year}")
    print(f"Future margin days   : {args.future_margin_days}")
    print(f"Action               : {'APPLY CHANGES' if args.apply else 'DRY-RUN ONLY'}")
    print(f"CSV log              : {log_path}")
    print("=" * 80)

    scanned = 0
    selected = 0
    changed = 0
    skipped = 0
    errors = 0

    for path in iter_paths(root, include_dirs=args.include_dirs):
        scanned += 1

        try:
            info = get_timestamp_info(path)
            is_suspicious, reason = suspicious_reasons(
                info,
                min_year=args.min_year,
                future_margin_days=args.future_margin_days,
            )

            if args.suspicious_only and not is_suspicious:
                continue

            selected += 1
            if not reason:
                reason = "all_items_mode"

            new_time, method = build_new_time_for_item(
                path=path,
                info=info,
                mode=args.mode,
                target_year=args.year,
                target_datetime=target_datetime,
            )

            if new_time is None:
                skipped += 1
                append_log(
                    log_path,
                    {
                        "path": str(path),
                        "item_type": "dir" if path.is_dir() else "file",
                        "status": "skipped",
                        "reason": reason,
                        "method": method,
                        "old_creation_time": safe_iso(info.created),
                        "old_modified_time": safe_iso(info.modified),
                        "old_accessed_time": safe_iso(info.accessed),
                        "new_time": "",
                        "creation_set": False,
                        "modified_set": False,
                        "accessed_set": False,
                        "error": "",
                    },
                )
                print(f"[SKIP] {path} | {method}")
                continue

            result = {
                "creation_set": False,
                "modified_set": False,
                "accessed_set": False,
            }

            if args.apply:
                result = set_all_times_cross_platform(path, new_time)
                if result.get("creation_set") or result.get("modified_set") or result.get("accessed_set"):
                    changed += 1

            append_log(
                log_path,
                {
                    "path": str(path),
                    "item_type": "dir" if path.is_dir() else "file",
                    "status": "changed" if args.apply else "dry_run",
                    "reason": reason,
                    "method": method,
                    "old_creation_time": safe_iso(info.created),
                    "old_modified_time": safe_iso(info.modified),
                    "old_accessed_time": safe_iso(info.accessed),
                    "new_time": safe_iso(new_time),
                    "creation_set": result.get("creation_set", False),
                    "modified_set": result.get("modified_set", False),
                    "accessed_set": result.get("accessed_set", False),
                    "error": "",
                },
            )

            print(
                f"[{'APPLY' if args.apply else 'DRY'}] {path} | "
                f"reason={reason} | new={safe_iso(new_time)} | method={method}"
            )

            if args.limit and selected >= args.limit:
                print(f"[INFO] Limit reached: {args.limit}")
                break

        except KeyboardInterrupt:
            print("\n[STOPPED] User interrupted.")
            break

        except Exception as exc:
            errors += 1
            with error_log_path.open("a", encoding="utf-8") as f:
                f.write(f"\n--- ERROR: {path} ---\n")
                f.write(traceback.format_exc())

            append_log(
                log_path,
                {
                    "path": str(path),
                    "item_type": "unknown",
                    "status": "error",
                    "reason": "",
                    "method": "",
                    "old_creation_time": "",
                    "old_modified_time": "",
                    "old_accessed_time": "",
                    "new_time": "",
                    "creation_set": False,
                    "modified_set": False,
                    "accessed_set": False,
                    "error": str(exc),
                },
            )

    print("=" * 80)
    print("Summary")
    print("=" * 80)
    print(f"Scanned items        : {scanned}")
    print(f"Selected items       : {selected}")
    print(f"Changed items        : {changed}")
    print(f"Skipped items        : {skipped}")
    print(f"Errors               : {errors}")
    print(f"CSV log              : {log_path}")
    if errors:
        print(f"Error log            : {error_log_path}")

    if not args.apply:
        print("")
        print("No timestamps were changed because this was DRY-RUN mode.")
        print("Review the CSV log first. Then rerun and choose Apply = Y when ready.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
