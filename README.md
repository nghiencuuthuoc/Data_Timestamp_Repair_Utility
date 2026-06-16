# Office Metadata Timestamp GUI v3

A Windows-friendly Python/Tkinter tool for repairing corrupted future-dated file timestamps by reading internal Office/PDF document metadata and applying the correct document creation date back to the file system.

This project was created for large document archives such as `D:\PharmSolu`, where files may show incorrect future dates in Windows Explorer, for example `2029`, `2062`, or other invalid timestamps.

---

## Key Features

- Paste one or many file paths directly into the GUI.
- Add files or folders using file/folder picker buttons.
- Recursively scan folders such as `D:\PharmSolu`.
- Filter and scan only files whose file-system date is greater than the current date.
- Read internal metadata from Office and PDF files.
- Prefer the document **Content created** date.
- Reject corrupted future metadata dates.
- Use filename/path date fallback, such as `17.05.2016`.
- Set Windows timestamps:
  - `CreationTime`
  - `ModifiedTime / LastWriteTime`
  - `AccessedTime`
- Save CSV logs for review and audit.
- No `pywin32` dependency required for timestamp writing; Windows API is called through `ctypes`.

---

## Why This Tool Is Needed

Some old file archives contain files whose Windows timestamps are incorrect because of backup, copy, restore, synchronization, or system clock errors.

Example:

```text
Windows file timestamp : 2062-05-16 22:57
Office Content created : 2016-05-11 12:43
```

In this case, the correct repair target should be:

```text
2016-05-11 12:43
```

The tool scans metadata first, shows the target date in a table, and only applies changes after user confirmation.

---

## Supported File Types

### Modern Office Files

These formats are read using Python standard libraries only:

```text
.docx .docm .dotx .dotm
.xlsx .xlsm .xltx .xltm
.pptx .pptm .potx .potm
```

The tool reads:

```text
docProps/core.xml
```

Important metadata fields:

```text
created
modified
creator
lastModifiedBy
```

### Legacy Microsoft Office Files

These formats require the optional `olefile` package:

```text
.doc .xls .ppt
```

The tool reads:

```text
create_time
last_saved_time
author
last_saved_by
```

### PDF Files

The tool scans PDF metadata fields:

```text
/CreationDate
/ModDate
```

### Filename or Path Date Fallback

If internal metadata is missing, the tool can extract dates from filenames or paths.

Supported examples:

```text
17.05.2016
17-05-2016
17_05_2016
2016-05-17
2016.05.17
```

---

## Recommended Settings for PharmSolu Repair

For large archives such as `D:\PharmSolu`, use the following settings:

```text
Scan folders recursively: ON
Only supported extensions: ON

Scan only files with filesystem date > current date: ON
Compare by date only, not exact time: ON
Show skipped non-future files: OFF

Prefer Content created date: ON
Allow Date last saved fallback: OFF
Reject future metadata dates: ON
Use filename/path date fallback: ON

CreationTime: ON
ModifiedTime / LastWriteTime: ON
AccessedTime: ON
```

Important:

Keep **Allow Date last saved fallback** turned **OFF** by default because some files may have corrupted future values in the `Date last saved` field.

---

## Project Structure

```text
office_metadata_timestamp_gui_v3/
├── office_metadata_timestamp_gui_v3.py
├── run_office_metadata_timestamp_gui_v3.bat
├── requirements.txt
├── README.md
├── readme__office_metadata_timestamp_gui_v3.md
└── context_pack__office_metadata_timestamp_gui_v3.md
```

---

## Requirements

Required:

```text
Python 3.9+
Tkinter
```

Tkinter is normally included with standard Python installations on Windows.

Optional but recommended:

```text
olefile>=0.47
```

Install dependencies:

```bat
pip install -r requirements.txt
```

Or install `olefile` manually:

```bat
pip install olefile
```

---

## How to Run

### Option 1: Run with BAT Launcher

Double-click:

```bat
run_office_metadata_timestamp_gui_v3.bat
```

The BAT launcher searches for portable Python 3.12 in the current and parent folders:

```text
Python312\python.exe
..\Python312\python.exe
..\..\Python312\python.exe
..\..\..\Python312\python.exe
```

### Option 2: Run with Python Directly

```bat
python office_metadata_timestamp_gui_v3.py
```

---

## Example Input

Paste one file or many files into the GUI:

```text
"D:\Doc1.docx"
```

You can also paste a folder path:

```text
"D:\PharmSolu"
```

Then click:

```text
Scan metadata
```

After checking the results, click:

```text
Apply timestamps
```

---

## Logs

The tool automatically creates CSV logs in:

```text
office_timestamp_logs_v3
```

The logs include:

```text
path
status
message
metadata_source
target_source
future_fs_fields
fs_time_error
fs_creation_time
fs_modified_time
fs_accessed_time
metadata_created
metadata_modified
author
last_saved_by
target_time
applied
creation_set
modified_set
accessed_set
```

Use these logs to verify which files were scanned and which timestamps were applied.

---

## Safety Workflow

Recommended workflow:

1. Paste file or folder paths.
2. Click **Scan metadata**.
3. Review the table.
4. Open the CSV log and verify target dates.
5. Click **Apply timestamps** only after confirming the results.

The tool never modifies files during scan. Timestamp changes only happen after the user clicks **Apply timestamps**.

---

## GitHub Upload Commands

Use these commands to upload the project to GitHub.

### 1. Open project folder

```bat
cd /d D:\path\to\office_metadata_timestamp_gui_v3
```

### 2. Initialize Git

```bat
git init
```

### 3. Add all files

```bat
git add .
```

### 4. Commit

```bat
git commit -m "Initial release: Office Metadata Timestamp GUI v3"
```

### 5. Rename default branch

```bat
git branch -M main
```

### 6. Add GitHub remote

Replace `YOUR_USERNAME` and `YOUR_REPOSITORY` with your GitHub account and repository name.

```bat
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
```

### 7. Push to GitHub

```bat
git push -u origin main
```

---

## Update Existing GitHub Repository

After editing files, use:

```bat
git status
git add .
git commit -m "Update Office Metadata Timestamp GUI"
git push
```

---

## Suggested `.gitignore`

Create a `.gitignore` file with:

```gitignore
__pycache__/
*.pyc
*.pyo
*.pyd
*.log
*.tmp
.DS_Store
Thumbs.db

office_timestamp_logs_v3/
timestamp_setter_logs/
timestamp_repair_logs/

.venv/
venv/
env/
```

---

## Notes for Developers

### Timestamp Writing

On Windows, the tool uses:

```text
kernel32.CreateFileW
kernel32.SetFileTime
```

through Python `ctypes`.

This allows setting:

```text
CreationTime
LastWriteTime
LastAccessTime
```

without requiring `pywin32`.

### Long Path Support

The app converts paths to long-path-safe Windows format:

```text
\\?\C:\path\to\file
```

and UNC paths to:

```text
\\?\UNC\server\share\path
```

### Metadata Priority

Default metadata priority:

```text
1. Office/PDF Content created date
2. Filename/path date fallback
3. Date last saved only if manually enabled
```

---

## Disclaimer

This tool changes file-system timestamps. Always scan and review the CSV log before applying changes to large folders. Make a backup when working with important archives.

---

## Author / Project

PharmSolu Timestamp Repair Utility  
NGHIEN CUU THUOC // RnD PHARMA PLUS  
Website: https://www.nghiencuuthuoc.com  
PharmApp Demo: https://www.pharmapp.dev
