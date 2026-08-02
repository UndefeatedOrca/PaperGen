# PaperGen — Setup Guide

Generate pre-filled academic paper files from your templates with a right-click.

---

## Requirements

- Windows 10 or 11
- Python 3.10+ installed and on your PATH
  - Download from https://python.org — check "Add Python to PATH" during install

---

## Installation (5 steps)

### 1. Install the Python dependency

Open a terminal (Win+R → `cmd`) and run:

```
pip install python-docx
```

### 2. Create the PaperGen folder

Create this folder (it likely already exists if you ran pip):

```
%APPDATA%\PaperGen\
```

Tip: paste `%APPDATA%\PaperGen` directly into File Explorer's address bar.

### 3. Copy files into place

Copy `papergen.py` into:
```
%APPDATA%\PaperGen\papergen.py
```

### 4. Add your templates

Place your `.docx` template files in:
```
%APPDATA%\PaperGen\templates\
```

**Using the sample templates (optional):**
Run `make_templates.py` once to generate starter `mla.docx` and `turabian.docx` files:
```
python make_templates.py
```

**Using your own existing templates:**
Open each `.docx` in Word and replace the blank fields with these placeholders exactly as written:

| Field | Placeholder |
|---|---|
| Your name | `{{your_name}}` |
| Professor's name | `{{professor_name}}` |
| Class name | `{{class_name}}` |
| Class number | `{{class_number}}` |
| School / college | `{{school}}` |
| Paper title | `{{paper_title}}` |
| Due date | `{{due_date}}` |

Save the file as e.g. `mla.docx` or `turabian.docx` in the templates folder.
The filename (without `.docx`) is what appears in the Style dropdown.

### 5. Install the right-click menu

**Important:** Before running, open `install_context_menu.reg` in Notepad and verify the
path in the `[command]` line matches where you put `papergen.py`. It should be:

```
C:\Users\YourActualUsername\AppData\Roaming\PaperGen\papergen.py
```

Then double-click `install_context_menu.reg` and click Yes when prompted.

---

## Usage

### New Paper
Right-click any folder background (not on a file) → **New Paper from Template**

1. Select your class from the dropdown
2. Select your formatting style (MLA, Turabian, etc.)
3. Enter the paper title
4. Confirm or change the due date
5. Click **Generate** — the `.docx` appears in that folder

### Manage Classes
Either click **Manage Classes** from the New Paper window, or run:

```
python papergen.py --manage
```

From there you can:
- **Add** a class (fill in the form)
- **Edit** an existing class
- **Delete** a single class
- **Clear Semester** — removes all classes at once (useful at semester end)

---

## File Locations

| File | Location |
|---|---|
| Main script | `%APPDATA%\PaperGen\papergen.py` |
| Class config | `%APPDATA%\PaperGen\classes.json` |
| Templates | `%APPDATA%\PaperGen\templates\*.docx` |

You can back up or edit `classes.json` directly in any text editor — it's plain JSON.

---

## Uninstalling the Right-Click Menu

Open Registry Editor (`regedit`), navigate to:

```
HKEY_CLASSES_ROOT\Directory\Background\shell\
```

Delete the `PaperGen` key.

---

## Troubleshooting

**"No templates found"** — Make sure `.docx` files are in `%APPDATA%\PaperGen\templates\`

**Nothing happens on right-click** — Make sure `pythonw.exe` is on your PATH.
Test by running `where pythonw` in a terminal. If not found, reinstall Python with PATH option checked.

**Placeholder not replaced** — Make sure the placeholder text is in a single run in Word.
If you typed it manually, delete it and retype `{{placeholder_name}}` fresh (don't paste from this doc).