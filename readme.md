# PaperGen

Generates pre-filled academic paper files from templates via a right-click context menu in Windows Explorer. Supports Word (`.docx`), Typst (`.typ`), and LaTeX (`.tex`) templates.

---

## Files

```
%APPDATA%\PaperGen\
├── papergen.py               — main application
├── templategen.py            — generates starter templates (run once)
├── install_context_menu.reg  — adds right-click menu entry to Explorer
├── profile.json              — your name and university (stable across classes)
├── classes.json              — your saved class profiles
├── config.json               — local filename configuration
└── templates\
    ├── mla.docx              — (example)
    ├── turabian.typ          — (example)
    └── mla.tex               — (example)
```

---

## Requirements

- Windows 10 or 11
- Python 3.10+ — download from https://python.org (check "Add Python to PATH")
- `python-docx` — install with:
  ```
  pip install python-docx
  ```
- For `.typ` templates: [Typst](https://typst.app) CLI or app
- For `.tex` templates: TeX Live or MiKTeX

---

## Setup

### 1. Place files

Copy `papergen.py` and `templategen.py` into `%APPDATA%\PaperGen\`.

Tip: paste `%APPDATA%\PaperGen` into the File Explorer address bar to navigate there.

### 2. Add templates

Place your `.docx`, `.typ`, or `.tex` template files in `%APPDATA%\PaperGen\templates\`.

**Option A — Generate starter templates:**
```
python templategen.py
```
This creates `mla.docx`, `turabian.docx`, `mla.typ`, `turabian.typ`, `turabian.tex`, and `mla.tex` in the templates folder.

**Option B — Use your own existing templates:**
Open each file and replace blank header fields with the placeholders listed below.

### 3. Install the right-click menu

Open `install_context_menu.reg` in Notepad and confirm the path on this line points to where you placed `papergen.py`:

```
@="C:\\Python314\\pythonw.exe \"C:\\Users\\YourName\\AppData\\Roaming\\PaperGen\\papergen.py\" \"%V\""
```

Update the Python path if needed — find yours by running:
```
python -c "import sys; print(sys.executable)"
```
Replace `python.exe` with `pythonw.exe` in that path.

Double-click `install_context_menu.reg` and confirm the prompt.

### 4. Set up your profile

Open Configuration (see Usage below) and click **Edit Profile** to enter your first name, last name, and university. You only do this once — it pre-fills every new class you add.

Example profile and class data are provided in `profile.example.json` and `classes.example.json`. Copy them to `profile.json` and `classes.json` if you want to try the application with sample data. The runtime JSON files are local user data and are excluded from version control.

To customize generated filenames, copy `config.example.json` to `config.json` and edit `filename_format`. The default is `{{last name}}{{title}}`, which produces names such as `SmithOnFreeWill.typ`. Supported tokens are `{{last name}}`, `{{first name}}`, `{{class code}}`, `{{class number}}`, and `{{title}}`; title uses the existing punctuation and space stripping logic. The template extension is appended automatically.

---

## Usage

### New Paper

Right-click any folder background (not on a file) → **New Paper from Template**

1. Select a class from the dropdown
2. Select a template/style (shows all `.docx`, `.typ`, and `.tex` files in your templates folder)
3. Enter the paper title
4. Confirm or change the due date
5. Click **Generate**

The output file appears in the folder you right-clicked. By default, the filename format is `LastTitleNoSpaces` with the same extension as the template (e.g. `SmithOnFreeWill.typ`). See `config.example.json` to customize the filename format.

### Configuration

Click **Configuration** from the New Paper window, or run:
```
python papergen.py --config
```

The older `--manage` option remains supported.

- **Add Class** — opens a form pre-filled with your profile name and university; fill in the rest
- **Edit** — modify a saved class
- **Delete** — remove a single class
- **Edit Profile** — update your name or university (affects pre-fill for new classes, and the `{{name_first}}` / `{{name_last}}` placeholders)
- **Clear Semester** — removes all classes at once
- **Filename Convention** — set the filename format using `{{last name}}`, `{{first name}}`, `{{class code}}`, `{{class number}}`, and `{{title}}`

---

## Placeholders

Place these in your templates where you want fields filled in. They work in `.docx`, `.typ`, and `.tex` files.

| Placeholder | Source | Example |
|---|---|---|
| `{{your_name}}` | class (pre-filled from profile) | `Patrick Smith` |
| `{{name_first}}` | profile | `Patrick` |
| `{{name_last}}` | profile | `Smith` |
| `{{professor_name}}` | class | `Dr. Jones` |
| `{{class_name}}` | class | `Modern History` |
| `{{class_code}}` | class | `HIST` |
| `{{class_number}}` | class | `220` |
| `{{class_full}}` | auto-composed | `HIST 220` |
| `{{school}}` | class | `College of Arts and Sciences` |
| `{{university}}` | class, falls back to profile | `State University` |
| `{{paper_title}}` | entered at generation time | `The Fall of Rome` |
| `{{due_date}}` | entered at generation time | `September 1, 2026` |

**Note for Word templates:** If you type placeholders directly in Word, it may split them across internal formatting runs and prevent replacement. If a field is not being replaced, delete it and retype it fresh in a single unformatted run. You can verify how Word stored it by running:
```
python -c "
from docx import Document
doc = Document(r'%APPDATA%\PaperGen\templates\your-template.docx')
for p in doc.paragraphs:
    for r in p.runs:
        if r.text.strip(): print(repr(r.text))
"
```
Each placeholder should appear as a single string like `'{{professor_name}}'`, not split across multiple entries.

---

## Typst Templates

Generated `.typ` files use `{{placeholders}}` in the header and are ready to write in immediately. Compile with:
```
typst compile filename.typ
```

The starter MLA Typst template (`mla.typ`) uses the `modern-mla` package from Typst Universe. Install it with:
```
typst init @preview/modern-mla
```
or use the [Typst web app](https://typst.app).

---

## LaTeX Templates

Generated `.tex` files use `{{placeholders}}` in the preamble. Compile with:
```
pdflatex filename.tex
```
Run twice for correct page numbers.

The Turabian LaTeX template (`turabian.tex`) uses the `turabian-formatting` package, which is included in TeX Live and MiKTeX. The MLA template (`mla.tex`) uses only standard packages (`geometry`, `setspace`, `fancyhdr`).

---

## Semester Turnover

At the end of each semester, open Configuration and click **Clear Semester** to remove all classes at once. Your profile (name, university) is preserved.

---

## Uninstalling the Right-Click Menu

Open Registry Editor (`regedit`), navigate to:
```
HKEY_CLASSES_ROOT\Directory\Background\shell\
```
Delete the `PaperGen` key.

---

## Troubleshooting

**Nothing happens on right-click** — verify `pythonw.exe` exists at the path in the `.reg` file. Find the correct path with `python -c "import sys; print(sys.executable)"` and replace `python.exe` with `pythonw.exe`.

**"No templates found"** — confirm `.docx`, `.typ`, or `.tex` files are in `%APPDATA%\PaperGen\templates\`.

**Placeholder not replaced in Word** — see the Note under Placeholders above.

**Typst package not found** — run `typst init @preview/modern-mla` in the folder containing your `.typ` file, or open the file in the Typst web app.

**LaTeX package not found** — run `tlmgr install turabian-formatting` (TeX Live) or use the MiKTeX package manager to install `turabian-formatting`.

---

## License

PaperGen is licensed under the [MIT License](LICENSE).
