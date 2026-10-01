# pptx_translator — translate PowerPoint files without breaking the formatting

This is a small, safe tool that helps you translate a PowerPoint (`.pptx`) file\
while keeping **everything else exactly the same** — fonts, colours, images,\
layout, animations, charts. It does not use the internet and it does not change\
your original file.

It works in **two simple steps**:

1. **Extract** — the tool pulls every piece of text out of your presentation\
  into one text file (a `.json` file).

2. **Patch** — after the text has been translated, the tool puts the\
  translations back and saves a **new** translated PowerPoint.

In between those two steps you (or an AI like ChatGPT) translate the text. The\
tool itself does not translate — it just takes the text out and puts it back\
perfectly.

```
 your.pptx  ──extract──▶  your.json  ──(you translate the text)──▶  your.json
                                                                       │
 your_translated.pptx  ◀──────────────────── patch ───────────────────┘
```

---

## What you need first: Python

This tool runs on **Python 3** (version 3.9 or newer). Most Macs already have\
it. On Windows you usually need to install it once.

* **Check if you have it** — open a terminal (see below) and type\
  `python3 --version` (Mac/Linux) or `py --version` (Windows). If you see a\
  number like `Python 3.11.6`, you are good.

* **If not installed** — download it free from\
  [https://www.python.org/downloads/](https://www.python.org/downloads/).\
  **On Windows, during installation tick the box that says\
  "Add Python to PATH".** This one checkbox saves a lot of trouble.

**How to open a terminal:**

* **Windows:** press the Start button, type `cmd`, open **Command Prompt**.

* **Mac:** press `Cmd`+`Space`, type `Terminal`, press Enter.

* **Linux:** open your **Terminal** app.

---

## One-time setup (install)

Put this whole folder somewhere easy to find, for example your Desktop. Then:

### Mac / Linux

Open Terminal, then type `cd ` (with a space), drag the folder into the Terminal\
window, and press Enter. Then run:

```bash
bash install.sh
```

### Windows

Simply **double-click `install.bat`**.

Or, in Command Prompt, `cd` into the folder and run:

```bat
install.bat
```

The installer creates a private, self-contained Python environment inside a\
`.venv` folder here. It does not affect anything else on your computer. When it\
finishes it prints "All set!".

---

## Step 1 — Extract the text

This reads your presentation and writes a `.json` file containing all the text.\
Replace `MyDeck.pptx` with the name of your file (keep the quotes).

### Mac / Linux

```bash
./.venv/bin/python -m pptx_translator extract "MyDeck.pptx" -o "MyDeck.json"
```

### Windows

```bat
.venv\Scripts\python.exe -m pptx_translator extract "MyDeck.pptx" -o "MyDeck.json"
```

> Tip: if the `.pptx` file is not in this folder, use its full path, e.g.\
> `"C:\Users\You\Desktop\MyDeck.pptx"` on Windows or\
> `"/Users/you/Desktop/MyDeck.pptx"` on Mac.

You now have **`MyDeck.json`**.

---

## Step 2 — Translate the text

Open `MyDeck.json` in any plain-text editor (Notepad, TextEdit, VS Code). You\
will see a list of entries that look like this:

```json
{
  "id": "slide1_tb0_p0",
  "source": "Quarterly Report",
  "target": "",
  "source_hash": "9f2c...",
  "context": { "part_label": "Slide 1", "role": "title", "character_limit_hint": 16 },
  "status": "new"
}
```

**Your only job:** type the translation into each empty `"target": ""` field.

```json
  "source": "Quarterly Report",
  "target": "Informe Trimestral",
```

Rules that keep everything working:

* **Do not change** `id` or `source_hash`. The tool uses them to put text back\
  in the right place and to detect if the slide changed.

* Leave `target` **empty** for anything you want to keep in the original\
  language — it will simply be left untouched.

* For a line break inside one text box, type `\n` in the target.

* `character_limit_hint` tells you roughly how many characters the original used,\
  in case space is tight.

> **Using an AI to translate?** Just give it the whole `MyDeck.json` file and\
> ask it to "fill in every `target` field with the <your language> translation\
> and return the same JSON unchanged otherwise." Save what it gives you back.

---

## Step 3 — Patch (build the translated PowerPoint)

This creates a **new** file and leaves your original alone.

### Mac / Linux

```bash
./.venv/bin/python -m pptx_translator patch "MyDeck.pptx" "MyDeck.json" -o "MyDeck_translated.pptx" --report "MyDeck_qa.json"
```

### Windows

```bat
.venv\Scripts\python.exe -m pptx_translator patch "MyDeck.pptx" "MyDeck.json" -o "MyDeck_translated.pptx" --report "MyDeck_qa.json"
```

Open **`MyDeck_translated.pptx`** — it is your presentation, translated, with all\
formatting intact. **`MyDeck_qa.json`** is a short quality report (see below).

---

## Understanding the quality report

After patching, the tool prints a summary and writes a `..._qa.json` report:

| Term | Meaning |
| --- | --- |
| **translated** | Text successfully replaced with your translation. |
| **empty** | You left the `target` blank, so the original text was kept. |
| **stale** | The slide's text changed since you extracted it, so that translation was **skipped** for safety. Re-run extract, or add `--force`. |
| **missing** | A translation id no longer exists in the presentation (slide deleted/edited). |
| **left untranslated** | Text in the file that was not in your JSON at all. |

### The `--force` option

If you see **stale** items, it means the original `.pptx` was edited after you\
extracted the text. The safe fix is to run **extract** again on the current file.\
If you are sure you want to apply the translations anyway, add `--force` to the\
patch command.

---

## What gets translated

Everything with text, including:

* All slides — titles, bullet lists, text boxes, placeholders

* Tables (every cell)

* Grouped shapes

* **Speaker notes**

* Slide **layouts** and **masters**

* Notes master and handout master

Auto-generated fields like slide numbers and dates are left alone on purpose\
(PowerPoint regenerates them).

---

## Common problems

* **"File not found"** — check the file name and that you are in the right\
  folder. Drag-and-drop the file into the terminal to get its exact path.

* **"does not look like a .pptx file"** — the first file must be your\
  PowerPoint, the second must be the translated `.json`.

* **"not valid JSON"** — the `.json` file got broken while editing (usually a\
  missing quote or comma). Re-extract and try again, or fix the editor error it\
  points to.

* **"output file must be different from the input file"** — choose a new name\
  after `-o`. The tool never overwrites your original.

---

## For developers

* No `python-pptx` is used to read or write the package — patching is done with\
  `zipfile` + `lxml` directly on the OOXML parts, so every unmodified byte of\
  the original ZIP is copied through unchanged.

* Translation units are **paragraph-level**, with stable IDs tied to the OOXML\
  location (`<part><index>_tb<textbody>_p<paragraph>`).

* Run the test suite:

  ```bash
  ./.venv/bin/python -m pip install python-pptx   # test-only dependency
  ./.venv/bin/python -m pytest tests/ -v
  # or, without pytest:
  ./.venv/bin/python tests/test_roundtrip.py
  ```

* Install as a command (optional): `pip install .` then use `pptx-translator ...`\
  instead of `python -m pptx_translator ...`.

### Project layout

```
pptx_translate/
├── pptx_translator/
│   ├── __init__.py
│   ├── __main__.py      # enables `python -m pptx_translator`
│   ├── cli.py           # command line interface + friendly errors
│   ├── extract.py       # package walker + extraction
│   ├── patch.py         # surgical OOXML patching + QA report
│   └── ooxml_utils.py   # namespaces & shared XML helpers
├── tests/test_roundtrip.py
├── requirements.txt
├── pyproject.toml
├── install.sh / install.bat
└── README.md
```

## License

MIT