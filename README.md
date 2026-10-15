# Books to PDF

A one-click Windows tool that converts the Word manuscripts of *The Guide* into PDFs.

## What it does

`Convert Word To PDF.ps1` looks at every `.docx` file in the source folder and, for each one:

1. **Skips it** if its PDF already exists and is newer than the Word file.
2. **Converts it** otherwise, using Microsoft Word in the background to export a PDF (with document bookmarks from headings, document properties and PDF/A-friendly settings).
3. **Shows live progress** in a console window: `[3/12] Rendering 'Chapter.docx'... (Elapsed: 42s) [ALIVE]`, then the size and time of each finished PDF.

Large books (500+ pages) can take minutes, so the script watches Word's CPU usage. If Word sits at 0% CPU for 45 seconds it is treated as hung: the Word process is killed and the run stops with an error.

Word's temporary lock files (`~$*.docx`) are ignored. The window stays open at the end so you can read the results; press Enter to close it.

## Folders

| Purpose | Path |
| --- | --- |
| Word source (`$wordFolderPath`) | `C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide\Editable format` |
| PDF output (`$pdfFolderPath`) | `C:\Users\iwan rahabok\Box\Private Cloud Management\The Guide` |

Edit the two variables at the top of the script to change them.

## Requirements

- Windows with Microsoft Word installed
- Windows PowerShell 5.1 (built in) or PowerShell 7

## Running it

- **Desktop icon:** double-click **Books to PDF** on the desktop.
- **Manually:**

  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Convert Word To PDF.ps1"
  ```

To recreate the desktop shortcut, run `.\Create Desktop Shortcut.ps1`.
