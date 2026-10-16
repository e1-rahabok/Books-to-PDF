# Books to PDF

A one-click Windows tool that converts the Word manuscripts of *The Guide* into PDFs and uploads them to Vercel Blob storage.

## What it does

`Convert Word To PDF.ps1` looks at every `.docx` file in the source folder and, for each one:

1. **Skips it** if its PDF already exists and is newer than the Word file.
2. **Converts it** otherwise, using Microsoft Word in the background to export a PDF (with document bookmarks from headings, document properties and PDF/A-friendly settings).
3. **Shows live progress** in a console window: `[3/12] Rendering 'Chapter.docx'... (Elapsed: 42s) [ALIVE]`, then the size and time of each finished PDF.

Large books (500+ pages) can take minutes, so the script watches Word's CPU usage. If Word sits at 0% CPU for 45 seconds it is treated as hung: the Word process is killed and the run stops with an error.

Word's temporary lock files (`~$*.docx`) are ignored. The window stays open at the end so you can read the results; press Enter to close it.

## Upload to Vercel Blob

After converting, the script runs `Upload PDFs To Vercel.ps1`, which uploads each PDF to the `isvara-media` Vercel Blob store as `downloads/<name>.pdf`:

- A PDF is uploaded only if it isn't in the store yet, or if the local file is newer than the copy in the store. So a failed upload is retried on the next run.
- An upload replaces the existing file in place, and the URL stays the same. The CDN cache is set to 60 seconds, so readers get the new version within a minute.
- You can run `Upload PDFs To Vercel.ps1` on its own to upload without converting.

### Credential setup (one time)

The upload needs the store's read-write token. **The token is never stored in this repository.** To save it:

1. In Vercel, open **Storage → isvara-media → .env.local** and copy the value of `BLOB_READ_WRITE_TOKEN`.
2. Run `.\Set Vercel Token.ps1` and paste the token when asked (input is hidden).

The token is encrypted with Windows DPAPI and saved to `%USERPROFILE%\.secrets\vercel-blob-isvara-media.xml`. Only your Windows account on this laptop can decrypt it. If the file is missing, the converter still runs and only skips the upload.

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
