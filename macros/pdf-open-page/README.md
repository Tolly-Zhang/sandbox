# PDF Open Page

[![Download ZIP](https://img.shields.io/badge/Download-ZIP-blue?style=for-the-badge&logo=github)](https://download-directory.github.io/?url=https://github.com/Tolly-Zhang/sandbox/tree/main/macros/pdf-open-page)

This macro opens a PDF directly to a requested page. You can launch from a file picker or from preconfigured PDF indexes in `config.json`, and it can resolve logical page labels (such as roman-numerals) when available.

## Prerequisites
- Windows 10/11
- PowerShell 5.1 or newer
- A PDF viewer (SumatraPDF, Adobe Acrobat Reader, or PDF-XChange)

### Dependencies
1. `pdftk` (optional): enables logical page label support. Without `pdftk`, the script still works with physical page numbers.

## Setup
1. Install a supported PDF viewer and (optionally) `pdftk`.

2. Open `config.json` and configure `pdfFiles`, `preferredViewer`, and `keepWindowOpen`.

3. Create custom launcher files from `run-template.vbs`. For each:
   - Replace `INDEX_VALUE` with a index starting from 1 that matches `pdfFiles`. For example, `-PdfIndex 1` opens the first file in `pdfFiles`.

4. Test by double-clicking `run-pdf-1.vbs` or running `run.ps1` directly. You should be prompted for a page number, and the PDF should open to that page.

## Configuration

| Parameter             | Type       | Usage                                                                                             |
| :-------------------- | :--------- | :------------------------------------------------------------------------------------------------ |
| `keepWindowOpen`      | `bool`     | `true` keeps the shell open after every run; `false` closes on success and pauses on errors only. |
| `pdfFiles`            | `string[]` | List of absolute PDF paths used with indexed launch (`-PdfIndex`).                                |
| `preferredViewer`     | `string`   | Preferred viewer key from `viewers` (for example `PDFXChange`).                                   |
| `viewerPath`          | `string`   | Direct path to a viewer executable. Used if `preferredViewer` cannot be resolved.                 |
| `viewerArgsTemplate`  | `string`   | Args template for `viewerPath`. Supports `{page}` and `{file}` placeholders.                      |
| `viewers.<name>.path` | `string`   | Executable path for each named viewer entry.                                                      |
| `viewers.<name>.args` | `string`   | Args template for that viewer. Supports `{page}` and `{file}` placeholders.                       |
| `pdfTkPath`           | `string`   | Optional path or command name for `pdftk` used to resolve logical page labels.                    |
| `tools.pdfTkPath`     | `string`   | Optional alternative location for `pdftk`.                                       |

## Usage

Double-click one of the `.vbs` launcher files, or run PowerShell directly.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -PdfIndex 1
```

You can also create shortcuts to `.vbs` launchers for easier access.

### Runtime Behavior
- PDF selection:
  - If `-PdfIndex` is provided and greater than `0`, the script opens `pdfFiles[PdfIndex - 1]`.
  - Otherwise, it shows a file picker dialog.
- Page input:
  - Enter page value in the terminal prompt.
  - If logical labels are available, you can enter either label or physical number.
- Viewer resolution order:
    1. `preferredViewer`
    2. `viewerPath`
    3. First valid entry in `viewers`
    4. Built-in fallback install locations