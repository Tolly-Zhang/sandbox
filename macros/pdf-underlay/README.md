# PDF Underlay

This macro applies a underlay to PDF files using a reference PDF. It can be used to add a background or template to existing PDFs. It works with both normal PDFs and scanned PDFs, using different processing flows for each.

## Prerequisites
- Windows 10/11
- PowerShell 5.1 or newer

### Dependencies
The script expects tool locations in `config.json` (or available on PATH where applicable).
| Dependency               | Use                                        |
| :----------------------- | :----------------------------------------- |
| `qpdf`                   | Normal mode underlay (`--underlay`)        |
| `pdftoppm` (Poppler)     | Converting PDF pages to PNG                |
| `pdfimages` (Poppler)    | Detecting source DPI                       |
| `ImageMagick` (`magick`) | Transparency processing and PDF conversion |
| `pdftk`                  | Page extraction, stamping, and final merge |

## Setup
1. Install `qpdf`, Poppler (`pdftoppm` and `pdfimages`), ImageMagick, and `pdftk`. Ensure they are in your system PATH or update their paths in `config.json`. 
   - `qpdf`:
        ```powershell
        winget install -e --id QPDF.QPDF
        ```
   - Poppler:
        ```powershell
        winget install -e --id oschwartz10612.Poppler
        ```
   - ImageMagick:
        ```powershell
        winget install -e --id ImageMagick.ImageMagick
        ```
   - `pdftk`:
        ```powershell
        winget install -e --id PDFLabs.PDFtk.Server
        ```
2. Update `QpdfPath`, `PdfToPpmPath`, `PdfImagesPath`, `MagickPath`, and `PdfTkPath` in `config.json` if your installations are not in the system PATH or if you want to specify custom locations.

3. Place the PDF you want to use as an underlay in the macro's folder and update `UnderlayFileName` in `config.json` if needed.

4. Test by double-clicking `run-normal.vbs` to run in normal mode or `run-scanned.vbs` for scanned PDFs. Follow the prompts to provide PDF paths.

## Configuration

| Parameter               | Type     | Usage                                                                                                           |
| :---------------------- | :------- | :-------------------------------------------------------------------------------------------------------------- |
| `Tools.QpdfPath`        | `string` | Path to the `qpdf` executable. If omitted, `qpdf` must be on `PATH`.                                            |
| `Tools.PdfToPpmPath`    | `string` | Command or path to `pdftoppm` (Poppler).                                                                        |
| `Tools.PdfImagesPath`   | `string` | Command or path to `pdfimages` (Poppler).                                                                       |
| `Tools.MagickPath`      | `string` | Command or path to ImageMagick (`magick`).                                                                      |
| `Tools.PdfTkPath`       | `string` | Command or path to `pdftk`.                                                                                     |
| `UnderlayFileName`      | `string` | Filename of the reference underlay PDF. Can be a relative or absolute path.                                     |
| `Defaults.Dpi`          | `int`    | Default rasterization DPI `(300 - 1200)` for scanned mode when DPI cannot be detected.                          |
| `Defaults.Fuzz`         | `int`    | ImageMagick fuzz percentage `(0 - 100)` used when converting white to transparent.                              |
| `Defaults.CreateBackup` | `bool`   | Whether to create a backup of the original file. Original files are copied to `*.bak` before being overwritten. |
| `Defaults.KeepTemp`     | `bool`   | Whether to keep temporary working folders for debugging.                                                        |

## Usage
Double-click `run-normal.vbs` to run in normal mode or `run-scanned.vbs` for scanned PDFs. You can also create a shortcuts to the .vbs files for easier access.