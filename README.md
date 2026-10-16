# Word Clone — Independent Microsoft Word Alternative

An independently implemented, full-featured document processor modeled after **Microsoft Word (`WINWORD.EXE` v16.0.20430.20146)**, built from evidence gathered via **REA (Reverse Engineer Anything)** static binary/resource analysis and live Windows UI Automation / COM observation.

- **Repository**: [https://github.com/dodoloiu1/word-clone](https://github.com/dodoloiu1/word-clone)
- **Live Web Application (GitHub Pages)**: [https://dodoloiu1.github.io/word-clone/](https://dodoloiu1.github.io/word-clone/)
- **Feature & Reference Specification**: [`SPEC.md`](./SPEC.md)

---

## Repository Branches

This repository maintains two primary platform branches in `C:/word`:

1. **`web` (Default Branch)** — Browser-based application built with Vite + ES2022 + `jszip` with automatic GitHub Pages deployment (`.github/workflows/deploy-web.yml`).
2. **`windows`** — Native Windows desktop application built with Electron + `electron-builder` featuring native Windows Open/Save file dialogs, direct local file I/O, native application menus, offline editing, native PDF export, and automated Windows release builds (`.github/workflows/build-windows.yml`).

---

## Core Features

- **Word-Inspired Interface**:
  - **Title Bar & Quick Access Toolbar**: AutoSave toggle, Save (`Ctrl+S`), Undo (`Ctrl+Z`), Redo (`Ctrl+Y`), New (`Ctrl+N`), Open (`Ctrl+O`), Print (`Ctrl+P`), document title, and **Tell Me / Microsoft Search** command box (`Alt+Q`).
  - **File Backstage Menu**: New (Blank & Templates), Open, Info / Document Properties, Save, Save As (`.docx`, `.html`, `.rtf`, `.txt`), Export to PDF, and Print.
  - **Complete Ribbon Interface**:
    - **Home**: Clipboard (Paste Rich/Plain, Cut, Copy, Format Painter, Clipboard History), Font (Font Family, Font Size, Grow/Shrink Font, Change Case, Clear Formatting, Bold, Italic, Underline styles, Strikethrough, Subscript, Superscript, Text Highlight Color, Font Color, Font Dialog), Paragraph (Bullets, Numbering, Multilevel List, Decrease/Increase Indent, Sort, Show/Hide `¶` Paragraph Marks, Align Left/Center/Right/Justify, Line & Paragraph Spacing, Shading, Borders, Paragraph Dialog), Styles (Quick Styles gallery with exact Word 16.0 metrics: `Normal`, `No Spacing`, `Heading 1`–`4`, `Title`, `Subtitle`, `Quote`, `Intense Quote`, `Strong`, `Emphasis`), Editing (Find, Replace, Go To, Select All).
    - **Insert**: Pages (Cover Page, Blank Page, Page Break `Ctrl+Enter`), Tables (interactive grid picker + Insert Table dialog + contextual Table Design/Layout tools: insert/delete rows & columns, merge/split cells, header row, cell shading & borders), Illustrations (Pictures from device, Shapes/Callouts, Horizontal Rule), Links (Hyperlink `Ctrl+K`, Bookmark), Comments, Header & Footer (Header, Footer, Page Numbers), Text & Symbols (Text Box, Date & Time, Equations, Symbol palette).
    - **Design**: Document Themes, Style Sets, Watermark (`CONFIDENTIAL`, `DRAFT`, Custom), Page Color, Page Borders.
    - **Layout**: Page Setup (Margins: Normal 1", Narrow 0.5", Moderate, Wide, Custom; Orientation: Portrait / Landscape; Paper Size: Letter 8.5"×11", A4, Legal, Tabloid, A5; Columns: 1, 2, 3; Breaks; Line Numbers; Page Setup Dialog), Paragraph Indent (Left/Right) & Spacing (Before/After) spinners.
    - **References**: Table of Contents generator (automatic from `Heading 1`–`3`), Insert Footnote / Endnote, Word Citations.
    - **Review**: Word Count statistics dialog, Read Aloud (Web Speech synthesis), Accessibility Checker, Comments pane, Track Changes highlight mode.
    - **View**: View modes (`Print Layout`, `Read Mode`, `Web Layout`, `Outline`, `Draft`, `Focus` mode, `Dark Mode`), Show/Hide (`Ruler`, `Gridlines`, `Navigation Pane`), Zoom (`10%`–`500%` slider, `100%`, `One Page`, `Page Width`, Zoom Dialog).
- **Navigation Pane & Status Bar**:
  - Docked left Navigation Pane with **Headings** (interactive outline), **Pages** (live page thumbnails), and **Results** (live Find matches).
  - Status Bar with **Page X of Y**, **Word Count**, **Character Count**, **Language**, **Accessibility Status**, View Mode buttons, and interactive **Zoom Slider**.
- **Genuine `.docx` Import & Export + PDF & Print**:
  - Exports standards-compliant OpenXML `.docx` archives (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `word/styles.xml`, `word/header1.xml`, `word/footer1.xml`, `word/_rels/document.xml.rels`, `docProps/core.xml`) preserving headings, text formatting, colors, highlights, alignments, lists, tables, images, hyperlinks, headers/footers, and page setup.
  - Imports `.docx` archives (parsing paragraphs, runs, styles, tables, hyperlinks, headers/footers, and embedded media images), as well as `.html`, `.rtf`, and `.txt` files.

---

## Quick Start & Development Instructions

### Prerequisites

- **Node.js** 18+ and **npm** 9+
- **Git**

### 1. Web Version (`web` branch)

```powershell
cd C:/word
git checkout web
npm install

# Start local development server
npm run dev

# Run automated test suite
npm test

# Build production bundle (configured for GitHub Pages /word-clone/)
npm run build

# Preview production bundle locally
npm run preview
```

### 2. Windows Desktop Version (`windows` branch)

```powershell
cd C:/word
git checkout windows
npm install

# Launch the Windows desktop application in development mode
npm run electron:start

# Run automated test suite
npm test

# Build the web renderer + package the runnable Windows desktop application (.exe)
npm run electron:pack

# Build full Windows distributable installer + portable executable
npm run electron:dist
```

Generated Windows desktop artifacts are placed in `C:/word/release/`:
- Unpacked runnable Windows application: `C:/word/release/win-unpacked/Word.exe`
- Windows installer / portable executable: `C:/word/release/*.exe`

---

## Differences Between `web` and `windows` Branches

| Capability | `web` Branch (Browser / GitHub Pages) | `windows` Branch (Electron Desktop) |
| :--- | :--- | :--- |
| **Application Host** | Browser SPA (`index.html`, Vite bundle) | Native Windows Electron process (`electron/main.cjs` + `electron/preload.cjs`) |
| **File Open / Save** | Browser File System Access API + `<input type="file">` / Blob download fallback | Native Windows OS Open/Save dialogs (`dialog.showOpenDialog` / `showSaveDialog`) + direct `fs` read/write to active file path |
| **Save (`Ctrl+S`)** | Saves to existing FileSystemFileHandle or triggers `.docx` download + `localStorage` AutoSave | Overwrites active file on disk directly without re-prompting (or opens native Save As dialog for untitled docs) |
| **PDF Export** | Client-side printable PDF layout / `jspdf` + `html2canvas` / print-to-PDF | Native Chromium `webContents.printToPDF()` writing directly to a chosen `.pdf` path on disk |
| **Window & Menu** | In-page Ribbon & Backstage menu | Native Windows application menu bar, keyboard accelerators, window title sync with dirty indicator (`*`), and offline desktop execution |
