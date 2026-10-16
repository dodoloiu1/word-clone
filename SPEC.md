# Microsoft Word Reference Analysis & Feature Specification (`SPEC.md`)

## 1. Executive Summary & Reference Identity

This specification documents the reverse-engineering and live behavioral analysis of the installed **Microsoft Word (`WINWORD.EXE`)** reference application and defines the feature roadmap, architecture, and implementation status for our independent Microsoft Word alternative across the **`web`** and **`windows`** development branches.

### 1.1 Primary Reference Artifacts Inspected (Read-Only)

| Artifact Path | Role | Size (Bytes) | File Version | SHA-256 Digest | REA Evidence IDs |
| :--- | :--- | ---: | :--- | :--- | :--- |
| `C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE` | Main Word x86-64 executable & Win32/WinUI host | 1,557,224 | `16.0.20430.20146` | `a02ec8e28e1210a13f1c8d1874c9f10d6ed987a18365cbaae72d8403ba2d906d` | `ev_380adc1e175f40c1e2340b0844ddcfa9486ca45b467680fdd5576ae95038edfb` (`binary_overview`)<br>`ev_68952ffea8ca19cae7ceea52772e08b64e9075983b9d1d361ba0b5d2f9b908ea` (`list_strings`)<br>`ev_bace08a3278e92c99a5802a4018dcda2ec35643f65239963ec66457bac6ce19e` (`list_procedures`)<br>`ev_89f0d36764be20dc8a2daba54e59cbd6032226d842fbb9e7c44f8ad3457af4ae` (`inspect_pe_resources`)<br>`ev_b144dac2193ddc56efccff92c276e3cf221dce1439062631f4b38181c634b1b6` (`search_strings:"Word"`)<br>`ev_edb63e20611e266ea1e042293b4be58394d4e0dd817d366797a046ff93ac4197` (`search_strings:"WinWord"`) |
| `C:\Program Files\Microsoft Office\root\Office16\WWLIB.DLL` | Core Word document engine, layout (PTS/LS), UI resources & filters | 59,527,400 | `16.0.20430.20146` | `314ef7da77a61ab554a7809eb64818766816127ea4f9b023af3c714e39e819dd` | `ev_59cc2da07469549315abe91cb2300b69cdfd94a22c8fde93c11594fabde5e218` (`inspect_pe_resources`) |
| `C:\Program Files\Microsoft Office\root\Office16\1033\WWINTL.DLL` | English (US, LCID `1033`) localized string bundles & HTML/JS dialogs | 873,704 | `16.0.20430.20032` | `899ae69467078110ef8df0ea6fc2208277936070557e43979a5ff2fad37d4f58` | `ev_af1ae13021bbb1464f68239dad050547624a31902d192b3ef110ad23f6648b4c` (`binary_overview`)<br>`ev_64adedfe6a5550e3373c914d0a9a27847208806fbe1e36ac0f5cdbae5533c5ba` (`inspect_pe_resources`) |
| `C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB` | Word Object Library (`MSFT` COM TypeLib: interfaces, enums, dialogs, keys) | 957,168 | `16.0` | *(Read-only static TypeLib)* | Static string/identifier extraction (5,342 clean identifiers) |
| `C:\Program Files\Microsoft Office\root\Office16\1033\TellMeWord.nrr` | Tell Me / Microsoft Search command lexicon & `msotcid` mappings | 625,856 | `16.0` | *(Read-only static lexicon)* | Static UTF-16LE extraction (1,190 entries, 636 `msotcid` commands) |
| `C:\Program Files\Microsoft Office\root\Office16\1033\WordNaiveBayesCommandRanker.txt` | Naive Bayes command ranker mapping `<TabSet_*>` and search tokens to TCIDs | 630,884 | `16.0` | *(Read-only text model)* | 15,816 lines of `<TabSet_*>` and token-to-TCID weights |
| `C:\Program Files\Microsoft Office\root\Office16\wordEtw.man` | Event Tracing for Windows (ETW) manifest for `Microsoft-Office-Word` | 1,128,352 | `16.0` | *(Read-only XML manifest)* | Provider `{DAF0B914-9C1C-450a-81B2-FEA7244F6FFA}` (`wwlib.dll`) |

---

## 2. Evidence Classification Taxonomy

Every finding in this specification is categorized into one of four explicit evidence classes:

1. **[CAT-1: Observed in Running Application]** — Directly observed by launching `WINWORD.EXE` (v16.0.20430.20146) and inspecting its live Windows UI Automation accessibility tree (`System.Windows.Automation`) and live COM Automation state (`Word.Application` / `ActiveDocument` / `ActiveWindow`).
2. **[CAT-2: Discovered via REA / Static Resources]** — Directly extracted from `WINWORD.EXE`, `WWLIB.DLL`, `WWINTL.DLL`, `MSWORD.OLB`, `TellMeWord.nrr`, `WordNaiveBayesCommandRanker.txt`, or `wordEtw.man` via REA (`binary_overview`, `list_strings`, `search_strings`, `list_procedures`, `inspect_pe_resources`) and static offset-anchored resource parsing.
3. **[CAT-3: Inferred from Available Evidence]** — Architectural or behavioral conclusions synthesized from multiple static and runtime observations (e.g., correlating `WINWORD.EXE` string `FMain` + `wwlib.dll` with stub-loader architecture, or correlating `msotcid` IDs with Ribbon galleries).
4. **[CAT-4: Unverified / Out of Scope for Static Verification]** — Proprietary internal algorithms or cloud-dependent behaviors that could not be verified statically or without cloud credentials (documented in Section 14).

---

## 3. Binary Architecture & PE Resource Findings (REA)

### 3.1 `WINWORD.EXE` Binary Overview (`ev_380adc1e175f40c1e2340b0844ddcfa9486ca45b467680fdd5576ae95038edfb`)
- **[CAT-2]** `WINWORD.EXE` is a 64-bit PE32+ executable (`x86_64`, machine `34404` / `0x8664`) with **11 segments**, **227 procedures** (158 internal, 69 external imports, `ev_bace08a3278e92c99a5802a4018dcda2ec35643f65239963ec66457bac6ce19e`), and **147 defined strings** (`ev_68952ffea8ca19cae7ceea52772e08b64e9075983b9d1d361ba0b5d2f9b908ea`):
  - `Headers`: `0x140000000`–`0x140000400` (1,024 bytes)
  - `.text`: `0x140001000`–`0x140004400` (13,312 bytes) — Entry point at `0x140001540` (`entry`)
  - `.rdata`: `0x140005000`–`0x140007000` (8,192 bytes)
  - `.data`: `0x140007000`–`0x140007654` (1,620 bytes)
  - `.pdata`: `0x140008000`–`0x140008600` (1,536 bytes)
  - `.didat`: `0x140009000`–`0x140009200` (512 bytes) — Delay-load import descriptors
  - `.c2r`: `0x14000a000`–`0x14000a200` (512 bytes) — Click-to-Run virtualization section (`AppVIsvSubsystems64.dll` at `0x14000a0f0`)
  - `.rsrc`: `0x14000b000`–`0x14017b600` (1,508,864 bytes — 96.9% of the executable size)
  - `.reloc`: `0x14017c000`–`0x14017c200` (512 bytes)

### 3.2 Key Strings Discovered in `WINWORD.EXE` (`ev_68952ffea8ca19cae7ceea52772e08b64e9075983b9d1d361ba0b5d2f9b908ea`)
- **[CAT-2]** Core loader and Office client strings:
  - `0x1400059b0` (UTF-16): `"wwlib.dll"`
  - `0x140005998` (UTF-16): `"wwlibt.dll"` / `0x140005980`: `"WwlibtDll"`
  - `0x140005a20` (US-ASCII): `"FMain"`
  - `0x140005930` (UTF-16): `"Software\\Microsoft\\Office\\%d.%d\\Word"`
  - `0x140005bd0` (UTF-16): `"Software\\Microsoft\\Office\\16.0\\Common\\FilesPaths"`
  - `0x140005b18` (UTF-16): `"\\Microsoft Shared\\Office16\\"`
  - `0x140005620` (UTF-16): `"mso20win32client.dll"`
  - `0x1400056d8` (UTF-16): `"mso30win32client.dll"`
  - `0x140005720` (UTF-16): `"mso40uiwin32client.dll"`
  - `0x140005768` (UTF-16): `"mso50win32client.dll"`
  - `0x1400057b0` (UTF-16): `"mso98win32client.dll"`
  - `0x140005698` (UTF-16): `"mso.dll"` / `0x140005d59`: `"MSO.DllGetLCID"`
  - `0x1400059d0` (UTF-16): `"{DC5CCACD-A7AC-4FD3-9F70-9454B5DE5161}"`
  - `0x140005a70` (UTF-16): `"{5812C571-53F0-4467-BEFA-0A4F47A9437C}"`
  - `0x140158d8c` (UTF-16): `"Microsoft Word"`, `0x140158dcc`: `"16.0.20430.20146"`, `0x140158e10`: `"WinWord"`, `0x140158f94`: `"WinWord.exe"`
- **[CAT-3]** **Inferred Architecture**: `WINWORD.EXE` is a lightweight native stub loader (only 13 KB of `.text` code) that initializes Click-to-Run (`AppVIsvSubsystems64.dll`), locates the Office shared libraries (`mso*win32client.dll`, `mso.dll`), dynamically loads `wwlib.dll` (`LoadLibraryExW` at `0x140006368`), resolves `FMain` (`GetProcAddress` at `0x140006348`), and delegates the entire Word application lifecycle to `wwlib.dll::FMain`.

### 3.3 PE Resource Inventory of `WINWORD.EXE`, `WWINTL.DLL`, and `WWLIB.DLL` (`inspect_pe_resources`)
- **[CAT-2]** `WINWORD.EXE` (`ev_89f0d36764be20dc8a2daba54e59cbd6032226d842fbb9e7c44f8ad3457af4ae`):
  - **453 directory entries**, **223 leaf resources**:
    - `RT_ICON` (`ID:3`): 205 icon payloads (1,352,920 bytes)
    - `RT_GROUP_ICON` (`ID:14`): 13 icon groups (IDs `1, 5, 6, 7, 8, 12, 13, 14, 15, 16, 17, 18, 19`, 2,948 bytes)
    - `RT_VERSION` (`ID:16`, Name `1`, Lang `1033`): Offset `1393792`, `2,044` bytes (`sha256: 0205464295d45d6f91f39950a91facccd3f74a528e821ed576594e4582eb6034`)
    - `RT_MANIFEST` (`ID:24`, Name `1`, Lang `1033`): Offset `1395836`, `139,607` bytes (`sha256: 0c4bcf07913cfe3be118f88ba905c24731827c7d466e4a9a8564bc4332380afe`) — Declares `assemblyIdentity name="winword" version="16.0.0.0" processorArchitecture="AMD64"`, 13 COM classes in `offxml.dll` (`Msxml2.DOMDocument.5.0`, `Msxml2.SAXXMLReader.5.0`, etc.), and 965 WinRT/WinAppSDK activatable classes (`Microsoft.UI.Xaml.*`, `DWriteCore.dll`, `WinUIEdit.dll`, `Microsoft.Windows.Storage.Pickers.FileOpenPicker` / `FileSavePicker`).
    - `LIMITEDACCESSFEATURE` (`IDENTITY`): Offset `37796`, `70` bytes (`"Microsoft365ClientApp_pj5pefh34bg8g"`).
- **[CAT-2]** `1033\WWINTL.DLL` (`ev_64adedfe6a5550e3373c914d0a9a27847208806fbe1e36ac0f5cdbae5533c5ba`):
  - **524 directory entries**, **259 leaf resources**:
    - `RT_HTML` (`ID:23`): 3 embedded WebView2/Trident HTML/JS resources:
      1. `TEXTPREDICTIONFEEDBACK.HTM` (offset `15316`, `5,440` bytes, `sha256: dc744155a08b40b34d7be9809b6c83d14264c3dc867d71d1e31b09ad05998d14`) — Uses Office Fluent UI tokens (`font-family: Segoe UI`, primary brand blue `#106EBE`, neutral dark `#252423`, neutral light `#F3F2F1`, border `#8A8886`).
      2. `TEXTPREDICTIONFEEDBACK.JS` (offset `20756`, `8,741` bytes, `sha256: 1fe660aac0fdf94c3a64fe1b4d3ab0a88aebeb58c49ffc828348850ab204d5b8`).
      3. `TEXTPREDICTIONFEEDBACKSENT.HTM` (offset `29500`, `8,081` bytes, `sha256: b1acefbf80c6cd68118baf09302dbe7fddf824a44c6a27420f71ae8249d7bdc7`).
    - `ID:221`: 246 localized string-table blocks (227,680 bytes).
    - `ID:234`: 1 master packed localization resource (offset `265308`, `575,353` bytes, `sha256: 0389e34446e371faaf261821df33d8bc78088ff6b8bd53506b18bbbbae266c03`).
- **[CAT-2]** `WWLIB.DLL` (`ev_59cc2da07469549315abe91cb2300b69cdfd94a22c8fde93c11594fabde5e218`):
  - **1,012 leaf resources**, including **136 `UIFILE` resources** (NetUI Ribbon & dialog schemas), **4 `SVG` resources**, **5 `IMAGE` resources**, **463 `RT_CURSOR` (`ID:1`)**, **91 `RT_GROUP_CURSOR` (`ID:12`)**, **205 `RT_ICON` (`ID:3`)**, **13 `RT_GROUP_ICON` (`ID:14`)**, **72 `ID:998` resources**, and **1 `WEVT_TEMPLATE`**.

---

## 4. Ribbon Tabs, Groups, Commands, Menus & Contextual Toolbars

### 4.1 Application Frame, Quick Access Toolbar & Top-Level Ribbon Tabs
- **[CAT-1]** Directly observed in the running `WINWORD.EXE` window (`Class="OpusApp"`, `TitleBar` `Class="NetUIOfficeCaption"`):
  - **Quick Access Toolbar (`Class="NetUIElement"`)**:
    - `AutoSave` toggle (`Class="NetUISwitch"`, `AutomationId="AutoSaveSwitch"`)
    - `Save` (`Class="NetUIRibbonButton"`, `AutomationId="FileSave"`)
    - `Undo` (`Can't Undo` / `Undo`, `Class="NetUIRibbonButton"`)
    - `Redo / Repeat` (`Repeat New...`, `Class="NetUIRibbonButton"`, `AutomationId="RedoOrRepeat"`)
    - `Customize Quick Access Toolbar` (`Class="NetUIAnchor"`)
  - **Title Bar & Search**:
    - Document title (`Document1 - Word`)
    - Microsoft Search / Tell Me box (`Class="NetUISearchBoxTextbox"`, `AutomationId="TellMeTextBoxAutomationId"`, prompt `"Type to search and use the up and down arrow keys to navigate"`)
  - **Ribbon Tab Strip (`Name="Ribbon Tabs"`, `Class="NetUIPanViewer"`)**:
    1. `File Tab` (`AutomationId="FileTabButton"`) — Backstage menu (`New`, `Open`, `Info`, `Save`, `Save As`, `Print`, `Share`, `Export`, `Close`, `Options`)
    2. `Home` (`AutomationId="TabHome"`)
    3. `Insert` (`AutomationId="TabInsert"`)
    4. `Draw` (`AutomationId="TabDrawInk"`)
    5. `Design` (`AutomationId="TabWordDesign"`)
    6. `Layout` (`AutomationId="TabPageLayoutWord"`)
    7. `References` (`AutomationId="TabReferences"`)
    8. `Mailings` (`AutomationId="TabMailings"`)
    9. `Review` (`AutomationId="TabReviewWord"`)
    10. `View` (`AutomationId="TabView"`)
    11. `Developer` (`AutomationId="TabDeveloper"`)
    12. `Help` (`AutomationId="HelpTab"`)

### 4.2 `Home` Tab (`AutomationId="TabHome"`)
- **[CAT-1]** Directly observed groups (`Class="NetUIChunk"`) and controls:
  - **Group: `Clipboard`**
    - `Paste` split button (`Paste`, `Keep Source Formatting`, `Merge Formatting`, `Keep Text Only` — corroborated by `MSWORD.OLB` `wdKeepSourceFormatting`, `wdMatchDestinationFormatting` at offsets `299228`, `299264`)
    - `Cut` (`AutomationId="Cut"`)
    - `Copy` (`AutomationId="Copy"`, `TellMeWord.nrr` `19 msotcidCopy`)
    - `Format Painter` (`AutomationId="FormatPainter"`)
    - `Office Clipboard...` dialog launcher (`AutomationId="ShowClipboard"`)
  - **Group: `Font`**
    - `Font` family combo box (`AutomationId="Font"`, default `"Calibri"`)
    - `Font Size` combo box (`AutomationId="FontSize"`, default `"11"`)
    - `Grow Font` (`AutomationId="FontSizeIncreaseWord"`)
    - `Shrink Font` (`AutomationId="FontSizeDecreaseWord"`)
    - `Change Case` menu (`AutomationId="ChangeCaseGallery"` — Sentence case, lowercase, UPPERCASE, Capitalize Each Word, tOGGLE cASE; `MSWORD.OLB` `wdDialogFormatChangeCase` at `172884`)
    - `Clear Formatting` (`AutomationId="ClearFormatting"`)
    - `Bold` (`AutomationId="Bold"`)
    - `Italic` (`AutomationId="Italic"`)
    - `Underline` split button (`AutomationId="UnderlineGallery"`, `MSWORD.OLB` `wdUnderlineDouble`, `wdUnderlineThick`, `wdUnderlineDottedHeavy`, `wdUnderlineDashHeavy`)
    - `Strikethrough` (`AutomationId="Strikethrough"`)
    - `Subscript` (`AutomationId="Subscript"`)
    - `Superscript` (`AutomationId="Superscript"`)
    - `Text Effects and Typography` (`AutomationId="TextEffectsGallery"`)
    - `Text Highlight Color` split button (`AutomationId="TextHighlightColorPicker"`, default Yellow)
    - `Font Color` split button (`AutomationId="FontColorPicker"`, `WWLIB.DLL` offset `38172832`)
    - `Font...` dialog launcher (`AutomationId="FontDialog"`, `MSWORD.OLB` `wdDialogFormatFont` at `171200`)
  - **Group: `Paragraph`**
    - `Bullets` split button (`AutomationId="BulletsGalleryWord"`)
    - `Numbering` split button (`AutomationId="NumberingGalleryWord"`)
    - `Multilevel List` (`AutomationId="MultilevelListGallery"`)
    - `Decrease Indent` (`AutomationId="IndentDecreaseWord"`)
    - `Increase Indent` (`AutomationId="IndentIncreaseWord"`)
    - `Sort...` (`AutomationId="SortDialogClassic"`)
    - `Show All` (`¶` Paragraph Marks toggle, `AutomationId="ParagraphMarks"`)
    - `Align Left` (`AutomationId="AlignLeft"`, `120 msotcidLeftJustify`)
    - `Center` (`AutomationId="AlignCenter"`)
    - `Align Right` (`AutomationId="AlignRight"`, `121 msotcidRightJustify`)
    - `Justify` (`AutomationId="AlignJustify"`, `123 msotcidFullJustify`)
    - `Line and Paragraph Spacing` (`AutomationId="LineSpacingGallery"`, options `1.0`, `1.15`, `1.5`, `2.0`, `2.5`, `3.0`, Space Before/After)
    - `Shading` split button (`AutomationId="ShadingColorPicker"`)
    - `Borders` split button (`AutomationId="BordersSelectionGallery"`, `24118 msotcidBordersGallery`)
    - `Paragraph...` dialog launcher (`AutomationId="ParagraphDialog"`, `MSWORD.OLB` `wdDialogFormatParagraph` at `171232`)
  - **Group: `Styles`**
    - `Styles` Quick Styles gallery (`AutomationId="QuickStylesGallery"`, `14839 msotcidTextStyleGallery`)
    - `Styles...` pane launcher (`AutomationId="StylesPane"`)
  - **Group: `Editing`**
    - `Editing` menu (`Find`, `Advanced Find...`, `Go To...`, `Replace...`, `Select All`, `Select Objects`)

### 4.3 `Insert` Tab (`AutomationId="TabInsert"`)
- **[CAT-1]** Directly observed groups and controls:
  - **Group: `Pages`** — `Cover Page` (`12255 msotcidLegoCoverPageGallery`), `Blank Page`, `Page Break`
  - **Group: `Tables`** — `Table` gallery & menu (`AutomationId="TableInsertGallery"`, grid selector + `Insert Table...` dialog + `Draw Table`, `WWLIB.DLL` `TableInsertTable` at `38128000`)
  - **Group: `Illustrations`** — `Pictures` (`AutomationId="FlyoutAnchorInsertPictures"`, This Device / Online), `Shapes` (`AutomationId="ShapesInsertGallery"`), `Icons` (`AutomationId="IconInsertFromFile"`), `3D Models` (`AutomationId="Insert3DModelDefault"`), `SmartArt...` (`AutomationId="SmartArtInsert"`), `Chart...` (`AutomationId="ChartInsert"`), `Screenshot` (`AutomationId="ScreenshotInsertGallery"`)
  - **Group: `Media`** — `Online Videos...` (`AutomationId="MovieFromClipOrganizerInsert"`)
  - **Group: `Links`** — `Link` / `Hyperlink`, `Bookmark`, `Cross-reference` (`WWLIB.DLL` `HyperlinkOpen` at `38153384`)
  - **Group: `Comments`** — `Comment` (`AutomationId="InsertNewComment"`)
  - **Group: `Header & Footer`** — `Header` (`AutomationId="HeaderInsertGallery"`, `12253 msotcidLegoHeaderGallery`), `Footer` (`AutomationId="FooterInsertGallery"`, `12254 msotcidLegoFooterGallery`), `Page Number` (`AutomationId="HeaderFooterPageNumberInsert"`, `14108 msotcidLegoPageNumBottomGallery`)
  - **Group: `Text`** — `Text Box` (`AutomationId="TextBoxInsertGallery"`), `Quick Parts` (`AutomationId="QuickPartsInsertGallery"`), `WordArt` (`AutomationId="WordArtInsertGallery"`), `Drop Cap` (`AutomationId="DropCapInsertGallery"`), `Signature Line` (`AutomationId="SignatureLineInsert"`), `Date & Time...` (`AutomationId="DateAndTimeInsert"`), `Object...` (`AutomationId="OleObjectctInsert"`)
  - **Group: `Symbols`** — `Equation` (`15301 msotcidInsertLegoEquationsGallery`), `Symbol` (`12322 msotcidSymbolMRUGallery`)

### 4.4 `Layout` Tab (`AutomationId="TabPageLayoutWord"`)
- **[CAT-1]** Directly observed groups and controls:
  - **Group: `Page Setup`**
    - `Margins` (`AutomationId="PageMarginsGallery"` — Normal 1", Narrow 0.5", Moderate, Wide, Mirrored, Custom Margins...)
    - `Orientation` (`AutomationId="PageOrientationGallery"` — Portrait `wdOrientPortrait`, Landscape `wdOrientLandscape`)
    - `Size` (`AutomationId="PageSizeGallery"` — Letter 8.5"x11", Legal 8.5"x14", A4 210x297mm, A3, A5, Executive, Tabloid)
    - `Columns` (`AutomationId="TableColumnsGallery"` / `WWLIB.DLL` `ColumnsGallery` at `38172864` — One, Two, Three, Left, Right)
    - `Breaks` (`AutomationId="BreaksGallery"` — Page, Column, Text Wrapping, Section Breaks: Next Page, Continuous, Even Page, Odd Page)
    - `Line Numbers` (`AutomationId="LineNumbersMenu"`)
    - `Hyphenation` (`AutomationId="HyphenationMenu"`)
    - `Page Setup...` dialog launcher (`AutomationId="PageSetupDialog"`)
  - **Group: `Paragraph`**
    - Indent `Left` / `Right` spinners (`More`/`Less` repeat buttons)
    - Spacing `Before` / `After` spinners (`More`/`Less` repeat buttons)
    - `Paragraph...` dialog launcher (`AutomationId="ParagraphDialog"`)
  - **Group: `Arrange`**
    - `Position`, `Wrap Text`, `Bring Forward`, `Send Backward`, `Selection Pane...`, `Align`, `Group`, `Rotate`

### 4.5 `View` Tab (`AutomationId="TabView"`), `Design` Tab, `References` Tab, and `Review` Tab
- **[CAT-1]** Directly observed on `View` tab:
  - **Group: `Views`** — `Read Mode` (`AutomationId="ViewFullScreenReadingView"`), `Print Layout` (`AutomationId="ViewPrintLayoutView"`), `Web Layout` (`AutomationId="ViewWebLayoutView"`), `Outline` (`AutomationId="ViewOutlineView"`), `Draft` (`AutomationId="ViewDraftView"`)
  - **Group: `Immersive`** — `Focus` (`AutomationId="ViewFocusModeView"`, `27829 msotcidFocusModeView`), `Immersive Reader` (`AutomationId="ToggleLearningTools"`)
  - **Group: `Dark Mode`** — `Switch Modes` (`AutomationId="DarkModeOn"`)
  - **Group: `Page Movement`** — `Vertical` (`AutomationId="VerticalPageMovement"`), `Side to Side` (`AutomationId="TwoPageMode"`)
  - **Group: `Show`** — `Ruler`, `Gridlines`, `Navigation Pane` (`17641 msotcidShowNavigationPaneNew`)
  - **Group: `Zoom`** — `Zoom` dialog (`174512 wdDialogViewZoom`), `100%`, `One Page`, `Multiple Pages`, `Page Width`
  - **Group: `Window`** — `New Window`, `Arrange All`, `Split`, `View Side by Side`, `Switch Windows`
  - **Group: `SharePoint` / Properties** — `Properties` (`AutomationId="DocumentInformationShowHide"`)
- **[CAT-2]** Contextual Ribbon Tab Sets discovered in `WordNaiveBayesCommandRanker.txt` (lines 2–18):
  - `<TabSet_13262>` — **Table Tools (Table Design & Layout)** (`3683 msotcidInsertRowsBelow`, `12936 msotcidTableSelectMenu`, `12937 msotcidTableDeleteMenu`, `12939 msotcidTableAutoFitMenu`, `14125 msotcidWordTableStylesHeaderRow`, `14126 msotcidWordTableStylesTotalRow`, `2068 msotcidDistributeRow`, `WWLIB.DLL` `TableMergeCells`, `TableSplitCells`, `TableInsertRowAbove`, `TableInsertRowBelow`, `TableInsertColumnRight`, `TableDeleteRow`, `TableDeleteColumn`, `TableDeleteTable`).
  - `<TabSet_13264>` — **Header & Footer Tools** (`12253 msotcidLegoHeaderGallery`, `12254 msotcidLegoFooterGallery`, `WWLIB.DLL` `HeaderPositionSpinner`, `FooterPositionSpinner`).
  - `<TabSet_13263>` — **Picture Format Tools** (`33891 msotcidImagePlay`, `17227 msotcidGalTransparency`, `2175 msotcidToggleShadow`).
  - `<TabSet_15456>` — **Equation Tools** (`15433 msotcidEquationSymbolsGallery` through `15444 msotcidEquationMatrixGallery`).

---

## 5. Typography, Paragraph Formatting & Built-In Styles

- **[CAT-1]** Live COM inspection of `ActiveDocument.Styles` in `WINWORD.EXE` established the exact default typographic metrics:
  - **Default Normal Style**:
    - Font Family: `Calibri`, Size: `11 pt`, Bold: `0`, Italic: `0`, Color: `Automatic / #000000` (`-16777216`)
    - Alignment: `Left` (`0` = `wdAlignParagraphLeft`)
    - Space Before: `0 pt`, Space After: `8 pt`
    - Line Spacing: `12.95 pt` (`LineSpacingRule = 5` = `wdLineSpaceMultiple`, equivalent to `1.08` lines)
  - **Built-in Quick Styles Gallery Metrics**:
    | Style Name | Font Family | Size (pt) | Weight / Style | Color (Word RGB / Hex) | Space Before | Space After | Line Spacing | Alignment |
    | :--- | :--- | ---: | :--- | :--- | ---: | ---: | :--- | :--- |
    | `Normal` | Calibri | 11 | Regular | `#000000` (`-16777216`) | 0 pt | 8 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `No Spacing` | Calibri | 11 | Regular | `#000000` (`-16777216`) | 0 pt | 0 pt | 1.0x (12 pt) | Left (`0`) |
    | `Heading 1` | Calibri Light | 20 | Regular | `#2F5496` (`-738148353`) | 18 pt | 4 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `Heading 2` | Calibri Light | 16 | Regular | `#2F5496` (`-738148353`) | 8 pt | 4 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `Heading 3` | Calibri | 14 | Regular | `#2F5496` (`-738148353`) | 8 pt | 4 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `Heading 4` | Calibri | 11 | Italic | `#2F5496` (`-738148353`) | 4 pt | 2 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `Title` | Calibri Light | 28 | Regular | `#000000` (`-16777216`) | 0 pt | 4 pt | 1.0x (12 pt) | Left (`0`) |
    | `Subtitle` | Calibri | 14 | Regular | `#595959` (`-587137114`) | 0 pt | 8 pt | 1.08x (12.95 pt) | Left (`0`) |
    | `Subtle Emphasis` | Calibri | 11 | Italic | `#595959` (`-587137089`) | *(inline)* | *(inline)* | *(inline)* | *(inline)* |
    | `Emphasis` | Calibri | 11 | Italic | `#000000` (`-16777216`) | *(inline)* | *(inline)* | *(inline)* | *(inline)* |
    | `Intense Emphasis` | Calibri | 11 | Italic | `#2F5496` (`-738148353`) | *(inline)* | *(inline)* | *(inline)* | *(inline)* |
    | `Strong` | Calibri | 11 | Bold | `#000000` (`-16777216`) | *(inline)* | *(inline)* | *(inline)* | *(inline)* |
    | `Quote` | Calibri | 11 | Italic | `#404040` (`-587137089`) | 8 pt | 8 pt | 1.08x (12.95 pt) | Center (`1`) |
    | `Intense Quote` | Calibri | 11 | Italic | `#2F5496` (`-738148353`) | 18 pt | 18 pt | 1.08x (12.95 pt) | Center (`1`) |
    | `List Paragraph` | Calibri | 11 | Regular | `#000000` (`-16777216`) | 0 pt | 8 pt | 1.08x (12.95 pt) | Left (`0`) |

---

## 6. Page Layout, Margins, Orientation, Columns, Headers & Footers

- **[CAT-1]** Live `ActiveDocument.PageSetup` observation on new blank document:
  - **Paper Size**: `PaperSize = 2` (`wdPaperLetter`): Width = `612 pt` (`8.5 in` = `816 px` at 96 DPI), Height = `792 pt` (`11.0 in` = `1056 px` at 96 DPI).
  - **Margins**: `TopMargin = 72 pt` (`1.0 in`), `BottomMargin = 72 pt` (`1.0 in`), `LeftMargin = 72 pt` (`1.0 in`), `RightMargin = 72 pt` (`1.0 in`).
  - **Orientation**: `Orientation = 0` (`wdOrientPortrait`; `1` = `wdOrientLandscape` swaps width/height to `792 pt x 612 pt`).
  - **Header & Footer Distance**: `HeaderDistance = 36 pt` (`0.5 in`), `FooterDistance = 36 pt` (`0.5 in`).
  - **Columns**: `TextColumns.Count = 1` (supports 1, 2, 3 columns; `wordEtw.man` event `EvtWordColumnsBeforeAfter` ID `252`).
- **[CAT-2]** Additional Paper Sizes & Breaks in `MSWORD.OLB`:
  - `wdPaperLetter` (`8.5" x 11"`), `wdPaperLegal` (`8.5" x 14"`, offset `190648`), `wdPaperA4` (`210mm x 297mm` = `8.27" x 11.69"`, offset `190724`), `wdPaperA3` (`190700`), `wdPaperA5` (`190776`), `wdPaperTabloid` (`11" x 17"`, offset `191168`, `WordNaiveBayesCommandRanker.txt` line 53 `11x17`).
  - Breaks: `wdPageBreak`, `wdColumnBreak` (offset `150188`), `wdSectionBreakNextPage` (`149996`), `wdSectionBreakEvenPage` (`150068`).
- **[CAT-2]** ETW Page Layout telemetry in `wordEtw.man`:
  - `EvtWordPageLayoutStart` (`0`), `EvtWordPageLayoutStop` (`1`), `EvtWordLineLayoutStart` (`6`), `EvtWordMarginsBeforeAfter` (`250`), `EvtWordPgSizeBeforeAfter` (`251`), `EvtWordColumnsBeforeAfter` (`252`).

---

## 7. Navigation Pane, Status Bar, Word Count, Zoom & View Modes

- **[CAT-1]** **Navigation Pane** (`Class="MsoCommandBar"`, `Name="Navigation"`, docked in `MsoDockLeft`):
  - Search input: `Edit` `Name="Search document"` (`Class="NetUITextbox"`)
  - Three sub-tabs (`Class="NetUITabHeader"`):
    1. `Headings` — Interactive outline built from document `Heading 1`–`Heading 4` elements; clicking a heading scrolls to that section.
    2. `Pages` — Thumbnail previews of each page in the document.
    3. `Results` — Highlighted matches from `Search document` / Find.
- **[CAT-1]** **Status Bar** (`Class="MsoCommandBar"` / `Class="NetUInetpane"`, docked in `MsoDockBottom`):
  - Left indicators:
    - `Page Number` (`"Page 1 of 1"`, opens Go To / Navigation)
    - `Word Count` (`"0 words"`, clicking opens `Word Count` dialog showing Pages, Words, Characters (no spaces), Characters (with spaces), Paragraphs, Lines)
    - `Language` (`"English (United States)"`)
    - `Accessibility Checker` (`"Accessibility: Good to go"`)
  - Right view & zoom controls:
    - `Focus` mode button (`27829 msotcidFocusModeView`)
    - `Read Mode` sticky button (`ViewFullScreenReadingView`)
    - `Print Layout` sticky button (`ViewPrintLayoutView`, default active `View.Type = 3`)
    - `Web Layout` sticky button (`ViewWebLayoutView`)
    - `Zoom Out (Ctrl -)` button, interactive `Zoom` slider (`Class="NetUISlider"`), `Zoom In (Ctrl +)` button, and `Zoom %` button (opens Zoom dialog: `200%`, `100%`, `75%`, `Page Width`, `Whole Page`, custom `10%`–`500%`).

---

## 8. Dialogs, Settings & Document Properties

- **[CAT-1 & CAT-2]** Dialogs identified via `MSWORD.OLB` (`wdDialog*`), `WWLIB.DLL`, and live UI Automation (`AutomationId` launchers):
  1. **Font Dialog** (`AutomationId="FontDialog"`, `wdDialogFormatFont` at `171200`, tabs `wdDialogFormatFontTabFont` at `166888`, `wdDialogFormatFontTabCharacterSpacing` at `166928`): Font family, Font style, Size, Font color, Underline style, Effects (Strikethrough, Double strikethrough, Superscript, Subscript, Small caps, All caps), Character spacing & live Preview box.
  2. **Paragraph Dialog** (`AutomationId="ParagraphDialog"`, `wdDialogFormatParagraph` at `171232`, `wdDialogFormatParagraphTabIndentsAndSpacing` at `167300`): General alignment, Indentation (Left, Right, Special: None / First line / Hanging), Spacing (Before, After, Line spacing: Single, 1.5 lines, Double, At least, Exactly, Multiple), and Preview.
  3. **Page Setup Dialog** (`AutomationId="PageSetupDialog"`, `wdDialogFilePageSetupTabPaperSource` at `169656`): Margins (Top, Bottom, Left, Right, Gutter), Orientation (Portrait, Landscape), Paper Size (Letter, A4, Legal, Tabloid, A3, A5, Custom width/height), Header/Footer distance, Columns.
  4. **Insert Table Dialog** (`AutomationId="TableInsertGallery"` -> `Insert Table...`, `WWLIB.DLL` `TableInsertTable` at `38128000`): Number of columns, Number of rows, AutoFit behavior (Fixed column width, AutoFit to contents, AutoFit to window), Header row toggle.
  5. **Insert Hyperlink Dialog** (`Ctrl+K`, `WWLIB.DLL` `HyperlinkOpen` at `38153384`): Text to display, Address (URL / `mailto:` / `#bookmark`), Target / Tooltip.
  6. **Find & Replace Dialog / Pane** (`Ctrl+F`, `Ctrl+H`, `WWLIB.DLL` `FindFirst`/`FindNext`/`FindPrevious` at `37828800`–`37828872`): Find what, Replace with, Match case, Whole words only, Find Next, Find Previous, Replace, Replace All, live match count.
  7. **Word Count Dialog** (`AutomationId="WordCount"` on Review tab & Status Bar button): Statistics table (Pages, Words, Characters (no spaces), Characters (with spaces), Paragraphs, Lines).
  8. **Document Properties / Info Dialog** (`AutomationId="DocumentInformationShowHide"`, `WWLIB.DLL` `Microsoft.WordServices.DocumentProperties` at `43107152`): Title, Author, Subject, Keywords/Tags, Comments, Created/Modified timestamps, template (`Normal.dotm`).

---

## 9. Keyboard Shortcuts & Command Behavior

- **[CAT-1 & CAT-2]** Corroborated via `MSWORD.OLB` (`WdKey` enum offsets `194484`–`195720`: `wdKeyControl`, `wdKeyShift`, `wdKeyAlt`, `wdKeyA`–`wdKeyZ`, `wdKey0`–`wdKey9`, `wdKeyReturn`, `wdKeyEsc`, `wdKeyPageDown`, `wdKeyEnd`) and live UI tooltips:
  | Shortcut | Command / Action | Evidence Source |
  | :--- | :--- | :--- |
  | `Ctrl+N` | New Blank Document (`18 msotcidNew`) | [CAT-1] UI Tooltip, [CAT-2] `TellMeWord.nrr:22356`, `MSWORD.OLB:194912` (`wdKeyN`) |
  | `Ctrl+O` | Open Document (`wdDialogFileOpen`) | [CAT-1] UI Tooltip, [CAT-2] `MSWORD.OLB:169844`, `194932` (`wdKeyO`) |
  | `Ctrl+S` | Save Document (`AutomationId="FileSave"`) | [CAT-1] QAT Tooltip, [CAT-2] `MSWORD.OLB:195012` (`wdKeyS`) |
  | `Ctrl+Shift+S` / `F12` | Save As (`14477 msotcidFileSaveAsMenu`) | [CAT-2] `TellMeWord.nrr:22610` |
  | `Ctrl+P` | Print Document (`14478 msotcidFilePrintMenu`) | [CAT-1] File Menu, [CAT-2] `TellMeWord.nrr:22384`, `MSWORD.OLB:194952` (`wdKeyP`) |
  | `Ctrl+Z` | Undo (`UndoFormattingPreviousList`) | [CAT-1] QAT Tooltip, [CAT-2] `WWLIB.DLL:38169920`, `MSWORD.OLB:195152` (`wdKeyZ`) |
  | `Ctrl+Y` | Redo / Repeat (`AutomationId="RedoOrRepeat"`) | [CAT-1] QAT Tooltip, [CAT-2] `WWLIB.DLL:38169976`, `MSWORD.OLB:195132` (`wdKeyY`) |
  | `Ctrl+X` | Cut (`AutomationId="Cut"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:195112` (`wdKeyX`) |
  | `Ctrl+C` | Copy (`AutomationId="Copy"`, `19 msotcidCopy`) | [CAT-1] Home Tab, [CAT-2] `TellMeWord.nrr:19914`, `MSWORD.OLB:194692` (`wdKeyC`) |
  | `Ctrl+V` | Paste | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:195072` (`wdKeyV`) |
  | `Ctrl+A` | Select All | [CAT-1] Editing Menu, [CAT-2] `MSWORD.OLB:194652` (`wdKeyA`) |
  | `Ctrl+B` | Bold (`AutomationId="Bold"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:194672` (`wdKeyB`) |
  | `Ctrl+I` | Italic (`AutomationId="Italic"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:194812` (`wdKeyI`) |
  | `Ctrl+U` | Underline (`AutomationId="UnderlineGallery"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:195052` (`wdKeyU`) |
  | `Ctrl+L` | Align Left (`AutomationId="AlignLeft"`) | [CAT-1] Home Tab, [CAT-2] `TellMeWord.nrr:19992`, `MSWORD.OLB:194872` (`wdKeyL`) |
  | `Ctrl+E` | Align Center (`AutomationId="AlignCenter"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:194732` (`wdKeyE`) |
  | `Ctrl+R` | Align Right (`AutomationId="AlignRight"`) | [CAT-1] Home Tab, [CAT-2] `TellMeWord.nrr:19944`, `MSWORD.OLB:194992` (`wdKeyR`) |
  | `Ctrl+J` | Justify (`AutomationId="AlignJustify"`) | [CAT-1] Home Tab, [CAT-2] `TellMeWord.nrr:20104`, `MSWORD.OLB:194832` (`wdKeyJ`) |
  | `Ctrl+F` | Find / Navigation Search | [CAT-1] Editing Menu, [CAT-2] `MSWORD.OLB:194752` (`wdKeyF`) |
  | `Ctrl+H` | Find and Replace | [CAT-1] Editing Menu, [CAT-2] `MSWORD.OLB:194792` (`wdKeyH`) |
  | `Ctrl+K` | Insert Hyperlink | [CAT-1] Links Menu, [CAT-2] `MSWORD.OLB:194852` (`wdKeyK`) |
  | `Ctrl+Enter` | Insert Page Break (`wdMatchManualPageBreak`) | [CAT-2] `MSWORD.OLB:201000`, `195452` (`wdKeyReturn`) |
  | `Ctrl+Shift+>` / `Ctrl+]` | Grow Font (`FontSizeIncreaseWord`) | [CAT-1] Home Tab |
  | `Ctrl+Shift+<` / `Ctrl+[` | Shrink Font (`FontSizeDecreaseWord`) | [CAT-1] Home Tab |
  | `Ctrl+D` | Open Font Dialog (`AutomationId="FontDialog"`) | [CAT-1] Home Tab, [CAT-2] `MSWORD.OLB:194712` (`wdKeyD`) |

---

## 10. Supported Document & Export Formats

- **[CAT-1 & CAT-2]** Discovered in `WWLIB.DLL` (`offsets 39952790–40828000`), `MSWORD.OLB` (`offsets 147428–148128`, `302272`), and live `Word.Application.FileConverters`:
  1. **Word Document (`.docx`)** — OpenXML ZIP container (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `word/styles.xml`, `word/numbering.xml`, `word/header1.xml`, `word/footer1.xml`, `word/_rels/document.xml.rels`, `docProps/core.xml`, `docProps/app.xml`); `WWLIB.DLL` offset `39952968` (`.docx`).
  2. **PDF Document (`.pdf`)** — Fixed-layout export (`WWLIB.DLL` offset `39953000` `.pdf`, `MSWORD.OLB` offset `302272` `wdExportFormatPDF`, `178624` `wdDialogExportAsFixedFormat`).
  3. **Web Page (`.html`, `.htm`)** — Rich HTML document (`WWLIB.DLL` offsets `39952856` `.htm`, `39952888` `.html`, `MSWORD.OLB:147680` `wdOpenFormatWebPages`).
  4. **Rich Text Format (`.rtf`)** — RTF markup stream (`WWLIB.DLL` offset `39952872` `.rtf`).
  5. **Plain Text (`.txt`)** — UTF-8 text (`WWLIB.DLL` offset `40828000` `.txt`, `MSWORD.OLB:147548` `wdOpenFormatText`, `147576` `wdOpenFormatUnicodeText`).
  6. **OpenDocument Text (`.odt`)** / **Word 97–2003 (`.doc`)** / **Macro-Enabled (`.docm`, `.dotm`)** / **Single File Web Page (`.mht`, `.mhtml`)** — Discovered in `WWLIB.DLL` (`39952808`–`39952952`) and `FileConverters` (`WordMac`, `WordPerfect6x`, `Recover`).

---

## 11. Common User Workflows & Expected Behavior

1. **Blank Document Creation & Page-Based Editing**:
   - Launching Word or clicking `New` (`Ctrl+N`) opens a Letter-sized (`8.5" x 11"`) page canvas in `Print Layout` view with `1"` margins, `Calibri 11pt` font, `1.08` line spacing, and `8pt` paragraph spacing after.
   - As content exceeds the vertical printable height of a page (`11" - 2"` margins = `9"`), or when the user inserts a **Page Break** (`Ctrl+Enter` or `Insert > Page Break`), visual multi-page separation with distinct headers, footers, and automatic page numbers (`Page X of Y`) is rendered on the canvas.
2. **Rich Text & Paragraph Formatting**:
   - Selecting text and clicking Ribbon controls (`Bold`, `Italic`, `Underline`, `Strikethrough`, `Subscript`, `Superscript`, `Font Family`, `Font Size`, `Grow/Shrink Font`, `Change Case`, `Text Highlight Color`, `Font Color`, `Clear Formatting`) immediately updates the selection and reflects active toggle states in the Ribbon as the caret moves.
   - Applying `Quick Styles` (`Normal`, `No Spacing`, `Heading 1`–`Heading 4`, `Title`, `Subtitle`, `Quote`, `Intense Quote`, `Strong`, `Emphasis`, `Subtle Emphasis`) applies the exact Microsoft Word style metrics discovered in Section 5 and updates the `Navigation Pane > Headings` outline in real time.
3. **Header, Footer & Page Numbering**:
   - Double-clicking the header/footer zone or clicking `Insert > Header`, `Insert > Footer`, or `Insert > Page Number` activates Header/Footer editing across all pages in the document, supporting dynamic page number tokens (`Page {PAGE} of {NUMPAGES}`), left/center/right alignment, and date/title fields.
4. **Rich Object Insertion (Tables, Images, Hyperlinks, Shapes/Callouts, Symbols, Horizontal Rules, Date/Time)**:
   - `Insert > Table` opens an interactive grid picker (up to 10x8) plus an `Insert Table...` dialog; once caret is inside a table, contextual table operations (Insert Row Above/Below, Insert Column Left/Right, Delete Row/Column/Table, Merge Cells, Split Cells, Table Shading/Borders) become available.
   - `Insert > Pictures` embeds local images with interactive resize handles, caption support, and alignment.
   - `Insert > Link` (`Ctrl+K`) inserts or edits hyperlinks with custom display text and URL.
5. **DOCX Import/Export, PDF Export & Printing**:
   - **DOCX Export** generates a valid, standards-compliant OpenXML `.docx` ZIP package (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `word/styles.xml`, `word/header1.xml`, `word/footer1.xml`, `word/_rels/document.xml.rels`, `docProps/core.xml`) preserving headings, paragraphs, bold/italic/underline/strikethrough/sub/sup, font family, font size, colors, highlights, alignments, lists, tables, images, hyperlinks, page breaks, headers/footers, and page setup margins/size/orientation.
   - **DOCX Import** parses `.docx` ZIP archives (`word/document.xml`, `word/_rels/document.xml.rels`, embedded `word/media/*` images, `word/header*.xml`, `word/footer*.xml`) as well as `.html`, `.txt`, and `.rtf` files, reconstructing the formatted document in the editor.
   - **PDF Export & Print** generates a clean multi-page PDF download and supports native print preview (`window.print()`).

---

## 12. Feature Implementation Roadmap & Status Matrix

| ID | Feature Category | Feature Description | Priority | Evidence Class & Reference | Web Branch Status | Windows Branch Status |
| :--- | :--- | :--- | :---: | :--- | :---: | :---: |
| **F-01** | Application Shell | Word-inspired Title Bar, Quick Access Toolbar (AutoSave, Save, Undo, Redo, New, Open, Print), Tell Me search bar, File Backstage, and Ribbon (`Home`, `Insert`, `Design`, `Layout`, `References`, `Review`, `View`) | **P0** | [CAT-1] `live_word_tabs.json`, [CAT-2] `TellMeWord.nrr` | **Tested** | **Tested** |
| **F-02** | Editor Core | Rich-text editing, caret tracking, selection reflection in Ribbon, Cut, Copy, Paste (Rich & Plain), Format Painter, Select All, Undo, Redo | **P0** | [CAT-1] `Home_And_Shell`, [CAT-2] `MSWORD.OLB`, `WWLIB.DLL` | **Tested** | **Tested** |
| **F-03** | Font Formatting | Font family, Font size, Grow/Shrink font, Bold, Italic, Underline (single/double/dotted/dashed), Strikethrough, Subscript, Superscript, Change Case, Text Highlight Color, Font Color, Clear Formatting | **P0** | [CAT-1] `Home` Tab `Group="Font"`, [CAT-2] `MSWORD.OLB` | **Tested** | **Tested** |
| **F-04** | Paragraph Formatting | Align Left, Center, Align Right, Justify, Bulleted Lists, Numbered Lists, Multilevel Lists, Decrease/Increase Indent, Line & Paragraph Spacing (`1.0`–`3.0`, Space Before/After), Paragraph Shading & Borders, Sort Paragraphs, Show/Hide `¶` Marks | **P0** | [CAT-1] `Home` Tab `Group="Paragraph"`, [CAT-2] `MSWORD.OLB` | **Tested** | **Tested** |
| **F-05** | Styles & Headings | Built-in Word Quick Styles gallery (`Normal`, `No Spacing`, `Heading 1`–`4`, `Title`, `Subtitle`, `Quote`, `Intense Quote`, `Strong`, `Emphasis`) using exact Word font/color/spacing metrics | **P0** | [CAT-1] `live_word_ui_and_styles.json` | **Tested** | **Tested** |
| **F-06** | Page Layout & Pagination | Page-based document layout with rulers, visual page sheets, automatic and manual Page Breaks (`Ctrl+Enter`), Margins (`Normal 1"`, `Narrow 0.5"`, `Moderate`, `Wide`, Custom), Orientation (`Portrait`, `Landscape`), Paper Sizes (`Letter`, `A4`, `Legal`, `Tabloid`, `A5`), Columns (`1`, `2`, `3`), Watermark & Page Color | **P0** | [CAT-1] `Layout` & `Design` Tabs, `PageSetup`, [CAT-2] `wordEtw.man` | **Tested** | **Tested** |
| **F-07** | Headers, Footers & Page Numbers | Editable document Header & Footer zones on every page with preset templates and dynamic Page Numbers (`Page X of Y`, top/bottom, left/center/right) | **P0** | [CAT-1] `Insert` Tab `Group="Header & Footer"`, [CAT-2] `TellMeWord.nrr` | **Tested** | **Tested** |
| **F-08** | Document Insertion | Interactive Table grid & dialog + contextual Table operations (insert/delete row/col, merge/split cells, shading/borders), Image insertion & resizing, Hyperlinks (`Ctrl+K`), Cover Page, Blank Page, Horizontal Line, Date & Time, Symbols & Equations, Comments, Footnotes | **P0** | [CAT-1] `Insert` Tab, [CAT-2] `WWLIB.DLL`, `TellMeWord.nrr` | **Tested** | **Tested** |
| **F-09** | Find, Replace & Navigation Pane | Docked left Navigation Pane (`Headings` interactive outline, `Pages` thumbnails, `Results` search matches) and Find & Replace dialog/bar (Match case, Whole words, Next/Prev, Replace, Replace All) | **P0** | [CAT-1] `MsoDockLeft` `Navigation`, [CAT-2] `WWLIB.DLL:37828800` | **Tested** | **Tested** |
| **F-10** | Status Bar, Word Count, Zoom & Views | Bottom Status Bar with live Page X of Y, Word Count, Character Count, Language, Accessibility status, View switchers (`Print Layout`, `Read Mode`, `Web Layout`, `Outline`, `Draft`, `Focus` mode, `Dark Mode`), and Zoom controls (`10%`–`500%` slider, `-`/`+`, Zoom dialog) | **P0** | [CAT-1] `Status Bar`, `View` Tab, [CAT-2] `MSWORD.OLB` | **Tested** | **Tested** |
| **F-11** | Functional Dialogs | Interactive modal dialogs for `Font...` (`Ctrl+D`), `Paragraph...`, `Page Setup...`, `Insert Table...`, `Insert Hyperlink...` (`Ctrl+K`), `Find & Replace` (`Ctrl+H`), `Word Count`, `Document Properties`, `Symbols`, `Date & Time`, and `Save As` | **P0** | [CAT-1] Dialog launchers, [CAT-2] `MSWORD.OLB` `wdDialog*` | **Tested** | **Tested** |
| **F-12** | File I/O, DOCX Import/Export & PDF | New, Open, Save, Save As, AutoSave to local storage, genuine OpenXML `.docx` import and export (preserving styles, formatting, tables, images, hyperlinks, headers/footers, page setup), `.html`/`.txt`/`.rtf` support, PDF export, and Print | **P0** | [CAT-2] `WWLIB.DLL:39952808–39953000`, `MSWORD.OLB:302272` | **Tested** | **Tested** |
| **F-13** | Keyboard Shortcuts | Comprehensive keyboard shortcut handler (`Ctrl+N/O/S/P/Z/Y/X/C/V/A/B/I/U/L/E/R/J/F/H/K/D/Enter`, `Ctrl+Shift+S`, `Ctrl+]`/`Ctrl+[`) | **P0** | [CAT-2] `MSWORD.OLB` `WdKey` enum | **Tested** | **Tested** |
| **F-14** | Windows Desktop Native Integration | Native Electron Windows desktop app (`windows` branch) with native application menu, IPC bridge (`preload.js`), native Windows Open/Save file dialogs (`.docx`, `.html`, `.rtf`, `.txt`, `.pdf`), native PDF generation (`webContents.printToPDF`), recent files, window controls, offline operation, and packaged `.exe` build | **P0** | [CAT-2] `WINWORD.EXE` `RT_MANIFEST` (`FileOpenPicker`/`FileSavePicker`) | N/A (`web`) | **Tested** |

---

## 13. Branch Architecture & Platform Differences

1. **`web` Branch (Browser Application & GitHub Pages)**:
   - Built with Vite + modular ES2022 JavaScript/CSS (`src/`), bundling `jszip` for authentic client-side OpenXML `.docx` ZIP creation and parsing.
   - Configured with `base: './'` (and `/word-clone/` compatibility) so all assets, scripts, stylesheets, and icons load cleanly on GitHub Pages (`https://dodoloiu1.github.io/word-clone/`) as well as from any static host or local preview server.
   - Uses Browser File System Access API (`showOpenFilePicker` / `showSaveFilePicker`) with automatic `<input type="file">` and Blob download fallback, plus `localStorage` AutoSave and recent documents list.
2. **`windows` Branch (Windows Desktop Application)**:
   - Built with **Electron** (`electron`, `electron-builder`) wrapping the shared document engine alongside a dedicated Windows main process (`electron/main.cjs`) and secure context-isolated preload bridge (`electron/preload.cjs`).
   - Uses native Windows OS file dialogs (`dialog.showOpenDialog`, `dialog.showSaveDialog`), direct `node:fs` local file reading/writing (remembering the active file path on disk so `Ctrl+S` saves directly without re-prompting), native Windows application menus & accelerators, native Chromium `webContents.printToPDF()` for high-fidelity vector PDF export, and `electron-builder` packaging producing a runnable Windows `.exe` executable and installer under `release/`.

---

## 14. Known Limitations & Unverified Reference Features ([CAT-4])

- **Legacy Binary `.doc` (Word 97–2003 BIFF/FIB Format)**: While `WWLIB.DLL` (`offset 39952920`) and `MSWORD.OLB` (`wdOpenFormatTemplate97`) reference binary `.doc` files, the proprietary OLE2 Compound File FIB/PieceTable binary parser inside `WWLIB.DLL` was not decompiled or copied. Our application focuses on OpenXML `.docx`, `.html`, `.rtf`, `.txt`, and `.pdf`.
- **Cloud Co-Authoring (ODSP / SharePoint Real-Time Merge)**: `WORD_WHATSNEW.XML` (`Microsoft.Office.Word.EnableAutosaveWin32ODSP`) and `wordEtw.man` (`EvtWordCOAHrApplyContentStart`, `ThreeWayMerge`) show cloud co-authoring internals, which depend on Microsoft 365 cloud servers and could not be verified offline.
- **3D Models, SmartArt & VBA Macro Execution (`VBE7.DLL`)**: Observed in `Insert` and `View > Macros` tabs (`3627 msotcidVBASecurity`), but VBA execution and 3D GLB scene rendering are outside the scope of our independent document processor.
- **Line-Breaking Engine (PTS / LineServices)**: Word uses its internal Page & Table Services (`EvtWordPTSErr`) and Line Services (`EvtWordLSErr`) engines for sub-pixel glyph shaping. Our implementation uses standards-based CSS paged layout + DOM height pagination for multi-page visual sheets.
