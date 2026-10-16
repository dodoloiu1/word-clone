import {
  exportDocx,
  importDocx,
  exportHtmlDocument,
  exportRtfDocument,
  importRtfDocument
} from '../docx/docxEngine.js';

const DEFAULT_DOCUMENT_HTML = `<p class="doc-title">Project Proposal &amp; Specification</p>
<p class="doc-subtitle">Prepared with Word — Office 16 Reference Edition</p>
<h1>1. Executive Overview</h1>
<p>Welcome to <strong>Word</strong>. This document editor is modeled directly after the installed Microsoft Word 2024 / Microsoft 365 desktop reference application (<code>WINWORD.EXE 16.0.20430.20146</code>), featuring the authentic Fluent Ribbon, Aptos 11pt typography, multi-page sheet layout, and OpenXML (<code>.docx</code>) document interoperability.</p>
<h2>1.1 Core Capabilities</h2>
<ul>
  <li><strong>Rich Text &amp; Typography:</strong> Format paragraphs with built-in Quick Styles, custom fonts, sizes, highlight colors, lists, and line spacing.</li>
  <li><strong>Page Layout &amp; Structure:</strong> Customize margins, portrait/landscape orientation, paper sizes, multi-column flow, headers, footers, watermarks, and page breaks.</li>
  <li><strong>Document Objects:</strong> Insert and edit structured tables, pictures, hyperlinks, equations, footnotes, comments, and automatic Tables of Contents.</li>
</ul>
<h2>1.2 Milestone Schedule</h2>
<table>
  <tbody>
    <tr>
      <th>Phase</th>
      <th>Deliverable</th>
      <th>Status</th>
    </tr>
    <tr>
      <td>Phase 1</td>
      <td>Reference Binary &amp; UI Analysis (<code>SPEC.md</code>)</td>
      <td>Completed</td>
    </tr>
    <tr>
      <td>Phase 2</td>
      <td>Web Application &amp; OpenXML <code>.docx</code> Engine</td>
      <td>Completed</td>
    </tr>
    <tr>
      <td>Phase 3</td>
      <td>Windows Desktop Application (Electron Native Shell)</td>
      <td>Completed</td>
    </tr>
  </tbody>
</table>
<blockquote>"Design is not just what it looks like and feels like. Design is how it works."</blockquote>`;

const TEMPLATES = {
  blank: {
    title: 'Document1',
    headerText: '',
    footerText: '',
    html: '<p><br></p>'
  },
  default: {
    title: 'Project Proposal',
    headerText: 'Project Proposal & Specification',
    footerText: 'Confidential',
    html: DEFAULT_DOCUMENT_HTML
  },
  report: {
    title: 'Quarterly Business Report',
    headerText: 'Q4 Executive Business Review',
    footerText: 'Internal Use Only',
    html: `<p class="doc-title">Quarterly Business Report</p>
<p class="doc-subtitle">Financial &amp; Operational Performance Summary</p>
<h1>Executive Summary</h1>
<p>Revenue grew by <strong>18.4%</strong> year-over-year driven by strong enterprise adoption and improved customer retention across all regional segments.</p>
<h2>Key Performance Indicators</h2>
<table>
  <tbody>
    <tr><th>Metric</th><th>Q3 Actual</th><th>Q4 Actual</th><th>YoY Change</th></tr>
    <tr><td>Annual Recurring Revenue</td><td>$42.1M</td><td>$49.8M</td><td>+18.4%</td></tr>
    <tr><td>Net Dollar Retention</td><td>112%</td><td>116%</td><td>+400 bps</td></tr>
    <tr><td>Operating Margin</td><td>21.0%</td><td>24.5%</td><td>+350 bps</td></tr>
  </tbody>
</table>
<h2>Strategic Priorities</h2>
<ol>
  <li>Expand self-service onboarding workflow for mid-market accounts.</li>
  <li>Accelerate cross-platform desktop and web synchronization.</li>
  <li>Maintain 99.99% service availability across all deployment regions.</li>
</ol>`
  },
  academic: {
    title: 'Research Paper',
    headerText: 'Running head: DISTRIBUTED DOCUMENT SYSTEMS',
    footerText: 'Manuscript Preprint',
    html: `<p class="doc-title" style="text-align: center">Deterministic OpenXML Serialization in Modern Web Environments</p>
<p class="doc-subtitle" style="text-align: center">Department of Computer Science</p>
<h1>Abstract</h1>
<p><em>This paper examines client-side OpenXML (ECMA-376) document synthesis and round-trip fidelity across browser and native desktop runtimes.</em></p>
<h1>1. Introduction</h1>
<p>Modern word processors require strict separation between semantic document structure and visual pagination. By combining standardized XML packaging with responsive page layout rules, applications achieve high interoperability.</p>
<h2>2. Methodology</h2>
<p>We evaluate paragraph style inheritance, run-level formatting attributes, and section geometry across diverse document workloads.</p>`
  },
  resume: {
    title: 'Executive Resume',
    headerText: 'Alex Morgan — Curriculum Vitae',
    footerText: 'alex.morgan@example.com',
    html: `<p class="doc-title">Alex Morgan</p>
<p class="doc-subtitle">Principal Systems Architect • Seattle, WA • alex.morgan@example.com</p>
<h1>Professional Summary</h1>
<p>Systems architect with 12+ years of experience designing high-performance productivity applications, document engines, and cross-platform desktop software.</p>
<h1>Experience</h1>
<h2>Principal Software Engineer — Cloud Productivity Suite (2021–Present)</h2>
<ul>
  <li>Architected collaborative rich-text editing engine serving 4M+ daily active users.</li>
  <li>Reduced OpenXML import/export latency by 62% through streaming ZIP compression.</li>
</ul>
<h2>Senior Desktop Engineer — Enterprise Systems Corp (2017–2021)</h2>
<ul>
  <li>Led Windows desktop client modernization and automated release pipelines.</li>
</ul>
<h1>Education</h1>
<p><strong>M.S. in Computer Science</strong> — University of Washington</p>`
  }
};

const QUICK_STYLES = [
  { id: 'normal', name: 'Normal', tag: 'P', className: '', previewStyle: 'font-family: Aptos, sans-serif; font-size: 11pt; color: #201F1E;' },
  { id: 'no-spacing', name: 'No Spacing', tag: 'P', className: 'no-spacing', previewStyle: 'font-family: Aptos, sans-serif; font-size: 11pt; margin: 0;' },
  { id: 'heading1', name: 'Heading 1', tag: 'H1', className: '', previewStyle: 'font-family: "Aptos Display", sans-serif; font-size: 15pt; color: #0F4761;' },
  { id: 'heading2', name: 'Heading 2', tag: 'H2', className: '', previewStyle: 'font-family: "Aptos Display", sans-serif; font-size: 13pt; color: #0F4761;' },
  { id: 'heading3', name: 'Heading 3', tag: 'H3', className: '', previewStyle: 'font-size: 12pt; color: #0F4761;' },
  { id: 'title', name: 'Title', tag: 'P', className: 'doc-title', previewStyle: 'font-family: "Aptos Display", sans-serif; font-size: 16pt; color: #262626;' },
  { id: 'subtitle', name: 'Subtitle', tag: 'P', className: 'doc-subtitle', previewStyle: 'font-size: 11pt; color: #595959;' },
  { id: 'quote', name: 'Quote', tag: 'BLOCKQUOTE', className: '', previewStyle: 'font-style: italic; color: #404040;' }
];

const PAPER_SIZES_PX = {
  letter: { width: 816, height: 1056, label: 'Letter (8.5" × 11")' },
  legal: { width: 816, height: 1344, label: 'Legal (8.5" × 14")' },
  a4: { width: 794, height: 1123, label: 'A4 (210mm × 297mm)' },
  executive: { width: 696, height: 1008, label: 'Executive (7.25" × 10.5")' }
};

const MARGIN_PRESETS_PX = {
  normal: { top: 96, right: 96, bottom: 96, left: 96, label: 'Normal (1" all)' },
  narrow: { top: 48, right: 48, bottom: 48, left: 48, label: 'Narrow (0.5" all)' },
  moderate: { top: 96, right: 72, bottom: 96, left: 72, label: 'Moderate (1" T/B, 0.75" L/R)' },
  wide: { top: 96, right: 192, bottom: 96, left: 192, label: 'Wide (1" T/B, 2" L/R)' }
};

export function createWordApp(container, options = {}) {
  const isDesktop = Boolean(typeof window !== 'undefined' && window.wordDesktop);

  const state = {
    activeTab: 'home',
    ribbonCollapsed: false,
    backstageOpen: false,
    backstageSection: 'new',
    autoSave: true,
    dirty: false,
    currentFilePath: null,
    documentTitle: 'Document1',
    headerText: 'Project Proposal & Specification',
    footerText: 'Confidential',
    showPageNumbers: true,
    showParagraphMarks: false,
    showRuler: true,
    showGridlines: false,
    showNavPane: false,
    navPaneTab: 'headings',
    navSearchQuery: '',
    showCommentsPane: false,
    trackChanges: false,
    darkMode: false,
    focusMode: false,
    viewMode: 'print', // 'print' | 'read' | 'web' | 'outline' | 'draft'
    zoom: 100,
    watermark: '',
    pageColor: '#FFFFFF',
    pageBorder: false,
    pageSetup: {
      paperSize: 'letter',
      orientation: 'portrait',
      margins: 'normal',
      customMarginsPx: null,
      columns: 1
    },
    metadata: {
      title: 'Project Proposal & Specification',
      author: 'Word User',
      subject: 'Reference Implementation',
      comments: ''
    },
    comments: [],
    footnotes: [],
    formatPainterStyle: null,
    history: [],
    historyIndex: -1,
    activeCell: null,
    recentFiles: []
  };

  // Load saved local state if available
  try {
    const savedRecent = localStorage.getItem('word_recent_files');
    if (savedRecent) state.recentFiles = JSON.parse(savedRecent);
  } catch (_) {}

  container.innerHTML = `
    <button class="exit-focus-btn" data-action="toggle-focus">Exit Focus Mode (Esc)</button>

    <!-- Top Title Bar -->
    <header class="title-bar">
      <div class="qat-left">
        <div class="app-brand" title="Word — Office 16 Reference Edition">
          <svg viewBox="0 0 64 64">
            <rect x="14" y="6" width="44" height="52" rx="4" fill="#2B7CD3"/>
            <rect x="6" y="16" width="32" height="32" rx="4" fill="#0C3E8C"/>
            <path d="M13.2 23h3.6l2.3 11.4L21.7 23h3.1l2.5 11.4L29.7 23h3.5l-4.1 18h-3.5l-2.4-11-2.5 11h-3.5L13.2 23z" fill="#ffffff"/>
          </svg>
          <span>Word</span>
        </div>
        <div class="autosave-toggle" data-action="toggle-autosave" title="Toggle AutoSave to local storage">
          <span>AutoSave</span>
          <div class="autosave-pill on" id="autosave-pill"></div>
        </div>
        <button class="qat-btn" data-action="save-doc" title="Save (Ctrl+S)">💾</button>
        <button class="qat-btn" data-action="undo" title="Undo (Ctrl+Z)">↶</button>
        <button class="qat-btn" data-action="redo" title="Redo (Ctrl+Y)">↷</button>
        <button class="qat-btn" data-action="print-doc" title="Print (Ctrl+P)">🖨️</button>
      </div>

      <div class="title-center">
        <input type="text" class="doc-title-input" id="doc-title-input" value="${state.documentTitle}" aria-label="Document Title" />
        <span class="save-badge" id="save-badge">Saved</span>
        <div class="tell-me-wrapper">
          <span class="tell-me-icon">💡</span>
          <input type="text" class="tell-me-input" id="tell-me-input" placeholder="Tell me what you want to do (Alt+Q)..." autocomplete="off" />
          <div class="tell-me-dropdown" id="tell-me-dropdown" style="display: none;"></div>
        </div>
      </div>

      <div class="title-right">
        <span class="platform-pill" id="platform-pill">${isDesktop ? 'Windows Desktop' : 'Web Edition'}</span>
      </div>
    </header>

    <!-- Ribbon Tab Strip -->
    <nav class="ribbon-tabs" id="ribbon-tabs" role="tablist">
      <button class="ribbon-tab file-tab" data-action="open-backstage">File</button>
      <button class="ribbon-tab active" data-tab="home">Home</button>
      <button class="ribbon-tab" data-tab="insert">Insert</button>
      <button class="ribbon-tab" data-tab="design">Design</button>
      <button class="ribbon-tab" data-tab="layout">Layout</button>
      <button class="ribbon-tab" data-tab="references">References</button>
      <button class="ribbon-tab" data-tab="review">Review</button>
      <button class="ribbon-tab" data-tab="view">View</button>
      <button class="ribbon-tab contextual-tab" data-tab="table" id="table-contextual-tab" style="display: none;">Table Layout</button>
      <div class="ribbon-spacer"></div>
      <button class="r-btn" data-action="toggle-ribbon" title="Collapse or expand the Ribbon">⌃</button>
    </nav>

    <!-- Ribbon Surface -->
    <section class="ribbon-surface" id="ribbon-surface"></section>

    <!-- Horizontal Ruler -->
    <div class="ruler-bar" id="ruler-bar">
      <div class="ruler-track" id="ruler-track">
        <div class="ruler-margin-left" id="ruler-margin-left" style="width: 96px;"></div>
        <div class="ruler-margin-right" id="ruler-margin-right" style="width: 96px;"></div>
        <div class="ruler-ticks">
          <span>1</span><span>2</span><span>3</span><span>4</span><span>5</span><span>6</span><span>7</span>
        </div>
      </div>
    </div>

    <!-- Main Workspace -->
    <div class="workspace">
      <!-- Left Navigation Pane -->
      <aside class="nav-pane hidden" id="nav-pane">
        <div class="nav-pane-header">
          <span>Navigation</span>
          <button class="r-btn r-btn-icon" data-action="toggle-nav-pane" title="Close Navigation Pane">✕</button>
        </div>
        <div class="nav-search-box">
          <input type="text" id="nav-search-input" placeholder="Search document..." />
        </div>
        <div class="nav-tabs">
          <button class="nav-tab active" data-nav-tab="headings">Headings</button>
          <button class="nav-tab" data-nav-tab="pages">Pages</button>
          <button class="nav-tab" data-nav-tab="results">Results</button>
        </div>
        <div class="nav-content" id="nav-content"></div>
      </aside>

      <!-- Document Canvas Viewport -->
      <main class="canvas-viewport view-print" id="canvas-viewport">
        <div class="zoom-stage" id="zoom-stage">
          <div class="document-page" id="document-page" style="padding: 48px 96px;">
            <div class="watermark-overlay"><span class="watermark-text" id="watermark-text"></span></div>

            <!-- Header Zone -->
            <div class="header-zone" id="header-zone">
              <span class="zone-tag">Header</span>
              <div class="header-editable" id="header-editable" contenteditable="true" aria-label="Document Header">${state.headerText}</div>
            </div>

            <!-- Editable Rich-Text Document Body -->
            <div class="editor-body" id="editor-body" contenteditable="true" spellcheck="true" role="textbox" aria-multiline="true" aria-label="Document Body">${DEFAULT_DOCUMENT_HTML}</div>

            <!-- Footer Zone -->
            <div class="footer-zone" id="footer-zone">
              <span class="zone-tag">Footer</span>
              <div class="footer-editable" id="footer-editable" contenteditable="true" aria-label="Document Footer">${state.footerText}</div>
              <div class="page-number-slot" id="page-number-slot">Page 1 of 1</div>
            </div>
          </div>
        </div>
      </main>

      <!-- Right Comments Pane -->
      <aside class="comments-pane hidden" id="comments-pane">
        <div class="nav-pane-header">
          <span>Comments</span>
          <button class="r-btn r-btn-icon" data-action="toggle-comments-pane">✕</button>
        </div>
        <div class="nav-content" id="comments-list"></div>
      </aside>
    </div>

    <!-- Floating Selection Mini-Toolbar -->
    <div class="mini-toolbar hidden" id="mini-toolbar">
      <button class="r-btn r-btn-icon" data-action="bold" title="Bold (Ctrl+B)"><b>B</b></button>
      <button class="r-btn r-btn-icon" data-action="italic" title="Italic (Ctrl+I)"><i>I</i></button>
      <button class="r-btn r-btn-icon" data-action="underline" title="Underline (Ctrl+U)"><u>U</u></button>
      <button class="r-btn r-btn-icon" data-action="highlight-yellow" title="Highlight">🖍️</button>
      <button class="r-btn r-btn-icon" data-action="open-hyperlink-dialog" title="Insert Link (Ctrl+K)">🔗</button>
      <button class="r-btn r-btn-icon" data-action="add-comment" title="New Comment">💬</button>
    </div>

    <!-- Status Bar -->
    <footer class="status-bar" id="status-bar">
      <div class="status-left">
        <button class="status-btn" data-action="toggle-nav-pane" id="status-page-btn" title="Open Navigation Pane">Page 1 of 1</button>
        <button class="status-btn" data-action="open-word-count-dialog" id="status-words-btn" title="Click to open Word Count statistics">0 words</button>
        <span class="status-btn" id="status-chars-label">0 chars</span>
        <span class="status-btn" title="Proofing Language">English (United States)</span>
        <span class="status-btn" id="status-track-indicator" style="display: none;">Track Changes: ON</span>
      </div>
      <div class="status-right">
        <button class="status-btn" data-action="toggle-focus" title="Focus Mode">🎯 Focus</button>
        <button class="status-btn" data-action="set-view" data-view="read" title="Read Mode">📖</button>
        <button class="status-btn active" data-action="set-view" data-view="print" title="Print Layout">📄</button>
        <button class="status-btn" data-action="set-view" data-view="web" title="Web Layout">🌐</button>
        <button class="status-btn" data-action="zoom-out" title="Zoom Out">−</button>
        <input type="range" class="zoom-slider" id="zoom-slider" min="10" max="500" step="10" value="100" aria-label="Zoom Percentage" />
        <button class="status-btn" data-action="zoom-in" title="Zoom In">+</button>
        <button class="status-btn" data-action="reset-zoom" id="status-zoom-btn" title="Reset Zoom to 100%">100%</button>
      </div>
    </footer>

    <!-- Backstage File Menu Overlay -->
    <div class="backstage-overlay hidden" id="backstage-overlay">
      <div class="backstage-nav">
        <button class="backstage-back-btn" data-action="close-backstage">← Back</button>
        <button class="backstage-nav-item active" data-backstage="new">New</button>
        <button class="backstage-nav-item" data-backstage="open">Open</button>
        <button class="backstage-nav-item" data-backstage="info">Info</button>
        <button class="backstage-nav-item" data-action="save-doc">Save</button>
        <button class="backstage-nav-item" data-backstage="save-as">Save As</button>
        <button class="backstage-nav-item" data-backstage="export">Export</button>
        <button class="backstage-nav-item" data-action="print-doc">Print</button>
      </div>
      <div class="backstage-content" id="backstage-content"></div>
    </div>

    <!-- Modal Dialog Container -->
    <div class="modal-backdrop hidden" id="modal-backdrop">
      <div class="modal-dialog" id="modal-dialog" role="dialog" aria-modal="true"></div>
    </div>

    <!-- Hidden File Inputs -->
    <input type="file" id="file-open-input" accept=".docx,.html,.htm,.rtf,.txt" style="display: none;" />
    <input type="file" id="image-upload-input" accept="image/*" style="display: none;" />
  `;

  // Element references
  const editorBody = container.querySelector('#editor-body');
  const headerEditable = container.querySelector('#header-editable');
  const footerEditable = container.querySelector('#footer-editable');
  const documentPage = container.querySelector('#document-page');
  const canvasViewport = container.querySelector('#canvas-viewport');
  const zoomStage = container.querySelector('#zoom-stage');
  const ribbonSurface = container.querySelector('#ribbon-surface');
  const docTitleInput = container.querySelector('#doc-title-input');
  const saveBadge = container.querySelector('#save-badge');
  const statusPageBtn = container.querySelector('#status-page-btn');
  const statusWordsBtn = container.querySelector('#status-words-btn');
  const statusCharsLabel = container.querySelector('#status-chars-label');
  const statusZoomBtn = container.querySelector('#status-zoom-btn');
  const zoomSlider = container.querySelector('#zoom-slider');
  const navPane = container.querySelector('#nav-pane');
  const navContent = container.querySelector('#nav-content');
  const navSearchInput = container.querySelector('#nav-search-input');
  const commentsPane = container.querySelector('#comments-pane');
  const commentsList = container.querySelector('#comments-list');
  const backstageOverlay = container.querySelector('#backstage-overlay');
  const backstageContent = container.querySelector('#backstage-content');
  const modalBackdrop = container.querySelector('#modal-backdrop');
  const modalDialog = container.querySelector('#modal-dialog');
  const fileOpenInput = container.querySelector('#file-open-input');
  const imageUploadInput = container.querySelector('#image-upload-input');
  const tellMeInput = container.querySelector('#tell-me-input');
  const tellMeDropdown = container.querySelector('#tell-me-dropdown');
  const miniToolbar = container.querySelector('#mini-toolbar');
  const tableContextualTab = container.querySelector('#table-contextual-tab');

  // History Management (Undo / Redo)
  function pushHistory() {
    const snapshot = {
      html: editorBody.innerHTML,
      headerText: headerEditable.innerHTML,
      footerText: footerEditable.innerHTML
    };
    if (
      state.historyIndex >= 0 &&
      state.history[state.historyIndex] &&
      state.history[state.historyIndex].html === snapshot.html &&
      state.history[state.historyIndex].headerText === snapshot.headerText &&
      state.history[state.historyIndex].footerText === snapshot.footerText
    ) {
      return;
    }
    state.history = state.history.slice(0, state.historyIndex + 1);
    state.history.push(snapshot);
    if (state.history.length > 80) state.history.shift();
    state.historyIndex = state.history.length - 1;
  }

  function undo() {
    if (state.historyIndex > 0) {
      state.historyIndex--;
      const snap = state.history[state.historyIndex];
      editorBody.innerHTML = snap.html;
      headerEditable.innerHTML = snap.headerText;
      footerEditable.innerHTML = snap.footerText;
      refreshDocumentMetrics();
    }
  }

  function redo() {
    if (state.historyIndex < state.history.length - 1) {
      state.historyIndex++;
      const snap = state.history[state.historyIndex];
      editorBody.innerHTML = snap.html;
      headerEditable.innerHTML = snap.headerText;
      footerEditable.innerHTML = snap.footerText;
      refreshDocumentMetrics();
    }
  }

  // Word Count & Metrics Calculation
  function getDocumentStats() {
    const rawText = (editorBody.innerText || editorBody.textContent || '')
      .replace(/\u200B/g, '')
      .replace(/¶/g, '')
      .trim();
    const words = rawText ? rawText.split(/\s+/).filter(Boolean).length : 0;
    const charsWithSpaces = rawText.length;
    const charsNoSpaces = rawText.replace(/\s+/g, '').length;
    const paragraphs = Math.max(
      1,
      editorBody.querySelectorAll('p, h1, h2, h3, h4, li, blockquote').length
    );
    const lines = rawText ? rawText.split(/\r?\n/).length : 1;
    const manualBreaks = editorBody.querySelectorAll('hr.page-break').length;
    const estimatedHeightPages = Math.max(1, Math.ceil((editorBody.scrollHeight || 400) / 860));
    const pages = Math.max(1, manualBreaks + estimatedHeightPages);

    return {
      pages,
      words,
      charsWithSpaces,
      charsNoSpaces,
      paragraphs,
      lines
    };
  }

  function refreshDocumentMetrics() {
    const stats = getDocumentStats();
    statusPageBtn.textContent = `Page 1 of ${stats.pages}`;
    statusWordsBtn.textContent = `${stats.words} ${stats.words === 1 ? 'word' : 'words'}`;
    statusCharsLabel.textContent = `${stats.charsWithSpaces} chars`;

    const pageSlot = container.querySelector('#page-number-slot');
    if (pageSlot) {
      pageSlot.style.display = state.showPageNumbers ? 'block' : 'none';
      pageSlot.textContent = `Page 1 of ${stats.pages}`;
    }

    if (state.showNavPane) {
      renderNavPane();
    }

    if (state.autoSave) {
      try {
        localStorage.setItem('word_autosave_html', editorBody.innerHTML);
        localStorage.setItem('word_autosave_title', state.documentTitle);
      } catch (_) {}
    }

    if (isDesktop && window.wordDesktop.setWindowState) {
      window.wordDesktop.setWindowState({
        title: `${state.documentTitle}${state.dirty ? ' *' : ''} - Word`,
        dirty: state.dirty
      });
    } else if (typeof document !== 'undefined') {
      document.title = `${state.documentTitle}${state.dirty ? ' *' : ''} - Word`;
    }
  }

  function markDirty() {
    state.dirty = true;
    saveBadge.textContent = state.autoSave ? 'Saved to browser' : 'Modified';
    refreshDocumentMetrics();
  }

  // Apply Page Setup (size, orientation, margins, columns, background, borders, watermark)
  function applyPageSetupToDom() {
    const sizeInfo = PAPER_SIZES_PX[state.pageSetup.paperSize] || PAPER_SIZES_PX.letter;
    const isLandscape = state.pageSetup.orientation === 'landscape';
    const widthPx = isLandscape ? sizeInfo.height : sizeInfo.width;
    const heightPx = isLandscape ? sizeInfo.width : sizeInfo.height;

    const marginPreset = MARGIN_PRESETS_PX[state.pageSetup.margins] || MARGIN_PRESETS_PX.normal;
    const margins = state.pageSetup.customMarginsPx || marginPreset;

    documentPage.style.width = `${widthPx}px`;
    documentPage.style.minHeight = `${heightPx}px`;
    documentPage.style.paddingTop = `${ Math.max(24, margins.top - 36) }px`;
    documentPage.style.paddingBottom = `${ Math.max(24, margins.bottom - 36) }px`;
    documentPage.style.paddingLeft = `${margins.left}px`;
    documentPage.style.paddingRight = `${margins.right}px`;
    documentPage.style.backgroundColor = state.pageColor || '#FFFFFF';
    documentPage.style.border = state.pageBorder ? '2px solid #185ABD' : 'none';

    editorBody.style.columnCount = state.pageSetup.columns > 1 ? String(state.pageSetup.columns) : '1';
    editorBody.style.columnGap = state.pageSetup.columns > 1 ? '36px' : '0px';

    const rulerTrack = container.querySelector('#ruler-track');
    const rulerLeft = container.querySelector('#ruler-margin-left');
    const rulerRight = container.querySelector('#ruler-margin-right');
    if (rulerTrack && rulerLeft && rulerRight) {
      rulerTrack.style.width = `${widthPx}px`;
      rulerLeft.style.width = `${margins.left}px`;
      rulerRight.style.width = `${margins.right}px`;
    }

    const wmText = container.querySelector('#watermark-text');
    if (wmText) {
      wmText.textContent = state.watermark || '';
    }
  }

  // Render Ribbon Surface for Active Tab
  function renderRibbon() {
    const tab = state.activeTab;
    container.querySelectorAll('.ribbon-tab[data-tab]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === tab);
    });

    ribbonSurface.classList.toggle('collapsed', state.ribbonCollapsed);

    if (tab === 'home') {
      ribbonSurface.innerHTML = `
        <!-- Clipboard Group -->
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="paste" title="Paste (Ctrl+V)">
              <span class="r-icon">📋</span>
              <span class="r-label">Paste</span>
            </button>
            <div class="ribbon-col">
              <button class="r-btn" data-action="cut" title="Cut (Ctrl+X)">✂️ Cut</button>
              <button class="r-btn" data-action="copy" title="Copy (Ctrl+C)">📄 Copy</button>
              <button class="r-btn ${state.formatPainterStyle ? 'active' : ''}" data-action="format-painter" title="Copy formatting from selection">🖌️ Format Painter</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Clipboard</div>
        </div>

        <!-- Font Group -->
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <div class="ribbon-row">
                <select class="r-select" id="ribbon-font-family" aria-label="Font Family" style="width: 132px;">
                  <option value="Aptos">Aptos (Body)</option>
                  <option value="Aptos Display">Aptos Display</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Calibri Light">Calibri Light</option>
                  <option value="Segoe UI">Segoe UI</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Arial">Arial</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Cambria">Cambria</option>
                  <option value="Consolas">Consolas</option>
                  <option value="Verdana">Verdana</option>
                  <option value="Courier New">Courier New</option>
                </select>
                <select class="r-select" id="ribbon-font-size" aria-label="Font Size" style="width: 54px;">
                  <option value="8">8</option>
                  <option value="9">9</option>
                  <option value="10">10</option>
                  <option value="11" selected>11</option>
                  <option value="12">12</option>
                  <option value="14">14</option>
                  <option value="16">16</option>
                  <option value="18">18</option>
                  <option value="20">20</option>
                  <option value="24">24</option>
                  <option value="28">28</option>
                  <option value="36">36</option>
                  <option value="48">48</option>
                  <option value="72">72</option>
                </select>
                <button class="r-btn r-btn-icon" data-action="font-size-up" title="Increase Font Size (Ctrl+Shift+>)">A▲</button>
                <button class="r-btn r-btn-icon" data-action="font-size-down" title="Decrease Font Size (Ctrl+Shift+<)">A▼</button>
                <select class="r-select" id="ribbon-change-case" aria-label="Change Case" title="Change Case" style="width: 54px;">
                  <option value="">Aa</option>
                  <option value="sentence">Sentence case</option>
                  <option value="lower">lowercase</option>
                  <option value="upper">UPPERCASE</option>
                  <option value="capitalize">Capitalize Each Word</option>
                </select>
                <button class="r-btn r-btn-icon" data-action="clear-formatting" title="Clear All Formatting">🧹</button>
              </div>
              <div class="ribbon-row">
                <button class="r-btn r-btn-icon" data-action="bold" title="Bold (Ctrl+B)"><b>B</b></button>
                <button class="r-btn r-btn-icon" data-action="italic" title="Italic (Ctrl+I)"><i>I</i></button>
                <button class="r-btn r-btn-icon" data-action="underline" title="Underline (Ctrl+U)"><u>U</u></button>
                <button class="r-btn r-btn-icon" data-action="strikethrough" title="Strikethrough"><s>ab</s></button>
                <button class="r-btn r-btn-icon" data-action="subscript" title="Subscript (Ctrl+=)">x₂</button>
                <button class="r-btn r-btn-icon" data-action="superscript" title="Superscript (Ctrl+Shift++)">x²</button>
                <label class="color-swatch-btn" title="Text Highlight Color">
                  <span>🖍️</span>
                  <input type="color" id="ribbon-highlight-color" value="#FFF100" />
                </label>
                <label class="color-swatch-btn" title="Font Color">
                  <span>A</span>
                  <input type="color" id="ribbon-font-color" value="#C00000" />
                </label>
              </div>
            </div>
          </div>
          <div class="ribbon-group-footer">
            <span>Font</span>
            <button class="dialog-launcher" data-action="open-font-dialog" title="Font Dialog (Ctrl+D)">↘</button>
          </div>
        </div>

        <!-- Paragraph Group -->
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <div class="ribbon-row">
                <button class="r-btn r-btn-icon" data-action="bullet-list" title="Bullets">•≡</button>
                <button class="r-btn r-btn-icon" data-action="number-list" title="Numbering">1≡</button>
                <button class="r-btn r-btn-icon" data-action="outdent" title="Decrease Indent">⇤</button>
                <button class="r-btn r-btn-icon" data-action="indent" title="Increase Indent">⇥</button>
                <button class="r-btn r-btn-icon" data-action="sort-paragraphs" title="Sort Paragraphs A to Z">A↓Z</button>
                <button class="r-btn r-btn-icon ${state.showParagraphMarks ? 'active' : ''}" data-action="toggle-paragraph-marks" title="Show/Hide Paragraph Marks (¶)">¶</button>
              </div>
              <div class="ribbon-row">
                <button class="r-btn r-btn-icon" data-action="align-left" title="Align Left (Ctrl+L)">⫷</button>
                <button class="r-btn r-btn-icon" data-action="align-center" title="Center (Ctrl+E)">⫼</button>
                <button class="r-btn r-btn-icon" data-action="align-right" title="Align Right (Ctrl+R)">⫸</button>
                <button class="r-btn r-btn-icon" data-action="align-justify" title="Justify (Ctrl+J)">☰</button>
                <select class="r-select" id="ribbon-line-spacing" aria-label="Line Spacing" title="Line and Paragraph Spacing" style="width: 62px;">
                  <option value="1">↕ 1.0</option>
                  <option value="1.15" selected>↕ 1.15</option>
                  <option value="1.5">↕ 1.5</option>
                  <option value="2">↕ 2.0</option>
                  <option value="2.5">↕ 2.5</option>
                  <option value="3">↕ 3.0</option>
                </select>
                <button class="r-btn r-btn-icon" data-action="paragraph-shading" title="Paragraph Shading">🪣</button>
                <button class="r-btn r-btn-icon" data-action="paragraph-border" title="Toggle Paragraph Box Border">▣</button>
              </div>
            </div>
          </div>
          <div class="ribbon-group-footer">
            <span>Paragraph</span>
            <button class="dialog-launcher" data-action="open-paragraph-dialog" title="Paragraph Settings">↘</button>
          </div>
        </div>

        <!-- Styles Group -->
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="styles-gallery">
              ${QUICK_STYLES.map(s => `
                <button class="style-card" data-action="apply-style" data-style-id="${s.id}" title="Apply ${s.name} style">
                  <span class="style-preview" style="${s.previewStyle}">AaBbCc</span>
                  <span class="style-name">${s.name}</span>
                </button>
              `).join('')}
            </div>
          </div>
          <div class="ribbon-group-footer">Styles</div>
        </div>

        <!-- Editing Group -->
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <button class="r-btn" data-action="open-find-dialog" title="Find (Ctrl+F)">🔍 Find</button>
              <button class="r-btn" data-action="open-replace-dialog" title="Replace (Ctrl+H)">⇄ Replace</button>
              <button class="r-btn" data-action="select-all" title="Select All (Ctrl+A)">▣ Select All</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Editing</div>
        </div>
      `;
    } else if (tab === 'insert') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="insert-cover-page" title="Insert Cover Page">
              <span class="r-icon">📘</span>
              <span class="r-label">Cover Page</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-blank-page" title="Insert Blank Page">
              <span class="r-icon">📄</span>
              <span class="r-label">Blank Page</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-page-break" title="Insert Page Break (Ctrl+Enter)">
              <span class="r-icon">✂️</span>
              <span class="r-label">Page Break</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Pages</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="open-table-dialog" title="Insert Table">
              <span class="r-icon">⊞</span>
              <span class="r-label">Table</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Tables</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="trigger-image-upload" title="Insert Picture from Device">
              <span class="r-icon">🖼️</span>
              <span class="r-label">Pictures</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-callout-shape" title="Insert Callout Box">
              <span class="r-icon">💠</span>
              <span class="r-label">Shapes</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-chart-graphic" title="Insert Sample SVG Chart">
              <span class="r-icon">📊</span>
              <span class="r-label">Chart</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Illustrations</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="open-hyperlink-dialog" title="Insert Hyperlink (Ctrl+K)">
              <span class="r-icon">🔗</span>
              <span class="r-label">Link</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-bookmark" title="Insert Bookmark Anchor">
              <span class="r-icon">🔖</span>
              <span class="r-label">Bookmark</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Links</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="add-comment" title="Insert New Comment">
              <span class="r-icon">💬</span>
              <span class="r-label">Comment</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Comments</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="focus-header" title="Edit Document Header">
              <span class="r-icon">⬆️</span>
              <span class="r-label">Header</span>
            </button>
            <button class="r-btn r-btn-large" data-action="focus-footer" title="Edit Document Footer">
              <span class="r-icon">⬇️</span>
              <span class="r-label">Footer</span>
            </button>
            <button class="r-btn r-btn-large ${state.showPageNumbers ? 'active' : ''}" data-action="toggle-page-numbers" title="Toggle Page Numbers">
              <span class="r-icon">#️⃣</span>
              <span class="r-label">Page Number</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Header &amp; Footer</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="insert-date-time" title="Insert Current Date & Time">
              <span class="r-icon">📅</span>
              <span class="r-label">Date &amp; Time</span>
            </button>
            <button class="r-btn r-btn-large" data-action="open-equation-dialog" title="Insert Mathematical Equation">
              <span class="r-icon">π</span>
              <span class="r-label">Equation</span>
            </button>
            <button class="r-btn r-btn-large" data-action="open-symbol-dialog" title="Insert Special Symbol">
              <span class="r-icon">Ω</span>
              <span class="r-label">Symbol</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Symbols</div>
        </div>
      `;
    } else if (tab === 'design') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="apply-theme" data-theme="office">
              <span class="r-icon">🎨</span>
              <span class="r-label">Office (Aptos)</span>
            </button>
            <button class="r-btn r-btn-large" data-action="apply-theme" data-theme="classic">
              <span class="r-icon">🏛️</span>
              <span class="r-label">Classic (Calibri)</span>
            </button>
            <button class="r-btn r-btn-large" data-action="apply-theme" data-theme="editorial">
              <span class="r-icon">📰</span>
              <span class="r-label">Editorial (Georgia)</span>
            </button>
            <button class="r-btn r-btn-large" data-action="apply-theme" data-theme="academic">
              <span class="r-icon">🎓</span>
              <span class="r-label">Academic (Times)</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Document Formatting</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="open-watermark-dialog">
              <span class="r-icon">💧</span>
              <span class="r-label">Watermark</span>
            </button>
            <label class="r-btn r-btn-large" style="cursor: pointer;">
              <span class="r-icon">🎨</span>
              <span class="r-label">Page Color</span>
              <input type="color" id="design-page-color" value="${state.pageColor}" style="width: 24px; height: 14px; border: none; padding: 0;" />
            </label>
            <button class="r-btn r-btn-large ${state.pageBorder ? 'active' : ''}" data-action="toggle-page-border">
              <span class="r-icon">🖼️</span>
              <span class="r-label">Page Borders</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Page Background</div>
        </div>
      `;
    } else if (tab === 'layout') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <label style="font-size: 10.5px; color: var(--text-secondary);">Margins</label>
              <select class="r-select" id="layout-margins-select">
                <option value="normal" ${state.pageSetup.margins === 'normal' ? 'selected' : ''}>Normal (1" all)</option>
                <option value="narrow" ${state.pageSetup.margins === 'narrow' ? 'selected' : ''}>Narrow (0.5" all)</option>
                <option value="moderate" ${state.pageSetup.margins === 'moderate' ? 'selected' : ''}>Moderate</option>
                <option value="wide" ${state.pageSetup.margins === 'wide' ? 'selected' : ''}>Wide (1" T/B, 2" L/R)</option>
              </select>
            </div>
            <div class="ribbon-col">
              <label style="font-size: 10.5px; color: var(--text-secondary);">Orientation</label>
              <select class="r-select" id="layout-orientation-select">
                <option value="portrait" ${state.pageSetup.orientation === 'portrait' ? 'selected' : ''}>Portrait</option>
                <option value="landscape" ${state.pageSetup.orientation === 'landscape' ? 'selected' : ''}>Landscape</option>
              </select>
            </div>
            <div class="ribbon-col">
              <label style="font-size: 10.5px; color: var(--text-secondary);">Size</label>
              <select class="r-select" id="layout-size-select">
                <option value="letter" ${state.pageSetup.paperSize === 'letter' ? 'selected' : ''}>Letter (8.5" × 11")</option>
                <option value="legal" ${state.pageSetup.paperSize === 'legal' ? 'selected' : ''}>Legal (8.5" × 14")</option>
                <option value="a4" ${state.pageSetup.paperSize === 'a4' ? 'selected' : ''}>A4 (210 × 297 mm)</option>
                <option value="executive" ${state.pageSetup.paperSize === 'executive' ? 'selected' : ''}>Executive (7.25" × 10.5")</option>
              </select>
            </div>
            <div class="ribbon-col">
              <label style="font-size: 10.5px; color: var(--text-secondary);">Columns</label>
              <select class="r-select" id="layout-columns-select">
                <option value="1" ${state.pageSetup.columns === 1 ? 'selected' : ''}>One Column</option>
                <option value="2" ${state.pageSetup.columns === 2 ? 'selected' : ''}>Two Columns</option>
                <option value="3" ${state.pageSetup.columns === 3 ? 'selected' : ''}>Three Columns</option>
              </select>
            </div>
            <button class="r-btn r-btn-large" data-action="insert-page-break">
              <span class="r-icon">✂️</span>
              <span class="r-label">Breaks</span>
            </button>
          </div>
          <div class="ribbon-group-footer">
            <span>Page Setup</span>
            <button class="dialog-launcher" data-action="open-page-setup-dialog" title="Page Setup Dialog">↘</button>
          </div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <button class="r-btn" data-action="indent">⇥ Indent Left +0.5"</button>
              <button class="r-btn" data-action="outdent">⇤ Indent Left -0.5"</button>
            </div>
            <div class="ribbon-col">
              <button class="r-btn" data-action="space-after-increase">↕ Add Space After Paragraph</button>
              <button class="r-btn" data-action="space-after-remove">↕ Remove Space After Paragraph</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Paragraph</div>
        </div>
      `;
    } else if (tab === 'references') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="insert-toc" title="Generate Table of Contents from Headings">
              <span class="r-icon">📑</span>
              <span class="r-label">Table of Contents</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Table of Contents</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="insert-footnote" title="Insert Footnote at End of Document">
              <span class="r-icon">¹</span>
              <span class="r-label">Insert Footnote</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Footnotes</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="insert-citation" title="Insert Formatted Citation">
              <span class="r-icon">❝</span>
              <span class="r-label">Insert Citation</span>
            </button>
            <button class="r-btn r-btn-large" data-action="insert-caption" title="Insert Figure or Table Caption">
              <span class="r-icon">🏷️</span>
              <span class="r-label">Insert Caption</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Citations &amp; Captions</div>
        </div>
      `;
    } else if (tab === 'review') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="open-word-count-dialog" title="Word Count Statistics">
              <span class="r-icon">🔢</span>
              <span class="r-label">Word Count</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Proofing</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="add-comment">
              <span class="r-icon">💬</span>
              <span class="r-label">New Comment</span>
            </button>
            <button class="r-btn r-btn-large ${state.showCommentsPane ? 'active' : ''}" data-action="toggle-comments-pane">
              <span class="r-icon">🗂️</span>
              <span class="r-label">Show Comments</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Comments</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large ${state.trackChanges ? 'active' : ''}" data-action="toggle-track-changes">
              <span class="r-icon">✍️</span>
              <span class="r-label">Track Changes</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Tracking</div>
        </div>
      `;
    } else if (tab === 'view') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large ${state.viewMode === 'read' ? 'active' : ''}" data-action="set-view" data-view="read">
              <span class="r-icon">📖</span>
              <span class="r-label">Read Mode</span>
            </button>
            <button class="r-btn r-btn-large ${state.viewMode === 'print' ? 'active' : ''}" data-action="set-view" data-view="print">
              <span class="r-icon">📄</span>
              <span class="r-label">Print Layout</span>
            </button>
            <button class="r-btn r-btn-large ${state.viewMode === 'web' ? 'active' : ''}" data-action="set-view" data-view="web">
              <span class="r-icon">🌐</span>
              <span class="r-label">Web Layout</span>
            </button>
            <button class="r-btn r-btn-large ${state.viewMode === 'outline' ? 'active' : ''}" data-action="set-view" data-view="outline">
              <span class="r-icon">📑</span>
              <span class="r-label">Outline</span>
            </button>
            <button class="r-btn r-btn-large ${state.viewMode === 'draft' ? 'active' : ''}" data-action="set-view" data-view="draft">
              <span class="r-icon">📝</span>
              <span class="r-label">Draft</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Views</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large ${state.focusMode ? 'active' : ''}" data-action="toggle-focus">
              <span class="r-icon">🎯</span>
              <span class="r-label">Focus</span>
            </button>
            <button class="r-btn r-btn-large ${state.darkMode ? 'active' : ''}" data-action="toggle-dark-mode">
              <span class="r-icon">🌙</span>
              <span class="r-label">Dark Mode</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Immersive</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                <input type="checkbox" id="view-ruler-chk" ${state.showRuler ? 'checked' : ''} /> Ruler
              </label>
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                <input type="checkbox" id="view-gridlines-chk" ${state.showGridlines ? 'checked' : ''} /> Gridlines
              </label>
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                <input type="checkbox" id="view-navpane-chk" ${state.showNavPane ? 'checked' : ''} /> Navigation Pane
              </label>
            </div>
          </div>
          <div class="ribbon-group-footer">Show</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="reset-zoom">
              <span class="r-icon">🔍</span>
              <span class="r-label">100%</span>
            </button>
            <div class="ribbon-col">
              <button class="r-btn" data-action="zoom-page-width">↔ Page Width (120%)</button>
              <button class="r-btn" data-action="zoom-one-page">📄 One Page (80%)</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Zoom</div>
        </div>
      `;
    } else if (tab === 'table') {
      ribbonSurface.innerHTML = `
        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <button class="r-btn r-btn-large" data-action="table-insert-row-above">
              <span class="r-icon">⬆️</span>
              <span class="r-label">Insert Above</span>
            </button>
            <button class="r-btn r-btn-large" data-action="table-insert-row-below">
              <span class="r-icon">⬇️</span>
              <span class="r-label">Insert Below</span>
            </button>
            <button class="r-btn r-btn-large" data-action="table-insert-col-left">
              <span class="r-icon">⬅️</span>
              <span class="r-label">Insert Left</span>
            </button>
            <button class="r-btn r-btn-large" data-action="table-insert-col-right">
              <span class="r-icon">➡️</span>
              <span class="r-label">Insert Right</span>
            </button>
          </div>
          <div class="ribbon-group-footer">Rows &amp; Columns</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <button class="r-btn" data-action="table-delete-row">🗑️ Delete Row</button>
              <button class="r-btn" data-action="table-delete-col">🗑️ Delete Column</button>
              <button class="r-btn" data-action="table-delete-table">❌ Delete Table</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Delete</div>
        </div>

        <div class="ribbon-group">
          <div class="ribbon-group-body">
            <div class="ribbon-col">
              <button class="r-btn" data-action="table-merge-right">⊞ Merge Right</button>
              <button class="r-btn" data-action="table-split-cell">⊟ Split Cell</button>
              <button class="r-btn" data-action="table-shade-cell">🪣 Shade Cell</button>
            </div>
          </div>
          <div class="ribbon-group-footer">Merge &amp; Shading</div>
        </div>
      `;
    }

    bindRibbonDynamicInputs();
  }

  function bindRibbonDynamicInputs() {
    const fontSelect = container.querySelector('#ribbon-font-family');
    if (fontSelect) {
      fontSelect.addEventListener('change', e => {
        applyInlineStyle('fontFamily', e.target.value);
      });
    }

    const sizeSelect = container.querySelector('#ribbon-font-size');
    if (sizeSelect) {
      sizeSelect.addEventListener('change', e => {
        applyFontSizePt(Number(e.target.value));
      });
    }

    const caseSelect = container.querySelector('#ribbon-change-case');
    if (caseSelect) {
      caseSelect.addEventListener('change', e => {
        if (e.target.value) {
          transformSelectedCase(e.target.value);
          e.target.value = '';
        }
      });
    }

    const highlightInput = container.querySelector('#ribbon-highlight-color');
    if (highlightInput) {
      highlightInput.addEventListener('input', e => {
        applyInlineStyle('backgroundColor', e.target.value);
      });
    }

    const fontColorInput = container.querySelector('#ribbon-font-color');
    if (fontColorInput) {
      fontColorInput.addEventListener('input', e => {
        applyInlineStyle('color', e.target.value);
      });
    }

    const lineSpacingSelect = container.querySelector('#ribbon-line-spacing');
    if (lineSpacingSelect) {
      lineSpacingSelect.addEventListener('change', e => {
        applyBlockStyle('lineHeight', e.target.value);
      });
    }

    const pageColorInput = container.querySelector('#design-page-color');
    if (pageColorInput) {
      pageColorInput.addEventListener('input', e => {
        state.pageColor = e.target.value;
        applyPageSetupToDom();
      });
    }

    const marginsSelect = container.querySelector('#layout-margins-select');
    if (marginsSelect) {
      marginsSelect.addEventListener('change', e => {
        state.pageSetup.margins = e.target.value;
        state.pageSetup.customMarginsPx = null;
        applyPageSetupToDom();
        markDirty();
      });
    }

    const orientSelect = container.querySelector('#layout-orientation-select');
    if (orientSelect) {
      orientSelect.addEventListener('change', e => {
        state.pageSetup.orientation = e.target.value;
        applyPageSetupToDom();
        markDirty();
      });
    }

    const sizeLayoutSelect = container.querySelector('#layout-size-select');
    if (sizeLayoutSelect) {
      sizeLayoutSelect.addEventListener('change', e => {
        state.pageSetup.paperSize = e.target.value;
        applyPageSetupToDom();
        markDirty();
      });
    }

    const colsSelect = container.querySelector('#layout-columns-select');
    if (colsSelect) {
      colsSelect.addEventListener('change', e => {
        state.pageSetup.columns = Number(e.target.value);
        applyPageSetupToDom();
        markDirty();
      });
    }

    const rulerChk = container.querySelector('#view-ruler-chk');
    if (rulerChk) {
      rulerChk.addEventListener('change', e => {
        state.showRuler = e.target.checked;
        container.querySelector('#ruler-bar').classList.toggle('hidden', !state.showRuler);
      });
    }

    const gridChk = container.querySelector('#view-gridlines-chk');
    if (gridChk) {
      gridChk.addEventListener('change', e => {
        state.showGridlines = e.target.checked;
        canvasViewport.classList.toggle('gridlines-on', state.showGridlines);
      });
    }

    const navChk = container.querySelector('#view-navpane-chk');
    if (navChk) {
      navChk.addEventListener('change', e => {
        state.showNavPane = e.target.checked;
        navPane.classList.toggle('hidden', !state.showNavPane);
        if (state.showNavPane) renderNavPane();
      });
    }
  }

  // Formatting & Block Manipulation Helpers
  function getActiveBlockElement() {
    const sel = window.getSelection ? window.getSelection() : null;
    if (sel && sel.rangeCount > 0) {
      let node = sel.anchorNode;
      while (node && node !== editorBody) {
        if (node.nodeType === 1 && /^(P|H1|H2|H3|H4|BLOCKQUOTE|LI|TD|TH)$/i.test(node.tagName)) {
          return node;
        }
        node = node.parentNode;
      }
    }
    return editorBody.firstElementChild || editorBody;
  }

  function applyInlineStyle(cssProp, value) {
    const sel = window.getSelection ? window.getSelection() : null;
    if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
      const range = sel.getRangeAt(0);
      const span = document.createElement('span');
      span.style[cssProp] = value;
      try {
        range.surroundContents(span);
      } catch (_) {
        const frag = range.extractContents();
        span.appendChild(frag);
        range.insertNode(span);
      }
    } else {
      const block = getActiveBlockElement();
      if (block && block.style) block.style[cssProp] = value;
    }
    pushHistory();
    markDirty();
  }

  function applyFontSizePt(pt) {
    applyInlineStyle('fontSize', `${pt}pt`);
  }

  function applyBlockStyle(cssProp, value) {
    const block = getActiveBlockElement();
    if (block && block !== editorBody) {
      block.style[cssProp] = value;
    } else {
      editorBody.style[cssProp] = value;
    }
    pushHistory();
    markDirty();
  }

  function applyQuickStyle(styleId) {
    const styleDef = QUICK_STYLES.find(s => s.id === styleId);
    if (!styleDef) return;

    const block = getActiveBlockElement();
    if (!block || block === editorBody) return;

    const replacement = document.createElement(styleDef.tag);
    if (styleDef.className) replacement.className = styleDef.className;
    if (styleId === 'no-spacing') replacement.style.margin = '0';
    replacement.innerHTML = block.innerHTML;
    block.replaceWith(replacement);

    pushHistory();
    markDirty();
  }

  function transformSelectedCase(mode) {
    const block = getActiveBlockElement();
    if (!block) return;
    const text = block.textContent || '';
    if (!text) return;

    let next = text;
    if (mode === 'upper') next = text.toUpperCase();
    else if (mode === 'lower') next = text.toLowerCase();
    else if (mode === 'capitalize') {
      next = text.replace(/\b\w/g, c => c.toUpperCase());
    } else if (mode === 'sentence') {
      next = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
    }
    block.textContent = next;
    pushHistory();
    markDirty();
  }

  function insertHtmlAtCursor(htmlSnippet) {
    editorBody.focus();
    const sel = window.getSelection ? window.getSelection() : null;
    if (sel && sel.rangeCount > 0 && editorBody.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const temp = document.createElement('div');
      temp.innerHTML = htmlSnippet;
      const frag = document.createDocumentFragment();
      let lastNode = null;
      while (temp.firstChild) {
        lastNode = frag.appendChild(temp.firstChild);
      }
      range.insertNode(frag);
      if (lastNode) {
        range.setStartAfter(lastNode);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    } else {
      editorBody.insertAdjacentHTML('beforeend', htmlSnippet);
    }
    pushHistory();
    markDirty();
  }

  // Table Operations
  function getActiveTableCell() {
    if (state.activeCell && editorBody.contains(state.activeCell)) {
      return state.activeCell;
    }
    return editorBody.querySelector('td, th');
  }

  function insertTable(rows = 3, cols = 3) {
    const rCount = Math.max(1, Math.min(30, Number(rows) || 3));
    const cCount = Math.max(1, Math.min(15, Number(cols) || 3));
    let html = '<table><tbody>';
    for (let r = 0; r < rCount; r++) {
      html += '<tr>';
      for (let c = 0; c < cCount; c++) {
        if (r === 0) {
          html += `<th>Header ${c + 1}</th>`;
        } else {
          html += `<td>Cell ${r},${c + 1}</td>`;
        }
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    insertHtmlAtCursor(html);
  }

  function handleTableAction(action) {
    const cell = getActiveTableCell();
    if (!cell) return;
    const row = cell.parentElement;
    const table = cell.closest('table');
    if (!row || !table) return;
    const colIndex = Array.from(row.children).indexOf(cell);

    if (action === 'table-insert-row-above' || action === 'table-insert-row-below') {
      const newRow = document.createElement('tr');
      const count = row.children.length;
      for (let i = 0; i < count; i++) {
        const td = document.createElement('td');
        td.innerHTML = '<br>';
        newRow.appendChild(td);
      }
      if (action === 'table-insert-row-above') row.before(newRow);
      else row.after(newRow);
    } else if (action === 'table-insert-col-left' || action === 'table-insert-col-right') {
      table.querySelectorAll('tr').forEach((tr, rIdx) => {
        const refCell = tr.children[colIndex] || tr.lastElementChild;
        const newCell = document.createElement(rIdx === 0 && refCell.tagName === 'TH' ? 'th' : 'td');
        newCell.innerHTML = rIdx === 0 ? 'Header' : '<br>';
        if (action === 'table-insert-col-left') refCell.before(newCell);
        else refCell.after(newCell);
      });
    } else if (action === 'table-delete-row') {
      if (table.querySelectorAll('tr').length <= 1) table.remove();
      else row.remove();
    } else if (action === 'table-delete-col') {
      table.querySelectorAll('tr').forEach(tr => {
        if (tr.children[colIndex]) tr.children[colIndex].remove();
      });
      if (table.querySelectorAll('td, th').length === 0) table.remove();
    } else if (action === 'table-delete-table') {
      table.remove();
    } else if (action === 'table-merge-right') {
      const nextCell = cell.nextElementSibling;
      if (nextCell) {
        const span = (Number(cell.getAttribute('colspan')) || 1) + (Number(nextCell.getAttribute('colspan')) || 1);
        cell.setAttribute('colspan', String(span));
        cell.innerHTML += ' ' + nextCell.innerHTML;
        nextCell.remove();
      }
    } else if (action === 'table-split-cell') {
      const span = Number(cell.getAttribute('colspan')) || 1;
      if (span > 1) {
        cell.setAttribute('colspan', String(span - 1));
        const newCell = document.createElement('td');
        newCell.innerHTML = '<br>';
        cell.after(newCell);
      }
    } else if (action === 'table-shade-cell') {
      cell.style.backgroundColor = cell.style.backgroundColor ? '' : '#D0E4F9';
    }

    pushHistory();
    markDirty();
  }

  // Find and Replace Engine
  function findTextMatches(query, matchCase = false) {
    if (!query) return 0;
    const content = editorBody.innerText || editorBody.textContent || '';
    const flags = matchCase ? 'g' : 'gi';
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const matches = content.match(new RegExp(escaped, flags));
    return matches ? matches.length : 0;
  }

  function replaceAllMatches(findStr, replaceStr, matchCase = false) {
    if (!findStr) return 0;
    const flags = matchCase ? 'g' : 'gi';
    const escaped = findStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, flags);

    let count = 0;
    const walker = document.createTreeWalker(editorBody, NodeFilter.SHOW_TEXT);
    const textNodes = [];
    while (walker.nextNode()) textNodes.push(walker.currentNode);

    for (const node of textNodes) {
      const val = node.nodeValue || '';
      const matches = val.match(regex);
      if (matches) {
        count += matches.length;
        node.nodeValue = val.replace(regex, replaceStr);
      }
    }
    if (count > 0) {
      pushHistory();
      markDirty();
    }
    return count;
  }

  // Navigation Pane Renderer
  function renderNavPane() {
    container.querySelectorAll('.nav-tab').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.navTab === state.navPaneTab);
    });

    if (state.navPaneTab === 'headings') {
      const headings = Array.from(editorBody.querySelectorAll('h1, h2, h3, h4'));
      if (headings.length === 0) {
        navContent.innerHTML = `<p style="color: var(--text-muted); font-size: 12px;">Add Heading 1, Heading 2, or Heading 3 styles to build an interactive document outline.</p>`;
        return;
      }
      navContent.innerHTML = headings.map((h, idx) => {
        const level = h.tagName === 'H1' ? 1 : h.tagName === 'H2' ? 2 : 3;
        return `<div class="nav-item level-${level}" data-heading-idx="${idx}">${h.textContent || 'Untitled Heading'}</div>`;
      }).join('');
      navContent.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
          const target = headings[Number(item.dataset.headingIdx)];
          if (target && target.scrollIntoView) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
      });
    } else if (state.navPaneTab === 'pages') {
      const stats = getDocumentStats();
      const cards = [];
      for (let p = 1; p <= stats.pages; p++) {
        cards.push(`<div class="page-thumb"><strong>Page ${p}</strong><div style="color:var(--text-muted);margin-top:4px;">Print Layout Sheet ${p}</div></div>`);
      }
      navContent.innerHTML = cards.join('');
    } else if (state.navPaneTab === 'results') {
      const q = state.navSearchQuery.trim();
      if (!q) {
        navContent.innerHTML = `<p style="color: var(--text-muted); font-size: 12px;">Type a phrase in the search box above to find matches.</p>`;
        return;
      }
      const blocks = Array.from(editorBody.querySelectorAll('p, h1, h2, h3, h4, li, td, blockquote'));
      const matched = blocks.filter(b => (b.textContent || '').toLowerCase().includes(q.toLowerCase()));
      if (matched.length === 0) {
        navContent.innerHTML = `<p style="color: var(--text-muted); font-size: 12px;">No matches found for "${q}".</p>`;
        return;
      }
      navContent.innerHTML = `<div style="margin-bottom:8px;font-weight:600;">${matched.length} result(s)</div>` +
        matched.map((m, idx) => `<div class="nav-item" data-match-idx="${idx}">${(m.textContent || '').slice(0, 70)}</div>`).join('');
      navContent.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
          const target = matched[Number(item.dataset.matchIdx)];
          if (target && target.scrollIntoView) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
      });
    }
  }

  // Comments Sidebar
  function addComment(text = '') {
    const commentText = text || 'Review section formatting and metrics.';
    const commentObj = {
      id: `c_${Date.now()}_${state.comments.length + 1}`,
      author: state.metadata.author || 'Word User',
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      text: commentText
    };
    state.comments.push(commentObj);
    state.showCommentsPane = true;
    commentsPane.classList.remove('hidden');
    renderComments();
    markDirty();
  }

  function renderComments() {
    if (state.comments.length === 0) {
      commentsList.innerHTML = `<p style="color: var(--text-muted); font-size: 12px;">No comments in this document yet.</p>`;
      return;
    }
    commentsList.innerHTML = state.comments.map(c => `
      <div class="comment-card">
        <div class="comment-card-header">
          <span>${c.author}</span>
          <span>${c.timestamp}</span>
        </div>
        <div>${c.text}</div>
      </div>
    `).join('');
  }

  // Table of Contents & Footnotes
  function insertTableOfContents() {
    const headings = Array.from(editorBody.querySelectorAll('h1, h2, h3'));
    const itemsHtml = headings.length > 0
      ? headings.map((h, idx) => {
          const indent = h.tagName === 'H2' ? '16px' : h.tagName === 'H3' ? '32px' : '0px';
          return `<div style="display:flex;justify-content:space-between;padding-left:${indent};margin:3px 0;">
            <span>${h.textContent}</span>
            <span style="color:var(--text-muted);">p. 1</span>
          </div>`;
        }).join('')
      : `<div>1. Executive Overview ........................................ p. 1</div>`;

    const tocHtml = `<div class="word-toc" style="border:1px solid #E1DFDD;background:#FAF9F8;padding:14px 18px;margin:12px 0;border-radius:4px;">
      <p style="font-family:'Aptos Display',sans-serif;font-size:16pt;color:#0F4761;margin:0 0 8px 0;">Table of Contents</p>
      ${itemsHtml}
    </div><p><br></p>`;

    editorBody.insertAdjacentHTML('afterbegin', tocHtml);
    pushHistory();
    markDirty();
  }

  function insertFootnote(noteText = 'Reference source verified against Microsoft Word 16.0 specification.') {
    const num = state.footnotes.length + 1;
    state.footnotes.push({ num, text: noteText });
    insertHtmlAtCursor(`<sup class="footnote-ref">[${num}]</sup>`);

    let fnSection = editorBody.querySelector('.footnotes-section');
    if (!fnSection) {
      fnSection = document.createElement('div');
      fnSection.className = 'footnotes-section';
      fnSection.style.borderTop = '1px solid #8A8886';
      fnSection.style.marginTop = '32px';
      fnSection.style.paddingTop = '8px';
      fnSection.style.fontSize = '9.5pt';
      editorBody.appendChild(fnSection);
    }
    const p = document.createElement('p');
    p.innerHTML = `<sup>[${num}]</sup> ${noteText}`;
    fnSection.appendChild(p);
    pushHistory();
    markDirty();
  }

  // Modal Dialogs
  function openModal(title, bodyHtml, onConfirm, confirmLabel = 'OK') {
    modalDialog.innerHTML = `
      <div class="modal-header">
        <span>${title}</span>
        <button class="r-btn r-btn-icon" id="modal-close-x">✕</button>
      </div>
      <div class="modal-body">${bodyHtml}</div>
      <div class="modal-footer">
        <button class="btn-secondary" id="modal-cancel-btn">Cancel</button>
        ${onConfirm ? `<button class="btn-primary" id="modal-confirm-btn">${confirmLabel}</button>` : ''}
      </div>
    `;
    modalBackdrop.classList.remove('hidden');

    const closeModal = () => modalBackdrop.classList.add('hidden');
    modalDialog.querySelector('#modal-close-x').addEventListener('click', closeModal);
    modalDialog.querySelector('#modal-cancel-btn').addEventListener('click', closeModal);
    const confirmBtn = modalDialog.querySelector('#modal-confirm-btn');
    if (confirmBtn && onConfirm) {
      confirmBtn.addEventListener('click', () => {
        const shouldClose = onConfirm(modalDialog);
        if (shouldClose !== false) closeModal();
      });
    }
  }

  function openFontDialog() {
    openModal(
      'Font',
      `
      <div class="form-grid">
        <div class="form-field">
          <label>Font Family</label>
          <select id="dlg-font-family">
            <option value="Aptos">Aptos</option>
            <option value="Aptos Display">Aptos Display</option>
            <option value="Calibri">Calibri</option>
            <option value="Segoe UI">Segoe UI</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Arial">Arial</option>
            <option value="Georgia">Georgia</option>
            <option value="Consolas">Consolas</option>
          </select>
        </div>
        <div class="form-field">
          <label>Font Size (pt)</label>
          <input type="number" id="dlg-font-size" value="11" min="6" max="144" />
        </div>
        <div class="form-field">
          <label>Font Color</label>
          <input type="color" id="dlg-font-color" value="#201F1E" />
        </div>
        <div class="form-field">
          <label>Effects</label>
          <div style="display:flex;gap:10px;padding-top:6px;">
            <label><input type="checkbox" id="dlg-font-bold" /> Bold</label>
            <label><input type="checkbox" id="dlg-font-italic" /> Italic</label>
            <label><input type="checkbox" id="dlg-font-underline" /> Underline</label>
          </div>
        </div>
      </div>
      `,
      dlg => {
        const fam = dlg.querySelector('#dlg-font-family').value;
        const sz = Number(dlg.querySelector('#dlg-font-size').value) || 11;
        const col = dlg.querySelector('#dlg-font-color').value;
        const bold = dlg.querySelector('#dlg-font-bold').checked;
        const italic = dlg.querySelector('#dlg-font-italic').checked;
        const under = dlg.querySelector('#dlg-font-underline').checked;

        const block = getActiveBlockElement();
        if (block) {
          block.style.fontFamily = fam;
          block.style.fontSize = `${sz}pt`;
          block.style.color = col;
          if (bold) block.style.fontWeight = 'bold';
          if (italic) block.style.fontStyle = 'italic';
          if (under) block.style.textDecoration = 'underline';
        }
        pushHistory();
        markDirty();
      },
      'Apply'
    );
  }

  function openParagraphDialog() {
    openModal(
      'Paragraph',
      `
      <div class="form-grid">
        <div class="form-field">
          <label>Alignment</label>
          <select id="dlg-para-align">
            <option value="left">Left</option>
            <option value="center">Centered</option>
            <option value="right">Right</option>
            <option value="justify">Justified</option>
          </select>
        </div>
        <div class="form-field">
          <label>Line Spacing</label>
          <select id="dlg-para-line">
            <option value="1">Single (1.0)</option>
            <option value="1.15" selected>Multiple (1.15)</option>
            <option value="1.5">1.5 lines</option>
            <option value="2">Double (2.0)</option>
          </select>
        </div>
        <div class="form-field">
          <label>Space Before (pt)</label>
          <input type="number" id="dlg-para-before" value="0" min="0" max="120" />
        </div>
        <div class="form-field">
          <label>Space After (pt)</label>
          <input type="number" id="dlg-para-after" value="8" min="0" max="120" />
        </div>
      </div>
      `,
      dlg => {
        const block = getActiveBlockElement();
        if (block) {
          block.style.textAlign = dlg.querySelector('#dlg-para-align').value;
          block.style.lineHeight = dlg.querySelector('#dlg-para-line').value;
          block.style.marginTop = `${Number(dlg.querySelector('#dlg-para-before').value) || 0}pt`;
          block.style.marginBottom = `${Number(dlg.querySelector('#dlg-para-after').value) || 0}pt`;
        }
        pushHistory();
        markDirty();
      },
      'Apply'
    );
  }

  function openPageSetupDialog() {
    const currentMar = state.pageSetup.customMarginsPx || MARGIN_PRESETS_PX[state.pageSetup.margins] || MARGIN_PRESETS_PX.normal;
    openModal(
      'Page Setup',
      `
      <div class="form-grid">
        <div class="form-field">
          <label>Top Margin (inches)</label>
          <input type="number" step="0.25" id="dlg-mar-top" value="${(currentMar.top / 96).toFixed(2)}" />
        </div>
        <div class="form-field">
          <label>Bottom Margin (inches)</label>
          <input type="number" step="0.25" id="dlg-mar-bottom" value="${(currentMar.bottom / 96).toFixed(2)}" />
        </div>
        <div class="form-field">
          <label>Left Margin (inches)</label>
          <input type="number" step="0.25" id="dlg-mar-left" value="${(currentMar.left / 96).toFixed(2)}" />
        </div>
        <div class="form-field">
          <label>Right Margin (inches)</label>
          <input type="number" step="0.25" id="dlg-mar-right" value="${(currentMar.right / 96).toFixed(2)}" />
        </div>
        <div class="form-field">
          <label>Orientation</label>
          <select id="dlg-pg-orient">
            <option value="portrait" ${state.pageSetup.orientation === 'portrait' ? 'selected' : ''}>Portrait</option>
            <option value="landscape" ${state.pageSetup.orientation === 'landscape' ? 'selected' : ''}>Landscape</option>
          </select>
        </div>
        <div class="form-field">
          <label>Paper Size</label>
          <select id="dlg-pg-size">
            <option value="letter" ${state.pageSetup.paperSize === 'letter' ? 'selected' : ''}>Letter (8.5" × 11")</option>
            <option value="legal" ${state.pageSetup.paperSize === 'legal' ? 'selected' : ''}>Legal (8.5" × 14")</option>
            <option value="a4" ${state.pageSetup.paperSize === 'a4' ? 'selected' : ''}>A4</option>
            <option value="executive" ${state.pageSetup.paperSize === 'executive' ? 'selected' : ''}>Executive</option>
          </select>
        </div>
      </div>
      `,
      dlg => {
        const topIn = parseFloat(dlg.querySelector('#dlg-mar-top').value) || 1;
        const bottomIn = parseFloat(dlg.querySelector('#dlg-mar-bottom').value) || 1;
        const leftIn = parseFloat(dlg.querySelector('#dlg-mar-left').value) || 1;
        const rightIn = parseFloat(dlg.querySelector('#dlg-mar-right').value) || 1;
        state.pageSetup.customMarginsPx = {
          top: Math.round(topIn * 96),
          bottom: Math.round(bottomIn * 96),
          left: Math.round(leftIn * 96),
          right: Math.round(rightIn * 96)
        };
        state.pageSetup.orientation = dlg.querySelector('#dlg-pg-orient').value;
        state.pageSetup.paperSize = dlg.querySelector('#dlg-pg-size').value;
        applyPageSetupToDom();
        renderRibbon();
        markDirty();
      },
      'OK'
    );
  }

  function openTableDialog() {
    openModal(
      'Insert Table',
      `
      <div class="form-grid">
        <div class="form-field">
          <label>Number of columns</label>
          <input type="number" id="dlg-tbl-cols" value="3" min="1" max="12" />
        </div>
        <div class="form-field">
          <label>Number of rows</label>
          <input type="number" id="dlg-tbl-rows" value="4" min="1" max="30" />
        </div>
      </div>
      `,
      dlg => {
        const cols = Number(dlg.querySelector('#dlg-tbl-cols').value) || 3;
        const rows = Number(dlg.querySelector('#dlg-tbl-rows').value) || 4;
        insertTable(rows, cols);
      },
      'Insert'
    );
  }

  function openHyperlinkDialog() {
    openModal(
      'Insert Hyperlink',
      `
      <div class="form-field">
        <label>Text to display</label>
        <input type="text" id="dlg-link-text" value="Microsoft Word Reference Documentation" />
      </div>
      <div class="form-field">
        <label>Address (URL)</label>
        <input type="url" id="dlg-link-url" value="https://learn.microsoft.com/en-us/office/" />
      </div>
      `,
      dlg => {
        const text = dlg.querySelector('#dlg-link-text').value.trim() || 'Link';
        const url = dlg.querySelector('#dlg-link-url').value.trim() || 'https://example.com';
        insertHtmlAtCursor(`<a href="${url}" target="_blank" rel="noopener noreferrer">${text}</a>`);
      },
      'Insert'
    );
  }

  function openFindReplaceDialog(focusReplace = false) {
    openModal(
      'Find and Replace',
      `
      <div class="form-field">
        <label>Find what</label>
        <input type="text" id="dlg-find-input" placeholder="Enter search text..." />
      </div>
      <div class="form-field">
        <label>Replace with</label>
        <input type="text" id="dlg-replace-input" placeholder="Enter replacement text..." />
      </div>
      <label style="display:flex;align-items:center;gap:6px;font-size:12px;">
        <input type="checkbox" id="dlg-match-case" /> Match case
      </label>
      <div id="dlg-find-status" style="font-size:12px;color:var(--word-blue-primary);min-height:18px;"></div>
      <div style="display:flex;gap:8px;">
        <button type="button" class="btn-secondary" id="dlg-find-count-btn">Count Matches</button>
        <button type="button" class="btn-primary" id="dlg-replace-all-btn">Replace All</button>
      </div>
      `,
      null
    );

    const findInput = modalDialog.querySelector('#dlg-find-input');
    const replaceInput = modalDialog.querySelector('#dlg-replace-input');
    const matchCaseChk = modalDialog.querySelector('#dlg-match-case');
    const statusEl = modalDialog.querySelector('#dlg-find-status');

    modalDialog.querySelector('#dlg-find-count-btn').addEventListener('click', () => {
      const count = findTextMatches(findInput.value, matchCaseChk.checked);
      statusEl.textContent = `Found ${count} match(es) in document.`;
    });

    modalDialog.querySelector('#dlg-replace-all-btn').addEventListener('click', () => {
      const replaced = replaceAllMatches(findInput.value, replaceInput.value, matchCaseChk.checked);
      statusEl.textContent = `All done. We made ${replaced} replacement(s).`;
    });

    setTimeout(() => {
      if (focusReplace) replaceInput.focus();
      else findInput.focus();
    }, 20);
  }

  function openWordCountDialog() {
    const stats = getDocumentStats();
    openModal(
      'Word Count',
      `
      <div style="font-size:12.5px;display:flex;flex-direction:column;gap:8px;">
        <div style="border-bottom:1px solid var(--ribbon-border);padding-bottom:6px;font-weight:600;">Statistics:</div>
        <div style="display:flex;justify-content:space-between;"><span>Pages</span><strong>${stats.pages}</strong></div>
        <div style="display:flex;justify-content:space-between;"><span>Words</span><strong>${stats.words}</strong></div>
        <div style="display:flex;justify-content:space-between;"><span>Characters (no spaces)</span><strong>${stats.charsNoSpaces}</strong></div>
        <div style="display:flex;justify-content:space-between;"><span>Characters (with spaces)</span><strong>${stats.charsWithSpaces}</strong></div>
        <div style="display:flex;justify-content:space-between;"><span>Paragraphs</span><strong>${stats.paragraphs}</strong></div>
        <div style="display:flex;justify-content:space-between;"><span>Lines</span><strong>${stats.lines}</strong></div>
      </div>
      `,
      () => true,
      'Close'
    );
  }

  function openEquationDialog() {
    const equations = [
      { label: 'Area of a Circle', tex: 'A = πr²' },
      { label: 'Quadratic Formula', tex: 'x = (−b ± √(b² − 4ac)) / 2a' },
      { label: 'Pythagorean Theorem', tex: 'a² + b² = c²' },
      { label: "Euler's Identity", tex: 'e^(iπ) + 1 = 0' },
      { label: 'Gaussian Integral', tex: '∫ e^(−x²) dx = √π' }
    ];
    openModal(
      'Insert Equation',
      `
      <div class="form-field">
        <label>Select Built-In Equation or Customize</label>
        <select id="dlg-eq-preset">
          ${equations.map(e => `<option value="${e.tex}">${e.label}: ${e.tex}</option>`).join('')}
        </select>
      </div>
      <div class="form-field">
        <label>Equation Expression</label>
        <input type="text" id="dlg-eq-custom" value="A = πr²" />
      </div>
      `,
      dlg => {
        const expr = dlg.querySelector('#dlg-eq-custom').value || dlg.querySelector('#dlg-eq-preset').value;
        insertHtmlAtCursor(`<span class="word-equation" style="font-family:'Cambria Math',Georgia,serif;font-style:italic;background:#F3F2F1;padding:2px 6px;border-radius:3px;">${expr}</span> `);
      },
      'Insert'
    );
    const presetSel = modalDialog.querySelector('#dlg-eq-preset');
    const customInp = modalDialog.querySelector('#dlg-eq-custom');
    presetSel.addEventListener('change', () => {
      customInp.value = presetSel.value;
    });
  }

  function openSymbolDialog() {
    const symbols = ['©', '®', '™', '§', '¶', '†', '‡', '•', '±', '×', '÷', '≠', '≈', '≤', '≥', '∞', '∑', '∏', '√', '∫', 'Δ', 'Ω', 'π', 'μ', '€', '£', '¥', '¢', '←', '→', '↑', '↓'];
    openModal(
      'Insert Symbol',
      `
      <div style="display:grid;grid-template-columns:repeat(8, 1fr);gap:6px;">
        ${symbols.map(sym => `<button type="button" class="btn-secondary sym-pick-btn" data-sym="${sym}" style="font-size:16px;padding:8px;">${sym}</button>`).join('')}
      </div>
      `,
      null
    );
    modalDialog.querySelectorAll('.sym-pick-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        insertHtmlAtCursor(btn.dataset.sym);
        modalBackdrop.classList.add('hidden');
      });
    });
  }

  function openWatermarkDialog() {
    openModal(
      'Watermark',
      `
      <div class="form-field">
        <label>Watermark Text</label>
        <select id="dlg-wm-select">
          <option value="">(No Watermark)</option>
          <option value="CONFIDENTIAL">CONFIDENTIAL</option>
          <option value="DRAFT">DRAFT</option>
          <option value="DO NOT COPY">DO NOT COPY</option>
          <option value="SAMPLE">SAMPLE</option>
        </select>
      </div>
      `,
      dlg => {
        state.watermark = dlg.querySelector('#dlg-wm-select').value;
        applyPageSetupToDom();
      },
      'Apply'
    );
  }

  // File Operations (New, Open, Save, Save As, Export, Print)
  function loadTemplate(templateKey = 'blank') {
    const tpl = TEMPLATES[templateKey] || TEMPLATES.blank;
    state.documentTitle = tpl.title;
    state.headerText = tpl.headerText;
    state.footerText = tpl.footerText;
    state.currentFilePath = null;
    state.dirty = false;
    docTitleInput.value = state.documentTitle;
    headerEditable.innerHTML = state.headerText;
    footerEditable.innerHTML = state.footerText;
    editorBody.innerHTML = tpl.html;
    pushHistory();
    refreshDocumentMetrics();
    backstageOverlay.classList.add('hidden');
  }

  async function getSerializedState() {
    return {
      html: editorBody.innerHTML,
      headerText: headerEditable.textContent || '',
      footerText: footerEditable.textContent || '',
      showPageNumbers: state.showPageNumbers,
      pageSetup: state.pageSetup,
      metadata: {
        ...state.metadata,
        title: state.documentTitle
      }
    };
  }

  function triggerBrowserDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function saveDocument(format = 'docx', forceSaveAs = false) {
    const docState = await getSerializedState();
    const safeBase = (state.documentTitle || 'Document1').replace(/[^a-zA-Z0-9_-]/g, '_');

    if (isDesktop && window.wordDesktop) {
      let dataBytes;
      let ext = format;
      if (format === 'docx') {
        dataBytes = await exportDocx(docState);
      } else if (format === 'html') {
        dataBytes = new TextEncoder().encode(exportHtmlDocument(docState));
      } else if (format === 'rtf') {
        dataBytes = new TextEncoder().encode(exportRtfDocument(docState));
      } else {
        dataBytes = new TextEncoder().encode(editorBody.innerText || editorBody.textContent || '');
        ext = 'txt';
      }

      const res = await window.wordDesktop.saveFile({
        defaultPath: state.currentFilePath && !forceSaveAs ? state.currentFilePath : `${safeBase}.${ext}`,
        forceSaveAs,
        format: ext,
        data: Array.from(dataBytes)
      });

      if (res && res.filePath) {
        state.currentFilePath = res.filePath;
        state.dirty = false;
        saveBadge.textContent = 'Saved to PC';
        addRecentFile(res.filePath);
        refreshDocumentMetrics();
      }
      backstageOverlay.classList.add('hidden');
      return;
    }

    // Browser Save / Download
    if (format === 'docx') {
      const uint8 = await exportDocx(docState);
      const blob = new Blob([uint8], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      triggerBrowserDownload(blob, `${safeBase}.docx`);
    } else if (format === 'html') {
      const htmlStr = exportHtmlDocument(docState);
      triggerBrowserDownload(new Blob([htmlStr], { type: 'text/html' }), `${safeBase}.html`);
    } else if (format === 'rtf') {
      const rtfStr = exportRtfDocument(docState);
      triggerBrowserDownload(new Blob([rtfStr], { type: 'application/rtf' }), `${safeBase}.rtf`);
    } else if (format === 'txt') {
      const txtStr = editorBody.innerText || editorBody.textContent || '';
      triggerBrowserDownload(new Blob([txtStr], { type: 'text/plain' }), `${safeBase}.txt`);
    }

    state.dirty = false;
    saveBadge.textContent = 'Saved';
    addRecentFile(`${safeBase}.${format}`);
    backstageOverlay.classList.add('hidden');
  }

  function addRecentFile(nameOrPath) {
    state.recentFiles = [
      { name: nameOrPath, date: new Date().toLocaleString() },
      ...state.recentFiles.filter(f => f.name !== nameOrPath)
    ].slice(0, 10);
    try {
      localStorage.setItem('word_recent_files', JSON.stringify(state.recentFiles));
    } catch (_) {}
  }

  async function openDocumentFromBuffer(name, arrayBuffer) {
    const lower = name.toLowerCase();
    const baseTitle = name.replace(/\.[^.]+$/, '');
    state.documentTitle = baseTitle;
    docTitleInput.value = baseTitle;

    if (lower.endsWith('.docx')) {
      const imported = await importDocx(arrayBuffer);
      editorBody.innerHTML = imported.html;
      headerEditable.textContent = imported.headerText || '';
      footerEditable.textContent = imported.footerText || '';
      state.pageSetup = { ...state.pageSetup, ...imported.pageSetup };
      state.metadata = { ...state.metadata, ...imported.metadata };
      applyPageSetupToDom();
    } else if (lower.endsWith('.rtf')) {
      const text = new TextDecoder().decode(arrayBuffer);
      editorBody.innerHTML = importRtfDocument(text);
    } else if (lower.endsWith('.html') || lower.endsWith('.htm')) {
      const text = new TextDecoder().decode(arrayBuffer);
      const doc = new DOMParser().parseFromString(text, 'text/html');
      const mainEl = doc.querySelector('main') || doc.body;
      editorBody.innerHTML = mainEl.innerHTML;
    } else {
      const text = new TextDecoder().decode(arrayBuffer);
      editorBody.innerHTML = text
        .split(/\r?\n/)
        .map(line => `<p>${line || '<br>'}</p>`)
        .join('');
    }

    state.dirty = false;
    saveBadge.textContent = 'Opened';
    addRecentFile(name);
    pushHistory();
    refreshDocumentMetrics();
    backstageOverlay.classList.add('hidden');
  }

  async function triggerOpenDocument() {
    if (isDesktop && window.wordDesktop && window.wordDesktop.openFile) {
      const res = await window.wordDesktop.openFile();
      if (res && res.data && res.filePath) {
        state.currentFilePath = res.filePath;
        const uint8 = new Uint8Array(res.data);
        await openDocumentFromBuffer(res.fileName || res.filePath, uint8.buffer);
      }
      return;
    }
    fileOpenInput.click();
  }

  async function exportPdf() {
    backstageOverlay.classList.add('hidden');
    if (isDesktop && window.wordDesktop && window.wordDesktop.exportPdf) {
      const safeBase = (state.documentTitle || 'Document1').replace(/[^a-zA-Z0-9_-]/g, '_');
      const res = await window.wordDesktop.exportPdf({ defaultPath: `${safeBase}.pdf` });
      if (res && res.filePath) {
        saveBadge.textContent = 'Exported PDF';
      }
      return;
    }
    window.print();
  }

  // Backstage View Renderer
  function renderBackstage() {
    const sec = state.backstageSection;
    container.querySelectorAll('.backstage-nav-item[data-backstage]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.backstage === sec);
    });

    if (sec === 'new') {
      backstageContent.innerHTML = `
        <h2 style="margin-top:0;">New Document</h2>
        <p style="color:var(--text-secondary);">Choose a document template modeled on Microsoft Word 16.0:</p>
        <div class="template-grid">
          <div class="template-card" data-template="blank">
            <div class="template-preview"><strong>Blank document</strong><br><br>Aptos 11pt<br>1" Normal Margins</div>
            <div><strong>Blank Document</strong></div>
          </div>
          <div class="template-card" data-template="default">
            <div class="template-preview"><strong>Project Proposal</strong><br>1. Executive Overview<br>Table &amp; Milestones</div>
            <div><strong>Project Proposal</strong></div>
          </div>
          <div class="template-card" data-template="report">
            <div class="template-preview"><strong>Quarterly Report</strong><br>KPI Table &amp; Financial Highlights</div>
            <div><strong>Business Report</strong></div>
          </div>
          <div class="template-card" data-template="academic">
            <div class="template-preview"><strong>Research Paper</strong><br>Abstract, Introduction, Methodology</div>
            <div><strong>Academic Paper</strong></div>
          </div>
          <div class="template-card" data-template="resume">
            <div class="template-preview"><strong>Executive Resume</strong><br>Experience &amp; Education</div>
            <div><strong>Executive Resume</strong></div>
          </div>
        </div>
      `;
      backstageContent.querySelectorAll('.template-card').forEach(card => {
        card.addEventListener('click', () => loadTemplate(card.dataset.template));
      });
    } else if (sec === 'open') {
      backstageContent.innerHTML = `
        <h2 style="margin-top:0;">Open</h2>
        <button class="btn-primary" id="bs-open-btn">📂 Browse Local Files (.docx, .html, .rtf, .txt)</button>
        <h3 style="margin-top:28px;">Recent Documents</h3>
        <div>
          ${
            state.recentFiles.length === 0
              ? `<p style="color:var(--text-muted);">No recent files opened yet.</p>`
              : state.recentFiles.map(f => `<div style="padding:8px 0;border-bottom:1px solid var(--ribbon-border);display:flex;justify-content:space-between;"><span>📄 ${f.name}</span><span style="color:var(--text-muted);">${f.date}</span></div>`).join('')
          }
        </div>
      `;
      backstageContent.querySelector('#bs-open-btn').addEventListener('click', triggerOpenDocument);
    } else if (sec === 'info') {
      const stats = getDocumentStats();
      backstageContent.innerHTML = `
        <h2 style="margin-top:0;">Document Information &amp; Properties</h2>
        <div class="form-grid" style="max-width:640px;">
          <div class="form-field">
            <label>Title</label>
            <input type="text" id="info-title" value="${state.documentTitle}" />
          </div>
          <div class="form-field">
            <label>Author</label>
            <input type="text" id="info-author" value="${state.metadata.author}" />
          </div>
          <div class="form-field">
            <label>Subject</label>
            <input type="text" id="info-subject" value="${state.metadata.subject}" />
          </div>
          <div class="form-field">
            <label>Comments</label>
            <input type="text" id="info-comments" value="${state.metadata.comments}" />
          </div>
        </div>
        <div style="margin-top:16px;">
          <button class="btn-primary" id="info-save-props">Update Document Properties</button>
        </div>
        <h3 style="margin-top:28px;">Document Statistics</h3>
        <p>Pages: <strong>${stats.pages}</strong> • Words: <strong>${stats.words}</strong> • Characters: <strong>${stats.charsWithSpaces}</strong> • Paragraphs: <strong>${stats.paragraphs}</strong></p>
      `;
      backstageContent.querySelector('#info-save-props').addEventListener('click', () => {
        state.documentTitle = backstageContent.querySelector('#info-title').value || 'Document1';
        docTitleInput.value = state.documentTitle;
        state.metadata.author = backstageContent.querySelector('#info-author').value;
        state.metadata.subject = backstageContent.querySelector('#info-subject').value;
        state.metadata.comments = backstageContent.querySelector('#info-comments').value;
        backstageOverlay.classList.add('hidden');
        refreshDocumentMetrics();
      });
    } else if (sec === 'save-as') {
      backstageContent.innerHTML = `
        <h2 style="margin-top:0;">Save As</h2>
        <p style="color:var(--text-secondary);">Choose a document format derived from Microsoft Word's Save As dialog:</p>
        <div style="display:flex;flex-direction:column;gap:10px;max-width:420px;margin-top:16px;">
          <button class="btn-primary" data-save-fmt="docx">📘 Word Document (*.docx) — OpenXML Package</button>
          <button class="btn-secondary" data-save-fmt="html">🌐 Web Page (*.htm; *.html)</button>
          <button class="btn-secondary" data-save-fmt="rtf">📝 Rich Text Format (*.rtf)</button>
          <button class="btn-secondary" data-save-fmt="txt">📄 Plain Text (*.txt)</button>
        </div>
      `;
      backstageContent.querySelectorAll('[data-save-fmt]').forEach(btn => {
        btn.addEventListener('click', () => saveDocument(btn.dataset.saveFmt, true));
      });
    } else if (sec === 'export') {
      backstageContent.innerHTML = `
        <h2 style="margin-top:0;">Export</h2>
        <p style="color:var(--text-secondary);">Create a fixed-layout PDF/XPS document or OpenXML archive:</p>
        <div style="display:flex;gap:12px;margin-top:16px;">
          <button class="btn-primary" id="bs-export-pdf">📕 Create PDF/XPS Document (*.pdf)</button>
          <button class="btn-secondary" id="bs-export-docx">📘 Export OpenXML (*.docx)</button>
        </div>
      `;
      backstageContent.querySelector('#bs-export-pdf').addEventListener('click', exportPdf);
      backstageContent.querySelector('#bs-export-docx').addEventListener('click', () => saveDocument('docx', true));
    }
  }

  // Tell Me Command Registry (modeled on TellMeWord.nrr)
  const COMMAND_PALETTE = [
    { label: 'Save Document (.docx)', shortcut: 'Ctrl+S', action: () => saveDocument('docx') },
    { label: 'Open Document...', shortcut: 'Ctrl+O', action: () => triggerOpenDocument() },
    { label: 'Export to PDF / Print', shortcut: 'Ctrl+P', action: () => exportPdf() },
    { label: 'Bold', shortcut: 'Ctrl+B', action: () => executeAction('bold') },
    { label: 'Italic', shortcut: 'Ctrl+I', action: () => executeAction('italic') },
    { label: 'Underline', shortcut: 'Ctrl+U', action: () => executeAction('underline') },
    { label: 'Insert Table...', shortcut: '', action: () => openTableDialog() },
    { label: 'Insert Page Break', shortcut: 'Ctrl+Enter', action: () => executeAction('insert-page-break') },
    { label: 'Insert Hyperlink...', shortcut: 'Ctrl+K', action: () => openHyperlinkDialog() },
    { label: 'Insert Table of Contents', shortcut: '', action: () => insertTableOfContents() },
    { label: 'Insert Footnote', shortcut: '', action: () => insertFootnote() },
    { label: 'Insert Equation', shortcut: '', action: () => openEquationDialog() },
    { label: 'Find and Replace...', shortcut: 'Ctrl+H', action: () => openFindReplaceDialog(false) },
    { label: 'Word Count Statistics...', shortcut: 'Ctrl+Shift+G', action: () => openWordCountDialog() },
    { label: 'Page Setup (Margins & Orientation)...', shortcut: '', action: () => openPageSetupDialog() },
    { label: 'Toggle Dark Mode', shortcut: '', action: () => executeAction('toggle-dark-mode') },
    { label: 'Toggle Focus Mode', shortcut: '', action: () => executeAction('toggle-focus') },
    { label: 'Toggle Navigation Pane', shortcut: 'Ctrl+F', action: () => executeAction('toggle-nav-pane') }
  ];

  tellMeInput.addEventListener('input', () => {
    const q = tellMeInput.value.trim().toLowerCase();
    const filtered = q
      ? COMMAND_PALETTE.filter(c => c.label.toLowerCase().includes(q))
      : COMMAND_PALETTE.slice(0, 8);

    if (filtered.length === 0) {
      tellMeDropdown.style.display = 'none';
      return;
    }
    tellMeDropdown.innerHTML = filtered
      .map((c, idx) => `<div class="tell-me-item" data-cmd-idx="${idx}"><span>${c.label}</span><span class="tell-me-shortcut">${c.shortcut}</span></div>`)
      .join('');
    tellMeDropdown.style.display = 'block';
    tellMeDropdown.querySelectorAll('.tell-me-item').forEach(el => {
      el.addEventListener('click', () => {
        const cmd = filtered[Number(el.dataset.cmdIdx)];
        tellMeDropdown.style.display = 'none';
        tellMeInput.value = '';
        if (cmd) cmd.action();
      });
    });
  });

  tellMeInput.addEventListener('focus', () => {
    tellMeInput.dispatchEvent(new Event('input'));
  });

  // Central Action Dispatcher
  function executeAction(action, dataset = {}) {
    switch (action) {
      case 'open-backstage':
        state.backstageOpen = true;
        backstageOverlay.classList.remove('hidden');
        renderBackstage();
        break;
      case 'close-backstage':
        state.backstageOpen = false;
        backstageOverlay.classList.add('hidden');
        break;
      case 'toggle-ribbon':
        state.ribbonCollapsed = !state.ribbonCollapsed;
        ribbonSurface.classList.toggle('collapsed', state.ribbonCollapsed);
        break;
      case 'toggle-autosave':
        state.autoSave = !state.autoSave;
        container.querySelector('#autosave-pill').classList.toggle('on', state.autoSave);
        break;
      case 'save-doc':
        saveDocument('docx', false);
        break;
      case 'print-doc':
        exportPdf();
        break;
      case 'undo':
        undo();
        break;
      case 'redo':
        redo();
        break;
      case 'cut':
        document.execCommand && document.execCommand('cut');
        pushHistory();
        markDirty();
        break;
      case 'copy':
        document.execCommand && document.execCommand('copy');
        break;
      case 'paste':
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then(txt => {
            if (txt) insertHtmlAtCursor(txt);
          }).catch(() => {});
        }
        break;
      case 'format-painter': {
        const block = getActiveBlockElement();
        if (state.formatPainterStyle) {
          state.formatPainterStyle = null;
        } else if (block) {
          state.formatPainterStyle = block.getAttribute('style') || '';
        }
        renderRibbon();
        break;
      }
      case 'bold':
        if (document.execCommand) document.execCommand('bold');
        pushHistory();
        markDirty();
        break;
      case 'italic':
        if (document.execCommand) document.execCommand('italic');
        pushHistory();
        markDirty();
        break;
      case 'underline':
        if (document.execCommand) document.execCommand('underline');
        pushHistory();
        markDirty();
        break;
      case 'strikethrough':
        if (document.execCommand) document.execCommand('strikeThrough');
        pushHistory();
        markDirty();
        break;
      case 'subscript':
        if (document.execCommand) document.execCommand('subscript');
        pushHistory();
        markDirty();
        break;
      case 'superscript':
        if (document.execCommand) document.execCommand('superscript');
        pushHistory();
        markDirty();
        break;
      case 'highlight-yellow':
        applyInlineStyle('backgroundColor', '#FFF100');
        break;
      case 'font-size-up': {
        const sel = container.querySelector('#ribbon-font-size');
        const next = Math.min(72, (sel ? Number(sel.value) : 11) + 2);
        if (sel) sel.value = String(next);
        applyFontSizePt(next);
        break;
      }
      case 'font-size-down': {
        const sel = container.querySelector('#ribbon-font-size');
        const next = Math.max(8, (sel ? Number(sel.value) : 11) - 2);
        if (sel) sel.value = String(next);
        applyFontSizePt(next);
        break;
      }
      case 'clear-formatting': {
        const block = getActiveBlockElement();
        if (block && block !== editorBody) {
          block.removeAttribute('style');
          block.className = '';
        }
        if (document.execCommand) document.execCommand('removeFormat');
        pushHistory();
        markDirty();
        break;
      }
      case 'bullet-list':
        if (document.execCommand && !document.execCommand('insertUnorderedList')) {
          insertHtmlAtCursor('<ul><li>List item</li></ul>');
        }
        pushHistory();
        markDirty();
        break;
      case 'number-list':
        if (document.execCommand && !document.execCommand('insertOrderedList')) {
          insertHtmlAtCursor('<ol><li>Numbered item</li></ol>');
        }
        pushHistory();
        markDirty();
        break;
      case 'indent': {
        const block = getActiveBlockElement();
        if (block && block !== editorBody) {
          const cur = parseInt(block.style.marginLeft || '0', 10);
          block.style.marginLeft = `${cur + 36}px`;
          pushHistory();
          markDirty();
        }
        break;
      }
      case 'outdent': {
        const block = getActiveBlockElement();
        if (block && block !== editorBody) {
          const cur = parseInt(block.style.marginLeft || '0', 10);
          block.style.marginLeft = `${Math.max(0, cur - 36)}px`;
          pushHistory();
          markDirty();
        }
        break;
      }
      case 'sort-paragraphs': {
        const paras = Array.from(editorBody.querySelectorAll(':scope > p'));
        if (paras.length > 1) {
          paras.sort((a, b) => (a.textContent || '').localeCompare(b.textContent || ''));
          paras.forEach(p => editorBody.appendChild(p));
          pushHistory();
          markDirty();
        }
        break;
      }
      case 'toggle-paragraph-marks':
        state.showParagraphMarks = !state.showParagraphMarks;
        editorBody.classList.toggle('show-pilcrows', state.showParagraphMarks);
        renderRibbon();
        break;
      case 'align-left':
        applyBlockStyle('textAlign', 'left');
        break;
      case 'align-center':
        applyBlockStyle('textAlign', 'center');
        break;
      case 'align-right':
        applyBlockStyle('textAlign', 'right');
        break;
      case 'align-justify':
        applyBlockStyle('textAlign', 'justify');
        break;
      case 'paragraph-shading': {
        const block = getActiveBlockElement();
        if (block && block !== editorBody) {
          block.style.backgroundColor = block.style.backgroundColor ? '' : '#E8F1FB';
          block.style.padding = block.style.backgroundColor ? '6px 10px' : '';
          pushHistory();
          markDirty();
        }
        break;
      }
      case 'paragraph-border': {
        const block = getActiveBlockElement();
        if (block && block !== editorBody) {
          block.style.border = block.style.border ? '' : '1px solid #185ABD';
          block.style.padding = block.style.border ? '6px 10px' : '';
          pushHistory();
          markDirty();
        }
        break;
      }
      case 'apply-style':
        applyQuickStyle(dataset.styleId);
        break;
      case 'open-font-dialog':
        openFontDialog();
        break;
      case 'open-paragraph-dialog':
        openParagraphDialog();
        break;
      case 'open-page-setup-dialog':
        openPageSetupDialog();
        break;
      case 'open-find-dialog':
        openFindReplaceDialog(false);
        break;
      case 'open-replace-dialog':
        openFindReplaceDialog(true);
        break;
      case 'select-all': {
        const range = document.createRange();
        range.selectNodeContents(editorBody);
        const sel = window.getSelection();
        if (sel) {
          sel.removeAllRanges();
          sel.addRange(range);
        }
        break;
      }
      case 'insert-cover-page':
        editorBody.insertAdjacentHTML(
          'afterbegin',
          `<div style="padding:72px 0;border-bottom:2px solid #185ABD;margin-bottom:32px;">
            <p class="doc-title">${state.documentTitle}</p>
            <p class="doc-subtitle">${state.metadata.subject || 'Executive Document'}</p>
            <p style="margin-top:48px;color:#595959;">Prepared by ${state.metadata.author} • ${new Date().toLocaleDateString()}</p>
          </div>
          <hr class="page-break" />`
        );
        pushHistory();
        markDirty();
        break;
      case 'insert-blank-page':
        insertHtmlAtCursor('<hr class="page-break" /><p><br></p><hr class="page-break" /><p><br></p>');
        break;
      case 'insert-page-break':
        insertHtmlAtCursor('<hr class="page-break" /><p><br></p>');
        break;
      case 'open-table-dialog':
        openTableDialog();
        break;
      case 'trigger-image-upload':
        imageUploadInput.click();
        break;
      case 'insert-callout-shape':
        insertHtmlAtCursor(
          `<div style="border-left:4px solid #185ABD;background:#E8F1FB;padding:12px 16px;margin:10px 0;border-radius:0 6px 6px 0;">
            <strong>Key Note:</strong> Highlighted callout block inserted from Shapes gallery.
          </div><p><br></p>`
        );
        break;
      case 'insert-chart-graphic': {
        const svgData = `data:image/svg+xml;utf8,` + encodeURIComponent(
          `<svg xmlns="http://www.w3.org/2000/svg" width="460" height="200" viewBox="0 0 460 200">
            <rect width="460" height="200" rx="6" fill="#FAF9F8" stroke="#E1DFDD"/>
            <text x="20" y="28" font-family="Segoe UI, sans-serif" font-size="13" font-weight="bold" fill="#201F1E">Quarterly Growth Metrics</text>
            <rect x="45" y="110" width="55" height="60" fill="#185ABD"/>
            <rect x="135" y="85" width="55" height="85" fill="#106EBE"/>
            <rect x="225" y="65" width="55" height="105" fill="#2B7CD3"/>
            <rect x="315" y="42" width="55" height="128" fill="#0F4761"/>
            <line x1="30" y1="170" x2="420" y2="170" stroke="#8A8886"/>
          </svg>`
        );
        insertHtmlAtCursor(`<p><img src="${svgData}" alt="Quarterly Growth Chart" /></p>`);
        break;
      }
      case 'open-hyperlink-dialog':
        openHyperlinkDialog();
        break;
      case 'insert-bookmark':
        insertHtmlAtCursor(`<span id="bookmark-${Date.now()}" style="color:#185ABD;font-weight:600;">[🔖 Bookmark]</span> `);
        break;
      case 'add-comment':
        addComment();
        break;
      case 'focus-header':
        headerEditable.focus();
        break;
      case 'focus-footer':
        footerEditable.focus();
        break;
      case 'toggle-page-numbers':
        state.showPageNumbers = !state.showPageNumbers;
        refreshDocumentMetrics();
        renderRibbon();
        break;
      case 'insert-date-time':
        insertHtmlAtCursor(new Date().toLocaleString());
        break;
      case 'open-equation-dialog':
        openEquationDialog();
        break;
      case 'open-symbol-dialog':
        openSymbolDialog();
        break;
      case 'apply-theme': {
        const themeFonts = {
          office: 'Aptos, sans-serif',
          classic: 'Calibri, sans-serif',
          editorial: 'Georgia, serif',
          academic: '"Times New Roman", serif'
        };
        editorBody.style.fontFamily = themeFonts[dataset.theme] || themeFonts.office;
        markDirty();
        break;
      }
      case 'open-watermark-dialog':
        openWatermarkDialog();
        break;
      case 'toggle-page-border':
        state.pageBorder = !state.pageBorder;
        applyPageSetupToDom();
        renderRibbon();
        break;
      case 'space-after-increase':
        applyBlockStyle('marginBottom', '16pt');
        break;
      case 'space-after-remove':
        applyBlockStyle('marginBottom', '0pt');
        break;
      case 'insert-toc':
        insertTableOfContents();
        break;
      case 'insert-footnote':
        insertFootnote();
        break;
      case 'insert-citation':
        insertHtmlAtCursor(`<span>(Microsoft Office Reference, 2026)</span> `);
        break;
      case 'insert-caption':
        insertHtmlAtCursor(`<p style="font-size:9.5pt;font-style:italic;color:#0F4761;">Figure 1: Document Architecture Diagram</p>`);
        break;
      case 'open-word-count-dialog':
        openWordCountDialog();
        break;
      case 'toggle-comments-pane':
        state.showCommentsPane = !state.showCommentsPane;
        commentsPane.classList.toggle('hidden', !state.showCommentsPane);
        if (state.showCommentsPane) renderComments();
        renderRibbon();
        break;
      case 'toggle-track-changes':
        state.trackChanges = !state.trackChanges;
        container.querySelector('#status-track-indicator').style.display = state.trackChanges ? 'inline-flex' : 'none';
        renderRibbon();
        break;
      case 'set-view': {
        const v = dataset.view || 'print';
        state.viewMode = v;
        canvasViewport.className = `canvas-viewport view-${v}${state.showGridlines ? ' gridlines-on' : ''}`;
        editorBody.contentEditable = v === 'read' ? 'false' : 'true';
        container.querySelectorAll('.status-btn[data-view]').forEach(btn => {
          btn.classList.toggle('active', btn.dataset.view === v);
        });
        renderRibbon();
        break;
      }
      case 'toggle-focus':
        state.focusMode = !state.focusMode;
        document.body.classList.toggle('focus-mode', state.focusMode);
        renderRibbon();
        break;
      case 'toggle-dark-mode':
        state.darkMode = !state.darkMode;
        document.body.classList.toggle('dark-mode', state.darkMode);
        renderRibbon();
        break;
      case 'toggle-nav-pane':
        state.showNavPane = !state.showNavPane;
        navPane.classList.toggle('hidden', !state.showNavPane);
        if (state.showNavPane) renderNavPane();
        renderRibbon();
        break;
      case 'zoom-in':
        setZoom(Math.min(500, state.zoom + 10));
        break;
      case 'zoom-out':
        setZoom(Math.max(10, state.zoom - 10));
        break;
      case 'reset-zoom':
        setZoom(100);
        break;
      case 'zoom-page-width':
        setZoom(120);
        break;
      case 'zoom-one-page':
        setZoom(80);
        break;
      default:
        if (action && action.startsWith('table-')) {
          handleTableAction(action);
        }
        break;
    }
  }

  function setZoom(percent) {
    const clamped = Math.max(10, Math.min(500, Math.round(Number(percent) || 100)));
    state.zoom = clamped;
    zoomStage.style.transform = `scale(${clamped / 100})`;
    zoomSlider.value = String(clamped);
    statusZoomBtn.textContent = `${clamped}%`;
  }

  // Event Delegation for Clicks
  container.addEventListener('click', e => {
    if (!e.target.closest('.tell-me-wrapper')) {
      tellMeDropdown.style.display = 'none';
    }

    const tabBtn = e.target.closest('.ribbon-tab[data-tab]');
    if (tabBtn) {
      state.activeTab = tabBtn.dataset.tab;
      state.ribbonCollapsed = false;
      renderRibbon();
      return;
    }

    const navTabBtn = e.target.closest('.nav-tab[data-nav-tab]');
    if (navTabBtn) {
      state.navPaneTab = navTabBtn.dataset.navTab;
      renderNavPane();
      return;
    }

    const bsNavBtn = e.target.closest('.backstage-nav-item[data-backstage]');
    if (bsNavBtn) {
      state.backstageSection = bsNavBtn.dataset.backstage;
      renderBackstage();
      return;
    }

    const actionBtn = e.target.closest('[data-action]');
    if (actionBtn) {
      executeAction(actionBtn.dataset.action, actionBtn.dataset);
      return;
    }

    // Track contextual table cell selection
    const cell = e.target.closest('td, th');
    if (cell && editorBody.contains(cell)) {
      state.activeCell = cell;
      tableContextualTab.style.display = 'inline-flex';
    } else if (editorBody.contains(e.target)) {
      tableContextualTab.style.display = 'none';
    }

    // Format Painter application on click inside editor
    if (state.formatPainterStyle !== null && editorBody.contains(e.target)) {
      const block = getActiveBlockElement();
      if (block && block !== editorBody) {
        block.setAttribute('style', state.formatPainterStyle);
        state.formatPainterStyle = null;
        renderRibbon();
        pushHistory();
        markDirty();
      }
    }
  });

  // Document Title Input
  docTitleInput.addEventListener('input', () => {
    state.documentTitle = docTitleInput.value.trim() || 'Document1';
    markDirty();
  });

  // Editor Input & Track Changes
  editorBody.addEventListener('input', () => {
    pushHistory();
    markDirty();
  });

  headerEditable.addEventListener('input', () => {
    state.headerText = headerEditable.textContent || '';
    pushHistory();
    markDirty();
  });

  footerEditable.addEventListener('input', () => {
    state.footerText = footerEditable.textContent || '';
    pushHistory();
    markDirty();
  });

  // Zoom Slider
  zoomSlider.addEventListener('input', () => {
    setZoom(Number(zoomSlider.value));
  });

  // Navigation Search
  navSearchInput.addEventListener('input', () => {
    state.navSearchQuery = navSearchInput.value;
    state.navPaneTab = 'results';
    renderNavPane();
  });

  // File Upload Handlers
  fileOpenInput.addEventListener('change', async e => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    await openDocumentFromBuffer(file.name, buf);
    fileOpenInput.value = '';
  });

  imageUploadInput.addEventListener('change', e => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      insertHtmlAtCursor(`<p><img src="${reader.result}" alt="${file.name}" style="max-width:420px;" /></p>`);
    };
    reader.readAsDataURL(file);
    imageUploadInput.value = '';
  });

  // Keyboard Shortcuts (matching SPEC.md Section 3.6)
  window.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      if (state.focusMode) executeAction('toggle-focus');
      if (state.backstageOpen) executeAction('close-backstage');
      modalBackdrop.classList.add('hidden');
      miniToolbar.classList.add('hidden');
      return;
    }

    if (e.altKey && e.key.toLowerCase() === 'q') {
      e.preventDefault();
      tellMeInput.focus();
      return;
    }

    const mod = e.ctrlKey || e.metaKey;
    if (!mod) return;

    const k = e.key.toLowerCase();
    if (k === 's') {
      e.preventDefault();
      saveDocument('docx', e.shiftKey);
    } else if (k === 'o') {
      e.preventDefault();
      triggerOpenDocument();
    } else if (k === 'n') {
      e.preventDefault();
      loadTemplate('blank');
    } else if (k === 'p') {
      e.preventDefault();
      exportPdf();
    } else if (k === 'f') {
      e.preventDefault();
      state.showNavPane = true;
      navPane.classList.remove('hidden');
      state.navPaneTab = 'results';
      renderNavPane();
      navSearchInput.focus();
    } else if (k === 'h') {
      e.preventDefault();
      openFindReplaceDialog(true);
    } else if (k === 'k') {
      e.preventDefault();
      openHyperlinkDialog();
    } else if (k === 'd') {
      e.preventDefault();
      openFontDialog();
    } else if (k === 'z') {
      e.preventDefault();
      undo();
    } else if (k === 'y') {
      e.preventDefault();
      redo();
    } else if (k === 'l') {
      e.preventDefault();
      executeAction('align-left');
    } else if (k === 'e') {
      e.preventDefault();
      executeAction('align-center');
    } else if (k === 'r') {
      e.preventDefault();
      executeAction('align-right');
    } else if (k === 'j') {
      e.preventDefault();
      executeAction('align-justify');
    } else if (e.key === 'Enter') {
      e.preventDefault();
      executeAction('insert-page-break');
    } else if (e.shiftKey && k === 'g') {
      e.preventDefault();
      openWordCountDialog();
    }
  });

  // Desktop IPC Integration (when running inside Electron Windows Shell)
  if (isDesktop && window.wordDesktop && window.wordDesktop.onMenuCommand) {
    window.wordDesktop.onMenuCommand(async (cmd, payload) => {
      if (cmd === 'file:new') loadTemplate('blank');
      else if (cmd === 'file:open') triggerOpenDocument();
      else if (cmd === 'file:save') saveDocument('docx', false);
      else if (cmd === 'file:save-as') saveDocument('docx', true);
      else if (cmd === 'file:export-pdf') exportPdf();
      else if (cmd === 'file:open-buffer' && payload) {
        state.currentFilePath = payload.filePath;
        const uint8 = new Uint8Array(payload.data);
        await openDocumentFromBuffer(payload.fileName || payload.filePath, uint8.buffer);
      } else {
        executeAction(cmd);
      }
    });
  }

  // Initialize View
  renderRibbon();
  applyPageSetupToDom();
  pushHistory();
  refreshDocumentMetrics();

  // Public API for programmatic control & automated tests
  return {
    state,
    editorBody,
    headerEditable,
    footerEditable,
    executeAction,
    getDocumentStats,
    applyQuickStyle,
    applyInlineStyle,
    applyBlockStyle,
    insertTable,
    handleTableAction,
    findTextMatches,
    replaceAllMatches,
    insertTableOfContents,
    insertFootnote,
    addComment,
    setZoom,
    undo,
    redo,
    loadTemplate,
    getSerializedState,
    openDocumentFromBuffer,
    exportDocx: async () => exportDocx(await getSerializedState()),
    importDocx: async buf => openDocumentFromBuffer('Imported.docx', buf)
  };
}
