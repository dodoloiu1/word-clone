import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { createWordApp } from '../src/editor/wordApp.js';
import {
  exportDocx,
  importDocx,
  exportHtmlDocument,
  exportRtfDocument,
  importRtfDocument
} from '../src/docx/docxEngine.js';

describe('Word Application 1:1 Fidelity & Automated Feature Suite', () => {
  let container;
  let app;

  beforeEach(() => {
    document.body.innerHTML = '<div id="app"></div>';
    container = document.getElementById('app');
    app = createWordApp(container);
  });

  it('initializes with default document content, Fluent vector icons, ribbon tabs, and status bar metrics', () => {
    expect(container.querySelector('.title-bar')).not.toBeNull();
    expect(container.querySelectorAll('.ribbon-tab[data-tab]').length).toBeGreaterThanOrEqual(7);
    const stats = app.getDocumentStats();
    expect(stats.words).toBeGreaterThan(40);
    expect(stats.pages).toBeGreaterThanOrEqual(1);
    expect(stats.paragraphs).toBeGreaterThan(3);
    // Verify vector SVGs are present in QAT and ribbon
    expect(container.querySelector('.qat-btn svg')).not.toBeNull();
    expect(container.querySelector('.app-brand svg')).not.toBeNull();
  });

  it('switches across all Ribbon tabs (Home, Insert, Design, Layout, References, Review, View)', () => {
    const tabs = ['home', 'insert', 'design', 'layout', 'references', 'review', 'view'];
    for (const tabName of tabs) {
      const btn = container.querySelector(`.ribbon-tab[data-tab="${tabName}"]`);
      expect(btn).not.toBeNull();
      btn.click();
      expect(app.state.activeTab).toBe(tabName);
      expect(container.querySelector('#ribbon-surface').children.length).toBeGreaterThan(0);
    }
  });

  it('supports picture insertion, 8 resize handles, interactive resizing, and contextual Picture Format tab', () => {
    app.loadTemplate('blank');
    const img = document.createElement('img');
    img.src = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect fill="blue" width="100" height="100"/></svg>';
    img.style.width = '200px';
    img.style.height = '150px';
    app.editorBody.appendChild(img);

    const wrapper = app.wrapPictureElement(img);
    expect(wrapper).not.toBeNull();
    expect(wrapper.classList.contains('pic-wrapper')).toBe(true);

    // Verify 8 resize handles + rotation handle
    expect(wrapper.querySelectorAll('.resize-handle').length).toBe(9);
    expect(wrapper.querySelector('.handle-nw')).not.toBeNull();
    expect(wrapper.querySelector('.handle-se')).not.toBeNull();
    expect(wrapper.querySelector('.handle-rot')).not.toBeNull();
    expect(wrapper.querySelector('.pic-layout-btn')).not.toBeNull();

    // Select picture -> activates contextual Picture Format ribbon tab
    app.selectPicture(wrapper);
    expect(wrapper.classList.contains('selected')).toBe(true);
    expect(app.state.selectedPicture).toBe(wrapper);
    expect(app.state.activeTab).toBe('picture');

    const picTab = container.querySelector('#picture-contextual-tab');
    expect(picTab.style.display).not.toBe('none');

    // Apply picture formatting actions
    app.executeAction('pic-style-round');
    expect(img.style.borderRadius).toBe('16px');

    app.executeAction('pic-style-shadow');
    expect(img.style.boxShadow).toContain('rgba');

    app.executeAction('pic-rotate-90');
    expect(img.style.transform).toBe('rotate(90deg)');

    // Deselect picture -> hides contextual tab and switches back to home
    app.deselectAllPictures();
    expect(wrapper.classList.contains('selected')).toBe(false);
    expect(app.state.selectedPicture).toBeNull();
    expect(app.state.activeTab).toBe('home');
    expect(picTab.style.display).toBe('none');
  });

  it('activates contextual Table Layout tab when cursor is in table cell', () => {
    app.loadTemplate('blank');
    app.insertTable(3, 3);
    const table = app.editorBody.querySelector('table');
    expect(table).not.toBeNull();

    const cell = table.querySelector('td');
    cell.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const tableTab = container.querySelector('#table-contextual-tab');
    expect(tableTab.style.display).not.toBe('none');
  });

  it('applies Quick Styles (Heading 1, Heading 2, Title, Quote) and inline/paragraph formatting', () => {
    app.loadTemplate('blank');
    app.editorBody.innerHTML = '<p>Section Heading</p>';
    app.applyQuickStyle('heading1');
    expect(app.editorBody.querySelector('h1')).not.toBeNull();
    expect(app.editorBody.querySelector('h1').textContent).toBe('Section Heading');

    app.applyBlockStyle('textAlign', 'center');
    expect(app.editorBody.querySelector('h1').style.textAlign).toBe('center');

    app.applyInlineStyle('color', '#C00000');
    expect(app.editorBody.querySelector('h1').style.color).toBeTruthy();
  });

  it('inserts tables and performs row/column insertion, cell merging, and shading', () => {
    app.loadTemplate('blank');
    app.insertTable(2, 2);
    const table = app.editorBody.querySelector('table');
    expect(table).not.toBeNull();
    expect(table.querySelectorAll('tr').length).toBe(2);

    const firstCell = table.querySelector('th, td');
    app.state.activeCell = firstCell;

    app.handleTableAction('table-insert-row-below');
    expect(table.querySelectorAll('tr').length).toBe(3);

    app.handleTableAction('table-insert-col-right');
    expect(table.querySelector('tr').children.length).toBe(3);

    app.handleTableAction('table-merge-right');
    expect(firstCell.getAttribute('colspan')).toBe('2');
  });

  it('supports Find and Replace across document paragraphs', () => {
    app.loadTemplate('blank');
    app.editorBody.innerHTML = '<p>Alpha Beta Alpha</p><p>Gamma Alpha</p>';
    expect(app.findTextMatches('Alpha')).toBe(3);

    const replaced = app.replaceAllMatches('Alpha', 'Omega');
    expect(replaced).toBe(3);
    expect(app.editorBody.textContent).toContain('Omega Beta Omega');
    expect(app.findTextMatches('Alpha')).toBe(0);
  });

  it('supports Undo and Redo across document edits', () => {
    app.loadTemplate('blank');
    app.editorBody.innerHTML = '<p>First Version</p>';
    app.editorBody.dispatchEvent(new Event('input'));

    app.editorBody.innerHTML = '<p>Second Version</p>';
    app.editorBody.dispatchEvent(new Event('input'));

    app.undo();
    expect(app.editorBody.textContent).toBe('First Version');

    app.redo();
    expect(app.editorBody.textContent).toBe('Second Version');
  });

  it('updates page setup (orientation, margins, paper size, columns, page breaks) and view modes', () => {
    app.state.pageSetup.orientation = 'landscape';
    app.state.pageSetup.paperSize = 'legal';
    app.state.pageSetup.margins = 'narrow';
    app.state.pageSetup.columns = 2;
    app.executeAction('insert-page-break');
    expect(app.editorBody.querySelectorAll('hr.page-break').length).toBeGreaterThanOrEqual(1);

    app.setZoom(150);
    expect(app.state.zoom).toBe(150);

    app.executeAction('set-view', { view: 'web' });
    expect(app.state.viewMode).toBe('web');

    app.executeAction('toggle-dark-mode');
    expect(app.state.darkMode).toBe(true);
  });

  it('generates Table of Contents, Footnotes, and Comments', () => {
    app.loadTemplate('default');
    app.insertTableOfContents();
    expect(app.editorBody.querySelector('.word-toc')).not.toBeNull();

    app.insertFootnote('Test citation footnote.');
    expect(app.editorBody.querySelector('.footnote-ref')).not.toBeNull();
    expect(app.editorBody.querySelector('.footnotes-section')).not.toBeNull();

    app.addComment('Verify section metrics.');
    expect(app.state.comments.length).toBe(1);
  });

  it('exports a valid OpenXML .docx archive and round-trips content, styles, tables, hyperlinks, headers, and footers', async () => {
    const sampleState = {
      html: `
        <p class="doc-title">Architecture Spec</p>
        <h1>1. Introduction</h1>
        <p style="text-align: center"><strong>Bold text</strong> and <em>italic text</em> with <a href="https://example.com">Reference Link</a>.</p>
        <table><tbody><tr><th>Col A</th><th>Col B</th></tr><tr><td>Val 1</td><td>Val 2</td></tr></tbody></table>
      `,
      headerText: 'Spec Header',
      footerText: 'Spec Footer',
      showPageNumbers: true,
      pageSetup: {
        paperSize: 'letter',
        orientation: 'landscape',
        margins: 'normal',
        columns: 2
      },
      metadata: {
        title: 'Architecture Spec',
        author: 'Verification Suite',
        subject: 'OpenXML Roundtrip'
      }
    };

    const docxBytes = await exportDocx(sampleState);
    expect(docxBytes).toBeInstanceOf(Uint8Array);
    expect(docxBytes.byteLength).toBeGreaterThan(1000);

    const zip = await JSZip.loadAsync(docxBytes);
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('word/document.xml')).not.toBeNull();
    expect(zip.file('word/styles.xml')).not.toBeNull();
    expect(zip.file('word/header1.xml')).not.toBeNull();
    expect(zip.file('word/footer1.xml')).not.toBeNull();

    const roundTripped = await importDocx(docxBytes);
    expect(roundTripped.metadata.title).toBe('Architecture Spec');
    expect(roundTripped.metadata.author).toBe('Verification Suite');
    expect(roundTripped.headerText).toContain('Spec Header');
    expect(roundTripped.footerText).toContain('Spec Footer');
    expect(roundTripped.pageSetup.orientation).toBe('landscape');
    expect(roundTripped.pageSetup.columns).toBe(2);
    expect(roundTripped.html).toContain('<h1>1. Introduction</h1>');
    expect(roundTripped.html).toContain('<strong>Bold text</strong>');
    expect(roundTripped.html).toContain('https://example.com');
    expect(roundTripped.html).toContain('<table>');
  });

  it('exports and imports HTML and RTF document formats', () => {
    const docState = {
      html: '<h1>Report Title</h1><p>Paragraph content here.</p>',
      headerText: 'Header',
      footerText: 'Footer',
      metadata: { title: 'Test Report' }
    };
    const htmlOut = exportHtmlDocument(docState);
    expect(htmlOut).toContain('<!DOCTYPE html>');
    expect(htmlOut).toContain('Report Title');

    const rtfOut = exportRtfDocument(docState);
    expect(rtfOut).toContain('{\\rtf1');
    const rtfImportedHtml = importRtfDocument(rtfOut);
    expect(rtfImportedHtml).toContain('Report Title');
    expect(rtfImportedHtml).toContain('Paragraph content here.');
  });
});
