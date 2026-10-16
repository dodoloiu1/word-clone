import JSZip from 'jszip';

const PAPER_DIMENSIONS_TWIPS = {
  letter: { width: 12240, height: 15840 },    // 8.5" x 11"
  legal: { width: 12240, height: 20160 },     // 8.5" x 14"
  a4: { width: 11906, height: 16838 },        // 8.27" x 11.69"
  executive: { width: 10440, height: 15120 }  // 7.25" x 10.5"
};

const MARGIN_PRESETS_TWIPS = {
  normal: { top: 1440, right: 1440, bottom: 1440, left: 1440 },     // 1" all around (72pt)
  narrow: { top: 720, right: 720, bottom: 720, left: 720 },          // 0.5" all around (36pt)
  moderate: { top: 1440, right: 1080, bottom: 1440, left: 1080 },    // 1" T/B, 0.75" L/R
  wide: { top: 1440, right: 2880, bottom: 1440, left: 2880 }         // 1" T/B, 2" L/R
};

function escapeXml(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function normalizeHexColor(cssColor) {
  if (!cssColor) return null;
  const trimmed = String(cssColor).trim();
  if (trimmed.startsWith('#')) {
    const hex = trimmed.slice(1).toUpperCase();
    if (hex.length === 3) {
      return hex.split('').map(c => c + c).join('');
    }
    if (hex.length === 6) return hex;
  }
  const rgbMatch = trimmed.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgbMatch) {
    const r = Number(rgbMatch[1]).toString(16).padStart(2, '0');
    const g = Number(rgbMatch[2]).toString(16).padStart(2, '0');
    const b = Number(rgbMatch[3]).toString(16).padStart(2, '0');
    return (r + g + b).toUpperCase();
  }
  return null;
}

function parseFontSizeHalfPoints(sizeStr) {
  if (!sizeStr) return null;
  const ptMatch = String(sizeStr).match(/([\d.]+)\s*pt/i);
  if (ptMatch) return Math.round(parseFloat(ptMatch[1]) * 2);
  const pxMatch = String(sizeStr).match(/([\d.]+)\s*px/i);
  if (pxMatch) return Math.round((parseFloat(pxMatch[1]) * 0.75) * 2);
  return null;
}

function collectInlineRuns(node, inherited = {}, relCollector) {
  const runs = [];

  if (node.nodeType === 3) {
    const text = node.nodeValue || '';
    if (text.length > 0) {
      runs.push({ text, ...inherited });
    }
    return runs;
  }

  if (node.nodeType !== 1) return runs;

  const tag = node.tagName.toUpperCase();
  if (tag === 'BR') {
    runs.push({ breakLine: true });
    return runs;
  }
  if (tag === 'HR' && node.classList && node.classList.contains('page-break')) {
    runs.push({ breakPage: true });
    return runs;
  }

  const nextProps = { ...inherited };
  if (tag === 'B' || tag === 'STRONG') nextProps.bold = true;
  if (tag === 'I' || tag === 'EM') nextProps.italic = true;
  if (tag === 'U') nextProps.underline = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') nextProps.strike = true;
  if (tag === 'SUP') nextProps.vertAlign = 'superscript';
  if (tag === 'SUB') nextProps.vertAlign = 'subscript';

  const style = node.style || {};
  if (style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) nextProps.bold = true;
  if (style.fontStyle === 'italic') nextProps.italic = true;
  if (style.textDecoration && style.textDecoration.includes('underline')) nextProps.underline = true;
  if (style.textDecoration && style.textDecoration.includes('line-through')) nextProps.strike = true;

  if (style.fontFamily) {
    nextProps.fontFamily = style.fontFamily.replace(/['"]/g, '').split(',')[0].trim();
  } else if (node.getAttribute && node.getAttribute('face')) {
    nextProps.fontFamily = node.getAttribute('face');
  }

  const sizeHp = parseFontSizeHalfPoints(style.fontSize);
  if (sizeHp) nextProps.fontSizeHp = sizeHp;

  const colorHex = normalizeHexColor(style.color || (node.getAttribute && node.getAttribute('color')));
  if (colorHex) nextProps.color = colorHex;

  const bgHex = normalizeHexColor(style.backgroundColor);
  if (bgHex && bgHex !== 'transparent' && bgHex !== '000000') {
    nextProps.highlightHex = bgHex;
  }

  if (tag === 'A' && node.getAttribute('href')) {
    const href = node.getAttribute('href');
    const rId = `rIdLink${relCollector.links.length + 1}`;
    relCollector.links.push({ rId, href });
    nextProps.hyperlinkId = rId;
    if (!nextProps.color) nextProps.color = '467886';
    nextProps.underline = true;
  }

  for (const child of Array.from(node.childNodes)) {
    runs.push(...collectInlineRuns(child, nextProps, relCollector));
  }
  return runs;
}

function serializeRunXml(run) {
  if (run.breakPage) {
    return `<w:r><w:br w:type="page"/></w:r>`;
  }
  if (run.breakLine) {
    return `<w:r><w:br/></w:r>`;
  }

  const rPrParts = [];
  if (run.fontFamily) {
    const f = escapeXml(run.fontFamily);
    rPrParts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`);
  }
  if (run.bold) rPrParts.push(`<w:b/>`);
  if (run.italic) rPrParts.push(`<w:i/>`);
  if (run.underline) rPrParts.push(`<w:u w:val="single"/>`);
  if (run.strike) rPrParts.push(`<w:strike/>`);
  if (run.vertAlign) rPrParts.push(`<w:vertAlign w:val="${run.vertAlign}"/>`);
  if (run.color) rPrParts.push(`<w:color w:val="${run.color}"/>`);
  if (run.highlightHex) rPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${run.highlightHex}"/>`);
  if (run.fontSizeHp) {
    rPrParts.push(`<w:sz w:val="${run.fontSizeHp}"/><w:szCs w:val="${run.fontSizeHp}"/>`);
  }

  const rPrXml = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
  const runXml = `<w:r>${rPrXml}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`;

  if (run.hyperlinkId) {
    return `<w:hyperlink r:id="${run.hyperlinkId}">${runXml}</w:hyperlink>`;
  }
  return runXml;
}

function buildParagraphXml(el, relCollector, overrideStyle = null) {
  const tag = el.tagName ? el.tagName.toUpperCase() : 'P';
  if (tag === 'HR' && el.classList && el.classList.contains('page-break')) {
    return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  }

  let styleId = overrideStyle || 'Normal';
  if (!overrideStyle) {
    if (tag === 'H1') styleId = 'Heading1';
    else if (tag === 'H2') styleId = 'Heading2';
    else if (tag === 'H3') styleId = 'Heading3';
    else if (tag === 'H4') styleId = 'Heading4';
    else if (tag === 'BLOCKQUOTE') styleId = 'Quote';
    else if (el.classList && el.classList.contains('doc-title')) styleId = 'Title';
    else if (el.classList && el.classList.contains('doc-subtitle')) styleId = 'Subtitle';
  }

  const pPrParts = [`<w:pStyle w:val="${styleId}"/>`];
  const align = (el.style && el.style.textAlign) || (el.getAttribute && el.getAttribute('align'));
  if (align) {
    const jcMap = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    if (jcMap[align.toLowerCase()]) {
      pPrParts.push(`<w:jc w:val="${jcMap[align.toLowerCase()]}"/>`);
    }
  }

  if (el.dataset && el.dataset.listType) {
    const prefix = el.dataset.listType === 'ordered' ? `${el.dataset.listIndex || 1}. ` : '• ';
    const runs = collectInlineRuns(el, {}, relCollector);
    runs.unshift({ text: prefix });
    pPrParts.push(`<w:ind w:left="720" w:hanging="360"/>`);
    return `<w:p><w:pPr>${pPrParts.join('')}</w:pPr>${runs.map(serializeRunXml).join('')}</w:p>`;
  }

  const runs = collectInlineRuns(el, {}, relCollector);
  const runsXml = runs.map(serializeRunXml).join('');
  return `<w:p><w:pPr>${pPrParts.join('')}</w:pPr>${runsXml}</w:p>`;
}

function buildTableXml(tableEl, relCollector) {
  const rows = Array.from(tableEl.querySelectorAll('tr'));
  const trXmlList = rows.map(tr => {
    const cells = Array.from(tr.children).filter(c => c.tagName === 'TD' || c.tagName === 'TH');
    const tcXmlList = cells.map(cell => {
      const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
      const bgHex = normalizeHexColor(cell.style && cell.style.backgroundColor);
      const tcPrParts = [`<w:tcW w:w="2400" w:type="dxa"/>`];
      if (colspan > 1) tcPrParts.push(`<w:gridSpan w:val="${colspan}"/>`);
      if (bgHex) tcPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bgHex}"/>`);

      const cellRuns = collectInlineRuns(cell, cell.tagName === 'TH' ? { bold: true } : {}, relCollector);
      const pXml = `<w:p>${cellRuns.map(serializeRunXml).join('')}</w:p>`;
      return `<w:tc><w:tcPr>${tcPrParts.join('')}</w:tcPr>${pXml}</w:tc>`;
    });
    return `<w:tr>${tcXmlList.join('')}</w:tr>`;
  });

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="5000" w:type="pct"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="8A8886"/>
      </w:tblBorders>
    </w:tblPr>
    ${trXmlList.join('')}
  </w:tbl>`;
}

export async function exportDocx(docState = {}) {
  const {
    html = '<p></p>',
    headerText = '',
    footerText = '',
    showPageNumbers = true,
    pageSetup = {},
    metadata = {}
  } = docState;

  const parser = new DOMParser();
  const parsed = parser.parseFromString(`<div>${html}</div>`, 'text/html');
  const root = parsed.body.firstElementChild;

  const relCollector = { links: [] };
  const bodyBlocks = [];

  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType === 3) {
      const txt = child.nodeValue.trim();
      if (txt) {
        bodyBlocks.push(`<w:p><w:r><w:t xml:space="preserve">${escapeXml(txt)}</w:t></w:r></w:p>`);
      }
      continue;
    }
    if (child.nodeType !== 1) continue;
    const tag = child.tagName.toUpperCase();

    if (tag === 'TABLE') {
      bodyBlocks.push(buildTableXml(child, relCollector));
    } else if (tag === 'UL' || tag === 'OL') {
      const items = Array.from(child.querySelectorAll(':scope > li'));
      items.forEach((li, idx) => {
        li.dataset.listType = tag === 'OL' ? 'ordered' : 'bullet';
        li.dataset.listIndex = String(idx + 1);
        bodyBlocks.push(buildParagraphXml(li, relCollector, 'ListParagraph'));
      });
    } else {
      bodyBlocks.push(buildParagraphXml(child, relCollector));
    }
  }

  if (bodyBlocks.length === 0) {
    bodyBlocks.push('<w:p/>');
  }

  const paperKey = (pageSetup.paperSize || 'letter').toLowerCase();
  const baseDim = PAPER_DIMENSIONS_TWIPS[paperKey] || PAPER_DIMENSIONS_TWIPS.letter;
  const isLandscape = pageSetup.orientation === 'landscape';
  const pgW = isLandscape ? baseDim.height : baseDim.width;
  const pgH = isLandscape ? baseDim.width : baseDim.height;

  const marginPreset = MARGIN_PRESETS_TWIPS[pageSetup.margins || 'normal'] || MARGIN_PRESETS_TWIPS.normal;
  const customMar = pageSetup.customMarginsTwips || marginPreset;
  const colsCount = Number(pageSetup.columns || 1);

  const sectPrXml = `<w:sectPr>
    <w:headerReference w:type="default" r:id="rIdHeader1"/>
    <w:footerReference w:type="default" r:id="rIdFooter1"/>
    <w:pgSz w:w="${pgW}" w:h="${pgH}" ${isLandscape ? 'w:orient="landscape"' : ''}/>
    <w:pgMar w:top="${customMar.top}" w:right="${customMar.right}" w:bottom="${customMar.bottom}" w:left="${customMar.left}" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:num="${colsCount}" w:space="720"/>
  </w:sectPr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyBlocks.join('\n    ')}
    ${sectPrXml}
  </w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="278" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="360" w:after="80"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/>
      <w:color w:val="0F4761"/>
      <w:sz w:val="40"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/>
      <w:color w:val="0F4761"/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:color w:val="0F4761"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/>
      <w:color w:val="262626"/>
      <w:sz w:val="56"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:color w:val="595959"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:i/>
      <w:color w:val="404040"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>`;

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">${escapeXml(headerText)}</w:t></w:r>
  </w:p>
</w:hdr>`;

  const pageNumFieldXml = showPageNumbers
    ? `<w:r><w:t xml:space="preserve">${footerText ? '   —   Page ' : 'Page '}</w:t></w:r><w:fldSimple w:instr="PAGE"/>`
    : '';

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">${escapeXml(footerText)}</w:t></w:r>
    ${pageNumFieldXml}
  </w:p>
</w:ftr>`;

  const hyperlinkRelsXml = relCollector.links
    .map(l => `<Relationship Id="${l.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(l.href)}" TargetMode="External"/>`)
    .join('\n  ');

  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  ${hyperlinkRelsXml}
</Relationships>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

  const nowIso = new Date().toISOString();
  const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(metadata.title || 'Document1')}</dc:title>
  <dc:subject>${escapeXml(metadata.subject || '')}</dc:subject>
  <dc:creator>${escapeXml(metadata.author || 'Word User')}</dc:creator>
  <dc:description>${escapeXml(metadata.comments || '')}</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:modified>
</cp:coreProperties>`;

  const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml);
  zip.file('_rels/.rels', rootRelsXml);
  zip.file('docProps/core.xml', coreXml);
  zip.file('docProps/app.xml', appXml);
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/header1.xml', headerXml);
  zip.file('word/footer1.xml', footerXml);
  zip.file('word/_rels/document.xml.rels', docRelsXml);

  return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function getElementsByLocalName(parent, localName) {
  const matches = [];
  if (!parent || !parent.children) return matches;
  for (const child of Array.from(parent.children)) {
    const name = child.localName || child.tagName.replace(/^.*:/, '');
    if (name === localName) matches.push(child);
  }
  return matches;
}

function getFirstByLocalName(parent, localName) {
  return getElementsByLocalName(parent, localName)[0] || null;
}

function getAttr(el, attrLocalName) {
  if (!el || !el.attributes) return null;
  for (const attr of Array.from(el.attributes)) {
    const local = attr.localName || attr.name.replace(/^.*:/, '');
    if (local === attrLocalName) return attr.value;
  }
  return null;
}

function parseRunElementToHtml(rEl) {
  const brs = getElementsByLocalName(rEl, 'br');
  if (brs.length > 0) {
    const brType = getAttr(brs[0], 'type');
    if (brType === 'page') return '<hr class="page-break">';
    return '<br>';
  }

  const tEls = getElementsByLocalName(rEl, 't');
  const rawText = tEls.map(t => t.textContent || '').join('');
  if (!rawText) return '';

  let html = escapeXml(rawText);
  const rPr = getFirstByLocalName(rEl, 'rPr');
  if (!rPr) return html;

  const styles = [];
  const rFonts = getFirstByLocalName(rPr, 'rFonts');
  const asciiFont = getAttr(rFonts, 'ascii');
  if (asciiFont) styles.push(`font-family: '${asciiFont}'`);

  const sz = getFirstByLocalName(rPr, 'sz');
  const szVal = getAttr(sz, 'val');
  if (szVal) styles.push(`font-size: ${Number(szVal) / 2}pt`);

  const color = getFirstByLocalName(rPr, 'color');
  const colorVal = getAttr(color, 'val');
  if (colorVal && colorVal !== 'auto') styles.push(`color: #${colorVal}`);

  const shd = getFirstByLocalName(rPr, 'shd');
  const fillVal = getAttr(shd, 'fill');
  if (fillVal && fillVal !== 'auto') styles.push(`background-color: #${fillVal}`);

  if (styles.length > 0) {
    html = `<span style="${styles.join('; ')}">${html}</span>`;
  }

  const vertAlign = getFirstByLocalName(rPr, 'vertAlign');
  const vVal = getAttr(vertAlign, 'val');
  if (vVal === 'superscript') html = `<sup>${html}</sup>`;
  if (vVal === 'subscript') html = `<sub>${html}</sub>`;

  if (getFirstByLocalName(rPr, 'strike')) html = `<s>${html}</s>`;
  if (getFirstByLocalName(rPr, 'u')) html = `<u>${html}</u>`;
  if (getFirstByLocalName(rPr, 'i')) html = `<em>${html}</em>`;
  if (getFirstByLocalName(rPr, 'b')) html = `<strong>${html}</strong>`;

  return html;
}

function parseParagraphElementToHtml(pEl, relMap = {}) {
  const pPr = getFirstByLocalName(pEl, 'pPr');
  const pStyle = getFirstByLocalName(pPr, 'pStyle');
  const styleVal = getAttr(pStyle, 'val') || 'Normal';
  const jc = getFirstByLocalName(pPr, 'jc');
  const jcVal = getAttr(jc, 'val');

  let tag = 'p';
  let classAttr = '';
  if (/^Heading1$/i.test(styleVal)) tag = 'h1';
  else if (/^Heading2$/i.test(styleVal)) tag = 'h2';
  else if (/^Heading3$/i.test(styleVal)) tag = 'h3';
  else if (/^Heading4$/i.test(styleVal)) tag = 'h4';
  else if (/^Title$/i.test(styleVal)) { tag = 'p'; classAttr = ' class="doc-title"'; }
  else if (/^Subtitle$/i.test(styleVal)) { tag = 'p'; classAttr = ' class="doc-subtitle"'; }
  else if (/^Quote$/i.test(styleVal)) tag = 'blockquote';

  const alignMap = { left: 'left', center: 'center', right: 'right', both: 'justify' };
  const styleAttr = jcVal && alignMap[jcVal] ? ` style="text-align: ${alignMap[jcVal]}"` : '';

  const innerParts = [];
  for (const child of Array.from(pEl.children)) {
    const local = child.localName || child.tagName.replace(/^.*:/, '');
    if (local === 'r') {
      innerParts.push(parseRunElementToHtml(child));
    } else if (local === 'hyperlink') {
      const rId = getAttr(child, 'id');
      const href = relMap[rId] || '#';
      const linkRuns = getElementsByLocalName(child, 'r').map(parseRunElementToHtml).join('');
      innerParts.push(`<a href="${escapeXml(href)}">${linkRuns}</a>`);
    }
  }

  const innerHtml = innerParts.join('');
  if (innerHtml === '<hr class="page-break">') {
    return '<hr class="page-break">';
  }
  return `<${tag}${classAttr}${styleAttr}>${innerHtml || '<br>'}</${tag}>`;
}

function parseTableElementToHtml(tblEl, relMap = {}) {
  const rows = getElementsByLocalName(tblEl, 'tr');
  const trHtml = rows.map(tr => {
    const cells = getElementsByLocalName(tr, 'tc');
    const tcHtml = cells.map(tc => {
      const tcPr = getFirstByLocalName(tc, 'tcPr');
      const gridSpan = getFirstByLocalName(tcPr, 'gridSpan');
      const spanVal = getAttr(gridSpan, 'val');
      const shd = getFirstByLocalName(tcPr, 'shd');
      const fillVal = getAttr(shd, 'fill');

      const attrs = [];
      if (spanVal && Number(spanVal) > 1) attrs.push(`colspan="${spanVal}"`);
      if (fillVal && fillVal !== 'auto') attrs.push(`style="background-color: #${fillVal}"`);

      const paras = getElementsByLocalName(tc, 'p').map(p => {
        const pHtml = parseParagraphElementToHtml(p, relMap);
        return pHtml.replace(/^<p>(.*)<\/p>$/s, '$1');
      }).join('<br>');

      return `<td${attrs.length ? ' ' + attrs.join(' ') : ''}>${paras}</td>`;
    }).join('');
    return `<tr>${tcHtml}</tr>`;
  }).join('');

  return `<table><tbody>${trHtml}</tbody></table>`;
}

export async function importDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new DOMParser();

  const relMap = {};
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('string');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    for (const rel of Array.from(relsDoc.getElementsByTagName('Relationship'))) {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      if (id && target) relMap[id] = target;
    }
  }

  let html = '<p></p>';
  const pageSetup = {
    paperSize: 'letter',
    orientation: 'portrait',
    margins: 'normal',
    columns: 1
  };

  const docFile = zip.file('word/document.xml');
  if (docFile) {
    const docXml = await docFile.async('string');
    const xmlDoc = parser.parseFromString(docXml, 'application/xml');
    const body = getFirstByLocalName(xmlDoc.documentElement, 'body');
    if (body) {
      const blocks = [];
      for (const child of Array.from(body.children)) {
        const local = child.localName || child.tagName.replace(/^.*:/, '');
        if (local === 'p') {
          blocks.push(parseParagraphElementToHtml(child, relMap));
        } else if (local === 'tbl') {
          blocks.push(parseTableElementToHtml(child, relMap));
        } else if (local === 'sectPr') {
          const pgSz = getFirstByLocalName(child, 'pgSz');
          if (getAttr(pgSz, 'orient') === 'landscape') {
            pageSetup.orientation = 'landscape';
          }
          const cols = getFirstByLocalName(child, 'cols');
          const numCols = getAttr(cols, 'num');
          if (numCols) pageSetup.columns = Number(numCols);
        }
      }
      if (blocks.length > 0) html = blocks.join('\n');
    }
  }

  let headerText = '';
  const hdrFile = zip.file('word/header1.xml');
  if (hdrFile) {
    const hdrXml = await hdrFile.async('string');
    const hdrDoc = parser.parseFromString(hdrXml, 'application/xml');
    headerText = hdrDoc.documentElement.textContent.trim();
  }

  let footerText = '';
  const ftrFile = zip.file('word/footer1.xml');
  if (ftrFile) {
    const ftrXml = await ftrFile.async('string');
    const ftrDoc = parser.parseFromString(ftrXml, 'application/xml');
    footerText = ftrDoc.documentElement.textContent.replace(/—?\s*Page\s*$/i, '').trim();
  }

  const metadata = { title: 'Document1', author: 'Word User', subject: '', comments: '' };
  const coreFile = zip.file('docProps/core.xml');
  if (coreFile) {
    const coreXml = await coreFile.async('string');
    const coreDoc = parser.parseFromString(coreXml, 'application/xml');
    const titleEl = getFirstByLocalName(coreDoc.documentElement, 'title');
    const creatorEl = getFirstByLocalName(coreDoc.documentElement, 'creator');
    const subjectEl = getFirstByLocalName(coreDoc.documentElement, 'subject');
    if (titleEl && titleEl.textContent) metadata.title = titleEl.textContent;
    if (creatorEl && creatorEl.textContent) metadata.author = creatorEl.textContent;
    if (subjectEl && subjectEl.textContent) metadata.subject = subjectEl.textContent;
  }

  return {
    html,
    headerText,
    footerText,
    pageSetup,
    metadata
  };
}

export function exportHtmlDocument(docState = {}) {
  const { html = '', metadata = {}, headerText = '', footerText = '' } = docState;
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>${escapeXml(metadata.title || 'Document1')}</title>
  <style>
    body { font-family: 'Aptos', 'Calibri', 'Segoe UI', sans-serif; font-size: 11pt; line-height: 1.15; max-width: 816px; margin: 48px auto; padding: 0 48px; color: #201F1E; }
    h1 { font-family: 'Aptos Display', 'Calibri Light', sans-serif; font-size: 20pt; color: #0F4761; font-weight: 400; }
    h2 { font-family: 'Aptos Display', 'Calibri Light', sans-serif; font-size: 16pt; color: #0F4761; font-weight: 400; }
    h3 { font-size: 14pt; color: #0F4761; font-weight: 400; }
    table { border-collapse: collapse; width: 100%; margin: 8pt 0; }
    td, th { border: 1px solid #8A8886; padding: 6px 8px; }
    header, footer { color: #605E5C; font-size: 9.5pt; padding: 12px 0; }
  </style>
</head>
<body>
  ${headerText ? `<header>${escapeXml(headerText)}</header>` : ''}
  <main>${html}</main>
  ${footerText ? `<footer>${escapeXml(footerText)}</footer>` : ''}
</body>
</html>`;
}

export function exportRtfDocument(docState = {}) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${docState.html || ''}</div>`, 'text/html');
  const root = doc.body.firstElementChild;
  const paragraphs = [];

  for (const child of Array.from(root.children)) {
    const text = (child.textContent || '').replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
    if (child.tagName === 'H1') {
      paragraphs.push(`\\pard\\sa160\\fs40\\b ${text}\\b0\\fs22\\par`);
    } else if (child.tagName === 'H2') {
      paragraphs.push(`\\pard\\sa160\\fs32\\b ${text}\\b0\\fs22\\par`);
    } else {
      paragraphs.push(`\\pard\\sa160\\fs22 ${text}\\par`);
    }
  }
  return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Aptos;}{\\f1 Calibri;}}\n${paragraphs.join('\n')}\n}`;
}

export function importRtfDocument(rtfText = '') {
  const cleaned = String(rtfText)
    .replace(/\{\\fonttbl[^}]*\}/g, '')
    .replace(/\\par\b/g, '\n')
    .replace(/\\[a-zA-Z]+\d*\s?/g, '')
    .replace(/[{}]/g, '')
    .trim();

  const lines = cleaned.split(/\r?\n/).filter(Boolean);
  if (lines.length === 0) return '<p></p>';
  return lines.map(l => `<p>${escapeXml(l)}</p>`).join('\n');
}
