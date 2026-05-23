import { cleanCaptionText, startsLikeCaption, type CaptionKind, type CaptionTypography } from './captionUtils';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

export const TOC_FIELD = ' TOC \\o "1-3" \\h \\z \\u ';
export const TABLE_OF_FIGURES_FIELD = ' TOC \\h \\z \\c "Figure" ';
export const TABLE_OF_TABLES_FIELD = ' TOC \\h \\z \\c "Table" ';

export function localName(node: Node | null): string {
  return (node instanceof Element ? node.localName : undefined) ?? node?.nodeName.split(':').pop() ?? '';
}

export function childElements(element: Element): Element[] {
  return Array.from(element.childNodes).filter((node): node is Element => node.nodeType === Node.ELEMENT_NODE);
}

export function getBody(xmlDoc: Document): Element {
  const body = xmlDoc.getElementsByTagNameNS(W_NS, 'body')[0];
  if (!body) throw new Error('This document does not contain a Word body.');
  return body;
}

export function getParagraphText(paragraph: Element): string {
  const pieces: string[] = [];
  Array.from(paragraph.getElementsByTagNameNS(W_NS, 't')).forEach((node) => pieces.push(node.textContent ?? ''));
  Array.from(paragraph.getElementsByTagNameNS(W_NS, 'tab')).forEach(() => pieces.push('\t'));
  return pieces.join('').replace(/\s+/g, ' ').trim();
}

export function hasImage(element: Element): boolean {
  return element.getElementsByTagNameNS(W_NS, 'drawing').length > 0 ||
    element.getElementsByTagNameNS(W_NS, 'pict').length > 0;
}

export function hasSeqField(element: Element, kind: CaptionKind): boolean {
  const matcher = new RegExp(`\\bSEQ\\s+${kind}\\b`, 'i');
  const instrTexts = Array.from(element.getElementsByTagNameNS(W_NS, 'instrText'));
  if (instrTexts.some((node) => matcher.test(node.textContent ?? ''))) return true;

  const fldSimple = Array.from(element.getElementsByTagNameNS(W_NS, 'fldSimple'));
  return fldSimple.some((node) => matcher.test(node.getAttributeNS(W_NS, 'instr') ?? node.getAttribute('w:instr') ?? ''));
}

export function isCaptionParagraph(paragraph: Element, kind: CaptionKind): boolean {
  return hasSeqField(paragraph, kind) || startsLikeCaption(getParagraphText(paragraph), kind);
}

export function findNearbyCaption(node: Element, kind: CaptionKind, before: number, after: number): string | undefined {
  const paragraph = findNearbyCaptionParagraph(node, kind, before, after);
  return paragraph ? getParagraphText(paragraph) || `${kind} caption` : undefined;
}

export function findNearbyCaptionParagraph(node: Element, kind: CaptionKind, before: number, after: number): Element | undefined {
  const beforeCaption = scanParagraphSiblings(node, kind, 'previousSibling', before);
  const afterCaption = scanParagraphSiblings(node, kind, 'nextSibling', after);
  return beforeCaption ?? afterCaption;
}

export function scanParagraphSiblings(
  node: Node,
  kind: CaptionKind,
  direction: 'previousSibling' | 'nextSibling',
  maxParagraphs: number,
): Element | undefined {
  let current = node[direction];
  let seenParagraphs = 0;

  while (current && seenParagraphs < maxParagraphs) {
    if (current.nodeType === Node.ELEMENT_NODE) {
      const element = current as Element;
      if (localName(element) === 'p') {
        const text = getParagraphText(element);
        if (text || hasSeqField(element, kind)) {
          seenParagraphs += 1;
          if (isCaptionParagraph(element, kind)) return element;
        }
      } else if (localName(element) === 'tbl') {
        break;
      }
    }
    current = current[direction];
  }

  return undefined;
}

export function createTextParagraph(doc: Document, text: string, style?: string, options: { pageBreakBefore?: boolean } = {}): Element {
  const p = doc.createElementNS(W_NS, 'w:p');
  if (style || options.pageBreakBefore) {
    const pPr = doc.createElementNS(W_NS, 'w:pPr');
    if (style) {
      const pStyle = doc.createElementNS(W_NS, 'w:pStyle');
      pStyle.setAttributeNS(W_NS, 'w:val', style);
      pPr.appendChild(pStyle);
    }
    if (options.pageBreakBefore) {
      pPr.appendChild(doc.createElementNS(W_NS, 'w:pageBreakBefore'));
    }
    p.appendChild(pPr);
  }
  const r = doc.createElementNS(W_NS, 'w:r');
  const t = doc.createElementNS(W_NS, 'w:t');
  t.textContent = text;
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

export function createCaptionParagraph(
  doc: Document,
  kind: CaptionKind,
  caption: string,
  typography?: CaptionTypography,
  displayNumber = 1,
): Element {
  const p = doc.createElementNS(W_NS, 'w:p');
  const pPr = doc.createElementNS(W_NS, 'w:pPr');
  const pStyle = doc.createElementNS(W_NS, 'w:pStyle');
  pStyle.setAttributeNS(W_NS, 'w:val', 'Caption');
  pPr.appendChild(pStyle);

  const jc = doc.createElementNS(W_NS, 'w:jc');
  jc.setAttributeNS(W_NS, 'w:val', typography?.alignment ?? 'center');
  pPr.appendChild(jc);
  p.appendChild(pPr);

  p.appendChild(createTextRun(doc, `${kind} `, typography));
  p.appendChild(createFieldCharRun(doc, 'begin'));
  p.appendChild(createInstrRun(doc, ` SEQ ${kind} \\* ARABIC `));
  p.appendChild(createFieldCharRun(doc, 'separate'));
  p.appendChild(createTextRun(doc, String(displayNumber), typography));
  p.appendChild(createFieldCharRun(doc, 'end'));
  p.appendChild(createTextRun(doc, `: ${cleanCaptionText(caption, kind)}`, typography));
  return p;
}

export function createFieldSection(doc: Document, title: string, instruction: string, options: { pageBreakBefore?: boolean } = {}): DocumentFragment {
  const fragment = doc.createDocumentFragment();
  fragment.appendChild(createTextParagraph(doc, title, 'Heading1', { pageBreakBefore: options.pageBreakBefore }));

  const p = doc.createElementNS(W_NS, 'w:p');
  p.appendChild(createFieldCharRun(doc, 'begin', true));
  p.appendChild(createInstrRun(doc, instruction));
  p.appendChild(createFieldCharRun(doc, 'end'));
  fragment.appendChild(p);
  return fragment;
}

export function createPageBreak(doc: Document): Element {
  const p = doc.createElementNS(W_NS, 'w:p');
  const r = doc.createElementNS(W_NS, 'w:r');
  const br = doc.createElementNS(W_NS, 'w:br');
  br.setAttributeNS(W_NS, 'w:type', 'page');
  r.appendChild(br);
  p.appendChild(r);
  return p;
}

export function classifyTocInstruction(instruction: string): 'toc' | 'list-figures' | 'list-tables' | undefined {
  if (!/\bTOC\b/i.test(instruction)) return undefined;
  if (/\\c\s+"Figure"/i.test(instruction)) return 'list-figures';
  if (/\\c\s+"Table"/i.test(instruction)) return 'list-tables';
  return 'toc';
}

export function normalizeExistingTocFields(xmlDoc: Document): void {
  const replacements: Record<string, string> = {
    toc: TOC_FIELD,
    'list-figures': TABLE_OF_FIGURES_FIELD,
    'list-tables': TABLE_OF_TABLES_FIELD,
  };

  Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'instrText')).forEach((node) => {
    const kind = classifyTocInstruction(node.textContent ?? '');
    if (kind) node.textContent = replacements[kind];
  });

  Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'fldSimple')).forEach((node) => {
    const instruction = node.getAttributeNS(W_NS, 'instr') ?? node.getAttribute('w:instr') ?? '';
    const kind = classifyTocInstruction(instruction);
    if (kind) {
      node.setAttributeNS(W_NS, 'w:instr', replacements[kind]);
      node.setAttributeNS(W_NS, 'w:dirty', 'true');
    }
  });

  Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'fldChar')).forEach((node) => {
    if (node.getAttributeNS(W_NS, 'fldCharType') === 'begin') node.setAttributeNS(W_NS, 'w:dirty', 'true');
  });
}

export function ensureUpdateFields(settingsXml: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(settingsXml, 'application/xml');
  const settings = doc.documentElement;
  if (!settings || localName(settings) !== 'settings') {
    return `<w:settings xmlns:w="${W_NS}"><w:updateFields w:val="true"/></w:settings>`;
  }

  const existing = doc.getElementsByTagNameNS(W_NS, 'updateFields')[0];
  if (existing) {
    existing.setAttributeNS(W_NS, 'w:val', 'true');
  } else {
    const updateFields = doc.createElementNS(W_NS, 'w:updateFields');
    updateFields.setAttributeNS(W_NS, 'w:val', 'true');
    settings.appendChild(updateFields);
  }

  return new XMLSerializer().serializeToString(doc);
}

function createTextRun(doc: Document, text: string, typography?: CaptionTypography): Element {
  const r = doc.createElementNS(W_NS, 'w:r');
  const rPr = createRunProperties(doc, typography);
  if (rPr.childNodes.length > 0) r.appendChild(rPr);
  const t = doc.createElementNS(W_NS, 'w:t');
  if (/^\s|\s$/.test(text)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = text;
  r.appendChild(t);
  return r;
}

function createInstrRun(doc: Document, instruction: string): Element {
  const r = doc.createElementNS(W_NS, 'w:r');
  const instrText = doc.createElementNS(W_NS, 'w:instrText');
  instrText.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  instrText.textContent = instruction;
  r.appendChild(instrText);
  return r;
}

function createFieldCharRun(doc: Document, type: 'begin' | 'separate' | 'end', dirty = false): Element {
  const r = doc.createElementNS(W_NS, 'w:r');
  const fldChar = doc.createElementNS(W_NS, 'w:fldChar');
  fldChar.setAttributeNS(W_NS, 'w:fldCharType', type);
  if (dirty) fldChar.setAttributeNS(W_NS, 'w:dirty', 'true');
  r.appendChild(fldChar);
  return r;
}

function createRunProperties(doc: Document, typography?: CaptionTypography): Element {
  const rPr = doc.createElementNS(W_NS, 'w:rPr');
  const italic = doc.createElementNS(W_NS, 'w:i');
  rPr.appendChild(italic);

  if (typography?.size) {
    const size = doc.createElementNS(W_NS, 'w:sz');
    size.setAttributeNS(W_NS, 'w:val', String(typography.size * 2));
    rPr.appendChild(size);
  }

  if (typography?.font) {
    const fonts = doc.createElementNS(W_NS, 'w:rFonts');
    fonts.setAttributeNS(W_NS, 'w:ascii', typography.font);
    fonts.setAttributeNS(W_NS, 'w:hAnsi', typography.font);
    rPr.appendChild(fonts);
  }

  return rPr;
}
