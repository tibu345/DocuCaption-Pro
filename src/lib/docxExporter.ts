import JSZip from 'jszip';
import FileSaver from 'file-saver';
import {
  childElements,
  createCaptionParagraph,
  createFieldSection,
  createPageBreak,
  ensureUpdateFields,
  findNearbyCaptionParagraph,
  getBody,
  getParagraphText,
  hasSeqField,
  localName,
  normalizeExistingTocFields,
  scanParagraphSiblings,
  TABLE_OF_FIGURES_FIELD,
  TABLE_OF_TABLES_FIELD,
  TOC_FIELD,
  REL_NS,
  W_NS,
} from './docxFields';
import type { ProcessedDoc } from './documentProcessor';
import { extractParagraphImages, loadImageRelationships } from './docxParser';

export async function exportDocx(processed: ProcessedDoc, originalBuffer: ArrayBuffer): Promise<void> {
  const blob = await buildCaptionedDocxBlob(processed, originalBuffer);
  FileSaver.saveAs(blob, `${processed.title}_captioned.docx`);
}

export async function buildCaptionedDocxBlob(processed: ProcessedDoc, originalBuffer: ArrayBuffer): Promise<Blob> {
  const zip = await JSZip.loadAsync(originalBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) throw new Error('Could not find word/document.xml in the uploaded file.');

  const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml');
  const body = getBody(xmlDoc);
  normalizePageNumbering(xmlDoc);
  normalizeExistingTocFields(xmlDoc);
  stripGeneratedFieldPlaceholderText(body);
  removeRedundantBreaksAroundGlobalSections(body);
  ensureTocSourceHeadings(xmlDoc, body);
  insertGlobalFields(xmlDoc, body, processed);
  await insertAssetCaptions(zip, xmlDoc, body, processed);

  zip.file('word/document.xml', new XMLSerializer().serializeToString(xmlDoc));
  const settingsXml = await zip.file('word/settings.xml')?.async('string');
  zip.file('word/settings.xml', settingsXml ? ensureUpdateFields(settingsXml) : createSettingsXml());
  await ensureSettingsPart(zip);

  return zip.generateAsync({ type: 'blob' });
}

async function insertAssetCaptions(zip: JSZip, xmlDoc: Document, body: Element, processed: ProcessedDoc): Promise<void> {
  const imageRelationships = await loadImageRelationships(zip);
  let figureIndex = 0;
  let tableIndex = 0;

  for (const node of walkCaptionBlocks(body)) {
    if (localName(node) === 'p') {
      const images = await extractParagraphImages(node, imageRelationships, zip);
      let insertAfter: Node = node;
      for (const [_index, _image] of images.entries()) {
        const element = processed.content.find((item) => item.type === 'image' && item.id === `fig-${figureIndex}`);
        if (element?.type === 'image' && !element.excluded && element.caption) {
          const parent = getElementParent(node);
          const captionNode = createCaptionParagraph(xmlDoc, 'Figure', element.caption, processed.typography, element.order + 1);
          const existingCaption = _index === 0 ? findNearbyCaptionParagraph(node, 'Figure', 0, 3) : undefined;
          if (existingCaption && !hasSeqField(existingCaption, 'Figure')) {
            existingCaption.parentNode?.replaceChild(captionNode, existingCaption);
            insertAfter = captionNode;
          } else if (!existingCaption) {
            if (processed.figureCaptionPlacement === 'above') {
              parent.insertBefore(captionNode, node);
            } else {
              parent.insertBefore(captionNode, insertAfter.nextSibling);
              insertAfter = captionNode;
            }
          }
        }
        figureIndex += 1;
      }
    }

    if (localName(node) === 'tbl') {
      const element = processed.content.find((item) => item.type === 'table' && item.id === `tab-${tableIndex}`);
      if (element?.type === 'table' && !element.excluded && element.caption) {
        const parent = getElementParent(node);
        const captionNode = createCaptionParagraph(xmlDoc, 'Table', element.caption, processed.typography, element.order + 1);
        const existingCaption = findTableCaptionParagraph(node, processed.tableCaptionPlacement);
        if (existingCaption && !hasSeqField(existingCaption, 'Table')) {
          existingCaption.parentNode?.replaceChild(captionNode, existingCaption);
        } else if (!existingCaption) {
          if (processed.tableCaptionPlacement === 'below') {
            parent.insertBefore(captionNode, node.nextSibling);
          } else {
            parent.insertBefore(captionNode, node);
          }
        }
      }
      tableIndex += 1;
    }
  }
}

function findTableCaptionParagraph(node: Element, placement: ProcessedDoc['tableCaptionPlacement']): Element | undefined {
  if (placement === 'below') {
    return scanParagraphSiblings(node, 'Table', 'nextSibling', 3) ?? scanParagraphSiblings(node, 'Table', 'previousSibling', 1);
  }

  return scanParagraphSiblings(node, 'Table', 'previousSibling', 1) ?? scanParagraphSiblings(node, 'Table', 'nextSibling', 3);
}

function walkCaptionBlocks(container: Element): Element[] {
  const blocks: Element[] = [];

  for (const child of childElements(container)) {
    const name = localName(child);
    if (name === 'p' || name === 'tbl') {
      blocks.push(child);
      if (name === 'tbl') {
        for (const tableChild of childElements(child)) {
          blocks.push(...walkCaptionBlocks(tableChild));
        }
      }
      continue;
    }

    if (canContainCaptionBlocks(child)) {
      blocks.push(...walkCaptionBlocks(child));
    }
  }

  return blocks;
}

function canContainCaptionBlocks(element: Element): boolean {
  return ['body', 'sdt', 'sdtContent', 'tc', 'tr', 'txbxContent', 'hdr', 'ftr', 'footnote', 'endnote'].includes(localName(element));
}

function getElementParent(element: Element): Element {
  if (element.parentNode?.nodeType === Node.ELEMENT_NODE) return element.parentNode as Element;
  throw new Error('Could not find the document container for a caption target.');
}

function normalizePageNumbering(xmlDoc: Document): void {
  Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'pgNumType')).forEach((pageNumbering) => {
    const start = pageNumbering.getAttributeNS(W_NS, 'start') ?? pageNumbering.getAttribute('w:start');
    if (start === null) return;
    const parsed = Number(start);
    if (!Number.isFinite(parsed) || parsed > 0) return;
    pageNumbering.removeAttributeNS(W_NS, 'start');
    pageNumbering.removeAttribute('w:start');
  });
}

function insertGlobalFields(xmlDoc: Document, body: Element, processed: ProcessedDoc): void {
  const requested: GlobalFieldItem[] = [
    processed.toc ? { type: 'toc', title: 'Table of Contents', instruction: TOC_FIELD } : undefined,
    processed.listFigures ? { type: 'list-figures', title: 'List of Figures', instruction: TABLE_OF_FIGURES_FIELD } : undefined,
    processed.listTables ? { type: 'list-tables', title: 'List of Tables', instruction: TABLE_OF_TABLES_FIELD } : undefined,
  ].filter((item): item is GlobalFieldItem => Boolean(item));

  if (requested.length === 0) return;

  removeExistingGlobalSections(body, requested.map((item) => item.type));

  const inserted = new Map<GlobalFieldType, Element>();
  const existing = new Map<GlobalFieldType, Element>();
  const coverAnchor = findInsertionAnchor(body);

  for (const item of requested) {
    const insertion = getOrderedGlobalFieldInsertionPoint(body, item.type, existing, inserted, coverAnchor);
    const insertedEnd = insertFieldSectionAt(xmlDoc, body, item, insertion.beforeNode);
    inserted.set(item.type, insertedEnd);
  }

  const finalAnchor = inserted.get('list-tables') ?? inserted.get('list-figures') ?? inserted.get('toc');
  insertTrailingPageBreakWhenNeeded(xmlDoc, body, finalAnchor);
}

type GlobalFieldType = 'toc' | 'list-figures' | 'list-tables';
type GlobalFieldItem = { type: GlobalFieldType; title: string; instruction: string };

function removeExistingGlobalSections(body: Element, requestedTypes: GlobalFieldType[]): void {
  const requested = new Set(requestedTypes);
  const children = childElements(body);
  const ranges: Array<{ start: number; end: number }> = [];

  children.forEach((node, index) => {
    if (localName(node) !== 'p') return;
    const titleType = classifyGlobalFieldTitle(getParagraphText(node));
    const fieldType = classifyGlobalFieldInstruction(getFieldInstructionText(node));
    const type = titleType ?? fieldType;
    if (!type || !requested.has(type)) return;

    const endNode = titleType ? findTitleSectionEnd(body, node) : findFieldEndParagraph(node);
    const endIndex = Math.max(index, children.indexOf(endNode));
    ranges.push({ start: index, end: endIndex });
  });

  mergeRanges(ranges)
    .sort((a, b) => b.start - a.start)
    .forEach((range) => {
      for (let index = range.end; index >= range.start; index -= 1) {
        const node = children[index];
        if (node?.parentNode === body) body.removeChild(node);
      }
    });
}

function mergeRanges(ranges: Array<{ start: number; end: number }>): Array<{ start: number; end: number }> {
  const sorted = [...ranges].sort((a, b) => a.start - b.start);
  const merged: Array<{ start: number; end: number }> = [];
  sorted.forEach((range) => {
    const last = merged[merged.length - 1];
    if (!last || range.start > last.end + 1) {
      merged.push({ ...range });
    } else {
      last.end = Math.max(last.end, range.end);
    }
  });
  return merged;
}

function getExistingGlobalFields(body: Element): Map<GlobalFieldType, Element> {
  const found = new Map<GlobalFieldType, Element>();
  Array.from(body.getElementsByTagNameNS(W_NS, 'instrText')).forEach((node) => {
    const text = node.textContent ?? '';
    const type = classifyGlobalFieldInstruction(text);
    const paragraph = type ? findContainingParagraph(node) : undefined;
    if (type && paragraph && !found.has(type)) found.set(type, findFieldEndParagraph(paragraph));
  });
  Array.from(body.getElementsByTagNameNS(W_NS, 'fldSimple')).forEach((node) => {
    const text = node.getAttributeNS(W_NS, 'instr') ?? node.getAttribute('w:instr') ?? '';
    const type = classifyGlobalFieldInstruction(text);
    const paragraph = type ? findContainingParagraph(node) : undefined;
    if (type && paragraph && !found.has(type)) found.set(type, paragraph);
  });
  addTitleFallbackGlobalFields(body, found);
  return found;
}

function addTitleFallbackGlobalFields(body: Element, found: Map<GlobalFieldType, Element>): void {
  childElements(body).forEach((node) => {
    if (localName(node) !== 'p') return;
    const titleType = classifyGlobalFieldTitle(getParagraphText(node));
    if (titleType && !found.has(titleType)) found.set(titleType, findTitleSectionEnd(body, node));
  });
}

function classifyGlobalFieldTitle(text: string): GlobalFieldType | undefined {
  const normalized = text.replace(/\s+/g, ' ').trim().toLowerCase();
  if (/^table of contents?$/.test(normalized)) return 'toc';
  if (/^(list|table) of figures$/.test(normalized)) return 'list-figures';
  if (/^(list|table) of tables$/.test(normalized)) return 'list-tables';
  return undefined;
}

function findTitleSectionEnd(body: Element, titleParagraph: Element): Element {
  const children = childElements(body);
  const startIndex = children.indexOf(titleParagraph);
  if (startIndex === -1) return titleParagraph;

  for (let index = startIndex + 1; index < children.length; index += 1) {
    const candidate = children[index];
    if (localName(candidate) === 'sectPr') return children[Math.max(startIndex, index - 1)] ?? titleParagraph;
    if (localName(candidate) !== 'p') continue;
    if (index > startIndex + 1 && /^heading1$/i.test(getParagraphStyle(candidate))) {
      return children[index - 1] ?? titleParagraph;
    }
    const titleType = classifyGlobalFieldTitle(getParagraphText(candidate));
    if (titleType) return children[Math.max(startIndex, index - 1)] ?? titleParagraph;
    if (hasPageBreak(candidate) || hasSectionBreak(candidate)) return candidate;
  }

  return titleParagraph;
}

function getParagraphStyle(paragraph: Element): string {
  const style = paragraph.getElementsByTagNameNS(W_NS, 'pStyle')[0];
  return style?.getAttributeNS(W_NS, 'val') ?? style?.getAttribute('w:val') ?? '';
}

function ensureTocSourceHeadings(xmlDoc: Document, body: Element): void {
  if (bodyHasTocSourceHeadings(body)) return;

  for (const paragraph of childElements(body)) {
    if (localName(paragraph) !== 'p') continue;
    if (!isInferredHeadingCandidate(paragraph)) continue;
    setParagraphOutlineLevel(xmlDoc, paragraph, inferHeadingLevel(getParagraphText(paragraph)));
  }
}

function bodyHasTocSourceHeadings(body: Element): boolean {
  return childElements(body).some((node) => {
    if (localName(node) !== 'p') return false;
    const style = getParagraphStyle(node);
    return /^heading[1-9]$/i.test(style) || Boolean(getDirectChild(getDirectChild(node, 'pPr'), 'outlineLvl'));
  });
}

function isInferredHeadingCandidate(paragraph: Element): boolean {
  const text = getParagraphText(paragraph);
  if (text.length < 3 || text.length > 120) return false;
  if (/^(figure|fig\.|table)\s*\d*/i.test(text)) return false;
  if (classifyGlobalFieldTitle(text)) return false;
  if (getFieldInstructionText(paragraph)) return false;
  if (paragraph.getElementsByTagNameNS(W_NS, 'drawing').length > 0 || paragraph.getElementsByTagNameNS(W_NS, 'pict').length > 0) return false;
  if (hasPageBreak(paragraph) || hasSectionBreak(paragraph)) return false;

  const alignment = getParagraphAlignment(paragraph);
  const style = getParagraphStyle(paragraph);
  const maxFontSize = getMaxRunFontSize(paragraph);
  return alignment === 'center' || /title|heading/i.test(style) || hasBoldRun(paragraph) || maxFontSize >= 24 || /^\d+(\.\d+)*\s+\S/.test(text);
}

function inferHeadingLevel(text: string): number {
  const numericPrefix = text.match(/^(\d+(?:\.\d+)*)\s+/)?.[1];
  if (!numericPrefix) return 0;
  return Math.min(8, numericPrefix.split('.').length - 1);
}

function setParagraphOutlineLevel(xmlDoc: Document, paragraph: Element, level: number): void {
  const pPr = ensureParagraphProperties(xmlDoc, paragraph);
  let outline = getDirectChild(pPr, 'outlineLvl');
  if (!outline) {
    outline = xmlDoc.createElementNS(W_NS, 'w:outlineLvl');
    pPr.appendChild(outline);
  }
  outline.setAttributeNS(W_NS, 'w:val', String(level));
}

function ensureParagraphProperties(xmlDoc: Document, paragraph: Element): Element {
  const existing = getDirectChild(paragraph, 'pPr');
  if (existing) return existing;
  const pPr = xmlDoc.createElementNS(W_NS, 'w:pPr');
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

function getDirectChild(element: Element | undefined, name: string): Element | undefined {
  if (!element) return undefined;
  return childElements(element).find((child) => localName(child) === name);
}

function getParagraphAlignment(paragraph: Element): string {
  const jc = getDirectChild(getDirectChild(paragraph, 'pPr'), 'jc');
  return jc?.getAttributeNS(W_NS, 'val') ?? jc?.getAttribute('w:val') ?? '';
}

function hasBoldRun(paragraph: Element): boolean {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 'b')).some((node) => node.getAttributeNS(W_NS, 'val') !== 'false');
}

function getMaxRunFontSize(paragraph: Element): number {
  return Math.max(
    0,
    ...Array.from(paragraph.getElementsByTagNameNS(W_NS, 'sz'))
      .map((node) => Number(node.getAttributeNS(W_NS, 'val') ?? node.getAttribute('w:val') ?? 0))
      .filter((value) => Number.isFinite(value)),
  );
}

function classifyGlobalFieldInstruction(text: string): GlobalFieldType | undefined {
  if (/\\c\s+"Figure"/i.test(text)) return 'list-figures';
  if (/\\c\s+"Table"/i.test(text)) return 'list-tables';
  if (/\bTOC\b/i.test(text)) return 'toc';
  return undefined;
}

function removeRedundantBreaksAroundGlobalSections(body: Element): void {
  childElements(body).forEach((node) => {
    if (!isPageOrSectionBreakParagraph(node)) return;
    const nextMeaningful = getNextBodyElement(node);
    const nextTitle = nextMeaningful && localName(nextMeaningful) === 'p' ? classifyGlobalFieldTitle(getParagraphText(nextMeaningful)) : undefined;
    if (nextTitle && nextTitle !== 'toc' && previousGlobalFieldExists(node)) {
      body.removeChild(node);
      return;
    }

    const previousMeaningful = getPreviousBodyElement(node);
    const previousIsGlobal = previousMeaningful && localName(previousMeaningful) === 'p' &&
      (classifyGlobalFieldTitle(getParagraphText(previousMeaningful)) || classifyGlobalFieldInstruction(getFieldInstructionText(previousMeaningful)));
    const nextIsDocumentEnd = !nextMeaningful || localName(nextMeaningful) === 'sectPr';
    if (previousIsGlobal && nextIsDocumentEnd) {
      body.removeChild(node);
    }
  });
}

function previousGlobalFieldExists(node: Node): boolean {
  let current = node.previousSibling;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE) {
      const element = current as Element;
      if (localName(element) === 'p') {
        if (classifyGlobalFieldTitle(getParagraphText(element))) return true;
        if (classifyGlobalFieldInstruction(getFieldInstructionText(element))) return true;
      }
    }
    current = current.previousSibling;
  }
  return false;
}

function getFieldInstructionText(element: Element): string {
  const pieces = Array.from(element.getElementsByTagNameNS(W_NS, 'instrText')).map((node) => node.textContent ?? '');
  const simpleFields = Array.from(element.getElementsByTagNameNS(W_NS, 'fldSimple')).map((node) => node.getAttributeNS(W_NS, 'instr') ?? node.getAttribute('w:instr') ?? '');
  return [...pieces, ...simpleFields].join(' ');
}

function stripGeneratedFieldPlaceholderText(body: Element): void {
  const affectedParagraphs = new Set<Element>();
  Array.from(body.getElementsByTagNameNS(W_NS, 't')).forEach((node) => {
    const text = (node.textContent ?? '').trim().toLowerCase();
    if (text === 'field will update when opened in word.' || text === 'right-click and update field.') {
      const paragraph = findAncestorByLocalName(node, 'p');
      if (paragraph) affectedParagraphs.add(paragraph);
      const run = findAncestorByLocalName(node, 'r');
      run?.parentNode?.removeChild(run);
    }
  });
  affectedParagraphs.forEach(removeFieldSeparateRuns);
}

function findAncestorByLocalName(node: Node, name: string): Element | undefined {
  let current = node.parentNode;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE && localName(current) === name) return current as Element;
    current = current.parentNode;
  }
  return undefined;
}

function removeFieldSeparateRuns(paragraph: Element): void {
  Array.from(paragraph.getElementsByTagNameNS(W_NS, 'fldChar')).forEach((fieldChar) => {
    if (fieldChar.getAttributeNS(W_NS, 'fldCharType') !== 'separate') return;
    const run = findAncestorByLocalName(fieldChar, 'r');
    run?.parentNode?.removeChild(run);
  });
}

function findContainingParagraph(node: Node): Element | undefined {
  let current: Node | null = node;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE && localName(current) === 'p') return current as Element;
    current = current.parentNode;
  }
  return undefined;
}

function findFieldEndParagraph(startParagraph: Element): Element {
  let current: Node | null = startParagraph;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE) {
      const element = current as Element;
      if (
        localName(element) === 'p' &&
        Array.from(element.getElementsByTagNameNS(W_NS, 'fldChar')).some(
          (fieldChar) => fieldChar.getAttributeNS(W_NS, 'fldCharType') === 'end',
        )
      ) {
        return element;
      }
    }
    current = current.nextSibling;
  }
  return startParagraph;
}

function getFirstExistingFieldInBody(body: Element, nodes: Array<Element | undefined>): Element | undefined {
  const candidates = nodes.filter((node): node is Element => Boolean(node));
  return childElements(body).find((node) => candidates.includes(node));
}

function getOrderedGlobalFieldInsertionPoint(
  body: Element,
  type: string,
  existing: Map<GlobalFieldType, Element>,
  inserted: Map<GlobalFieldType, Element>,
  coverAnchor: Node | null,
): { beforeNode: Node | null } {
  if (type === 'toc') {
    return {
      beforeNode: getFirstExistingFieldInBody(body, [existing.get('list-figures'), existing.get('list-tables')]) ??
        getNodeAfterAnchor(body, coverAnchor),
    };
  }

  if (type === 'list-figures') {
    const anchor = existing.get('toc') ?? inserted.get('toc') ?? coverAnchor;
    return { beforeNode: getNodeAfterAnchor(body, anchor) };
  }

  const anchor = existing.get('list-figures') ?? inserted.get('list-figures') ?? existing.get('toc') ?? inserted.get('toc') ?? coverAnchor;
  return { beforeNode: getNodeAfterAnchor(body, anchor) };
}

function insertFieldSectionAt(
  xmlDoc: Document,
  body: Element,
  item: GlobalFieldItem,
  beforeNode: Node | null,
): Element {
  const section = createFieldSection(xmlDoc, item.title, item.instruction, { pageBreakBefore: item.type !== 'toc' });
  const lastNode = section.lastChild;
  insertBeforeBodyEnd(body, section, beforeNode);
  if (lastNode?.nodeType === Node.ELEMENT_NODE) return lastNode as Element;
  throw new Error(`Could not create ${item.title} section.`);
}

function getNodeAfterAnchor(body: Element, anchor: Node | null | undefined): Node | null {
  return normalizeBodyInsertionPoint(body, anchor?.nextSibling ?? null);
}

function insertBeforeBodyEnd(body: Element, node: Node, preferredBeforeNode: Node | null): void {
  body.insertBefore(node, normalizeBodyInsertionPoint(body, preferredBeforeNode));
}

function insertTrailingPageBreakWhenNeeded(xmlDoc: Document, body: Element, finalAnchor: Element | undefined): void {
  if (!finalAnchor) return;
  const nextElement = getNextBodyElement(finalAnchor);
  if (!nextElement || localName(nextElement) === 'sectPr' || isPageOrSectionBreakParagraph(nextElement)) return;
  insertBeforeBodyEnd(body, createPageBreak(xmlDoc), nextElement);
}

function getNextBodyElement(node: Node): Element | undefined {
  let current = node.nextSibling;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE) return current as Element;
    current = current.nextSibling;
  }
  return undefined;
}

function getPreviousBodyElement(node: Node): Element | undefined {
  let current = node.previousSibling;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE) return current as Element;
    current = current.previousSibling;
  }
  return undefined;
}

function normalizeBodyInsertionPoint(body: Element, preferredBeforeNode: Node | null): Node | null {
  const bodyChildren = childElements(body);
  const sectionProperties = bodyChildren.find((node) => localName(node) === 'sectPr');

  if (preferredBeforeNode && preferredBeforeNode.parentNode === body) {
    if (!sectionProperties) return preferredBeforeNode;
    if (preferredBeforeNode === sectionProperties || isBefore(preferredBeforeNode, sectionProperties)) return preferredBeforeNode;
  }

  return sectionProperties ?? null;
}

function isBefore(node: Node, other: Node): boolean {
  if (typeof node.compareDocumentPosition === 'function') {
    return Boolean(node.compareDocumentPosition(other) & Node.DOCUMENT_POSITION_FOLLOWING);
  }
  let current: Node | null = node;
  while (current) {
    if (current === other) return true;
    current = current.nextSibling;
  }
  return false;
}

function findInsertionAnchor(body: Element): Node | null {
  const children = childElements(body);
  const mainContentStart = findMainContentStart(children);
  if (mainContentStart) return getPreviousBodyElement(mainContentStart);

  const coverBreak = children.find((node) => localName(node) === 'p' && (hasPageBreak(node) || hasSectionBreak(node)));
  return coverBreak ?? children.find((node) => localName(node) === 'p') ?? children[0] ?? null;
}

function findMainContentStart(children: Element[]): Element | undefined {
  return children.find((node, index) => {
    if (localName(node) !== 'p') return false;
    if (index < 3) return false;
    const text = getParagraphText(node);
    if (!text || classifyGlobalFieldTitle(text) || getFieldInstructionText(node)) return false;
    if (hasPageBreak(node) || hasSectionBreak(node)) return false;
    if (/^references$/i.test(text)) return false;
    if (/^appendix\b/i.test(text)) return false;
    if (/^(\d+(\.\d+)*\.?\s+)?(abstract|introduction|background|overview)\b/i.test(text)) return true;
    return /^\d+(\.\d+)*\.?\s+\S/.test(text) && !/^\d{4}\b/.test(text);
  });
}

function hasPageBreak(paragraph: Element): boolean {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 'br')).some((breakNode) => breakNode.getAttributeNS(W_NS, 'type') === 'page');
}

function hasSectionBreak(paragraph: Element): boolean {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 'sectPr')).length > 0;
}

function isPageOrSectionBreakParagraph(node: Element): boolean {
  return localName(node) === 'p' && (hasPageBreak(node) || hasSectionBreak(node));
}

function createSettingsXml(): string {
  return `<w:settings xmlns:w="${W_NS}"><w:updateFields w:val="true"/></w:settings>`;
}

async function ensureSettingsPart(zip: JSZip): Promise<void> {
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const relationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
  const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

  const relsPath = 'word/_rels/document.xml.rels';
  const relsXml = await zip.file(relsPath)?.async('string');
  if (relsXml) {
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    const relationships = relsDoc.documentElement;
    const hasSettingsRel = Array.from(relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship')).some(
      (node) => node.getAttribute('Type') === relationshipType,
    );
    if (!hasSettingsRel) {
      const nextId = getNextRelationshipId(relsDoc);
      const relationship = relsDoc.createElementNS(REL_NS, 'Relationship');
      relationship.setAttribute('Id', nextId);
      relationship.setAttribute('Type', relationshipType);
      relationship.setAttribute('Target', 'settings.xml');
      relationships.appendChild(relationship);
      zip.file(relsPath, serializer.serializeToString(relsDoc));
    }
  }

  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
  if (contentTypesXml) {
    const contentTypesDoc = parser.parseFromString(contentTypesXml, 'application/xml');
    const types = contentTypesDoc.documentElement;
    const hasOverride = Array.from(contentTypesDoc.getElementsByTagName('Override')).some(
      (node) => node.getAttribute('PartName') === '/word/settings.xml',
    );
    if (!hasOverride) {
      const override = contentTypesDoc.createElementNS(types.namespaceURI, 'Override');
      override.setAttribute('PartName', '/word/settings.xml');
      override.setAttribute('ContentType', contentType);
      types.appendChild(override);
      zip.file('[Content_Types].xml', serializer.serializeToString(contentTypesDoc));
    }
  }
}

function getNextRelationshipId(relsDoc: Document): string {
  const ids = Array.from(relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship'))
    .map((node) => node.getAttribute('Id') ?? '')
    .map((id) => Number(id.replace(/^rId/i, '')))
    .filter((id) => Number.isFinite(id));
  return `rId${Math.max(0, ...ids) + 1}`;
}
