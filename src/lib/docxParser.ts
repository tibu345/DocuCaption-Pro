import JSZip from 'jszip';
import {
  childElements,
  classifyTocInstruction,
  findNearbyCaption,
  getBody,
  getParagraphText,
  localName,
  W_NS,
} from './docxFields';
import { DocumentAudit, fallbackCaption, sanitizeCaptionDescription } from './captionUtils';
import type { DocElement, ParsedDocument } from './documentProcessor';

export async function parseDocxFile(file: File): Promise<ParsedDocument> {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    throw new Error('Please choose a .docx file.');
  }

  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) throw new Error('This file is missing word/document.xml and is not a valid .docx document.');

  const imageRelationships = await loadImageRelationships(zip);
  const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml');
  const body = getBody(xmlDoc);
  const elements: DocElement[] = [];
  let figureIndex = 0;
  let tableIndex = 0;

  for (const node of walkDocumentBlocks(body)) {
    if (localName(node) === 'p') {
      const imageReferences = await extractParagraphImages(node, imageRelationships, zip);
      if (imageReferences.length > 0) {
        for (const image of imageReferences) {
          const existingCaption = findNearbyCaption(node, 'Figure', 0, 3);
          elements.push({
            type: 'image',
            id: `fig-${figureIndex}`,
            order: figureIndex,
            src: image.src,
            alt: image.alt,
            caption: existingCaption ?? fallbackCaption('Figure'),
            hasExistingCaption: Boolean(existingCaption),
          });
          figureIndex += 1;
        }
      } else {
        const text = getParagraphText(node);
        if (text) elements.push({ type: 'paragraph', text });
      }
    }

    if (localName(node) === 'tbl') {
      const existingCaption = findNearbyCaption(node, 'Table', 3, 3);
      elements.push({
        type: 'table',
        id: `tab-${tableIndex}`,
        order: tableIndex,
        rows: extractTableRows(node),
        caption: existingCaption ?? fallbackCaption('Table'),
        hasExistingCaption: Boolean(existingCaption),
      });
      tableIndex += 1;
    }
  }

  return {
    elements,
    audit: createAudit(elements, documentXml),
  };
}

function walkDocumentBlocks(container: Element): Element[] {
  const blocks: Element[] = [];

  for (const child of childElements(container)) {
    const name = localName(child);
    if (name === 'p' || name === 'tbl') {
      blocks.push(child);
      if (name === 'tbl') {
        for (const tableChild of childElements(child)) {
          blocks.push(...walkDocumentBlocks(tableChild));
        }
      }
      continue;
    }

    if (canContainWordBlocks(child)) {
      blocks.push(...walkDocumentBlocks(child));
    }
  }

  return blocks;
}

function canContainWordBlocks(element: Element): boolean {
  const name = localName(element);
  return ['body', 'sdt', 'sdtContent', 'tc', 'tr', 'txbxContent', 'hdr', 'ftr', 'footnote', 'endnote'].includes(name);
}

export function createAudit(elements: DocElement[], documentXml: string): DocumentAudit {
  const figures = elements.filter((element) => element.type === 'image');
  const tables = elements.filter((element) => element.type === 'table');
  const instructions = Array.from(documentXml.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>|w:instr="([^"]*)"/g))
    .map((match) => decodeXml(match[1] ?? match[2] ?? ''));

  return {
    totalFigures: figures.length,
    figuresWithCaptions: figures.filter((element) => element.hasExistingCaption).length,
    figuresMissingCaptions: figures.filter((element) => !element.hasExistingCaption).length,
    totalTables: tables.length,
    tablesWithCaptions: tables.filter((element) => element.hasExistingCaption).length,
    tablesMissingCaptions: tables.filter((element) => !element.hasExistingCaption).length,
    hasToc: instructions.some((instruction) => classifyTocInstruction(instruction) === 'toc'),
    hasTableOfFigures: instructions.some((instruction) => classifyTocInstruction(instruction) === 'list-figures'),
    hasTableOfTables: instructions.some((instruction) => classifyTocInstruction(instruction) === 'list-tables'),
  };
}

function extractTableRows(table: Element): string[][] {
  return Array.from(table.getElementsByTagNameNS(W_NS, 'tr')).map((row) =>
    Array.from(row.getElementsByTagNameNS(W_NS, 'tc')).map((cell) => getParagraphText(cell)),
  );
}

export async function loadImageRelationships(zip: JSZip): Promise<Map<string, string>> {
  const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
  const relationships = new Map<string, string>();
  if (!relsXml) return relationships;

  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((relationship) => {
    const id = relationship.getAttribute('Id');
    const target = relationship.getAttribute('Target');
    const type = relationship.getAttribute('Type') ?? '';
    if (!id || !target) return;
    if (!/\/image$/i.test(type) && !/\.(png|jpe?g|gif|bmp|webp|tiff?|emf|wmf)$/i.test(target)) return;
    if (/^https?:\/\//i.test(target)) return;
    relationships.set(id, resolveWordTarget(target));
  });

  return relationships;
}

export async function extractParagraphImages(
  paragraph: Element,
  relationships: Map<string, string>,
  zip: JSZip,
): Promise<Array<{ src: string; alt: string }>> {
  const seen = new Set<string>();
  const images: Array<{ src: string; alt: string }> = [];

  for (const reference of findImageRelationshipReferences(paragraph)) {
    const target = relationships.get(reference.relId);
    if (!target || seen.has(target)) continue;
    const file = zip.file(target);
    if (!file) continue;
    const mime = mimeTypeForPath(target);
    if (!mime) continue;

    seen.add(target);
    images.push({
      src: `data:${mime};base64,${await file.async('base64')}`,
      alt: reference.alt,
    });
  }

  return images;
}

function findImageRelationshipReferences(paragraph: Element): Array<{ relId: string; alt: string }> {
  const references: Array<{ relId: string; alt: string }> = [];

  Array.from(paragraph.getElementsByTagName('*')).forEach((element) => {
    const name = localName(element);
    if (name !== 'blip' && name !== 'imagedata') return;
    const relId = getRelationshipId(element);
    if (!relId) return;
    references.push({
      relId,
      alt: findNearbyAltText(element),
    });
  });

  return references;
}

function getRelationshipId(element: Element): string {
  const direct =
    element.getAttribute('r:embed') ??
    element.getAttribute('r:link') ??
    element.getAttribute('r:id') ??
    element.getAttribute('o:relid');
  if (direct) return direct;

  for (const attribute of Array.from(element.attributes)) {
    const name = localName(attribute);
    if (['embed', 'link', 'id', 'relid'].includes(name) && /^rId/i.test(attribute.value)) {
      return attribute.value;
    }
  }
  return '';
}

function findNearbyAltText(element: Element): string {
  const direct = readAltText(element);
  if (direct) return direct;

  let container: Element | null = element;
  while (container && !['drawing', 'pict'].includes(localName(container))) {
    const parent = container.parentNode;
    container = parent?.nodeType === Node.ELEMENT_NODE ? (parent as Element) : null;
  }

  if (container) {
    for (const candidate of Array.from(container.getElementsByTagName('*'))) {
      const text = readAltText(candidate);
      if (text) return text;
    }
  }
  return '';
}

function readAltText(element: Element): string {
  const direct = (
    element.getAttribute('descr') ??
    element.getAttribute('title') ??
    element.getAttribute('alt') ??
    element.getAttribute('o:title') ??
    ''
  ).trim();
  if (direct) return sanitizeAltText(direct);

  for (const attribute of Array.from(element.attributes)) {
    if (['descr', 'title', 'alt'].includes(localName(attribute)) && attribute.value.trim()) {
      return sanitizeAltText(attribute.value);
    }
  }
  return '';
}

function sanitizeAltText(value: string): string {
  const sanitized = sanitizeCaptionDescription(value);
  return sanitized === 'Add description' ? '' : sanitized;
}

function resolveWordTarget(target: string): string {
  const isPackageAbsolute = target.startsWith('/');
  const clean = target.replace(/^\/+/, '');
  if (clean.startsWith('word/')) return clean;
  const parts = isPackageAbsolute ? clean.split('/') : ['word', ...clean.split('/')];
  const normalized: string[] = [];
  for (const part of parts) {
    if (!part || part === '.') continue;
    if (part === '..') normalized.pop();
    else normalized.push(part);
  }
  return normalized.join('/');
}

function mimeTypeForPath(path: string): string | undefined {
  const extension = path.split('.').pop()?.toLowerCase();
  const types: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    emf: 'image/x-emf',
    wmf: 'image/wmf',
  };
  return extension ? types[extension] : undefined;
}

function decodeXml(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}
