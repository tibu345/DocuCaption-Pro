import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { buildCaptionedDocxBlob } from '../src/lib/docxExporter';
import type { ProcessedDoc } from '../src/lib/documentProcessor';

const sampleDocument = new DOMParser().parseFromString('<root/>', 'application/xml');
Object.assign(globalThis, {
  DOMParser,
  XMLSerializer,
  Node: { ELEMENT_NODE: 1, DOCUMENT_POSITION_FOLLOWING: 4 },
  Element: sampleDocument.documentElement.constructor,
});

const pngBytes = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

async function main() {
  const source = await createSourceDocx();
  const processed: ProcessedDoc = {
    title: 'regression-source',
    content: [
      { type: 'image', id: 'fig-0', order: 0, src: '', alt: 'local file path', caption: 'Figure 1: C:\\Users\\someone\\Desktop\\part1.png' },
      { type: 'table', id: 'tab-0', order: 0, rows: [['Kp', 'Rise time']], caption: 'Table 1: Control response summary' },
    ],
    toc: true,
    listFigures: true,
    listTables: true,
    figureCaptionPlacement: 'below',
    tableCaptionPlacement: 'above',
    typography: { font: 'Times New Roman', size: 11, alignment: 'center' },
  };

  const blob = await buildCaptionedDocxBlob(processed, source.buffer as ArrayBuffer);
  const outputZip = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml = await outputZip.file('word/document.xml')?.async('string');
  const settingsXml = await outputZip.file('word/settings.xml')?.async('string');

  assert(documentXml, 'Expected exported document.xml');
  assert(settingsXml, 'Expected exported settings.xml');
  assert(documentXml.includes('TOC \\o &quot;1-3&quot; \\h \\z \\u') || documentXml.includes('TOC \\o "1-3" \\h \\z \\u'), 'Expected Table of Contents field');
  assert(documentXml.includes('TOC \\h \\z \\c &quot;Figure&quot;') || documentXml.includes('TOC \\h \\z \\c "Figure"'), 'Expected List of Figures field');
  assert(documentXml.includes('TOC \\h \\z \\c &quot;Table&quot;') || documentXml.includes('TOC \\h \\z \\c "Table"'), 'Expected List of Tables field');
  assert(documentXml.includes('SEQ Figure'), 'Expected native Figure SEQ field');
  assert(documentXml.includes('SEQ Table'), 'Expected native Table SEQ field');
  assert(!/C:\\Users\\/i.test(documentXml), 'Expected local file paths to be removed from captions');
  assert(!/w:start="0"|w:start='0'/.test(documentXml), 'Expected invalid zero page numbering to be removed');
  assert(settingsXml.includes('updateFields'), 'Expected Word updateFields setting');

  console.log('docx-export-regression passed');
}

async function createSourceDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <w:body>
        <w:p><w:r><w:t>Cover Page</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>
        <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>Introduction</w:t></w:r></w:p>
        <w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="Architecture diagram"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
        <w:p><w:r><w:t>Figure 1: C:\\Users\\someone\\Desktop\\part1.png</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Kp</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Rise time</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <w:sectPr><w:pgNumType w:start="0"/></w:sectPr>
      </w:body>
    </w:document>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
    </Relationships>`,
  );
  zip.file('word/media/image1.png', pngBytes);
  return zip.generateAsync({ type: 'uint8array' });
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
