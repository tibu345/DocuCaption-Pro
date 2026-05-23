import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { parseDocxFile } from '../src/lib/docxParser';

const sampleDocument = new DOMParser().parseFromString('<root/>', 'application/xml');
Object.assign(globalThis, {
  DOMParser,
  XMLSerializer,
  Node: { ELEMENT_NODE: 1 },
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
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
      xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office">
      <w:body>
        <w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Decorative shape"/></wp:inline></w:drawing></w:r></w:p>
        <w:p><w:r><w:drawing><wp:inline><wp:docPr id="2" name="Picture 1" descr="Architecture diagram"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
        <w:p><w:r><w:pict><v:shape><v:imagedata r:id="rIdImage2" o:title="Legacy screenshot"/></v:shape></w:pict></w:r></w:p>
        <w:sdt><w:sdtContent><w:p><w:r><w:drawing><wp:inline><wp:docPr id="3" name="Picture 2" descr="Content control image"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage3"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:sdtContent></w:sdt>
        <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:drawing><wp:inline><wp:docPr id="4" name="Picture 3" descr="Image inside layout table"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc></w:tr></w:tbl>
      </w:body>
    </w:document>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
      <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
      <Relationship Id="rIdImage2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
      <Relationship Id="rIdImage3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/media/image3.png"/>
      <Relationship Id="rIdImage4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image4.png"/>
    </Relationships>`,
  );
  zip.file('word/media/image1.png', pngBytes);
  zip.file('word/media/image2.png', pngBytes);
  zip.file('media/image3.png', pngBytes);
  zip.file('word/media/image4.png', pngBytes);

  const buffer = await zip.generateAsync({ type: 'uint8array' });
  const file = new File([buffer], 'parser-smoke.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const parsed = await parseDocxFile(file);
  const figures = parsed.elements.filter((element) => element.type === 'image');
  const tables = parsed.elements.filter((element) => element.type === 'table');

  assert(figures.length === 4, `Expected 4 real figures, got ${figures.length}`);
  assert(figures.every((figure) => figure.src.startsWith('data:image/png;base64,')), 'Expected each figure to include a preview data URL');
  assert(figures[0].alt === 'Architecture diagram', `Expected DrawingML alt text, got "${figures[0].alt}"`);
  assert(figures[1].alt === 'Legacy screenshot', `Expected VML alt text, got "${figures[1].alt}"`);
  assert(figures[2].alt === 'Content control image', `Expected content control image alt text, got "${figures[2].alt}"`);
  assert(figures[3].alt === 'Image inside layout table', `Expected nested table image alt text, got "${figures[3].alt}"`);
  assert(tables.length === 1, `Expected 1 table, got ${tables.length}`);
  assert(parsed.audit.totalFigures === 4, `Expected audit to report 4 figures, got ${parsed.audit.totalFigures}`);
  assert(parsed.audit.totalTables === 1, `Expected audit to report 1 table, got ${parsed.audit.totalTables}`);

  console.log('parser-smoke-test passed');
}

function assert(condition: boolean, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
