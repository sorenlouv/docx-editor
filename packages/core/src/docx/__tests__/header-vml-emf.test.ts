import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';
import { extractMetafileRaster, isMetafileMimeType } from '../metafileRaster';
import type { DrawingContent, Run } from '../../types/content/run';
import type { Paragraph } from '../../types/content/paragraph';

// Synthetic fixture: first-page header (titlePg + headerReference type="first")
// containing <w:object><v:shape><v:imagedata r:id="rId1"/></w:object> where
// rId1 → media/image1.emf, plus a <w:smartTag>-wrapped run. Regenerate with
// `node scripts/gen-fixture-header-vml-emf.mjs`.
const FIXTURE = resolve(process.cwd(), 'e2e/fixtures/header-vml-emf.docx');
const DEFAULT_HEADER = resolve(process.cwd(), 'e2e/fixtures/sdt-header-content.docx');

function load(path: string): ArrayBuffer {
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe('extractMetafileRaster', () => {
  test('extracts the embedded PNG from an EMF', async () => {
    const zip = await JSZip.loadAsync(load(FIXTURE));
    const emf = await zip.file('word/media/image1.emf')!.async('uint8array');
    expect(isMetafileMimeType('image/x-emf')).toBe(true);

    const raster = extractMetafileRaster(emf);
    expect(raster).not.toBeNull();
    expect(raster!.mimeType).toBe('image/png');
    // PNG signature
    expect(raster!.bytes[0]).toBe(0x89);
    expect(raster!.bytes[1]).toBe(0x50);
    // IEND trailer present
    const tail = raster!.bytes.slice(-8);
    expect(String.fromCharCode(tail[0], tail[1], tail[2], tail[3])).toBe('IEND');
    // Slice owns its buffer (not a view into the parent EMF), so the data URL
    // built from `bytes.buffer` encodes only the PNG.
    expect(raster!.bytes.buffer.byteLength).toBe(raster!.bytes.byteLength);
  });

  test('returns null for non-metafile bytes', () => {
    const bytes = new Uint8Array(64).fill(0x20);
    expect(extractMetafileRaster(bytes)).toBeNull();
  });
});

describe('parseDocx — first-page header with VML/EMF image', () => {
  test('populates package.headers and decodes the EMF logo to a PNG data URL', async () => {
    const doc = await parseDocx(load(FIXTURE));

    // headerReferences parsed with type + titlePg
    const sectPr = doc.package.document.finalSectionProperties;
    expect(sectPr?.titlePg).toBe(true);
    expect(sectPr?.headerReferences).toEqual([{ type: 'first', rId: 'rId6' }]);

    // headers map is non-empty and keyed by relId
    expect(doc.package.headers?.size).toBe(1);
    const header = doc.package.headers!.get('rId6');
    expect(header).toBeDefined();
    expect(header!.type).toBe('header');

    // first paragraph carries the logo as a DrawingContent referencing image1
    const para0 = header!.content[0] as Paragraph;
    const run0 = para0.content.find((c): c is Run => c.type === 'run')!;
    const drawing = run0.content.find((c): c is DrawingContent => c.type === 'drawing')!;
    expect(drawing.image.rId).toBe('rId1');
    expect(drawing.image.filename).toBe('image1.emf');
    // EMF was rewritten to a browser-renderable PNG data URL for display
    expect(drawing.image.src?.startsWith('data:image/png;base64,')).toBe(true);
    // VML shape style width:72pt height:60pt → non-zero EMU
    expect(drawing.image.size.width).toBeGreaterThan(0);
    expect(drawing.image.size.height).toBeGreaterThan(0);

    // media map keeps the ORIGINAL EMF bytes (round-trip) under the part path
    const media = doc.package.media!.get('word/media/image1.emf');
    expect(media).toBeDefined();
    expect(media!.mimeType).toBe('image/x-emf');
    expect(media!.dataUrl?.startsWith('data:image/png')).toBe(true);
  });

  test('w:smartTag children are not dropped', async () => {
    const doc = await parseDocx(load(FIXTURE));
    const header = doc.package.headers!.get('rId6')!;
    const text = JSON.stringify(header.content);
    // "SMARTTAG-CITY" lives inside <w:smartTag> in the source XML.
    expect(text).toContain('SMARTTAG-CITY');
    expect(text).toContain(', STATE');
  });

  test('mediaResolver hook can override the display URL', async () => {
    const seen: string[] = [];
    const doc = await parseDocx(load(FIXTURE), {
      mediaResolver: async (file) => {
        seen.push(file.mimeType);
        return file.mimeType === 'image/x-emf' ? 'data:image/png;base64,OVERRIDE' : undefined;
      },
    });
    expect(seen).toContain('image/x-emf');
    const media = doc.package.media!.get('word/media/image1.emf')!;
    expect(media.dataUrl).toBe('data:image/png;base64,OVERRIDE');
    // The image node's src is resolved from media.dataUrl, so it picks up the override.
    const header = doc.package.headers!.get('rId6')!;
    const para0 = header.content[0] as Paragraph;
    const run0 = para0.content.find((c): c is Run => c.type === 'run')!;
    const drawing = run0.content.find((c): c is DrawingContent => c.type === 'drawing')!;
    expect(drawing.image.src).toBe('data:image/png;base64,OVERRIDE');
  });

  test('default-type header (no titlePg) also populates', async () => {
    const doc = await parseDocx(load(DEFAULT_HEADER));
    const sectPr = doc.package.document.finalSectionProperties;
    expect(sectPr?.titlePg).toBeFalsy();
    const ref = sectPr?.headerReferences?.[0];
    expect(ref?.type).toBe('default');
    expect(doc.package.headers?.get(ref!.rId)).toBeDefined();
  });
});

describe('repackDocx — header round-trip', () => {
  test('unedited header/footer XML is byte-identical after round-trip', async () => {
    const inBuf = load(FIXTURE);
    const inZip = await JSZip.loadAsync(inBuf);
    const headerIn = await inZip.file('word/header1.xml')!.async('string');
    const footerIn = await inZip.file('word/footer1.xml')!.async('string');
    const emfIn = await inZip.file('word/media/image1.emf')!.async('uint8array');

    const doc = await parseDocx(inBuf);
    const outBuf = await repackDocx(doc);
    const outZip = await JSZip.loadAsync(outBuf);

    const headerOut = await outZip.file('word/header1.xml')!.async('string');
    const footerOut = await outZip.file('word/footer1.xml')!.async('string');
    const emfOut = await outZip.file('word/media/image1.emf')!.async('uint8array');

    expect(headerOut).toBe(headerIn);
    expect(footerOut).toBe(footerIn);
    // EMF bytes survive (display URL rewrite doesn't touch the stored part).
    expect(emfOut.byteLength).toBe(emfIn.byteLength);
    expect(emfOut).toEqual(emfIn);
  });
});
