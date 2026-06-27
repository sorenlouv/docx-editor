/**
 * `cloneDocument` must preserve fields JSON can't represent.
 *
 * Regression for: any agent edit clones the document via `cloneDocument`, and
 * a naive `JSON.parse(JSON.stringify())` turned `originalBuffer` (ArrayBuffer)
 * and `package.media` (a Map) into `{}`. That broke export
 * (`repackDocx` → `JSZip.loadAsync({})` → "Can't read the data of 'the loaded
 * zip file'") and dropped every image on the first edit.
 */

import { describe, test, expect } from 'bun:test';
import { cloneDocument } from '../helpers';
import { createEmptyDocument } from '../../../utils/createDocument';
import type { Document } from '../../../types/document';
import type { MediaFile } from '../../../types/styles';

function docWithBinary(): Document {
  const doc = createEmptyDocument();
  doc.originalBuffer = new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer; // "PK\x03\x04"
  const media = new Map<string, MediaFile>();
  media.set('word/media/image1.png', {
    path: 'word/media/image1.png',
    mimeType: 'image/png',
    data: new Uint8Array([1, 2, 3, 4]).buffer,
  });
  doc.package.media = media;
  return doc;
}

describe('cloneDocument — non-JSON-serializable fields', () => {
  test('preserves originalBuffer as an ArrayBuffer with identical bytes', () => {
    const doc = docWithBinary();
    const cloned = cloneDocument(doc);

    expect(cloned.originalBuffer).toBeInstanceOf(ArrayBuffer);
    expect(Array.from(new Uint8Array(cloned.originalBuffer!))).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  test('preserves package.media as a Map carrying the binary MediaFile', () => {
    const doc = docWithBinary();
    const cloned = cloneDocument(doc);

    expect(cloned.package.media).toBeInstanceOf(Map);
    const img = cloned.package.media!.get('word/media/image1.png');
    expect(img?.mimeType).toBe('image/png');
    expect(img?.data).toBeInstanceOf(ArrayBuffer);
    expect(Array.from(new Uint8Array(img!.data))).toEqual([1, 2, 3, 4]);
  });

  test('media Map is a copy — adding to the clone does not leak into the original', () => {
    const doc = docWithBinary();
    const cloned = cloneDocument(doc);

    cloned.package.media!.set('word/media/image2.png', {
      path: 'word/media/image2.png',
      mimeType: 'image/png',
      data: new Uint8Array([9]).buffer,
    });

    expect(cloned.package.media!.size).toBe(2);
    expect(doc.package.media!.size).toBe(1);
  });

  test('still deep-clones the structural model (edits do not touch the original)', () => {
    const doc = docWithBinary();
    const cloned = cloneDocument(doc);

    cloned.package.document.content.push({ type: 'paragraph', content: [] });

    expect(cloned.package.document.content.length).toBe(doc.package.document.content.length + 1);
  });

  test('handles a document with neither originalBuffer nor media', () => {
    const doc = createEmptyDocument();
    const cloned = cloneDocument(doc);

    expect(cloned.originalBuffer).toBeUndefined();
    expect(cloned.package.document).toBeDefined();
  });
});
