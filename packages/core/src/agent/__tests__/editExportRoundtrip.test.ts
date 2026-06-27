/**
 * Edit → export round-trip must not crash.
 *
 * Regression for the headless workflow "generate a document, edit it, export
 * it": before the cloneDocument fix, any edit dropped `originalBuffer`, so the
 * subsequent `toBuffer()` threw inside JSZip ("Can't read the data of 'the
 * loaded zip file'"). No-edit export worked, masking the bug.
 */

import { describe, test, expect } from 'bun:test';
import { DocumentAgent } from '../DocumentAgent';
import { createDocx } from '../../docx/rezip';
import { createEmptyDocument } from '../../utils/createDocument';

async function seedDocxBytes(): Promise<ArrayBuffer> {
  // A real, valid .docx zip to act as the imported "original".
  return createDocx(createEmptyDocument());
}

describe('DocumentAgent edit → toBuffer round-trip', () => {
  test('exports successfully after a text edit (previously threw)', async () => {
    const bytes = await seedDocxBytes();
    const agent = await DocumentAgent.fromBuffer(bytes);

    const edited = agent.insertText({ paragraphIndex: 0, offset: 0 }, 'Hello ');
    const out = await edited.toBuffer();

    expect(out).toBeInstanceOf(ArrayBuffer);
    // A valid .docx is a ZIP — first two bytes are "PK".
    const head = new Uint8Array(out.slice(0, 2));
    expect(head[0]).toBe(0x50);
    expect(head[1]).toBe(0x4b);
  });

  test('exports successfully after inserting a table', async () => {
    const bytes = await seedDocxBytes();
    const agent = await DocumentAgent.fromBuffer(bytes);

    const edited = agent.insertTable({ paragraphIndex: 0, offset: 0 }, 2, 2, { hasHeader: true });
    const out = await edited.toBuffer();

    expect(out.byteLength).toBeGreaterThan(0);
  });
});
