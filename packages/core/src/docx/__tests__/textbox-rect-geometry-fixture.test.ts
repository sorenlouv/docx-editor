/**
 * Fixture regression — text inside a text box whose shape geometry is not exactly
 * "textBox" (here a `rect` AlternateContent box) must survive a save round-trip.
 *
 * `e2e/fixtures/textbox-rect-geometry.docx` is a tiny, fully-synthetic document
 * (made-up placeholder text, no real content). Before the fix, the serializer only
 * emitted `w:txbxContent` when `shapeType === 'textBox'`; the parser hard-codes
 * `rect` for these boxes, so their text was dropped on save even though it parsed.
 */
import { describe, test, expect } from 'bun:test';
import * as fs from 'fs';
import * as path from 'path';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';

const FIXTURE = path.join(__dirname, '../../../../../e2e/fixtures/textbox-rect-geometry.docx');

describe('text box with non-textBox (rect) geometry (fixture)', () => {
  test('text inside a rect-geometry text box survives a parse -> repack round-trip', async () => {
    const doc = await parseDocx(new Uint8Array(fs.readFileSync(FIXTURE)), { preloadFonts: false });
    const out = await repackDocx(doc, { updateModifiedDate: false });
    const reparsed = await parseDocx(new Uint8Array(out), { preloadFonts: false });
    expect(JSON.stringify(reparsed)).toContain('For internal review only');
  });
});
