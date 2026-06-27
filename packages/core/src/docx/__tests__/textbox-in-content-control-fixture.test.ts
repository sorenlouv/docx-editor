/**
 * Fixture regression — a text box anchored from a run nested inside an inline content
 * control (`w:sdt`) must not be dropped at parse time.
 *
 * `e2e/fixtures/textbox-in-content-control.docx` is a tiny, fully-synthetic document
 * (made-up placeholder text, no real content): a "STRICTLY CONFIDENTIAL" text box
 * whose run sits inside an inline `w:sdt`. Before the fix, the text-box enrichment
 * pass only walked direct `w:r` children of a paragraph, so a text box nested inside a
 * `w:sdt` / `w:hyperlink` was silently lost. Complements the inline-XML unit tests in
 * `wrapped-textbox-roundtrip.test.ts` by exercising the full unzip -> parseDocx pipeline.
 */
import { describe, test, expect } from 'bun:test';
import * as fs from 'fs';
import * as path from 'path';
import { parseDocx } from '../parser';

const FIXTURE = path.join(__dirname, '../../../../../e2e/fixtures/textbox-in-content-control.docx');

describe('text box anchored inside an inline content control (fixture)', () => {
  test('a text box nested in an inline content control is not dropped at parse', async () => {
    const doc = await parseDocx(new Uint8Array(fs.readFileSync(FIXTURE)), { preloadFonts: false });
    expect(JSON.stringify(doc)).toContain('STRICTLY CONFIDENTIAL');
  });
});
