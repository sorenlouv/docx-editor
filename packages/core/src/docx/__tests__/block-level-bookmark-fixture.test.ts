/**
 * Fixture regression — a bookmark whose `w:bookmarkStart`/`w:bookmarkEnd` markers sit
 * BETWEEN block elements (direct children of `w:body`, not inside a paragraph) must
 * survive a save round-trip.
 *
 * `e2e/fixtures/block-level-bookmark.docx` is a tiny, fully-synthetic document: a
 * "PaymentTerms" bookmark wraps a heading at body level, and a `REF` field above it
 * points at the bookmark. Before the fix, the block-content parser only handled
 * paragraphs/tables/SDTs (no default branch), so the block-level markers were dropped
 * — breaking every cross-reference that targets them.
 */
import { describe, test, expect } from 'bun:test';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';

const FIXTURE = path.join(__dirname, '../../../../../e2e/fixtures/block-level-bookmark.docx');

describe('block-level bookmark (fixture)', () => {
  test('a bookmark between block elements survives a parse -> repack round-trip', async () => {
    const doc = await parseDocx(new Uint8Array(fs.readFileSync(FIXTURE)), { preloadFonts: false });
    const out = await repackDocx(doc, { updateModifiedDate: false });
    const docXml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('text');
    // The bookmark element (not merely its name in the REF instruction) must be re-emitted.
    expect(docXml).toContain('w:name="PaymentTerms"');
    expect((docXml.match(/<w:bookmarkStart\b/g) ?? []).length).toBe(1);
  });
});
