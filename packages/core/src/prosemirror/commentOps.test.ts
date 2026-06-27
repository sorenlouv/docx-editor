/**
 * The id-bearing write-into-preview seam: proposeChange / addComment can adopt an
 * OOXML id assigned by an external source-of-truth model (e.g.
 * @beyondwork/docx-react-component) instead of minting their own, so the preview
 * mirrors a change/comment under the SAME id — and a later acceptChange/
 * rejectChange/resolve resolves it by that shared id. Without a supplied id the
 * allocator still mints, unchanged.
 */

import { describe, test, expect } from 'bun:test';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { addCommentToRange, applyProposedChange } from './commentOps';
import { createCommentIdAllocator } from './commentIdAllocator';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { paraId: { default: null }, pPrIns: { default: null }, pPrDel: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
  },
  marks: {
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: null } },
      toDOM: () => ['ins', 0],
    },
    deletion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: null } },
      toDOM: () => ['del', 0],
    },
    comment: { attrs: { commentId: { default: 0 } }, toDOM: () => ['span', 0] },
  },
});

/** A minimal EditorView stand-in: applyProposedChange/addCommentToRange only read
 *  `view.state` and call `view.dispatch(tr)`, so no DOM is needed. */
function fakeView(doc: PMNode): EditorView {
  let state = EditorState.create({ doc });
  return {
    get state() {
      return state;
    },
    dispatch(tr: ReturnType<typeof state.tr.scrollIntoView>) {
      state = state.apply(tr);
    },
  } as unknown as EditorView;
}

function para(paraId: string, text: string): PMNode {
  return schema.nodes.paragraph.create({ paraId }, text ? [schema.text(text)] : []);
}

/** Every distinct revisionId carried by insertion/deletion marks in the doc. */
function revisionIds(doc: PMNode): Set<number> {
  const ids = new Set<number>();
  doc.descendants((node) => {
    for (const m of node.marks) {
      if (
        (m.type === schema.marks.insertion || m.type === schema.marks.deletion) &&
        m.attrs.revisionId != null
      ) {
        ids.add(m.attrs.revisionId as number);
      }
    }
  });
  return ids;
}

function commentIds(doc: PMNode): Set<number> {
  const ids = new Set<number>();
  doc.descendants((node) => {
    for (const m of node.marks) {
      if (m.type === schema.marks.comment) ids.add(m.attrs.commentId as number);
    }
  });
  return ids;
}

describe('proposeChange: external revisionId', () => {
  test('adopts a supplied revisionId on both the ins and del marks, and seeds the allocator above it', () => {
    const allocator = createCommentIdAllocator();
    const view = fakeView(schema.nodes.doc.create({}, [para('P1', 'net 30 days')]));

    const ok = applyProposedChange(
      view,
      { paraId: 'P1', search: '30', replaceWith: '45', author: 'CLM', revisionId: 4242 },
      allocator
    );

    expect(ok).toBe(true);
    // The whole replacement (deletion of '30' + insertion of '45') carries the
    // supplied id — not a freshly-minted one.
    expect(revisionIds(view.state.doc)).toEqual(new Set([4242]));
    // Allocator is bumped past the adopted id so a later mint can't collide.
    expect(allocator.next()).toBe(4243);
  });

  test('still mints from the allocator when no revisionId is supplied (unchanged behavior)', () => {
    const allocator = createCommentIdAllocator(); // first next() === 1
    const view = fakeView(schema.nodes.doc.create({}, [para('P1', 'net 30 days')]));

    const ok = applyProposedChange(
      view,
      { paraId: 'P1', search: '30', replaceWith: '45', author: 'CLM' },
      allocator
    );

    expect(ok).toBe(true);
    expect(revisionIds(view.state.doc)).toEqual(new Set([1]));
    expect(allocator.next()).toBe(2);
  });
});

describe('addComment: external commentId', () => {
  test('adopts a supplied commentId and seeds the allocator above it', () => {
    const allocator = createCommentIdAllocator();
    const view = fakeView(schema.nodes.doc.create({}, [para('P2', 'Confidential Information')]));

    const comment = addCommentToRange(
      view,
      { paraId: 'P2', text: 'Please confirm the carve-out.', author: 'CLM', commentId: 9999 },
      allocator
    );

    expect(comment?.id).toBe(9999);
    expect(commentIds(view.state.doc)).toEqual(new Set([9999]));
    expect(allocator.next()).toBe(10000);
  });

  test('still mints from the allocator when no commentId is supplied (unchanged behavior)', () => {
    const allocator = createCommentIdAllocator(); // first next() === 1
    const view = fakeView(schema.nodes.doc.create({}, [para('P2', 'Confidential Information')]));

    const comment = addCommentToRange(
      view,
      { paraId: 'P2', text: 'note', author: 'CLM' },
      allocator
    );

    expect(comment?.id).toBe(1);
    expect(allocator.next()).toBe(2);
  });
});
