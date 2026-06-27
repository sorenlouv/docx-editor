/**
 * Accepting / rejecting a tracked change by its Word revision id, and undoing /
 * redoing those resolutions.
 *
 * This pins the behaviour the editor's imperative ref surfaces to embedders
 * (`acceptChange` / `rejectChange` / `undo` / `redo`): a host can resolve a
 * change — even a coalesced one whose sites are scattered across the document —
 * in a single transaction, then take it back with one `undo()`, all applied IN
 * PLACE through ProseMirror history (no document reload). These are command +
 * history tests: they exercise the same prosemirror-history mechanism the ref's
 * `undo()` drives, not the full React ref wiring.
 */
import { describe, test, expect } from 'bun:test';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history, undo, redo, undoDepth } from 'prosemirror-history';
import { acceptChangeById, rejectChangeById } from './comments';

// Minimal schema carrying the inline tracked-change marks and the paragraph-mark
// attrs the resolver reads (it keys off mark/attr names and `revisionId`).
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { pPrIns: { default: null }, pPrDel: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
  },
  marks: {
    insertion: { attrs: { revisionId: { default: 0 } }, toDOM: () => ['ins', 0] },
    deletion: { attrs: { revisionId: { default: 0 } }, toDOM: () => ['del', 0] },
  },
});

const ins = (text: string, id: number) =>
  schema.text(text, [schema.marks.insertion.create({ revisionId: id })]);
const del = (text: string, id: number) =>
  schema.text(text, [schema.marks.deletion.create({ revisionId: id })]);
const para = (content: PMNode[], attrs?: Record<string, unknown>) =>
  schema.node('paragraph', attrs ?? null, content);
const doc = (...paras: PMNode[]) => schema.node('doc', null, paras);

/** One paragraph: `keep `, then `tracked` carrying `markName`@`id`, then ` tail`. */
function docWith(markName: 'insertion' | 'deletion', id: number): PMNode {
  const run = markName === 'insertion' ? ins('tracked', id) : del('tracked', id);
  return doc(para([schema.text('keep '), run, schema.text(' tail')]));
}

/** An editor with history(), so undo/redo run the same mechanism as the ref's undo(). */
function editor(d: PMNode) {
  let state = EditorState.create({ doc: d, plugins: [history()] });
  return {
    get state() {
      return state;
    },
    run(cmd: Command) {
      return cmd(state, (tr) => {
        state = state.apply(tr);
      });
    },
  };
}

const text = (s: EditorState) => s.doc.textContent;
const hasMark = (s: EditorState, name: string) => {
  let found = false;
  s.doc.descendants((n) => {
    if (n.marks.some((m) => m.type.name === name)) found = true;
  });
  return found;
};

describe('acceptChangeById / rejectChangeById', () => {
  test('accepting an insertion keeps the text and drops the mark', () => {
    const ed = editor(docWith('insertion', 7));
    expect(ed.run(acceptChangeById(7))).toBe(true);
    expect(text(ed.state)).toBe('keep tracked tail');
    expect(hasMark(ed.state, 'insertion')).toBe(false);
  });

  test('rejecting an insertion removes the text', () => {
    const ed = editor(docWith('insertion', 7));
    expect(ed.run(rejectChangeById(7))).toBe(true);
    expect(text(ed.state)).toBe('keep  tail');
    expect(hasMark(ed.state, 'insertion')).toBe(false);
  });

  test('accepting a deletion removes the text', () => {
    const ed = editor(docWith('deletion', 9));
    expect(ed.run(acceptChangeById(9))).toBe(true);
    expect(text(ed.state)).toBe('keep  tail');
  });

  test('rejecting a deletion keeps the text and drops the mark', () => {
    const ed = editor(docWith('deletion', 9));
    expect(ed.run(rejectChangeById(9))).toBe(true);
    expect(text(ed.state)).toBe('keep tracked tail');
    expect(hasMark(ed.state, 'deletion')).toBe(false);
  });

  test('is a no-op (returns false) for an unknown revision id', () => {
    const ed = editor(docWith('insertion', 1));
    expect(ed.run(acceptChangeById(999))).toBe(false);
    expect(text(ed.state)).toBe('keep tracked tail');
    expect(hasMark(ed.state, 'insertion')).toBe(true);
  });
});

describe('undo / redo of a resolution — in-place reversal, no reload', () => {
  test('undo restores an accepted insertion mark; redo drops it again', () => {
    const ed = editor(docWith('insertion', 5));
    ed.run(acceptChangeById(5));
    expect(hasMark(ed.state, 'insertion')).toBe(false);

    ed.run(undo);
    expect(hasMark(ed.state, 'insertion')).toBe(true);
    expect(text(ed.state)).toBe('keep tracked tail');

    ed.run(redo);
    expect(hasMark(ed.state, 'insertion')).toBe(false);
  });

  test('undo restores text removed by accepting a deletion', () => {
    const ed = editor(docWith('deletion', 8));
    ed.run(acceptChangeById(8));
    expect(text(ed.state)).toBe('keep  tail');

    ed.run(undo);
    expect(text(ed.state)).toBe('keep tracked tail');
    expect(hasMark(ed.state, 'deletion')).toBe(true);
  });

  test('undo restores text + mark removed by rejecting an insertion; redo re-removes', () => {
    const ed = editor(docWith('insertion', 6));
    ed.run(rejectChangeById(6));
    expect(text(ed.state)).toBe('keep  tail');
    expect(hasMark(ed.state, 'insertion')).toBe(false);

    ed.run(undo);
    expect(text(ed.state)).toBe('keep tracked tail');
    expect(hasMark(ed.state, 'insertion')).toBe(true);

    ed.run(redo);
    expect(text(ed.state)).toBe('keep  tail');
  });

  test('resolves all sites of a coalesced revision in one undoable step', () => {
    // The same revisionId on two non-contiguous runs across two paragraphs.
    const ed = editor(
      doc(
        para([schema.text('a '), ins('one', 4), schema.text(' b')]),
        para([schema.text('c '), ins('two', 4), schema.text(' d')])
      )
    );
    const before = undoDepth(ed.state);
    expect(ed.run(acceptChangeById(4))).toBe(true);
    expect(hasMark(ed.state, 'insertion')).toBe(false); // every site resolved...
    expect(undoDepth(ed.state)).toBe(before + 1); // ...in a single history step,
    ed.run(undo);
    expect(hasMark(ed.state, 'insertion')).toBe(true); // so one undo restores them all.
  });

  test('a selection-only change is never recorded, so undo() still targets the last edit', () => {
    const ed = editor(docWith('insertion', 2));
    ed.run(acceptChangeById(2));
    expect(hasMark(ed.state, 'insertion')).toBe(false);
    // Move the selection — no document steps, so prosemirror-history ignores it
    // (the public undo()'s documented guarantee that scroll/locate isn't undone).
    ed.run((state, dispatch) => {
      dispatch?.(state.tr.setSelection(TextSelection.create(state.doc, 1)));
      return true;
    });
    ed.run(undo);
    expect(hasMark(ed.state, 'insertion')).toBe(true); // reverted the ACCEPT, not the selection
  });
});

describe('structural (paragraph-mark) revisions', () => {
  test('rejecting an inserted paragraph mark joins the paragraphs; accepting clears the marker', () => {
    const make = () =>
      editor(
        doc(
          para([schema.text('first')], { pPrIns: { revisionId: 11 } }),
          para([schema.text('second')])
        )
      );

    const rej = make();
    expect(rej.run(rejectChangeById(11))).toBe(true);
    expect(rej.state.doc.childCount).toBe(1); // the inserted break is undone → paragraphs join

    const acc = make();
    expect(acc.run(acceptChangeById(11))).toBe(true);
    expect(acc.state.doc.childCount).toBe(2); // break kept...
    expect(acc.state.doc.firstChild?.attrs.pPrIns).toBeNull(); // ...marker cleared
  });
});
