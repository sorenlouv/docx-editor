import type { RefObject } from 'react';
import {
  acceptChangeById,
  rejectChangeById,
} from '@eigenpal/docx-editor-core/prosemirror/commands';
import { extractTrackedChanges } from '../../hooks/useTrackedChanges';
import type { PagedEditorRef } from './PagedEditor';

/**
 * Accept or reject a tracked change by its Word revision `w:id`, resolving it in
 * whichever editor view actually holds it — the document body or a header /
 * footer — in a single in-place transaction. A coalesced revision is matched by
 * its `revisionId`, `insertionRevisionId`, or any of its `coalescedRevisionIds`.
 *
 * Header/footer views paginate separately from the body, so when the change
 * lived in one, `onHeaderFooterResolved` is invoked to let the caller repaint
 * those bands.
 *
 * @returns `false` when the revision id is not present in any view.
 */
export function resolveTrackedChangeById(
  pagedEditorRef: RefObject<PagedEditorRef | null>,
  revisionId: number,
  mode: 'accept' | 'reject',
  onHeaderFooterResolved?: () => void
): boolean {
  const resolve = mode === 'accept' ? acceptChangeById(revisionId) : rejectChangeById(revisionId);

  const hfViews = pagedEditorRef.current?.getHfPmViews?.();
  if (hfViews) {
    for (const view of hfViews.values()) {
      const { entries } = extractTrackedChanges(view.state);
      const holdsRevision = entries.some(
        (e) =>
          e.revisionId === revisionId ||
          e.insertionRevisionId === revisionId ||
          e.coalescedRevisionIds?.includes(revisionId)
      );
      if (holdsRevision) {
        const resolved = resolve(view.state, view.dispatch);
        onHeaderFooterResolved?.();
        return resolved;
      }
    }
  }

  const body = pagedEditorRef.current?.getView();
  return body ? resolve(body.state, body.dispatch) : false;
}
