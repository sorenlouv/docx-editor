/**
 * Block-level bookmark marker serialization.
 *
 * `w:bookmarkStart`/`w:bookmarkEnd` markers can sit as direct children of a
 * block container (`w:body`/`w:tc`/`w:sdtContent`) between paragraphs/tables/
 * SDTs (e.g. `</w:p><w:bookmarkEnd/><w:p>`). The block content model has no
 * slot for them, so the parser rides them on the adjacent block via
 * `leadingBlockMarkers` / `trailingBlockMarkers` (see `BlockMarkerCollector`).
 * This wrapper re-emits them in their original position around the block's XML.
 */

import type { BookmarkStart, BookmarkEnd } from '../../types/document';
import { serializeBookmarkStart, serializeBookmarkEnd } from './paragraphSerializer/content';

/** A block that may carry orphaned block-level bookmark markers. */
interface BlockWithMarkers {
  leadingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
  trailingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
}

/** Serialize a single block-level bookmark marker. */
function serializeBlockMarker(marker: BookmarkStart | BookmarkEnd): string {
  return marker.type === 'bookmarkStart'
    ? serializeBookmarkStart(marker)
    : serializeBookmarkEnd(marker);
}

/**
 * Wrap a block's serialized XML with its leading/trailing block-level bookmark
 * markers, reproducing the original `<w:bookmarkStart/><w:p>...</w:p><w:bookmarkEnd/>`
 * shape. A no-op for blocks without markers (the common case).
 */
export function wrapBlockMarkers(block: BlockWithMarkers, innerXml: string): string {
  const leading = block.leadingBlockMarkers;
  const trailing = block.trailingBlockMarkers;
  if (!leading?.length && !trailing?.length) return innerXml;

  const before = leading?.map(serializeBlockMarker).join('') ?? '';
  const after = trailing?.map(serializeBlockMarker).join('') ?? '';
  return `${before}${innerXml}${after}`;
}
