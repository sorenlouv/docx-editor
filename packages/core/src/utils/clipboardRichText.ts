/**
 * Clipboard rich-text detection.
 *
 * Word and many rich editors place a bitmap snapshot of the copied selection on
 * the clipboard *alongside* the real HTML. When that HTML is present the editor
 * must let the normal paste pipeline (ProseMirror parseDOM) handle it so text,
 * styles and structure survive, instead of dropping in the snapshot as an image.
 */

import { CLIPBOARD_TYPES } from './clipboard';

/**
 * Whether the clipboard carries a usable rich-text (HTML) representation.
 */
export function clipboardHasRichText(clipboardData: DataTransfer | null): boolean {
  if (!clipboardData) return false;
  if (!clipboardData.types || !Array.from(clipboardData.types).includes(CLIPBOARD_TYPES.HTML)) {
    return false;
  }
  const html = clipboardData.getData(CLIPBOARD_TYPES.HTML);
  return typeof html === 'string' && html.trim().length > 0;
}
