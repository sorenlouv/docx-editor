import { describe, expect, test } from 'bun:test';
import { ImagePasteExtension } from './ImagePasteExtension';

/**
 * Pull the plugin's `paste` DOM-event handler out of the built extension so we
 * can drive it directly with a faked view/clipboard.
 */
function getPasteHandler(): (view: unknown, event: unknown) => boolean {
  const runtime = ImagePasteExtension().onSchemaReady({} as never);
  const plugin = runtime.plugins?.[0] as unknown as {
    props?: { handleDOMEvents?: { paste?: (view: unknown, event: unknown) => boolean } };
  };
  const paste = plugin?.props?.handleDOMEvents?.paste;
  if (!paste) throw new Error('paste handler not found on ImagePasteExtension plugin');
  return paste;
}

describe('ImagePasteExtension paste guard (issue #981)', () => {
  test('defers to the HTML pipeline when the clipboard carries both rich HTML and an image', () => {
    const paste = getPasteHandler();

    // Word puts a bitmap snapshot on the clipboard next to the real HTML.
    const snapshot = new File([new Uint8Array([1, 2, 3])], 'snapshot.png', { type: 'image/png' });
    let prevented = false;
    const clipboardData = {
      types: ['text/html', 'text/plain', 'Files'],
      getData: (type: string) =>
        type === 'text/html' ? '<p class="MsoNormal">Hello from Word</p>' : 'Hello from Word',
      files: [snapshot],
      items: [{ kind: 'file', type: 'image/png', getAsFile: () => snapshot }],
    };
    const event = {
      clipboardData,
      preventDefault: () => {
        prevented = true;
      },
    };
    // schema has an image node, so an unguarded handler WOULD intercept here.
    const view = { state: { schema: { nodes: { image: {} } } } };

    const result = paste(view, event);

    // Returning false (without preventDefault) lets ProseMirror's parseDOM pipeline
    // parse the Word HTML instead of inserting the snapshot image.
    expect(result).toBe(false);
    expect(prevented).toBe(false);
  });

  test('does not intercept when there is no image on the clipboard', () => {
    const paste = getPasteHandler();

    let prevented = false;
    const clipboardData = {
      types: ['text/html'],
      getData: () => '<p>plain rich text</p>',
      files: [],
      items: [],
    };
    const event = {
      clipboardData,
      preventDefault: () => {
        prevented = true;
      },
    };
    const view = { state: { schema: { nodes: { image: {} } } } };

    expect(paste(view, event)).toBe(false);
    expect(prevented).toBe(false);
  });
});
