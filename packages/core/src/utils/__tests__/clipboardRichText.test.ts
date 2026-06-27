import { describe, expect, test } from 'bun:test';
import { clipboardHasRichText } from '../clipboardRichText';

describe('clipboardHasRichText', () => {
  const makeClipboard = (types: string[], data: Record<string, string>): DataTransfer =>
    ({
      types,
      getData: (type: string) => data[type] ?? '',
    }) as unknown as DataTransfer;

  test('is true when non-empty HTML is present (Word snapshot + HTML case)', () => {
    const clipboard = makeClipboard(['text/html', 'text/plain'], {
      'text/html': '<p class="MsoNormal">Hello</p>',
      'text/plain': 'Hello',
    });
    expect(clipboardHasRichText(clipboard)).toBe(true);
  });

  test('is false when HTML is absent (plain screenshot paste)', () => {
    const clipboard = makeClipboard(['Files'], {});
    expect(clipboardHasRichText(clipboard)).toBe(false);
  });

  test('is false when HTML type is present but empty', () => {
    const clipboard = makeClipboard(['text/html'], { 'text/html': '   ' });
    expect(clipboardHasRichText(clipboard)).toBe(false);
  });

  test('is false when clipboardData is null', () => {
    expect(clipboardHasRichText(null)).toBe(false);
  });
});
