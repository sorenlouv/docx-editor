/**
 * Extract a browser-renderable raster (PNG/JPEG/GIF) embedded inside an
 * EMF/WMF metafile.
 *
 * Word frequently stores header logos / OLE preview pictures as EMF — a
 * Windows GDI metafile browsers can't decode. In practice such an EMF almost
 * always carries the actual artwork as a single embedded PNG or JPEG (an
 * `EmfPlusObject` bitmap record, or a `StretchDIBits` payload). Rather than
 * implement a GDI renderer, this scans the byte stream for a raster signature,
 * brackets it to its container's end marker, and returns the slice. The caller
 * uses it as the media entry's `dataUrl` so `<img>` just works; the original
 * EMF bytes stay on `MediaFile.data` for round-trip.
 *
 * When no embedded raster is found this returns `null`; callers fall back to a
 * sized placeholder and/or the host-supplied `mediaResolver` hook.
 */

export interface ExtractedRaster {
  bytes: Uint8Array;
  mimeType: 'image/png' | 'image/jpeg';
}

function indexOfBytes(haystack: Uint8Array, needle: number[], from = 0): number {
  outer: for (let i = from; i + needle.length <= haystack.length; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (haystack[i + j] !== needle[j]) continue outer;
    }
    return i;
  }
  return -1;
}

function extractPng(bytes: Uint8Array): Uint8Array | null {
  const start = indexOfBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  if (start < 0) return null;
  // Walk PNG chunks until IEND so we stop exactly at the PNG boundary even
  // when more EMF records follow.
  let off = start + 8;
  while (off + 12 <= bytes.length) {
    const len =
      (bytes[off] << 24) | (bytes[off + 1] << 16) | (bytes[off + 2] << 8) | bytes[off + 3];
    const type = String.fromCharCode(
      bytes[off + 4],
      bytes[off + 5],
      bytes[off + 6],
      bytes[off + 7]
    );
    const next = off + 12 + (len >>> 0);
    if (type === 'IEND') return bytes.slice(start, off + 12);
    if (next <= off || next > bytes.length) break;
    off = next;
  }
  return null;
}

function extractJpeg(bytes: Uint8Array): Uint8Array | null {
  const start = indexOfBytes(bytes, [0xff, 0xd8, 0xff]);
  if (start < 0) return null;
  // JPEG ends at the EOI marker FF D9. An EXIF/APP1 thumbnail nests its own
  // SOI…EOI inside the outer stream, so the FIRST FF D9 may belong to the
  // thumbnail; take the LAST one before the EMF resumes (the EMF record that
  // wraps the JPEG sizes it exactly, so trailing EMF bytes after the outer
  // EOI are the next record, not image data).
  let end = -1;
  for (let i = bytes.length - 2; i > start + 1; i--) {
    if (bytes[i] === 0xff && bytes[i + 1] === 0xd9) {
      end = i + 2;
      break;
    }
  }
  return end > start ? bytes.slice(start, end) : null;
}

/**
 * Scan a metafile (EMF or WMF) for an embedded browser-renderable raster.
 * Returns the first PNG, else first JPEG found; `null` when none is present.
 */
export function extractMetafileRaster(data: ArrayBuffer | Uint8Array): ExtractedRaster | null {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.length < 32) return null;

  const png = extractPng(bytes);
  if (png) return { bytes: png, mimeType: 'image/png' };

  const jpeg = extractJpeg(bytes);
  if (jpeg) return { bytes: jpeg, mimeType: 'image/jpeg' };

  return null;
}

const METAFILE_MIME = new Set(['image/x-emf', 'image/emf', 'image/x-wmf', 'image/wmf']);

/** True for EMF/WMF MIME types — the formats browsers can't render natively. */
export function isMetafileMimeType(mimeType: string | undefined): boolean {
  return !!mimeType && METAFILE_MIME.has(mimeType);
}
