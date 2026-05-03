// test_scripts/util-signature-assets.spec.ts

import { describe, it, expect, vi } from 'vitest';

import {
  extractCidReferences,
  sanitizeContentIdForFile,
  saveSignatureAssets,
  loadManifest,
  loadSignatureAttachments,
} from '../src/util/signature-assets';

describe('extractCidReferences', () => {
  it('returns [] when no cid: refs in HTML', () => {
    expect(extractCidReferences('<p>hi</p>')).toEqual([]);
  });

  it('extracts a single cid reference', () => {
    const out = extractCidReferences('<img src="cid:image001.png@01DCD27B.DECD9E60">');
    expect(out).toEqual(['image001.png@01DCD27B.DECD9E60']);
  });

  it('handles multiple unique refs and deduplicates', () => {
    const html = '<img src="cid:logo">' + '<img src="cid:icon">' + '<img src="cid:logo">'; // duplicate
    expect(extractCidReferences(html)).toEqual(['logo', 'icon']);
  });

  it('accepts both single and double quotes', () => {
    const html = `<img src='cid:single'><img src="cid:double">`;
    expect(extractCidReferences(html)).toEqual(['single', 'double']);
  });

  it('case-insensitive on src= attribute', () => {
    const html = '<img SRC="cid:upper">';
    expect(extractCidReferences(html)).toEqual(['upper']);
  });

  it('preserves @ and . in contentId values', () => {
    const cid = 'image001.png@01DCD27B.DECD9E60';
    expect(extractCidReferences(`<img src="cid:${cid}">`)).toEqual([cid]);
  });
});

describe('sanitizeContentIdForFile', () => {
  it('preserves [A-Za-z0-9._@-]', () => {
    expect(sanitizeContentIdForFile('image001.png@01DCD27B.DECD9E60')).toBe(
      'image001.png@01DCD27B.DECD9E60',
    );
  });

  it('replaces unsafe chars with underscore', () => {
    expect(sanitizeContentIdForFile('weird/cid<with>spaces and|pipe')).toBe(
      'weird_cid_with_spaces_and_pipe',
    );
  });
});

describe('saveSignatureAssets + loadManifest round-trip', () => {
  it('writes assets + manifest, reads back equivalently', async () => {
    const written = new Map<string, string | Buffer>();
    const writeFile = vi.fn(async (p: string, data: Buffer | string) => {
      written.set(p, data);
    });
    const mkdir = vi.fn(async () => undefined);

    const result = await saveSignatureAssets({
      assetsDir: '/tmp/sig-assets',
      sourceMessageId: 'AAMk-source',
      attachments: [
        {
          contentId: 'image001.png@01DCD27B.DECD9E60',
          contentType: 'image/png',
          contentBytesBase64: Buffer.from('PNG-FAKE-BYTES').toString('base64'),
          name: 'logo.png',
        },
      ],
      writeFile,
      mkdir,
    });

    expect(result.version).toBe(1);
    expect(result.sourceMessageId).toBe('AAMk-source');
    expect(result.assets).toHaveLength(1);
    expect(result.assets[0]!.fileName).toBe('image001.png@01DCD27B.DECD9E60');
    expect(result.assets[0]!.contentType).toBe('image/png');

    // Verify writes
    expect(mkdir).toHaveBeenCalledWith('/tmp/sig-assets', { recursive: true, mode: 0o700 });
    expect(writeFile).toHaveBeenCalledTimes(2); // 1 binary + 1 manifest
    expect(written.has('/tmp/sig-assets/image001.png@01DCD27B.DECD9E60')).toBe(true);
    expect(written.has('/tmp/sig-assets/manifest.json')).toBe(true);

    // Round-trip via loadManifest
    const reader = vi.fn(async (p: string) => {
      const v = written.get(p);
      if (typeof v === 'string') return Buffer.from(v, 'utf-8');
      if (Buffer.isBuffer(v)) return v;
      throw new Error(`not in written: ${p}`);
    });
    const reloaded = await loadManifest('/tmp/sig-assets', reader);
    expect(reloaded?.sourceMessageId).toBe('AAMk-source');
    expect(reloaded?.assets[0]!.contentId).toBe('image001.png@01DCD27B.DECD9E60');
  });

  it('loadManifest returns null when manifest is missing', async () => {
    const reader = vi.fn(async () => {
      throw new Error('ENOENT');
    });
    expect(await loadManifest('/tmp/missing', reader)).toBeNull();
  });

  it('loadManifest returns null on malformed JSON', async () => {
    const reader = vi.fn(async () => Buffer.from('{not json'));
    expect(await loadManifest('/tmp/bad', reader)).toBeNull();
  });
});

describe('loadSignatureAttachments', () => {
  it('returns [] + empty unmatchedRefs when signature has no cid refs', async () => {
    const out = await loadSignatureAttachments({
      signatureHtml: '<p>just text signature</p>',
      assetsDir: '/tmp/sig-assets',
      reader: vi.fn(),
    });
    expect(out.attachments).toEqual([]);
    expect(out.unmatchedRefs).toEqual([]);
  });

  it('returns attachments matched against manifest', async () => {
    const written = new Map<string, Buffer>([
      [
        '/tmp/sig-assets/manifest.json',
        Buffer.from(
          JSON.stringify({
            version: 1,
            capturedAt: '2026-04-22T00:00:00Z',
            sourceMessageId: 'AAMk-src',
            assets: [
              {
                contentId: 'logo',
                fileName: 'logo',
                contentType: 'image/png',
                originalName: 'logo.png',
              },
            ],
          }),
        ),
      ],
      ['/tmp/sig-assets/logo', Buffer.from('IMAGE-BYTES')],
    ]);
    const reader = vi.fn(async (p: string) => {
      const v = written.get(p);
      if (!v) throw new Error(`ENOENT: ${p}`);
      return v;
    });

    const out = await loadSignatureAttachments({
      signatureHtml: '<img src="cid:logo">',
      assetsDir: '/tmp/sig-assets',
      reader,
    });
    expect(out.attachments).toHaveLength(1);
    expect(out.attachments[0]!.IsInline).toBe(true);
    expect(out.attachments[0]!.ContentId).toBe('logo');
    expect(out.attachments[0]!.Name).toBe('logo.png');
    expect(out.attachments[0]!.ContentType).toBe('image/png');
    expect(Buffer.from(out.attachments[0]!.ContentBytes, 'base64').toString('utf-8')).toBe(
      'IMAGE-BYTES',
    );
    expect(out.unmatchedRefs).toEqual([]);
  });

  it('reports unmatched refs when manifest is missing', async () => {
    const reader = vi.fn(async () => {
      throw new Error('ENOENT');
    });
    const out = await loadSignatureAttachments({
      signatureHtml: '<img src="cid:logo"><img src="cid:icon">',
      assetsDir: '/tmp/missing',
      reader,
    });
    expect(out.attachments).toEqual([]);
    expect(out.unmatchedRefs).toEqual(['logo', 'icon']);
  });

  it('partially resolves — some matched, some unmatched', async () => {
    const written = new Map<string, Buffer>([
      [
        '/tmp/sig-assets/manifest.json',
        Buffer.from(
          JSON.stringify({
            version: 1,
            capturedAt: '2026-04-22T00:00:00Z',
            sourceMessageId: 'AAMk-src',
            assets: [
              {
                contentId: 'logo',
                fileName: 'logo',
                contentType: 'image/png',
                originalName: 'logo.png',
              },
            ],
          }),
        ),
      ],
      ['/tmp/sig-assets/logo', Buffer.from('IMG')],
    ]);
    const reader = vi.fn(async (p: string) => {
      const v = written.get(p);
      if (!v) throw new Error(`ENOENT: ${p}`);
      return v;
    });
    const out = await loadSignatureAttachments({
      signatureHtml: '<img src="cid:logo"><img src="cid:icon">',
      assetsDir: '/tmp/sig-assets',
      reader,
    });
    expect(out.attachments).toHaveLength(1);
    expect(out.attachments[0]!.ContentId).toBe('logo');
    expect(out.unmatchedRefs).toEqual(['icon']);
  });
});
