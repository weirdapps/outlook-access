// test_scripts/util-open-outlook.spec.ts

import { describe, it, expect, vi } from 'vitest';
import { EventEmitter } from 'events';

import { activateOutlookApp } from '../src/util/open-outlook';

function fakeChild(closeCode: number | null, opts: { errFirst?: boolean } = {}) {
  const ee = new EventEmitter() as EventEmitter & { unref: () => void };
  ee.unref = () => {};
  queueMicrotask(() => {
    if (opts.errFirst) {
      ee.emit('error', new Error('ENOENT'));
    } else if (closeCode !== null) {
      ee.emit('close', closeCode);
    }
  });
  return ee;
}

describe('activateOutlookApp', () => {
  it('skips on non-darwin platforms (no spawn, resolves)', async () => {
    const spawnFn = vi.fn();
    await expect(
      activateOutlookApp({ platform: 'linux', spawnFn: spawnFn as never }),
    ).resolves.toBeUndefined();
    expect(spawnFn).not.toHaveBeenCalled();
  });

  it('spawns `open -a "Microsoft Outlook"` on darwin and resolves on close 0', async () => {
    const spawnFn = vi.fn(() => fakeChild(0)) as unknown as typeof import('child_process').spawn;
    await expect(activateOutlookApp({ platform: 'darwin', spawnFn })).resolves.toBeUndefined();
    expect(spawnFn).toHaveBeenCalledWith('open', ['-a', 'Microsoft Outlook'], {
      stdio: 'ignore',
      detached: true,
    });
  });

  it('keeps the open child referenced until it closes', async () => {
    // An unref'd child does not hold the event loop, so Node could exit 0
    // while this promise was still pending: the caller never printed its
    // --json result, and the MCP bridge read a created draft as a failure.
    const child = fakeChild(0);
    const unref = vi.spyOn(child, 'unref');
    const spawnFn = vi.fn(() => child) as unknown as typeof import('child_process').spawn;
    await activateOutlookApp({ platform: 'darwin', spawnFn });
    expect(unref).not.toHaveBeenCalled();
  });

  it('rejects with explanatory error when open exits non-zero', async () => {
    const spawnFn = vi.fn(() => fakeChild(1)) as unknown as typeof import('child_process').spawn;
    await expect(activateOutlookApp({ platform: 'darwin', spawnFn })).rejects.toThrow(
      /exited with code 1.*Microsoft Outlook installed/,
    );
  });

  it('rejects when spawn emits an error event (e.g. open binary not found)', async () => {
    const spawnFn = vi.fn(() =>
      fakeChild(null, { errFirst: true }),
    ) as unknown as typeof import('child_process').spawn;
    await expect(activateOutlookApp({ platform: 'darwin', spawnFn })).rejects.toThrow(
      /spawn failed: ENOENT/,
    );
  });
});
