// src/util/open-outlook.ts
//
// Activates Microsoft Outlook desktop on macOS via `open -a "Microsoft Outlook"`.
// Used by send-mail after creating a draft so the user can immediately switch
// to Drafts folder and review/send.
//
// macOS-only: on other platforms emits a one-line stderr warning and resolves
// without doing anything (so callers don't need to branch on platform).

import { spawn } from 'child_process';

const APP_NAME = 'Microsoft Outlook';

export interface ActivateOutlookOptions {
  /** Override platform detection — used by tests. */
  platform?: NodeJS.Platform;
  /**
   * Override the spawn function — used by tests. Real callers should leave
   * this unset so the default `child_process.spawn` is used.
   */
  spawnFn?: typeof spawn;
}

/**
 * Brings Microsoft Outlook to the foreground on macOS. Resolves on success
 * (open exited 0) or when running on a non-darwin platform (no-op).
 *
 * Rejects with an Error when `open` exits non-zero or fails to launch
 * (typically because Outlook is not installed in /Applications).
 */
export async function activateOutlookApp(opts: ActivateOutlookOptions = {}): Promise<void> {
  const platform = opts.platform ?? process.platform;
  if (platform !== 'darwin') {
    process.stderr.write(
      `open-outlook: skipping (platform=${platform}, only darwin is supported)\n`,
    );
    return;
  }
  const spawner = opts.spawnFn ?? spawn;
  return new Promise((resolve, reject) => {
    const child = spawner('open', ['-a', APP_NAME], {
      stdio: 'ignore',
      detached: true,
    });
    child.on('close', (code) => {
      if (code === 0) {
        resolve();
      } else {
        reject(
          new Error(
            `open -a "${APP_NAME}" exited with code ${code} ` + '(is Microsoft Outlook installed?)',
          ),
        );
      }
    });
    child.on('error', (err) => reject(new Error(`open-outlook: spawn failed: ${err.message}`)));
    child.unref();
  });
}
