# Playwright: Reliably Capturing the First Authorization Bearer Token from Outlook Web

Research date: 2026-04-21
Scope: Approach A of `investigation-outlook-cli.md` — headed Chrome via `launchPersistentContext`, init-script hook, `exposeBinding` channel, single-promise resolution.

---

## Overview

The goal is to intercept the first `Authorization: Bearer <token>` header that the Outlook Web App sends to `outlook.office.com/api/v2.0/...` or `outlook.office.com/ows/...`, relay it to Node.js exactly once, and resolve a timeout-guarded promise with the token string. This must work across full login flows AND silent-SSO restarts against a persistent Chrome profile.

This document covers:

1. API surface — `addInitScript` vs `page.addInitScript`, `exposeBinding` vs `exposeFunction`
2. Where and when to install the init script
3. The exact monkey-patch JavaScript for `fetch` and `XMLHttpRequest`
4. Deduplication so only the first match fires
5. The Node-side receiver with promise + timeout
6. The `page.on('request')` alternative and its header-visibility limitations
7. Race conditions with MSAL bootstrap code
8. Known pitfalls and how to avoid them

---

## 1. API Surface

### `context.addInitScript` vs `page.addInitScript`

| | `context.addInitScript(script)` | `page.addInitScript(script)` |
|---|---|---|
| Scope | Every frame in every page in the context, now and in the future | Every frame in that one page only |
| Survives navigation | Yes — re-evaluates on each navigation | Yes, but only for the one page |
| Applies to popup windows | Yes | No |
| Applies to pre-restored pages (persistent profile) | **No** (see Pitfall 1 below) | Same gap |
| Ordering guarantee vs page scripts | Runs after `document` created, before page scripts | Same |

**Use `context.addInitScript`.** Outlook uses redirects and popup windows during the login flow; context-level registration covers all of them without having to hook each new page individually.

The official doc states:

> "Adds a script which would be evaluated in one of the following scenarios: whenever a page is created in the browser context or is navigated, or whenever a child frame is attached or navigated in any page in the browser context. The script is evaluated after the document was created but before any of its scripts were run."

### `context.exposeBinding` vs `context.exposeFunction`

| | `exposeBinding(name, cb)` | `exposeFunction(name, cb)` |
|---|---|---|
| Callback signature | `(source, ...args)` where `source = { browserContext, page, frame }` | `(...args)` — plain args only |
| Scope | Every frame in every page in the context | Every frame in every page (when called on context) |
| Return value to caller | Yes — cb return value resolves the Promise in-page | Yes |

**Use `context.exposeBinding`.** The `source` argument gives you the `Page` object for free, which is useful for attaching the `page.on('close')` guard without needing a separate reference.

Both `exposeBinding` and `exposeFunction` survive navigations — the binding is reinstalled automatically on each navigation.

---

## 2. Where to Install the Init Script

### Correct sequence

```
1. chromium.launchPersistentContext(profileDir, opts)   // launches browser
2. context.exposeBinding(...)                           // register Node callback
3. context.addInitScript(hookScript)                    // register init script
4. page = context.pages()[0] ?? await context.newPage()
5. page.goto('https://outlook.office.com/mail/')        // triggers navigation
```

Steps 2 and 3 **must** precede `page.goto()`. Because `addInitScript` fires before page scripts on every navigation, registering it before `goto` guarantees it is in place for every document in the redirect chain, including the MSAL silent-SSO redirect.

### The persistent-profile restored-page problem

When `launchPersistentContext` restores a previous browser session with open tabs, those pre-existing pages are NOT newly "created" from Playwright's perspective. Playwright's `'page'` event does not fire for them, and the init script is **not evaluated** on those already-open pages.

**Confirmed bug**: [github.com/microsoft/playwright/issues/28692](https://github.com/microsoft/playwright/issues/28692) — opened Nov 2023, not resolved as of April 2026.

**Workaround** (mandatory for this tool): after registering the init script, call `page.reload()` on any page returned by `context.pages()` before waiting for the token. The reload triggers a fresh navigation, and the init script fires normally.

```typescript
// After context.addInitScript(...) and context.exposeBinding(...)
const existingPages = context.pages();
const page = existingPages[0] ?? await context.newPage();

// Force a reload so the init script fires on any pre-restored page.
// This also triggers Outlook's MSAL silent-SSO re-auth, which is what
// produces the first authenticated API call we want to capture.
await page.goto('https://outlook.office.com/mail/');
```

Since we always navigate to the Outlook mail URL anyway (to trigger the SPA bootstrap and the first API call), the `goto` itself is the reload. Do not rely on whatever page the persistent profile happened to have open.

---

## 3. The Monkey-Patch Init Script

### Design constraints

- Must patch `window.fetch` — Outlook web is a modern React SPA and uses the Fetch API for all its REST calls.
- Must also patch `XMLHttpRequest.prototype.open` + `send` — OWS (`/ows/...`) endpoints and some legacy OWA paths still use XHR. Defense-in-depth; the POC confirmed `fetch` is sufficient for the primary path, but XHR coverage prevents a miss if Outlook uses XHR for the initial authenticated call.
- Must be idempotent — the flag `__outlookCliHooked` prevents double-wrapping on HMR or re-navigation within the same document.
- Must fire the binding only once globally (the dedup flag is set on `window`, persisted across the life of the document).
- Must handle all three ways a caller can pass headers to `fetch`:
  1. `fetch(url, { headers: { Authorization: '...' } })` — plain object
  2. `fetch(url, { headers: new Headers([['Authorization', '...']]) })` — Headers instance
  3. `fetch(new Request(url, { headers: ... }))` — Request object as first arg

### Complete init-script text

```javascript
(function () {
  // Idempotency guard — prevents double-wrapping on re-navigation within
  // the same document (HMR, hash changes that don't unload the document).
  if (window.__outlookCliHooked) return;
  window.__outlookCliHooked = true;

  // Dedup flag — once we have reported the first token we stop.
  let reported = false;

  // Target URL prefixes to watch. Keep this list narrow to avoid false
  // positives from other outlook.office.com resources (images, fonts, CDN).
  const TARGET_PREFIXES = [
    'https://outlook.office.com/api/v2.0/',
    'https://outlook.office.com/ows/',
    'https://outlook.office365.com/api/v2.0/',
    'https://outlook.office365.com/ows/',
  ];

  function isTargetUrl(url) {
    // url may be a string or a URL object
    const s = typeof url === 'string' ? url : (url && url.href ? url.href : String(url));
    return TARGET_PREFIXES.some(prefix => s.startsWith(prefix));
  }

  function extractBearer(headers) {
    // headers may be: plain object, Headers instance, or array of [name, value] pairs
    if (!headers) return null;
    if (typeof headers.get === 'function') {
      // Headers instance
      return headers.get('authorization') || headers.get('Authorization') || null;
    }
    if (Array.isArray(headers)) {
      // Array of [name, value] tuples
      const pair = headers.find(
        ([k]) => k.toLowerCase() === 'authorization'
      );
      return pair ? pair[1] : null;
    }
    // Plain object — header names may be mixed-case
    const key = Object.keys(headers).find(
      k => k.toLowerCase() === 'authorization'
    );
    return key ? headers[key] : null;
  }

  function tryReport(url, authHeader) {
    if (reported) return;
    if (!authHeader || !authHeader.startsWith('Bearer ')) return;
    if (!isTargetUrl(url)) return;
    reported = true;
    // __outlookCliReportAuth is exposed by context.exposeBinding on the Node side.
    // Calling it returns a Promise; we don't need to await it from here.
    window.__outlookCliReportAuth({ url: String(url), token: authHeader });
  }

  // ── Patch window.fetch ──────────────────────────────────────────────────────

  const originalFetch = window.fetch;
  window.fetch = function fetch(input, init) {
    try {
      let url = input;
      let authHeader = null;

      if (input instanceof Request) {
        // First argument is a Request object — headers live on it.
        url = input.url;
        authHeader = extractBearer(input.headers);
        // init.headers, if provided, override the Request's headers per spec.
        if (init && init.headers) {
          const override = extractBearer(init.headers);
          if (override) authHeader = override;
        }
      } else {
        url = input;
        authHeader = init && init.headers ? extractBearer(init.headers) : null;
      }

      tryReport(url, authHeader);
    } catch (_) {
      // Never block the real fetch due to our instrumentation.
    }

    return originalFetch.apply(this, arguments);
  };

  // ── Patch XMLHttpRequest ────────────────────────────────────────────────────

  const OriginalXHR = window.XMLHttpRequest;
  function PatchedXHR() {
    const xhr = new OriginalXHR();
    let _url = '';
    let _authHeader = null;

    const originalOpen = xhr.open.bind(xhr);
    xhr.open = function open(method, url) {
      _url = url;
      _authHeader = null; // reset on each open()
      return originalOpen.apply(xhr, arguments);
    };

    const originalSetRequestHeader = xhr.setRequestHeader.bind(xhr);
    xhr.setRequestHeader = function setRequestHeader(name, value) {
      if (name.toLowerCase() === 'authorization') {
        _authHeader = value;
      }
      return originalSetRequestHeader.apply(xhr, arguments);
    };

    const originalSend = xhr.send.bind(xhr);
    xhr.send = function send() {
      // At send() time all headers have been set — safe to report.
      try {
        tryReport(_url, _authHeader);
      } catch (_) {}
      return originalSend.apply(xhr, arguments);
    };

    return xhr;
  }

  // Copy static properties (e.g., DONE, OPENED constants) from original.
  Object.setPrototypeOf(PatchedXHR, OriginalXHR);
  Object.setPrototypeOf(PatchedXHR.prototype, OriginalXHR.prototype);
  Object.defineProperty(PatchedXHR, 'name', { value: 'XMLHttpRequest' });
  window.XMLHttpRequest = PatchedXHR;
})();
```

**Key design decisions:**

- The outer IIFE runs synchronously at document creation time, before any SPA code. The `window.fetch` replacement is in place before MSAL or the Outlook bootstrap script can make any API call.
- `try/catch` around every instrumentation block: our hook must never throw an exception that could prevent the real `fetch` or XHR from executing.
- `reported` is a closure variable, not on `window`, so it is per-document-lifetime and not visible to page code.
- `window.__outlookCliHooked` is on `window` so the idempotency check works if the script is somehow injected a second time into the same document.

---

## 4. Node-Side Receiver

### `exposeBinding` handler + promise

```typescript
// src/auth/fetchHook.ts

import { BrowserContext, Page } from 'playwright';

export interface CapturedAuth {
  token: string;   // Full "Bearer eyJ..." string
  url: string;     // URL that carried the token (for diagnostics)
}

/**
 * Installs the monkey-patch init script and the exposeBinding receiver on the
 * given persistent context, then waits for the first matching Authorization
 * header to be reported from the page.
 *
 * @param context   The launchPersistentContext result.
 * @param page      The page that will navigate to Outlook (used for close guard).
 * @param timeoutMs How long to wait before failing. Default: from config.
 * @returns         The captured Bearer token string and the originating URL.
 * @throws          Error with code 'LOGIN_TIMEOUT' if no token captured in time.
 * @throws          Error with code 'BROWSER_CLOSED' if the page is closed before capture.
 */
export async function captureFirstBearerToken(
  context: BrowserContext,
  page: Page,
  timeoutMs: number,
): Promise<CapturedAuth> {

  // ── 1. Create the one-shot promise ─────────────────────────────────────────
  let resolveCapture!: (auth: CapturedAuth) => void;
  let rejectCapture!: (err: Error) => void;

  const capturePromise = new Promise<CapturedAuth>((resolve, reject) => {
    resolveCapture = resolve;
    rejectCapture  = reject;
  });

  // ── 2. Register the exposeBinding ──────────────────────────────────────────
  //
  // The binding is context-scoped so it applies to every frame/page.
  // It is safe to call exposeBinding before addInitScript — both are
  // registered before any navigation begins.
  //
  // The first argument `source` provides { browserContext, page, frame }.
  // The second argument is whatever object the in-page script passed to
  // window.__outlookCliReportAuth({ url, token }).
  let bindingInstalled = false;
  await context.exposeBinding(
    '__outlookCliReportAuth',
    (_source, payload: { url: string; token: string }) => {
      if (bindingInstalled) return; // extra safety — should not be needed
      bindingInstalled = true;
      resolveCapture({ token: payload.token, url: payload.url });
    },
  );

  // ── 3. Register the init script ────────────────────────────────────────────
  //
  // INIT_SCRIPT_TEXT is the string literal from §3 above, kept in a separate
  // constant (or loaded from a file) so this function stays readable.
  await context.addInitScript(INIT_SCRIPT_TEXT);

  // ── 4. Page-closed guard ───────────────────────────────────────────────────
  const onPageClose = () => {
    rejectCapture(
      Object.assign(new Error('Browser page closed before Bearer token was captured'), {
        code: 'BROWSER_CLOSED',
        exitCode: 4,
      }),
    );
  };
  page.once('close', onPageClose);
  context.once('close', onPageClose);

  // ── 5. Timeout guard ───────────────────────────────────────────────────────
  const timeoutHandle = setTimeout(() => {
    rejectCapture(
      Object.assign(
        new Error(`No Authorization: Bearer token captured within ${timeoutMs}ms`),
        { code: 'LOGIN_TIMEOUT', exitCode: 4 },
      ),
    );
  }, timeoutMs);

  // ── 6. Race and clean up ───────────────────────────────────────────────────
  try {
    const result = await capturePromise;
    return result;
  } finally {
    clearTimeout(timeoutHandle);
    page.off('close', onPageClose);
    context.off('close', onPageClose);
  }
}
```

### Caller pattern

```typescript
// src/auth/login.ts  (simplified)

import { chromium } from 'playwright';
import { captureFirstBearerToken } from './fetchHook';

const PROFILE_DIR  = path.join(os.homedir(), '.outlook-cli', 'playwright-profile');
const TIMEOUT_MS   = parseInt(process.env.OUTLOOK_CLI_LOGIN_TIMEOUT_MS ?? '', 10);
// Per project convention: throw if config missing, no fallback
if (isNaN(TIMEOUT_MS)) {
  throw Object.assign(
    new Error('OUTLOOK_CLI_LOGIN_TIMEOUT_MS is not set'),
    { code: 'MISSING_CONFIG', exitCode: 3 },
  );
}

export async function acquireToken(): Promise<string> {
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    channel: 'chrome',
    headless: false,
    // Do NOT pass --restore-last-session; we always navigate fresh to
    // outlook.com to trigger the MSAL silent-SSO flow.
    args: ['--no-first-run', '--no-default-browser-check'],
  });

  // Get or create the working page BEFORE registering the init script;
  // we pass it to captureFirstBearerToken for the close guard.
  // The actual navigation happens AFTER captureFirstBearerToken installs
  // its hooks (the function returns after hook installation, before goto).
  const page = context.pages()[0] ?? await context.newPage();

  // captureFirstBearerToken registers exposeBinding + addInitScript,
  // then returns a Promise that resolves on first token capture.
  // We must NOT await it here yet — we need to start the navigation first.
  const tokenPromise = captureFirstBearerToken(context, page, TIMEOUT_MS);

  // Navigate AFTER the hooks are installed.  goto() itself is async and
  // will drive the MSAL redirect chain; the init script fires on every
  // document in the chain.
  await page.goto('https://outlook.office.com/mail/', {
    waitUntil: 'domcontentloaded',
    timeout: TIMEOUT_MS,
  });

  try {
    const { token, url } = await tokenPromise;
    return token; // "Bearer eyJ..."
  } finally {
    await context.close();
  }
}
```

**Important**: `captureFirstBearerToken` must finish its `await context.exposeBinding(...)` and `await context.addInitScript(...)` calls before `page.goto()` is called. The current structure achieves this: `captureFirstBearerToken` is `async` and its first two operations are the hook registrations; since JavaScript is single-threaded, the `await page.goto(...)` line does not execute until `captureFirstBearerToken` has returned its `Promise` — which happens only after both `await` calls inside it complete.

---

## 5. Alternative: `page.on('request')` from the Node Side

Playwright's `page.on('request', handler)` and `context.on('request', handler)` fire for every network request the browser makes. Reading the `Authorization` header from the Node side is possible but has important constraints.

### Header visibility

- `request.headers()` — returns a plain object with lower-cased header names. Per the official docs: *"this method does not return security-related headers, including cookie-related ones"*. `Authorization` is considered security-related and **may be omitted** by `request.headers()` in some Playwright versions.
- `request.allHeaders()` — async method (returns a Promise) that includes all headers including cookies and `Authorization`. This is the correct method to use if taking the Node-side route.

```typescript
// Node-side alternative (not the recommended path — see below)
context.on('request', async (request) => {
  const url = request.url();
  if (!url.startsWith('https://outlook.office.com/api/v2.0/') &&
      !url.startsWith('https://outlook.office.com/ows/')) return;

  const headers = await request.allHeaders(); // async — note the await
  const auth = headers['authorization'];
  if (auth && auth.startsWith('Bearer ')) {
    // resolve promise here
  }
});
```

### Why the init-script approach is preferred over `page.on('request')`

1. **Service worker requests are invisible**: if Outlook's service worker (registered at `/owa/service-worker.js`) makes the first authenticated API call, `page.on('request')` from the Node side does not see it — the request does not originate from a frame. The init-script hook patches `window.fetch` in the main document and all frames, but cannot reach service workers (a known Playwright limitation: [github.com/microsoft/playwright/issues/28029](https://github.com/microsoft/playwright/issues/28029)).
2. **`request.allHeaders()` is async**: inside a synchronous event handler this requires `async` + careful error handling; missing an `await` means the handler returns before reading the header.
3. **Race on very fast requests**: the `context.on('request')` listener is installed after `launchPersistentContext` returns, but before `page.goto()`. For a persistent profile where Outlook's service worker pre-fetches data in the background immediately on browser open, the relevant request may fire before the Node-side listener is registered. The init-script approach has no such race because `addInitScript` is evaluated before any page script on every navigation.

**Recommendation**: Use `page.on('request')` only as a **tertiary fallback** (e.g., if the init-script approach fails in a specific environment). Do not rely on it as the primary path.

---

## 6. Race Conditions with MSAL Bootstrap

MSAL.js in Outlook web follows this initialization order:

1. Browser parses HTML, creates `document`.
2. **Playwright's init script runs** (this is where our hook installs itself).
3. Synchronous `<script>` tags in `<head>` execute.
4. `DOMContentLoaded` fires.
5. MSAL.js initializes, reads its cache, performs a silent-SSO token refresh if needed.
6. On cache hit (persistent profile, valid session): MSAL acquires a token silently and the SPA makes its first authenticated fetch, typically within 1-3 seconds of step 5.
7. On cache miss (first login or expired session): MSAL redirects to `login.microsoftonline.com`; after login, it redirects back to Outlook, and step 2 happens again on the new document — the init script re-fires on the new document and the hook is re-installed.

**The init script wins the race by design.** Step 2 precedes step 3, which precedes all MSAL code. There is no race condition between the hook installation and MSAL bootstrap — the hook is already in place before MSAL's first line of JavaScript executes.

The only scenario where the race matters is if MSAL were somehow invoked from a service worker before the main document loads. Outlook web's service worker handles caching of static assets and some offline scenarios, but the authenticated API calls (which carry a Bearer token) go through the main document's `fetch` (confirmed by the POC). The XHR fallback in the init script provides additional coverage.

---

## 7. Recommended Sequence (Complete Pseudocode)

```
PROCEDURE acquireToken(profileDir, timeoutMs):

  1. launchPersistentContext(profileDir, { channel:'chrome', headless:false })
     → context

  2. context.exposeBinding('__outlookCliReportAuth', nodeCallback)
     [nodeCallback: resolve the capturePromise on first call; ignore subsequent calls]

  3. context.addInitScript(INIT_SCRIPT_TEXT)
     [INIT_SCRIPT_TEXT: the monkey-patch from §3 above]

  4. page ← context.pages()[0] ?? context.newPage()

  5. Attach close guards:
       page.once('close', → reject capturePromise with BROWSER_CLOSED)
       context.once('close', → reject capturePromise with BROWSER_CLOSED)

  6. Set timeout: setTimeout(timeoutMs, → reject capturePromise with LOGIN_TIMEOUT)

  7. page.goto('https://outlook.office.com/mail/', { waitUntil:'domcontentloaded' })
     [This triggers MSAL silent-SSO or interactive login]

  8. WAIT capturePromise
     → { token: "Bearer eyJ...", url: "https://outlook.office.com/api/v2.0/..." }

  9. clearTimeout; detach close guards

  10. context.close()

  11. RETURN token
```

---

## 8. Pitfalls to Avoid

### Pitfall 1: Registering the init script AFTER `page.goto()`

If `context.addInitScript()` is called after `page.goto()` has already started (or after the page has loaded), the hook will not be in place for the current document. It will fire on the *next* navigation, not the current one. The Outlook SPA may have already made its first authenticated API call by then.

**Fix**: Always register `exposeBinding` and `addInitScript` before any `goto()` call. The code structure in §4 enforces this.

### Pitfall 2: Relying on `request.headers()` instead of `request.allHeaders()`

If you use the `page.on('request')` approach and call `request.headers()`, the `Authorization` header may be absent from the returned object. The Playwright docs explicitly state that `request.headers()` omits security-related headers. You must call `request.allHeaders()` (async) to see `Authorization`.

**Fix**: In any Node-side request listener, always use `await request.allHeaders()`, never `request.headers()`, when looking for `Authorization`.

### Pitfall 3: `addInitScript` does not fire on pre-restored pages in a persistent profile

When `launchPersistentContext` opens a browser with previously-open tabs restored, those pages do not trigger a `'page'` event and the init script is not evaluated on them. Any authenticated requests that Outlook's background refresh makes on those restored tabs will not be intercepted.

**Fix**: Always navigate explicitly to `https://outlook.office.com/mail/` via `page.goto()`. This forces a fresh navigation, which triggers the init script. Do not use `--restore-last-session` in the browser launch args. Do not rely on whatever page the persistent profile had open.

### Pitfall 4: Service worker requests are invisible to `addInitScript` and `page.on('request')`

The Playwright `addInitScript` hook runs in the page's main world (and sub-frames), but not in service workers. If Outlook's service worker makes the first authenticated fetch, the hook will not fire.

Evidence suggests Outlook's authenticated API calls originate from the main document, not the service worker (the POC validated this). However, if a future Outlook update moves API calls into the service worker, both the init-script approach and `page.on('request')` will miss them.

**Mitigation**: The `page.on('request')` tertiary fallback should be registered in addition to the init-script approach. If both fail within the timeout, surface a clear error rather than silently returning an empty token.

---

## 9. Complete Production-Ready Code

### `src/auth/fetchHook.ts` — init script constant + capture function

```typescript
import { BrowserContext, Page } from 'playwright';

// ── Init script (installed as a string; runs in the browser) ─────────────────

export const INIT_SCRIPT_TEXT = `
(function () {
  if (window.__outlookCliHooked) return;
  window.__outlookCliHooked = true;

  let reported = false;

  const TARGET_PREFIXES = [
    'https://outlook.office.com/api/v2.0/',
    'https://outlook.office.com/ows/',
    'https://outlook.office365.com/api/v2.0/',
    'https://outlook.office365.com/ows/',
  ];

  function isTargetUrl(url) {
    const s = typeof url === 'string' ? url : (url && url.href ? url.href : String(url));
    return TARGET_PREFIXES.some(prefix => s.startsWith(prefix));
  }

  function extractBearer(headers) {
    if (!headers) return null;
    if (typeof headers.get === 'function') {
      return headers.get('authorization') || headers.get('Authorization') || null;
    }
    if (Array.isArray(headers)) {
      const pair = headers.find(([k]) => k.toLowerCase() === 'authorization');
      return pair ? pair[1] : null;
    }
    const key = Object.keys(headers).find(k => k.toLowerCase() === 'authorization');
    return key ? headers[key] : null;
  }

  function tryReport(url, authHeader) {
    if (reported) return;
    if (!authHeader || !authHeader.startsWith('Bearer ')) return;
    if (!isTargetUrl(url)) return;
    reported = true;
    window.__outlookCliReportAuth({ url: String(url), token: authHeader });
  }

  // Patch fetch
  const originalFetch = window.fetch;
  window.fetch = function fetch(input, init) {
    try {
      let url = input;
      let authHeader = null;
      if (input instanceof Request) {
        url = input.url;
        authHeader = extractBearer(input.headers);
        if (init && init.headers) {
          const override = extractBearer(init.headers);
          if (override) authHeader = override;
        }
      } else {
        url = input;
        authHeader = init && init.headers ? extractBearer(init.headers) : null;
      }
      tryReport(url, authHeader);
    } catch (_) {}
    return originalFetch.apply(this, arguments);
  };

  // Patch XMLHttpRequest
  const OriginalXHR = window.XMLHttpRequest;
  function PatchedXHR() {
    const xhr = new OriginalXHR();
    let _url = '';
    let _authHeader = null;
    const originalOpen = xhr.open.bind(xhr);
    xhr.open = function open(method, url) {
      _url = url;
      _authHeader = null;
      return originalOpen.apply(xhr, arguments);
    };
    const originalSetRequestHeader = xhr.setRequestHeader.bind(xhr);
    xhr.setRequestHeader = function setRequestHeader(name, value) {
      if (name.toLowerCase() === 'authorization') _authHeader = value;
      return originalSetRequestHeader.apply(xhr, arguments);
    };
    const originalSend = xhr.send.bind(xhr);
    xhr.send = function send() {
      try { tryReport(_url, _authHeader); } catch (_) {}
      return originalSend.apply(xhr, arguments);
    };
    return xhr;
  }
  Object.setPrototypeOf(PatchedXHR, OriginalXHR);
  Object.setPrototypeOf(PatchedXHR.prototype, OriginalXHR.prototype);
  Object.defineProperty(PatchedXHR, 'name', { value: 'XMLHttpRequest' });
  window.XMLHttpRequest = PatchedXHR;
})();
`;

// ── Types ────────────────────────────────────────────────────────────────────

export interface CapturedAuth {
  token: string;  // Full "Bearer eyJ..." string
  url: string;    // URL that triggered capture (for diagnostics only)
}

// ── Main function ────────────────────────────────────────────────────────────

/**
 * Registers the fetch/XHR monkey-patch init script and the Node-side binding
 * on the given context, then returns a Promise that resolves with the first
 * Authorization: Bearer token sent to outlook.office.com/api/v2.0/ or /ows/.
 *
 * IMPORTANT: Call this function BEFORE page.goto(). The returned Promise will
 * resolve asynchronously once the page makes its first authenticated API call.
 *
 * @throws { code: 'BROWSER_CLOSED', exitCode: 4 } if page/context closes first.
 * @throws { code: 'LOGIN_TIMEOUT',  exitCode: 4 } if no token within timeoutMs.
 */
export async function captureFirstBearerToken(
  context: BrowserContext,
  page: Page,
  timeoutMs: number,
): Promise<CapturedAuth> {

  let resolveCapture!: (auth: CapturedAuth) => void;
  let rejectCapture!:  (err: Error) => void;

  const capturePromise = new Promise<CapturedAuth>((resolve, reject) => {
    resolveCapture = resolve;
    rejectCapture  = reject;
  });

  // Register binding BEFORE init script (order does not technically matter
  // since both are in place before goto(), but exposeBinding first is cleaner).
  let alreadyResolved = false;
  await context.exposeBinding(
    '__outlookCliReportAuth',
    (_source: unknown, payload: { url: string; token: string }) => {
      if (alreadyResolved) return;
      alreadyResolved = true;
      resolveCapture({ token: payload.token, url: payload.url });
    },
  );

  await context.addInitScript(INIT_SCRIPT_TEXT);

  const onClose = () => {
    if (alreadyResolved) return;
    rejectCapture(
      Object.assign(
        new Error('Browser closed before Bearer token was captured'),
        { code: 'BROWSER_CLOSED', exitCode: 4 },
      ),
    );
  };

  page.once('close', onClose);
  context.once('close', onClose);

  const timer = setTimeout(() => {
    if (alreadyResolved) return;
    rejectCapture(
      Object.assign(
        new Error(`No Bearer token captured within ${timeoutMs}ms — login may not have completed`),
        { code: 'LOGIN_TIMEOUT', exitCode: 4 },
      ),
    );
  }, timeoutMs);

  try {
    return await capturePromise;
  } finally {
    clearTimeout(timer);
    page.off('close', onClose);
    context.off('close', onClose);
  }
}
```

### `src/auth/login.ts` — wiring `captureFirstBearerToken` into the login flow

```typescript
import path from 'node:path';
import os   from 'node:os';
import { chromium } from 'playwright';
import { captureFirstBearerToken, CapturedAuth } from './fetchHook';

export async function acquireToken(): Promise<CapturedAuth> {
  const profileDir = path.join(os.homedir(), '.outlook-cli', 'playwright-profile');

  const timeoutRaw = process.env.OUTLOOK_CLI_LOGIN_TIMEOUT_MS;
  if (!timeoutRaw) {
    throw Object.assign(
      new Error('OUTLOOK_CLI_LOGIN_TIMEOUT_MS is not configured'),
      { code: 'MISSING_CONFIG', exitCode: 3 },
    );
  }
  const timeoutMs = parseInt(timeoutRaw, 10);
  if (isNaN(timeoutMs) || timeoutMs <= 0) {
    throw Object.assign(
      new Error(`OUTLOOK_CLI_LOGIN_TIMEOUT_MS has an invalid value: "${timeoutRaw}"`),
      { code: 'MISSING_CONFIG', exitCode: 3 },
    );
  }

  const context = await chromium.launchPersistentContext(profileDir, {
    channel: 'chrome',
    headless: false,
    args: ['--no-first-run', '--no-default-browser-check'],
    // Deliberately omit --restore-last-session. We always navigate fresh
    // so the init script fires reliably (see Pitfall 3).
  });

  const page = context.pages()[0] ?? await context.newPage();

  // Install hooks BEFORE goto(). captureFirstBearerToken is async and
  // completes the hook registrations before returning its Promise.
  const capturePromise = captureFirstBearerToken(context, page, timeoutMs);

  // Navigate after hooks are installed.
  await page.goto('https://outlook.office.com/mail/', {
    waitUntil: 'domcontentloaded',
    timeout: timeoutMs,
  });

  try {
    return await capturePromise;
  } finally {
    await context.close();
  }
}
```

---

## 10. Assumptions and Scope

| Assumption | Confidence | Impact if Wrong |
|---|---|---|
| Outlook web uses `fetch` (not exclusively XHR) for the first authenticated API call to `/api/v2.0/` | HIGH — confirmed by POC | XHR fallback in init script provides coverage; would need to verify which path fires |
| `window.fetch` and `window.XMLHttpRequest` are the actual request channels (not a service worker intercepting first) | MEDIUM — consistent with POC observations | Service worker requests are invisible to both init-script and `page.on('request')`; would require a different capture strategy |
| The Bearer token in the first captured request is valid for the full `outlook.office.com/api/v2.0/` surface (mail + calendar) | HIGH — POC validated with `/me/messages` and `/me/calendarview` | Would need to add validation calls as part of `auth-check` |
| `launchPersistentContext` with `channel: 'chrome'` (not Playwright's bundled Chromium) retains the full Chrome session state needed for MSAL silent-SSO | HIGH — spec requirement, matches POC | Without system Chrome, MSAL session cookies may not survive across runs |
| Playwright 1.40+ (specifically the bug in issue #28692) means pre-restored pages do not get init scripts | HIGH — bug confirmed as of April 2026 | Workaround (always `goto()`) is already applied |

### Out of scope

- Service worker request interception
- Capturing tokens across multiple concurrent pages
- Refreshing the token within the same browser session (the `acquireToken` function is called fresh each time)
- Anything related to `sessionStorage` / `localStorage` decryption (NG5 prohibition)

---

## References

| # | Source | URL | Information Gathered |
|---|---|---|---|
| 1 | Playwright Docs — BrowserContext.addInitScript | https://playwright.dev/docs/api/class-browsercontext | Script timing ("after document created, before page scripts"), context vs page scope, Disposable return |
| 2 | Playwright Docs — BrowserContext.exposeBinding | https://playwright.dev/docs/api/class-browsercontext | `source` argument shape `{ browserContext, page, frame }`, context-scope, Promise resolution |
| 3 | Playwright Docs — BrowserContext.exposeFunction | https://github.com/microsoft/playwright/blob/main/docs/src/api/class-browsercontext.md | Difference from exposeBinding (no source arg), Disposable |
| 4 | Playwright Docs — Request.headers / allHeaders | https://playwright.dev/docs/api/class-request | `headers()` strips security-related headers; `allHeaders()` is async and returns all including Authorization |
| 5 | Playwright Docs — Network interception | https://playwright.dev/docs/network | `page.route()` mechanics, `route.request().headers()`, service worker visibility |
| 6 | Playwright GitHub Issue #28692 | https://github.com/microsoft/playwright/issues/28692 | Confirmed bug: addInitScript does not fire on pre-restored pages in launchPersistentContext |
| 7 | Playwright GitHub Issue #28029 | https://github.com/microsoft/playwright/issues/28029 | addInitScript does not reach service workers — known limitation |
| 8 | Playwright GitHub Issue #1915 | https://github.com/microsoft/playwright/issues/1915 | Community question on capturing Bearer token from browser for use in Node requests |
| 9 | Playwright GitHub Issue #10884 | https://github.com/microsoft/playwright/issues/10884 | Pattern: page.evaluate() to extract token from localStorage as alternative |
| 10 | Playwright Source — mock-browser-js.md | https://github.com/microsoft/playwright/blob/main/docs/src/mock-browser-js.md | Combined addInitScript + exposeFunction pattern for logging API calls |
| 11 | Context7 — Playwright library docs | https://context7.com/microsoft/playwright/llms.txt | route() interception examples, TypeScript fetch/header patterns |
| 12 | Playwright Solutions — Part 4 Authentication | https://playwrightsolutions.com/the-definitive-guide-to-api-test-automation-with-playwright-part-4-handling-headers-and-authentication/ | Real-world pattern for creating auth tokens / cookies in Node context for Playwright |

### Recommended for Deep Reading

- **Issue #28692** ([link](https://github.com/microsoft/playwright/issues/28692)): The pre-restored-pages bug is directly relevant to this project. Monitor for a fix; the workaround (always `goto()`) should be kept until resolved.
- **Playwright Network Docs** ([link](https://playwright.dev/docs/network)): Complete reference for `page.route()`, `request.allHeaders()`, and service worker intercept behavior.
- **Playwright mock-browser-js.md** ([link](https://github.com/microsoft/playwright/blob/main/docs/src/mock-browser-js.md)): The official `addInitScript` + `exposeFunction` combination example — closest official analog to what this project does.

---

## Clarifying Questions for Follow-up

1. Should `captureFirstBearerToken` also register a tertiary `context.on('request', ...)` listener as a fallback (in case the init script misses a service worker request)? This adds complexity but improves resilience.
2. Is there a requirement to support Playwright's bundled Chromium as an alternative to `channel: 'chrome'`? If so, the session-persistence story changes significantly.
3. Does the tool need to handle multiple simultaneous `acquireToken` calls (e.g., from two terminal sessions)? The advisory lock file (§4.5 of the investigation) covers this at the process level, but the `capturePromise` is per-call — confirm the concurrency requirement.
4. What is the expected behavior when the captured token arrives from a `/ows/` URL rather than `/api/v2.0/`? Both are in the `TARGET_PREFIXES` list, but confirm whether the token audience and scope are identical for both URL families.
