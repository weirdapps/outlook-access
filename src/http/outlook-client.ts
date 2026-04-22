/**
 * Outlook REST v2 HTTP client.
 *
 * Normative contract: project-design.md §2.8.
 * Header construction rules: refined-request-outlook-cli.md §6.2.
 * 401 retry semantics: refined-request-outlook-cli.md §6.4.
 *
 * Responsibilities:
 *   - Build request URL, headers, and cookie header from the current session.
 *   - Apply per-request timeout via AbortSignal.
 *   - On 401, trigger a single automatic re-auth + retry (unless disabled).
 *   - Map non-2xx responses to typed errors from ./errors.
 *   - NEVER leak bearer tokens or cookies into error messages or logs.
 */

import type { Cookie, SessionFile } from '../session/schema';
import { UpstreamError } from '../config/errors';
import {
  MAX_FOLDER_PAGES,
  MAX_FOLDERS_VISITED,
  DEFAULT_LIST_TOP,
} from '../folders/types';
import {
  ApiError,
  AuthError,
  CollisionError,
  NetworkError,
  codeForStatus,
  isFolderExistsError,
  truncateAndRedactBody,
} from './errors';
import type {
  FolderCreateRequest,
  FolderSummary,
  MessageSummary,
  MoveMessageRequest,
  ODataListResponse,
} from './types';

// ---------------------------------------------------------------------------
// Public interface
// ---------------------------------------------------------------------------

/** Value types accepted in the `query` parameter bag. */
export type QueryValue = string | number;

/** Options for `listMessagesInFolder`. Mirrors the list-mail select/order API. */
export interface ListMessagesInFolderOptions {
  /** `$top` — 1..1000; when omitted the server default applies. */
  top?: number;
  /** `$select` field names. Serialised as a comma-joined OData `$select`. */
  select?: string[];
  /** `$orderby` clause, e.g. `'ReceivedDateTime desc'`. */
  orderBy?: string;
  /** `$filter` clause, e.g. `'ReceivedDateTime ge 2026-04-22T07:00:00Z'`. */
  filter?: string;
}

/** Options for `countMessagesInFolder`. */
export interface CountMessagesInFolderOptions {
  /** Raw OData `$filter` expression, passed through verbatim. */
  filter?: string;
}

/** Result shape for `countMessagesInFolder`. */
export interface CountMessagesResult {
  /** The message count. */
  count: number;
  /**
   * `true` when the server returned `@odata.count` (authoritative total
   * across all pages); `false` when it did not, in which case `count`
   * reflects only the first page and may underestimate the real total.
   */
  exact: boolean;
}

/** Options for `listMessagesByConversation`. */
export interface ListMessagesByConversationOptions {
  /** `$top` — 1..1000; when omitted the server default applies. */
  top?: number;
  /** `$select` field names. Serialised as a comma-joined OData `$select`. */
  select?: string[];
  /** `$orderby` clause. Default: `'ReceivedDateTime asc'` (oldest-first thread). */
  orderBy?: string;
}

/** Result envelope for the auto-paginating `listMessagesInFolderAll`. */
export interface ListMessagesInFolderAllResult {
  /** Collected message summaries (capped at `maxResults`). */
  messages: MessageSummary[];
  /** True when the cap was hit and more results were available. */
  truncated: boolean;
}

export interface OutlookClient {
  /**
   * GET a JSON resource. Returns the parsed body typed as T.
   *
   * @param path  Path starting with '/', e.g. '/api/v2.0/me/messages'.
   * @param query Optional query parameters. OData `$` keys are passed through
   *              verbatim; all values are URL-encoded.
   */
  get<T>(path: string, query?: Record<string, QueryValue>): Promise<T>;

  /**
   * List direct children of a mail folder via
   * `GET /api/v2.0/me/MailFolders/{parentId}/childfolders`. Pagination is
   * enforced by the internal `listAll<T>` generator: at most
   * `MAX_FOLDER_PAGES` (50) pages and `MAX_FOLDERS_VISITED` (5000) total
   * items before `UpstreamError{code:'UPSTREAM_PAGINATION_LIMIT'}` is thrown.
   *
   * `parentId` may be a raw opaque folder id or a well-known alias
   * (`Inbox`, `MsgFolderRoot`, …).
   *
   * @param parentId Parent folder id or well-known alias.
   * @param top      `$top` hint for the first page (default: DEFAULT_LIST_TOP).
   */
  listFolders(parentId: string, top?: number): Promise<FolderSummary[]>;

  /**
   * Fetch a single folder via `GET /api/v2.0/me/MailFolders/{idOrAlias}`.
   * `idOrAlias` is passed verbatim into the URL path segment (well-known
   * aliases are accepted per v2.0 contract).
   *
   * A 404 response is reclassified into
   * `UpstreamError{code:'UPSTREAM_FOLDER_NOT_FOUND'}` (exit 5).
   */
  getFolder(idOrAlias: string): Promise<FolderSummary>;

  /**
   * Create a folder under `parentId`. Uses
   * `POST /api/v2.0/me/MailFolders/{parentId}/childfolders` for any concrete
   * parent; when `parentId` is the synthetic `msgfolderroot` sentinel (or
   * the PascalCase `MsgFolderRoot` alias), creation targets the mailbox
   * root via `POST /api/v2.0/me/MailFolders` directly (§10.5, ADR-15).
   *
   * On HTTP 400 OR 409 whose parsed body matches `isFolderExistsError`, the
   * error is reclassified into
   * `CollisionError{code:'FOLDER_ALREADY_EXISTS'}` (exit 6). The caller is
   * responsible for any `--idempotent` recovery on top.
   */
  createFolder(parentId: string, displayName: string): Promise<FolderSummary>;

  /**
   * Move a message via `POST /api/v2.0/me/messages/{messageId}/move` with
   * body `{ DestinationId: destinationFolderId }`.
   *
   * The caller is responsible for pre-resolving any alias in
   * `destinationFolderId` to a raw id (ADR-16); this method sends the
   * value verbatim. The upstream response contains the moved message with
   * a NEW id — it is returned as-is.
   */
  moveMessage(
    messageId: string,
    destinationFolderId: string,
  ): Promise<MessageSummary>;

  /**
   * List messages across multiple pages by following `@odata.nextLink`.
   * Returns up to `maxResults` messages and a `truncated` flag indicating
   * whether more were available beyond the cap. Defense-in-depth host check
   * mirrors `listAll` (folders): nextLinks must stay on outlook.office.com.
   *
   * @param folderId    Folder id (raw or well-known alias).
   * @param opts        Same options as `listMessagesInFolder` (top is page size).
   * @param maxResults  Hard cap on total messages returned; protects against
   *                    runaway pagination on huge folders.
   */
  listMessagesInFolderAll(
    folderId: string,
    opts: ListMessagesInFolderOptions,
    maxResults: number,
  ): Promise<ListMessagesInFolderAllResult>;

  /**
   * List messages inside `folderId` via
   * `GET /api/v2.0/me/MailFolders/{folderId}/messages`, using the same
   * `$select` / `$orderby` / `$top` options as the existing list-mail code
   * path. `folderId` may be a raw id or a well-known alias.
   */
  listMessagesInFolder(
    folderId: string,
    opts: ListMessagesInFolderOptions,
  ): Promise<MessageSummary[]>;

  /**
   * Server-side count of messages matching `filter` inside `folderId`, via
   * `$count=true` on the same `/MailFolders/{folderId}/messages` endpoint.
   * Uses `$top=1&$select=Id` to minimize payload — the messages themselves
   * are discarded. Returns `{ count, exact }`; `exact: false` signals the
   * server did not return `@odata.count` and the count may be partial.
   */
  countMessagesInFolder(
    folderId: string,
    opts?: CountMessagesInFolderOptions,
  ): Promise<CountMessagesResult>;

  /**
   * List every message in a conversation (thread) regardless of folder, via
   * `GET /api/v2.0/me/messages?$filter=ConversationId eq '{id}'`. The caller
   * is responsible for providing the conversation id (usually extracted from
   * any single message's `ConversationId` field).
   */
  listMessagesByConversation(
    conversationId: string,
    opts?: ListMessagesByConversationOptions,
  ): Promise<MessageSummary[]>;
}

export interface CreateClientOptions {
  /** The active session. The client keeps a mutable reference and updates
   *  it after a successful re-auth. */
  session: SessionFile;
  /** Mandatory; from CliConfig.httpTimeoutMs. Applied per request. */
  httpTimeoutMs: number;
  /** Called exactly once on HTTP 401 before retrying. Must return a new
   *  SessionFile that the client then adopts for all subsequent calls. */
  onReauthNeeded: () => Promise<SessionFile>;
  /** When true, 401 throws AuthError immediately (no re-auth, no retry). */
  noAutoReauth: boolean;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const BASE_URL = 'https://outlook.office.com';
const ALLOWED_HOST = 'outlook.office.com';
const COOKIE_HOST_SUFFIXES = ['outlook.office.com', '.outlook.office.com'];
/**
 * Synthetic sentinel values (case-insensitive) that instruct `createFolder`
 * to target the mailbox root via `POST /me/MailFolders` instead of
 * `POST /me/MailFolders/{parent}/childfolders`. Both the Graph-style lowercase
 * `msgfolderroot` and the PascalCase v2.0 alias `MsgFolderRoot` are accepted
 * (research doc §2a / project-design §10.5).
 */
const MSG_FOLDER_ROOT_SENTINELS = new Set(['msgfolderroot']);

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

export function createOutlookClient(opts: CreateClientOptions): OutlookClient {
  if (!opts || !opts.session) {
    throw new Error('createOutlookClient: session is required');
  }
  if (typeof opts.httpTimeoutMs !== 'number' || opts.httpTimeoutMs <= 0) {
    throw new Error('createOutlookClient: httpTimeoutMs must be a positive number');
  }
  if (typeof opts.onReauthNeeded !== 'function') {
    throw new Error('createOutlookClient: onReauthNeeded is required');
  }

  // Mutable holder so a re-auth inside one call is visible to the next call.
  let session: SessionFile = opts.session;

  /**
   * Method-agnostic 401-retry-once envelope shared by GET, POST, and every
   * page fetched inside `listAll`. Accepts either a relative path (starting
   * with '/') or a fully-absolute `https://outlook.office.com/...` URL (used
   * by `listAll` for `@odata.nextLink` follow-through).
   */
  async function doRequest<T>(
    method: 'GET' | 'POST',
    urlOrPath: string,
    body?: unknown,
  ): Promise<T> {
    const url = urlOrPath.startsWith('http')
      ? urlOrPath
      : (() => {
          if (!urlOrPath.startsWith('/')) {
            throw new Error(
              `outlook-client: path must start with '/': ${urlOrPath}`,
            );
          }
          return `${BASE_URL}${urlOrPath}`;
        })();

    const firstResp = await executeFetch(
      method,
      url,
      body,
      session,
      opts.httpTimeoutMs,
    );

    if (firstResp.status === 401) {
      if (opts.noAutoReauth) {
        // Per design §2.8: noAutoReauth + 401 → AUTH_NO_REAUTH.
        await throwForResponse(firstResp, url, /*authReason*/ 'NO_AUTO_REAUTH');
      }
      // Drain the body so the underlying socket is released before we
      // potentially block on the user-driven re-auth flow.
      await safeDrainBody(firstResp);

      const refreshed = await opts.onReauthNeeded();
      session = refreshed;

      const retryResp = await executeFetch(
        method,
        url,
        body,
        session,
        opts.httpTimeoutMs,
      );
      if (retryResp.status === 401) {
        await throwForResponse(retryResp, url, /*authReason*/ 'AFTER_RETRY');
      }
      return await handleSuccessOrThrow<T>(retryResp, url);
    }

    return await handleSuccessOrThrow<T>(firstResp, url);
  }

  async function doGet<T>(
    path: string,
    query?: Record<string, QueryValue>,
  ): Promise<T> {
    if (!path.startsWith('/')) {
      throw new Error(`outlook-client: path must start with '/': ${path}`);
    }
    const url = buildUrl(path, query);
    return doRequest<T>('GET', url);
  }

  /**
   * Private `post<TBody, TRes>` helper. Mirrors the private GET helper's
   * contract — same 401-retry-once envelope, same error mapping. The JSON
   * body is serialised by `executeFetch`.
   */
  async function doPost<TBody, TRes>(
    path: string,
    body: TBody,
  ): Promise<TRes> {
    if (!path.startsWith('/')) {
      throw new Error(`outlook-client: path must start with '/': ${path}`);
    }
    const url = buildUrl(path, undefined);
    return doRequest<TRes>('POST', url, body);
  }

  /**
   * Private generic `listAll<T>` — follows `@odata.nextLink` verbatim up to
   * `MAX_FOLDER_PAGES` pages. Yields individual items as they are decoded.
   *
   * Enforces two safety rails:
   *   1. Off-host guard: any `@odata.nextLink` whose hostname is not
   *      `outlook.office.com` raises
   *      `UpstreamError{code:'UPSTREAM_PAGINATION_LIMIT'}`.
   *   2. Page cap: more than `MAX_FOLDER_PAGES` pages raises the same
   *      `UpstreamError{code:'UPSTREAM_PAGINATION_LIMIT'}`.
   *
   * Each page's GET rides the shared `doRequest` envelope, so a 401 on page
   * N is transparently retried after a single re-auth.
   */
  async function* listAll<T>(
    path: string,
    query?: Record<string, string>,
  ): AsyncGenerator<T> {
    if (!path.startsWith('/')) {
      throw new Error(`outlook-client: path must start with '/': ${path}`);
    }

    // First page: build the URL from path + query + $top default (callers may
    // override $top via query).
    const mergedQuery: Record<string, QueryValue> = {
      $top: String(DEFAULT_LIST_TOP),
      ...(query ?? {}),
    };
    let url: string | null = buildUrl(path, mergedQuery);
    let pageCount = 0;

    while (url !== null) {
      if (pageCount >= MAX_FOLDER_PAGES) {
        throw new UpstreamError({
          code: 'UPSTREAM_PAGINATION_LIMIT',
          message:
            `Exceeded ${MAX_FOLDER_PAGES}-page cap while paginating ${path}. ` +
            `Narrow the scope (e.g. --parent) or raise the cap.`,
        });
      }

      // Defense-in-depth host validation — the nextLink must stay on
      // outlook.office.com. The first-page URL is already controlled via
      // BASE_URL, but we re-check anyway for consistency.
      let parsed: URL;
      try {
        parsed = new URL(url);
      } catch (cause) {
        throw new UpstreamError({
          code: 'UPSTREAM_PAGINATION_LIMIT',
          message: `Malformed @odata.nextLink: ${String(url)}`,
          cause,
        });
      }
      if (parsed.hostname !== ALLOWED_HOST) {
        throw new UpstreamError({
          code: 'UPSTREAM_PAGINATION_LIMIT',
          message:
            `@odata.nextLink host '${parsed.hostname}' is not '${ALLOWED_HOST}'.`,
        });
      }

      const page: ODataListResponse<T> = await doRequest<ODataListResponse<T>>(
        'GET',
        url,
      );
      const items = Array.isArray(page.value) ? page.value : [];
      for (const item of items) {
        yield item;
      }
      pageCount++;

      // Follow the nextLink verbatim (it is absolute and already encodes
      // $skip / $top / $select); do NOT reconstruct the URL.
      url = page['@odata.nextLink'] ?? null;
    }
  }

  // -------------------------------------------------------------------------
  // Public semantic methods (folder feature — §10.4)
  // -------------------------------------------------------------------------

  async function listFolders(
    parentId: string,
    top?: number,
  ): Promise<FolderSummary[]> {
    if (typeof parentId !== 'string' || parentId.length === 0) {
      throw new Error('outlook-client: listFolders requires a non-empty parentId');
    }
    const query: Record<string, string> = {};
    if (typeof top === 'number' && Number.isFinite(top) && top > 0) {
      query.$top = String(Math.floor(top));
    }

    const path =
      `/api/v2.0/me/MailFolders/${encodeURIComponent(parentId)}/childfolders`;

    const collected: FolderSummary[] = [];
    try {
      for await (const item of listAll<FolderSummary>(path, query)) {
        if (collected.length >= MAX_FOLDERS_VISITED) {
          throw new UpstreamError({
            code: 'UPSTREAM_PAGINATION_LIMIT',
            message:
              `Exceeded ${MAX_FOLDERS_VISITED}-item cap while listing folders ` +
              `under parent '${parentId}'. Narrow the scope or raise the cap.`,
          });
        }
        collected.push(item);
      }
    } catch (err) {
      // UpstreamError already carries the CLI-layer shape — rethrow as-is.
      if (err instanceof UpstreamError) throw err;
      throw mapHttpToCliError(err);
    }
    return collected;
  }

  async function getFolder(idOrAlias: string): Promise<FolderSummary> {
    if (typeof idOrAlias !== 'string' || idOrAlias.length === 0) {
      throw new Error('outlook-client: getFolder requires a non-empty idOrAlias');
    }
    const path = `/api/v2.0/me/MailFolders/${encodeURIComponent(idOrAlias)}`;
    try {
      return await doGet<FolderSummary>(path);
    } catch (err) {
      // 404 gets reclassified to UPSTREAM_FOLDER_NOT_FOUND per §10.6.
      if (err instanceof ApiError && err.httpStatus === 404) {
        throw new UpstreamError({
          code: 'UPSTREAM_FOLDER_NOT_FOUND',
          message: `Folder '${idOrAlias}' was not found (404).`,
          httpStatus: 404,
          requestId: err.requestId,
          url: err.url,
          cause: err,
        });
      }
      throw mapHttpToCliError(err);
    }
  }

  async function createFolder(
    parentId: string,
    displayName: string,
  ): Promise<FolderSummary> {
    if (typeof parentId !== 'string' || parentId.length === 0) {
      throw new Error('outlook-client: createFolder requires a non-empty parentId');
    }
    if (typeof displayName !== 'string' || displayName.length === 0) {
      throw new Error('outlook-client: createFolder requires a non-empty displayName');
    }

    // Mailbox-root creation targets `POST /me/MailFolders` directly; every
    // other parent targets `POST /me/MailFolders/{parentId}/childfolders`.
    const isRoot = MSG_FOLDER_ROOT_SENTINELS.has(parentId.toLowerCase());
    const path = isRoot
      ? `/api/v2.0/me/MailFolders`
      : `/api/v2.0/me/MailFolders/${encodeURIComponent(parentId)}/childfolders`;
    const body: FolderCreateRequest = { DisplayName: displayName };

    try {
      return await doPost<FolderCreateRequest, FolderSummary>(path, body);
    } catch (err) {
      // 400/409 + OData error.code === 'ErrorFolderExists' → CollisionError.
      if (
        err instanceof ApiError &&
        (err.httpStatus === 400 || err.httpStatus === 409) &&
        isFolderExistsError(parseErrorBody(err))
      ) {
        throw new CollisionError({
          code: 'FOLDER_ALREADY_EXISTS',
          message: `A folder named '${displayName}' already exists under parent '${parentId}'.`,
          path: displayName,
          parentId,
          cause: err,
        });
      }
      throw mapHttpToCliError(err);
    }
  }

  async function moveMessage(
    messageId: string,
    destinationFolderId: string,
  ): Promise<MessageSummary> {
    if (typeof messageId !== 'string' || messageId.length === 0) {
      throw new Error('outlook-client: moveMessage requires a non-empty messageId');
    }
    if (typeof destinationFolderId !== 'string' || destinationFolderId.length === 0) {
      throw new Error(
        'outlook-client: moveMessage requires a non-empty destinationFolderId',
      );
    }
    const path =
      `/api/v2.0/me/messages/${encodeURIComponent(messageId)}/move`;
    const body: MoveMessageRequest = { DestinationId: destinationFolderId };
    try {
      return await doPost<MoveMessageRequest, MessageSummary>(path, body);
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }

  function buildMessagesQuery(opts: ListMessagesInFolderOptions): Record<string, QueryValue> {
    const query: Record<string, QueryValue> = {};
    if (typeof opts.top === 'number' && Number.isFinite(opts.top) && opts.top > 0) {
      query.$top = String(Math.floor(opts.top));
    }
    if (Array.isArray(opts.select) && opts.select.length > 0) {
      query.$select = opts.select.join(',');
    }
    if (typeof opts.orderBy === 'string' && opts.orderBy.length > 0) {
      query.$orderby = opts.orderBy;
    }
    if (typeof opts.filter === 'string' && opts.filter.length > 0) {
      query.$filter = opts.filter;
    }
    return query;
  }

  async function listMessagesInFolder(
    folderId: string,
    opts: ListMessagesInFolderOptions,
  ): Promise<MessageSummary[]> {
    if (typeof folderId !== 'string' || folderId.length === 0) {
      throw new Error(
        'outlook-client: listMessagesInFolder requires a non-empty folderId',
      );
    }
    const query = buildMessagesQuery(opts);
    const path =
      `/api/v2.0/me/MailFolders/${encodeURIComponent(folderId)}/messages`;
    try {
      const resp = await doGet<ODataListResponse<MessageSummary>>(path, query);
      return Array.isArray(resp.value) ? resp.value : [];
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }

  async function listMessagesInFolderAll(
    folderId: string,
    opts: ListMessagesInFolderOptions,
    maxResults: number,
  ): Promise<ListMessagesInFolderAllResult> {
    if (typeof folderId !== 'string' || folderId.length === 0) {
      throw new Error(
        'outlook-client: listMessagesInFolderAll requires a non-empty folderId',
      );
    }
    if (!Number.isInteger(maxResults) || maxResults < 1) {
      throw new Error(
        `outlook-client: maxResults must be a positive integer (got ${String(maxResults)})`,
      );
    }
    const query = buildMessagesQuery(opts);
    const path =
      `/api/v2.0/me/MailFolders/${encodeURIComponent(folderId)}/messages`;

    const messages: MessageSummary[] = [];
    let url: string | null = buildUrl(path, query);
    let truncated = false;

    try {
      while (url !== null && messages.length < maxResults) {
        // Defense-in-depth: nextLink must stay on outlook.office.com.
        let parsed: URL;
        try {
          parsed = new URL(url);
        } catch (cause) {
          throw new UpstreamError({
            code: 'UPSTREAM_PAGINATION_LIMIT',
            message: `Malformed @odata.nextLink: ${String(url)}`,
            cause,
          });
        }
        if (parsed.hostname !== ALLOWED_HOST) {
          throw new UpstreamError({
            code: 'UPSTREAM_PAGINATION_LIMIT',
            message:
              `@odata.nextLink host '${parsed.hostname}' is not '${ALLOWED_HOST}'.`,
          });
        }

        const page: ODataListResponse<MessageSummary> =
          await doRequest<ODataListResponse<MessageSummary>>('GET', url);
        const items = Array.isArray(page.value) ? page.value : [];
        const remaining = maxResults - messages.length;
        if (items.length > remaining) {
          messages.push(...items.slice(0, remaining));
          truncated = page['@odata.nextLink'] !== undefined || items.length > remaining;
          url = null;
          break;
        }
        messages.push(...items);
        url = page['@odata.nextLink'] ?? null;
      }
      // If we exited because maxResults was reached AND there's still a nextLink,
      // mark truncated.
      if (url !== null && messages.length >= maxResults) {
        truncated = true;
      }
    } catch (err) {
      throw mapHttpToCliError(err);
    }

    return { messages, truncated };
  }

  async function countMessagesInFolder(
    folderId: string,
    opts: CountMessagesInFolderOptions = {},
  ): Promise<CountMessagesResult> {
    if (typeof folderId !== 'string' || folderId.length === 0) {
      throw new Error(
        'outlook-client: countMessagesInFolder requires a non-empty folderId',
      );
    }
    const query: Record<string, QueryValue> = {
      $count: 'true',
      $top: '1',
      $select: 'Id',
    };
    if (typeof opts.filter === 'string' && opts.filter.length > 0) {
      query.$filter = opts.filter;
    }
    const path =
      `/api/v2.0/me/MailFolders/${encodeURIComponent(folderId)}/messages`;
    try {
      const resp = await doGet<ODataListResponse<MessageSummary>>(path, query);
      const serverCount = resp['@odata.count'];
      if (typeof serverCount === 'number' && Number.isFinite(serverCount)) {
        return { count: serverCount, exact: true };
      }
      return {
        count: Array.isArray(resp.value) ? resp.value.length : 0,
        exact: false,
      };
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }

  async function listMessagesByConversation(
    conversationId: string,
    opts: ListMessagesByConversationOptions = {},
  ): Promise<MessageSummary[]> {
    if (typeof conversationId !== 'string' || conversationId.length === 0) {
      throw new Error(
        'outlook-client: listMessagesByConversation requires a non-empty conversationId',
      );
    }
    const escaped = conversationId.replace(/'/g, "''");
    const query: Record<string, QueryValue> = {
      $filter: `ConversationId eq '${escaped}'`,
      $orderby:
        typeof opts.orderBy === 'string' && opts.orderBy.length > 0
          ? opts.orderBy
          : 'ReceivedDateTime asc',
    };
    if (typeof opts.top === 'number' && Number.isFinite(opts.top) && opts.top > 0) {
      query.$top = String(Math.floor(opts.top));
    }
    if (Array.isArray(opts.select) && opts.select.length > 0) {
      query.$select = opts.select.join(',');
    }
    try {
      const resp = await doGet<ODataListResponse<MessageSummary>>(
        '/api/v2.0/me/messages',
        query,
      );
      return Array.isArray(resp.value) ? resp.value : [];
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }

  return {
    get: doGet,
    listFolders,
    getFolder,
    createFolder,
    moveMessage,
    listMessagesInFolder,
    listMessagesInFolderAll,
    countMessagesInFolder,
    listMessagesByConversation,
  };
}

// ---------------------------------------------------------------------------
// Private helpers (module-scope, used by the factory)
// ---------------------------------------------------------------------------

/**
 * Best-effort JSON-parse of a stringified ApiError body. The existing
 * `throwForResponse` embeds the body as a raw string snippet into
 * `ApiError.message`; the OData `error.code` is therefore not readily
 * available on the thrown error. This helper attempts to recover the parsed
 * object by looking for a `{"error":{"code":...}}` JSON fragment in the
 * message string. It is intentionally conservative: any failure returns
 * `undefined`, which causes `isFolderExistsError` to return false.
 */
function parseErrorBody(err: ApiError): unknown {
  const text = err.message;
  if (typeof text !== 'string' || text.length === 0) return undefined;
  // Locate the first '{' and try to JSON-parse the remainder, tolerating a
  // trailing '...' truncation added by truncateAndRedactBody.
  const idx = text.indexOf('{');
  if (idx < 0) return undefined;
  let candidate = text.slice(idx);
  // Strip a possible '...' truncation suffix.
  if (candidate.endsWith('...')) {
    candidate = candidate.slice(0, -3);
  }
  // Try parse progressively by trimming trailing chars until JSON parses or
  // the candidate becomes too short to be worthwhile.
  for (let end = candidate.length; end > 0; end--) {
    try {
      return JSON.parse(candidate.slice(0, end));
    } catch {
      /* keep trying */
    }
  }
  return undefined;
}

/**
 * Translate HTTP-layer errors into the CLI's error taxonomy. Mirrors the
 * `mapHttpError` helper in `src/commands/list-mail.ts` so the new semantic
 * methods can emit the same CLI-layer shapes directly.
 */
function mapHttpToCliError(err: unknown): unknown {
  if (err instanceof UpstreamError || err instanceof CollisionError) {
    return err;
  }
  if (err instanceof AuthError) {
    // AuthError (HTTP-layer) is carried through to the command layer's
    // `mapHttpError`, which converts it into the CLI-layer AuthError. We
    // re-throw it unchanged so the existing wrapper in commands keeps
    // working.
    return err;
  }
  if (err instanceof ApiError) {
    return new UpstreamError({
      code: `UPSTREAM_HTTP_${err.httpStatus}`,
      message: err.message,
      httpStatus: err.httpStatus,
      requestId: err.requestId,
      url: err.url,
      cause: err,
    });
  }
  if (err instanceof NetworkError) {
    return new UpstreamError({
      code: err.timedOut ? 'UPSTREAM_TIMEOUT' : 'UPSTREAM_NETWORK',
      message: err.message,
      url: err.url,
      cause: err,
    });
  }
  return err;
}

// ---------------------------------------------------------------------------
// URL construction
// ---------------------------------------------------------------------------

function buildUrl(
  path: string,
  query: Record<string, QueryValue> | undefined,
): string {
  const base = `${BASE_URL}${path}`;
  if (!query) return base;

  const params = new URLSearchParams();
  for (const [k, v] of Object.entries(query)) {
    if (v === undefined || v === null) continue;
    params.append(k, String(v));
  }
  const qs = params.toString();
  return qs.length > 0 ? `${base}?${qs}` : base;
}

// ---------------------------------------------------------------------------
// Header / cookie construction
// ---------------------------------------------------------------------------

function buildHeaders(
  s: SessionFile,
  method: 'GET' | 'POST',
): Record<string, string> {
  const rawToken = s.bearer.token ?? '';
  const authValue = rawToken.startsWith('Bearer ')
    ? rawToken
    : `Bearer ${rawToken}`;

  const headers: Record<string, string> = {
    Authorization: authValue,
    'X-AnchorMailbox': s.anchorMailbox,
    Accept: 'application/json',
  };

  // Only body-bearing methods set Content-Type.
  if (method === 'POST') {
    headers['Content-Type'] = 'application/json';
  }

  const cookieHeader = serializeCookieJar(s.cookies ?? []);
  if (cookieHeader.length > 0) {
    headers.Cookie = cookieHeader;
  }

  return headers;
}

/**
 * Build a Cookie request header from the session jar.
 *
 * Rules:
 *   - Only include cookies whose `domain` matches outlook.office.com via
 *     RFC 6265 suffix rules. Cookies scoped to login.microsoftonline.com or
 *     other Microsoft hosts are NOT sent to outlook.office.com.
 *   - httpOnly and secure cookies ARE included (requests are HTTPS).
 *   - Values are joined "name=value" with "; " separators.
 *   - Values are NOT URL-encoded — the session stores them exactly as the
 *     browser received them.
 */
export function serializeCookieJar(jar: readonly Cookie[]): string {
  const parts: string[] = [];
  for (const c of jar) {
    if (!cookieDomainMatches(c.domain)) continue;
    if (!c.name) continue;
    parts.push(`${c.name}=${c.value}`);
  }
  return parts.join('; ');
}

function cookieDomainMatches(domain: string): boolean {
  if (!domain) return false;
  // Normalize the stored domain; Playwright stores a leading '.' for domain
  // cookies and the bare host for host-only cookies.
  const d = domain.toLowerCase();
  for (const suffix of COOKIE_HOST_SUFFIXES) {
    if (d === suffix) return true;
    // `.outlook.office.com` stored as '.outlook.office.com' — match either
    // `.outlook.office.com` or `outlook.office.com` literally.
    if (d === suffix.replace(/^\./, '')) return true;
  }
  // Also accept any sub-domain of outlook.office.com (sub.outlook.office.com).
  return d.endsWith('.outlook.office.com');
}

// ---------------------------------------------------------------------------
// Request execution
// ---------------------------------------------------------------------------

async function executeFetch(
  method: 'GET' | 'POST',
  url: string,
  body: unknown,
  s: SessionFile,
  timeoutMs: number,
): Promise<Response> {
  const headers = buildHeaders(s, method);

  // Serialise the body for POST (and any future body-bearing methods). GET
  // never carries a body.
  const init: RequestInit = {
    method,
    headers,
    signal: AbortSignal.timeout(timeoutMs),
  };
  if (method === 'POST' && body !== undefined) {
    init.body = JSON.stringify(body);
  }

  // Native fetch with per-request abort timeout.
  try {
    const response = await fetch(url, init);
    return response;
  } catch (cause: unknown) {
    throw mapFetchException(cause, url, timeoutMs);
  }
}

function mapFetchException(
  cause: unknown,
  url: string,
  timeoutMs: number,
): NetworkError {
  // AbortError from AbortSignal.timeout → timed-out request.
  if (isAbortLike(cause)) {
    return new NetworkError({
      message: `HTTP timeout after ${timeoutMs}ms`,
      url,
      cause,
      timedOut: true,
    });
  }
  // TypeError from fetch → DNS / TLS / connection failure.
  const detail =
    cause && typeof cause === 'object' && 'message' in cause
      ? String((cause as { message: unknown }).message ?? '')
      : String(cause ?? 'unknown');
  return new NetworkError({
    message: `Network error: ${detail}`,
    url,
    cause,
    timedOut: false,
  });
}

function isAbortLike(cause: unknown): boolean {
  if (!cause || typeof cause !== 'object') return false;
  const name = (cause as { name?: unknown }).name;
  return name === 'AbortError' || name === 'TimeoutError';
}

// ---------------------------------------------------------------------------
// Response handling
// ---------------------------------------------------------------------------

async function handleSuccessOrThrow<T>(
  response: Response,
  url: string,
): Promise<T> {
  if (response.ok) {
    // Treat a 204 or empty body as null. Callers that care use typed
    // command-level wrappers; this client just returns what JSON.parse gives.
    const text = await response.text();
    if (text.length === 0) {
      // Empty bodies have no schema; `null` is the least-surprising sentinel.
      return null as unknown as T;
    }
    try {
      return JSON.parse(text) as T;
    } catch (cause) {
      throw new ApiError({
        code: 'INVALID_JSON',
        message: `Upstream returned non-JSON body (status ${response.status})`,
        httpStatus: response.status,
        url,
        requestId: getRequestId(response),
      });
    }
  }

  await throwForResponse(response, url, /*authReason*/ 'NONE');
  // throwForResponse always throws; this line is unreachable but keeps
  // TypeScript's control-flow happy.
  throw new Error('unreachable');
}

/**
 * Consume the response body and throw the appropriate typed error. This is
 * only invoked on non-2xx paths.
 *
 * For 401, the caller passes the appropriate `authReason`:
 *   - 'NO_AUTO_REAUTH' → --no-auto-reauth was set and we saw the first 401.
 *   - 'AFTER_RETRY'    → the single automatic retry also failed.
 * When status is not 401, `authReason` is ignored.
 */
async function throwForResponse(
  response: Response,
  url: string,
  authReason: 'NO_AUTO_REAUTH' | 'AFTER_RETRY' | 'NONE',
): Promise<never> {
  const bodyText = await safeReadText(response);
  const requestId = getRequestId(response);
  const status = response.status;
  const snippet = truncateAndRedactBody(bodyText);

  if (status === 401) {
    const reason = authReason === 'NO_AUTO_REAUTH' ? 'NO_AUTO_REAUTH' : 'AFTER_RETRY';
    const message =
      reason === 'AFTER_RETRY'
        ? `Outlook rejected credentials (401) after re-auth retry.${snippet ? ` ${snippet}` : ''}`
        : `Session is missing or expired and --no-auto-reauth was set (401).${snippet ? ` ${snippet}` : ''}`;
    throw new AuthError({
      message,
      url,
      httpStatus: status,
      requestId,
      reason,
    });
  }

  // 429 → include Retry-After for the caller to surface.
  if (status === 429) {
    const retryAfter = response.headers.get('retry-after') ?? undefined;
    const ra = retryAfter ? ` Retry-After: ${retryAfter}s.` : '';
    throw new ApiError({
      code: 'RATE_LIMITED',
      message: `Outlook rate limited the request (429).${ra}${snippet ? ` ${snippet}` : ''}`,
      httpStatus: status,
      url,
      requestId,
    });
  }

  // Everything else is a flat ApiError; the `code` field carries the
  // semantic distinction.
  throw new ApiError({
    code: codeForStatus(status),
    message: `Outlook request failed with HTTP ${status}.${snippet ? ` ${snippet}` : ''}`,
    httpStatus: status,
    url,
    requestId,
  });
}

function getRequestId(response: Response): string | undefined {
  // Outlook emits `request-id`; Graph-bridge sometimes emits
  // `x-ms-request-id`. Check both.
  const headers = response.headers;
  return (
    headers.get('request-id') ?? headers.get('x-ms-request-id') ?? undefined
  );
}

async function safeReadText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

async function safeDrainBody(response: Response): Promise<void> {
  // `response.text()` consumes the stream; if it fails we don't care — the
  // GC will release the underlying socket.
  try {
    await response.text();
  } catch {
    /* ignore */
  }
}
