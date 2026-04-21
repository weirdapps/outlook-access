# Plan 002 — Outlook CLI Folder Management

Plan date: 2026-04-21
Inputs consumed (in priority order):

1. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/refined-request-folders.md`
2. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/investigation-folders.md`
3. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/research/outlook-v2-folder-pagination-filter.md`
4. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/research/outlook-v2-move-destination-alias.md`
5. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/research/outlook-v2-folder-duplicate-error.md`
6. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/reference/codebase-scan-folders.md`
7. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/project-design.md`
8. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/plan-001-outlook-cli.md` (structural template)
9. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/CLAUDE.md`

Plan 002 is **strictly additive** on top of the shipped Plan 001 codebase. Phase A
of Plan 001 (scaffolding) and Phases B-H (config, session, auth, HTTP, seven
read-only commands, CLI wiring, ACs) are assumed complete and stable. This plan
does not revisit them.

---

## 0. Open Questions (require user decisions BEFORE Phase P1 begins)

The following four items are blockers unless the defaults are explicitly
confirmed. Each maps to a refined-request open question (§13) or an
investigation risk that must be pinned before code ships.

1. **OQ-1: Collision class choice.** Should `FOLDER_ALREADY_EXISTS` be routed
   through a **new** `CollisionError extends OutlookCliError (exitCode = 6)` or
   through the existing `IoError` class with a folder-specific code?
   **Default picked: new `CollisionError` class** (investigation §4.3,
   refined §13 Q1). Matches the "cause is not filesystem IO" argument. Sets
   the shape of Phase P3 (Errors) and its consumers.

2. **OQ-2: `--first-match` tiebreaker.** When ambiguity is resolved, do we
   order candidates by `CreatedDateTime asc` then `Id asc`, or by
   `DisplayName asc` then `Id asc` (refined §13 Q2)?
   **Default picked: `CreatedDateTime asc, Id asc`.** Requires the resolver
   to include `CreatedDateTime` in the default `$select`.

3. **OQ-3: Default parent for `create-folder` / `list-folders`.** Confirm
   `MsgFolderRoot` (mailbox root — sibling of Inbox) rather than `Inbox`
   (refined §13 Q3). **Default picked: `MsgFolderRoot`** — matches Outlook
   web "New folder" UX at the root.

4. **OQ-4: Move-destination alias pass-through.** Research
   (`outlook-v2-move-destination-alias.md`) concludes that v2.0
   `DestinationId` alias acceptance is **uncertain**; recommendation is
   **resolve-first for every alias** (one extra `GET` per alias). Is a single
   extra `GET` on every `move-mail --to <alias>` invocation acceptable?
   **Default picked: yes — always resolve alias → raw id before
   `POST /move`** (matches the investigation recommendation and risk §2).
   A future `--raw-alias` opt-in is out of scope for this iteration.

If the user does NOT override these defaults within the usual review window,
the plan proceeds as written.

---

## 1. Goals & Non-goals

### Goals (from `refined-request-folders.md §2`)

- **G1.** Four new subcommands (`list-folders`, `find-folder`, `create-folder`,
  `move-mail`) plus `--folder-id` / `--folder-parent` extensions on `list-mail`.
- **G2.** A single **canonical folder resolver** (`src/folders/resolver.ts`)
  owning every piece of path / alias / NFC / case-fold / ambiguity / well-known
  precedence logic.
- **G3.** JSON default + `--table` for every new command.
- **G4.** Deterministic JSON shapes documented per subcommand in refined §5.
- **G5.** Reuse existing `OutlookClient`, auth, 401-retry-once, error classes.
  No new exit-code values.
- **G6.** Update `CLAUDE.md` `<outlook-cli>` block with per-subcommand entries.
- **G7.** Update `project-design.md` + `project-functions.MD` + `Issues - Pending
  Items.md`.

### Non-goals (from `refined-request-folders.md §3`, verbatim)

- NG1 rename, NG2 delete, NG3 copy, NG4 `$batch`, NG5 delta/sync,
  NG6 search-folder creation, NG7 move-folder parent change, NG8 IsHidden
  mutation, NG9 shared/archive-mailbox access, NG10 concurrent moves.

---

## 2. Overall Architecture Delta

Added to the system diagram (cf. `project-design.md §1`). New components
shown with `NEW`. Changed components shown with `CHG`.

```
                        ┌──────────────────────────────────────────┐
                        │              user / shell                │
                        └──────────────────┬───────────────────────┘
                                           │
                                           ▼
                        ┌──────────────────────────────────────────┐
                        │        src/cli.ts  (bin entry)           │  CHG
                        │  + 4 new subcommands                     │
                        │  + 2 new flags on list-mail              │
                        │  + CollisionError branch in             │
                        │    formatErrorJson / exitCodeFor         │
                        └──────────────────┬───────────────────────┘
                                           │
              ┌────────────────────────────┼──────────────────────────┐
              ▼                            ▼                          ▼
   ┌────────────────────┐   ┌───────────────────────────┐  ┌─────────────────────┐
   │  src/config/       │   │  src/commands/            │  │  src/output/        │
   │  errors.ts   CHG   │   │   list-folders.ts   NEW   │  │  formatter.ts       │
   │  (+ CollisionError)│   │   find-folder.ts    NEW   │  │  (unchanged; new    │
   │  (+ new code       │   │   create-folder.ts  NEW   │  │   ColumnSpecs in    │
   │   strings on       │   │   move-mail.ts      NEW   │  │   cli.ts)           │
   │   UsageError /     │   │   list-mail.ts      CHG   │  └─────────────────────┘
   │   UpstreamError)   │   │   (+ --folder-id /        │
   └──────────┬─────────┘   │    --folder-parent /      │
              │             │    widened --folder)      │
              │             └───────────┬───────────────┘
              │                         │
              │                         ▼
              │          ┌───────────────────────────────┐
              │          │  src/folders/        NEW       │
              │          │    resolver.ts                 │
              │          │    types.ts                    │
              │          │  (parseFolderPath,             │
              │          │   buildFolderPath,             │
              │          │   matchesWellKnownAlias,       │
              │          │   listChildren,                │
              │          │   resolveFolder,               │
              │          │   createFolderPath,            │
              │          │   isFolderExistsError)         │
              │          └──────────┬─────────────────────┘
              │                     │
              ▼                     ▼
   ┌────────────────────┐  ┌──────────────────────────┐
   │ src/session/       │  │ src/http/                │  CHG
   │ (unchanged)        │  │   outlook-client.ts       │
   │                    │  │   + post<TBody,TRes>(...) │
   │                    │  │   + listAll<T>(...)       │
   │                    │  │   (refactor doGet         │
   │                    │  │    → doRequest)           │
   │                    │  │   errors.ts (unchanged)   │
   │                    │  │   types.ts   CHG          │
   │                    │  │   (+ FolderSummary,       │
   │                    │  │     FolderCreateRequest)  │
   └────────────────────┘  └──────────────────────────┘
```

Runtime dataflow for a folder-aware command:

```
cli.ts → command → ensureSession → createClient → resolver.resolveFolder
       └→ (walks) client.listAll<FolderSummary>(/me/MailFolders/{id}/childfolders)
       ↓
client.get / client.post  ← REST v2.0 call
       ↓
command returns typed result → cli.ts emitResult → formatter (JSON | table)
```

---

## 3. Dependency Graph of Implementation Phases

```
                P1 (types)  P2 (errors)
                     │          │
                     └────┬─────┘
                          ▼
                     P3 (HTTP client: post + listAll)
                          │
                          ▼
                     P4 (resolver + folder-path utils)
                          │
          ┌────────┬──────┼───────┬──────────┐
          ▼        ▼      ▼       ▼          ▼
          P5a     P5b    P5c     P5d         P5e
      list-folders find  create  move-mail   list-mail
                  folder folder  (move)      extension
          └────────┴──────┴───────┴──────────┘
                          │
                          ▼
                     P6 (cli.ts wiring + ColumnSpecs + CollisionError branch)
                          │
                          ▼
                     P7 (docs: CLAUDE.md, project-design, project-functions,
                         Issues - Pending Items)
                          │
                          ▼
                     P8 (tests: unit + AC + smoke)
```

- **P1 + P2** independent — can start together.
- **P3** needs P1 (uses `ODataListResponse<T>` from `src/http/types.ts`) and P2
  (raises `ApiError` for pagination cap; folder-specific codes route through
  P2's extended vocabulary).
- **P4** needs P3 (`client.listAll`, `client.post`).
- **P5a-e** can all run in parallel after P4 — each command file is a leaf.
  `list-mail` (P5e) touches `list-mail.ts` exclusively; the other four each
  create a new file.
- **P6** is the single integration point that touches `cli.ts`. Strictly
  sequential after P5a-e (all five must export their `run()` + option shapes).
- **P7** is documentation — can run in parallel with P8 once P6 lands.
- **P8** is tests — gated on P6 (needs working `node dist/cli.js <verb>`).

---

## 4. Phase-by-Phase Breakdown

Each phase lists: Goal, Files created / modified (with symbols), Dependencies,
Parallelization, Acceptance-criteria coverage, and Verification steps.

### Phase P1 — Types (wire + CLI shapes)

- **Goal.** Declare every new TypeScript type used by the rest of the phases.
  No behaviour.
- **Files modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/http/types.ts`
    - **New exports**: `FolderSummary` (REST wire shape, PascalCase),
      `FolderCreateRequest` (`{ DisplayName: string }`), `MoveMessageRequest`
      (`{ DestinationId: string }`).
    - No changes to existing exports.
- **Files created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/folders/types.ts`
    - **Exports**:
      - `FolderSpec` (discriminated: `{ kind: 'wellknown' | 'path' | 'id',
        value: string, parent?: FolderSpec }`).
      - `ResolvedFolder` (`FolderSummary & { Path: string; ResolvedVia:
        'wellknown'|'path'|'id' }`).
      - `CreateFolderResult` (`{ created: CreateSegment[]; leaf:
        CreatedSegment; idempotent: boolean }` + nested segment shape).
      - `MoveMailResult` (`{ destination: MoveDestination; moved: MoveEntry[];
        failed: FailedEntry[]; summary: { requested, moved, failed } }`).
      - `WELL_KNOWN_ALIASES: readonly string[]` (frozen list from refined
        §6.2, PascalCase matching v2.0 URL conventions).
      - `MAX_PATH_SEGMENTS = 16`, `MAX_FOLDER_PAGES = 50`,
        `MAX_FOLDERS_VISITED = 5000`, `DEFAULT_LIST_TOP = 250`,
        `DEFAULT_LIST_FOLDERS_TOP = 100`.
- **Dependencies.** none (fresh types). Runs first.
- **Parallel with.** P2.
- **Acceptance-criteria covered.** none directly; gating for AC-*.
- **Verification.**
  - `npx tsc --noEmit` passes.
  - `grep -n "FolderSummary" src/http/types.ts` returns one export.
  - `grep -n "ResolvedFolder" src/folders/types.ts` returns one export.

---

### Phase P2 — Errors (CollisionError + extended code vocabularies)

- **Goal.** Introduce the one new error class justified by OQ-1 and extend
  the `code` string vocabularies on `UsageError` / `UpstreamError`. No CLI
  wiring yet (that is P6).
- **Files modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/config/errors.ts`
    - **New class**: `CollisionError extends OutlookCliError { exitCode = 6;
      code: string; path?: string; parentId?: string; }` — mirrors `IoError`
      structure but with its own `instanceof` discriminator.
    - **No changes** to `ConfigurationError`, `AuthError`, `UpstreamError`,
      `IoError` class bodies. `UsageError` / `UpstreamError` `code` fields
      already accept free strings — new codes are just documented in the
      block comment above each class:
      - `UsageError.code`: add `FOLDER_AMBIGUOUS`, `FOLDER_MISSING_PARENT`,
        `FOLDER_PATH_INVALID`.
      - `UpstreamError.code`: add `UPSTREAM_FOLDER_NOT_FOUND`,
        `UPSTREAM_FOLDER_AMBIGUOUS`, `UPSTREAM_PAGINATION_LIMIT`.
- **Dependencies.** none.
- **Parallel with.** P1.
- **Acceptance-criteria covered.** AC-CREATE-COLLISION (the class exists),
  AC-CREATE-MISSING-PARENT (code string defined), AC-FOLDER-AMBIGUOUS,
  AC-FOLDER-NOT-FOUND, AC-PATH-DEPTH-CAP. Full end-to-end coverage in P5 / P8.
- **Verification.**
  - `npx tsc --noEmit` passes.
  - Unit: `new CollisionError('FOLDER_ALREADY_EXISTS', 'A/B', 'parentId')`
    has `.exitCode === 6`.

---

### Phase P3 — HTTP client extensions (`post` + `listAll`)

- **Goal.** Hoist the existing `doGet` 401-retry-once envelope into a shared
  `doRequest`, add a public `post<TBody, TRes>`, and add `listAll<T>` with
  the 50-page safety cap.
- **Files modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/http/outlook-client.ts`
    - **Interface change** (`OutlookClient`):
      - **Add**: `post<TBody, TRes>(path: string, body: TBody, query?:
        Record<string, QueryValue>): Promise<TRes>`.
      - **Add**: `listAll<T>(path: string, query?: Record<string, QueryValue>,
        opts?: { maxPages?: number; top?: number }): Promise<T[]>`.
      - `get<T>(...)` unchanged.
    - **Implementation change**:
      - Refactor private `doGet` → `doRequest(method, path, body?, query?)` —
        hoist the 401-retry-once envelope (today at lines 93-110) into the
        method-agnostic shape. `buildUrl`, `buildHeaders`, `executeFetch`,
        `handleSuccessOrThrow`, `throwForResponse`, `mapFetchException`
        **unchanged**.
      - `buildHeaders`: emit `Content-Type: application/json` only when
        `method === 'POST' | 'PATCH'`. Keep `Accept: application/json`,
        `Authorization`, `X-AnchorMailbox`, `Cookie` untouched.
      - `listAll<T>` per the snippet in `outlook-v2-folder-pagination-filter.md
        §5` — first `GET` uses caller query + `$top` default
        (`DEFAULT_LIST_TOP = 250`), subsequent pages follow
        `@odata.nextLink` **verbatim** (no query re-merge).
      - Off-host guard: reject any `@odata.nextLink` whose hostname is not
        `outlook.office.com` — raise
        `ApiError('PAGINATION_OFF_HOST', ...)`.
      - Page-cap guard: on `pageCount >= maxPages` (default 50) raise
        `ApiError('PAGINATION_LIMIT', ...)`.
  - **No edits** to `src/http/errors.ts` (the `codeForStatus` 409 → `CONFLICT`
    mapping already covers the duplicate-folder path).
- **Dependencies.** P1 (uses `ODataListResponse<T>` from `src/http/types.ts`).
- **Parallel with.** — (P4 and P5 need this).
- **Acceptance-criteria covered.** AC-401-RETRY-FOLDERS (the shared envelope
  automatically covers POST + listAll), basis for every folder AC.
- **Verification.**
  - `npx tsc --noEmit` passes.
  - `npm test -- outlook-client` passes (new tests: `post` issues POST with
    JSON body; `listAll` follows two pages then stops; `listAll` caps at 50
    pages; off-host nextLink rejected; 401 retry covers `post` and `listAll`).

---

### Phase P4 — Folder resolver module

- **Goal.** One module owning every piece of path / alias / NFC /
  case-fold / ambiguity / well-known precedence / collision-error
  classification. Zero duplication across commands.
- **Files created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/folders/resolver.ts`
    - **Exports**:
      - `parseFolderPath(input: string): string[]` — split on `/`; unescape
        `\/` → `/`, `\\` → `\`; reject empty segments; reject paths > 16
        segments (raises `UsageError('FOLDER_PATH_INVALID', ...)`);
        NFC-normalize every segment.
      - `buildFolderPath(segments: string[]): string` — inverse. Escapes
        `/` and `\` inside segments.
      - `matchesWellKnownAlias(input: string): string | null` — exact match
        against `WELL_KNOWN_ALIASES` (PascalCase). Returns canonical form
        or null.
      - `listChildren(client, parentId, opts): Promise<FolderSummary[]>` —
        thin wrapper over `client.listAll<FolderSummary>` for
        `/me/MailFolders/{parentId}/childfolders`. Default `$select`:
        `Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,
        TotalItemCount,WellKnownName,CreatedDateTime,IsHidden`. Honors
        `includeHidden` via `includeHiddenFolders=true` query param.
      - `resolveFolder(client, spec: FolderSpec, opts: { caseSensitive?,
        includeHidden?, firstMatch? }): Promise<ResolvedFolder>` — the
        path-walk workhorse. Contract:
        - `spec.kind === 'id'` → one `GET /me/MailFolders/{id}`; map 404 →
          `UpstreamError('UPSTREAM_FOLDER_NOT_FOUND', ...)`.
        - `spec.kind === 'wellknown'` → no REST call; constructed
          `ResolvedFolder` with `Id = value`, `DisplayName = value`,
          `ResolvedVia = 'wellknown'`. Reason: Outlook accepts the alias in
          subsequent URL paths verbatim.
        - `spec.kind === 'path'` → walk segment-by-segment with
          `listChildren`, match client-side using NFC + simple case-fold
          (case-sensitive if `caseSensitive`). Ambiguity rules per
          refined §6.4 (raises `UsageError('FOLDER_AMBIGUOUS', ...)`
          unless `firstMatch` — then sort by `CreatedDateTime asc, Id asc`
          per OQ-2).
      - `createFolderPath(client, { anchorId, segments, createParents,
        idempotent }): Promise<CreateFolderResult>` — for each segment:
        (1) lookup under current parent;
        (2) if present → advance with `PreExisting: true`;
        (3) if not and not leaf and not `createParents` → raise
            `UsageError('FOLDER_MISSING_PARENT', ...)`;
        (4) else `client.post<FolderCreateRequest, FolderSummary>`;
        (5) on `ApiError` where `isFolderExistsError(err)` is true:
            - if `idempotent` → re-list children, locate by DisplayName,
              advance with `PreExisting: true`.
            - else raise `CollisionError('FOLDER_ALREADY_EXISTS',
              segmentPath, parentId)`.
      - `isFolderExistsError(err): boolean` — `instanceof ApiError` +
        (status 400 OR 409) + `err.body?.error?.code === 'ErrorFolderExists'`
        (per `outlook-v2-folder-duplicate-error.md §4.1`).
- **Dependencies.** P1 (types), P2 (errors), P3 (client.post, client.listAll).
- **Parallel with.** — (P5a-e all need this).
- **Acceptance-criteria covered.** Building block for every folder AC.
- **Verification.**
  - `npx tsc --noEmit` passes.
  - `npm test -- resolver` passes (unit tests with mocked client:
    `parseFolderPath` escape rules, depth cap; `matchesWellKnownAlias`
    precedence; `resolveFolder` wellknown shortcut / path walk / ambiguity /
    firstMatch ordering; `createFolderPath` idempotent 400+code /
    idempotent 409+code / non-idempotent collision → `CollisionError`;
    missing-parent without `createParents` → `UsageError`).

---

### Phase P5 — Command modules (five commands, four parallelizable)

All five sub-phases start concurrently after P4 completes. Each touches its
own file — zero file overlap (see §5 Parallel-safety matrix).

#### Phase P5a — `list-folders.ts` (new)

- **Goal.** Enumerate top-level folders, or children of any parent; recursive
  mode materializes `Path`.
- **File created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/list-folders.ts`
    - **Exports**:
      - `ListFoldersDeps` / `ListFoldersOptions` / `run(deps, opts):
        Promise<FolderSummary[]>`.
      - Options: `parent?: string` (well-known / path / `id:...`, default
        `MsgFolderRoot`), `recursive: boolean`, `includeHidden: boolean`,
        `top?: number` (default 100, range 1..250; validated — raises
        `UsageError('BAD_USAGE', ...)` outside range).
    - **Shape**: canonical per `codebase-scan-folders.md §2.1`:
      `ensureSession` → `createClient` → resolve parent via
      `resolveFolder(...)` (or short-circuit on `MsgFolderRoot` → no
      REST call; path is `/me/MailFolders`) → `client.listAll` →
      if `recursive`, recurse DFS while materializing `Path`, bounded by
      `MAX_FOLDERS_VISITED = 5000` (exceeded → raises
      `UpstreamError('UPSTREAM_PAGINATION_LIMIT', ...)`).
      Wraps every `client.listAll` call in `try { ... } catch (err) {
      throw mapHttpError(err); }`.
- **Dependencies.** P1-P4.
- **Parallel with.** P5b, P5c, P5d, P5e.
- **Acceptance-criteria covered.** AC-LISTFOLDERS-ROOT, AC-LISTFOLDERS-CHILDREN.
- **Verification.** `npx tsc --noEmit` passes; unit test with mocked client
  asserting recursive fan-out and `Path` construction.

#### Phase P5b — `find-folder.ts` (new)

- **Goal.** Resolve a query (well-known / path / `id:...`) to a single
  `ResolvedFolder`, surfacing ambiguity and not-found correctly.
- **File created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/find-folder.ts`
    - Exports `FindFolderDeps` / `FindFolderOptions` / `run(deps, query,
      opts): Promise<ResolvedFolder>`.
    - Options: `parent?: string`, `caseSensitive: boolean`, `includeHidden:
      boolean`, `firstMatch: boolean`.
    - Positional `<query>` — missing → `UsageError('BAD_USAGE', ...)`.
    - `id:` prefix detection → `FolderSpec { kind: 'id' }`; else alias →
      `kind: 'wellknown'`; else path → `parseFolderPath` + `kind: 'path'`.
    - Single call to `resolver.resolveFolder`.
- **Dependencies.** P1-P4.
- **Parallel with.** P5a, P5c, P5d, P5e.
- **Acceptance-criteria covered.** AC-FIND-WELLKNOWN, AC-FIND-PATH, AC-FIND-ID,
  AC-FOLDER-NOT-FOUND, AC-FOLDER-AMBIGUOUS, AC-PATH-ESCAPE,
  AC-WELLKNOWN-PRECEDENCE, AC-PATH-DEPTH-CAP.
- **Verification.** `npx tsc --noEmit` passes; unit tests cover the three
  `FolderSpec.kind` branches + ambiguity + `firstMatch` tiebreaker.

#### Phase P5c — `create-folder.ts` (new)

- **Goal.** Create a folder path (optionally nested) under a parent, with
  idempotent-on-collision semantics.
- **File created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/create-folder.ts`
    - Exports `CreateFolderDeps` / `CreateFolderOptions` / `run(deps, path,
      opts): Promise<CreateFolderResult>`.
    - Options: `parent?: string`, `createParents: boolean`, `idempotent:
      boolean`, `displayName?: string` (override the last segment's
      DisplayName).
    - Positional `<path>` — missing → `UsageError('BAD_USAGE', ...)`.
    - Flow: resolve parent → `parseFolderPath` → `resolver.createFolderPath`.
    - Reject well-known aliases as top-level DisplayNames per refined §5.3
      ("cannot create Inbox at root"): raise `UsageError('BAD_USAGE', ...)`
      if the last segment normalizes to a well-known alias AND the resolved
      anchor is `MsgFolderRoot`.
- **Dependencies.** P1-P4.
- **Parallel with.** P5a, P5b, P5d, P5e.
- **Acceptance-criteria covered.** AC-CREATE-TOPLEVEL, AC-CREATE-NESTED,
  AC-CREATE-IDEMPOTENT, AC-CREATE-COLLISION, AC-CREATE-MISSING-PARENT.
- **Verification.** `npx tsc --noEmit` passes; unit tests with mocked client
  (`POST` returns 201 / `POST` returns 400+ErrorFolderExists /
  `POST` returns 409+ErrorFolderExists / `POST` returns 400 with a
  different code — only the first two + third route to `CollisionError`
  when `--idempotent=false` and to `PreExisting=true` when `--idempotent=true`).

#### Phase P5d — `move-mail.ts` (new)

- **Goal.** Move one or more messages to a destination folder, always
  pre-resolving the destination to a raw id (OQ-4).
- **File created**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/move-mail.ts`
    - Exports `MoveMailDeps` / `MoveMailOptions` / `run(deps, id?, opts):
      Promise<MoveMailResult>`.
    - Options: `to?: string` (well-known or path), `toId?: string` (raw id),
      `toParent?: string`, `idsFrom?: string` (path, or `-` for stdin),
      `continueOnError: boolean`, `stopAt: number` (default 1000, range
      1..10000), `firstMatch: boolean`.
    - Validation:
      - `<id>` positional XOR `--ids-from` — both / neither →
        `UsageError('BAD_USAGE', ...)`.
      - `--to` XOR `--to-id` — both / neither →
        `UsageError('BAD_USAGE', ...)`.
      - `--stop-at` out of range → `UsageError('BAD_USAGE', ...)`.
      - When `--ids-from` yields > `--stop-at` ids → exit 2
        `BAD_USAGE` per AC-MOVE-STOPAT.
    - Flow:
      1. Resolve destination id once (pre-move): if `--to-id` → use as-is;
         else `resolver.resolveFolder` and use `.Id`. Aliases still go
         through the resolver (short-circuits to the alias string for
         `GET /MailFolders/{alias}` → returns raw `Id`), per OQ-4.
      2. For each source id (one or many): `client.post<MoveMessageRequest,
         { Id: string }>('/me/messages/{srcId}/move', { DestinationId:
         destId })`. On success push `{ sourceId, newId }` to `moved[]`.
         On `ApiError`:
         - if `--continue-on-error` → push to `failed[]` + continue.
         - else throw `mapHttpError(err)`.
      3. Assemble `MoveMailResult` with `summary: { requested, moved,
         failed }`.
    - Exit-code semantics per refined §5.4: partial failure under
      `--continue-on-error` still surfaces exit 5 (done in P6 by
      re-throwing the last `UpstreamError` after emission — the payload
      still contains `moved[]`/`failed[]`). Pattern mirrored from
      `download-attachments`.
- **Dependencies.** P1-P4.
- **Parallel with.** P5a, P5b, P5c, P5e.
- **Acceptance-criteria covered.** AC-MOVE-SINGLE, AC-MOVE-MANY,
  AC-MOVE-BAD-DEST, AC-MOVE-BAD-SOURCE, AC-MOVE-PARTIAL, AC-MOVE-STOPAT.
- **Verification.** `npx tsc --noEmit` passes; unit tests with mocked
  client: single move happy path; batch move 3 items (1 fails) with
  `continueOnError` asserts result shape + throws at end; `--stop-at` cap
  raises `UsageError`; mutual-exclusion validations.

#### Phase P5e — `list-mail.ts` extension

- **Goal.** Accept `--folder-id` and widen `--folder` to accept paths and
  any well-known alias, while preserving the five-alias fast path.
- **File modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/list-mail.ts`
    - **ListMailOptions change**: add `folderId?: string`, `folderParent?:
      string`.
    - **`ALLOWED_FOLDERS`** (line 37) — keep as-is; re-interpret as
      "fast-path alias list" (no resolver hop). Extend the per-`--folder`
      validation (currently at line 73) to:
      - If `opts.folderId` set AND `opts.folder` set → `UsageError`.
      - If `opts.folderId` set → path = `/me/MailFolders/{folderId}/messages`.
      - Else if `opts.folder` is in `ALLOWED_FOLDERS` → existing fast path,
        no change.
      - Else if `opts.folder` matches any other `WELL_KNOWN_ALIASES` entry
        (`JunkEmail`, `Outbox`, `MsgFolderRoot`, `RecoverableItemsDeletions`)
        → fast path with that alias.
      - Else → `resolver.resolveFolder(FolderSpec { kind: 'path', value,
        parent: opts.folderParent })` → use `.Id`.
    - **Everything else unchanged** (`$orderby`, `$select`, `--top`, table
      output).
    - **`ensureSession` / `mapHttpError` / `UsageError`** — unchanged, still
      re-exported for sibling commands.
- **Dependencies.** P1-P4.
- **Parallel with.** P5a, P5b, P5c, P5d.
- **Acceptance-criteria covered.** AC-LISTMAIL-PATH, AC-LISTMAIL-ID,
  AC-LISTMAIL-WELLKNOWN-BACKCOMPAT.
- **Verification.** `npx tsc --noEmit` passes; existing `list-mail` tests
  keep passing; new tests cover `--folder-id`, path-based `--folder`,
  mutual-exclusion.

---

### Phase P6 — CLI wiring (`cli.ts` integration)

- **Goal.** Register the four new subcommands, add the two new flags on
  `list-mail`, add `ColumnSpec` constants, wire `CollisionError` into
  `formatErrorJson` / `exitCodeFor`.
- **Files modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/cli.ts`
    - **Additions** (next to existing `LIST_MAIL_COLUMNS` at line 199):
      - `LIST_FOLDERS_COLUMNS`: `Path | Unread | Total | Children | Id`
        (no `maxWidth` on `Id`).
      - `CREATE_FOLDER_COLUMNS` (applied to `result.created`):
        `Path | Id | PreExisting`.
      - `MOVE_MAIL_COLUMNS` (applied to `result.moved` concatenated with
        `result.failed` after a status-mapping step):
        `Source Id | New Id | Status | Error`.
      - `FIND_FOLDER_COLUMNS`: **none** — `find-folder` returns a single
        object; `emitResult` falls back to JSON when no column spec is
        passed (acceptable per `codebase-scan-folders.md §7`).
    - **Subcommand registrations** (mirror the existing `list-mail` block at
      lines 486-507):
      - `.command('list-folders')` → options `--parent`, `--recursive`,
        `--include-hidden`, `--top` → calls `list-folders.run`, passes
        `LIST_FOLDERS_COLUMNS`.
      - `.command('find-folder <query>')` → options `--parent`,
        `--case-sensitive`, `--include-hidden`, `--first-match` → calls
        `find-folder.run`, no columns.
      - `.command('create-folder <path>')` → options `--parent`,
        `--create-parents`, `--idempotent`, `--display-name` → calls
        `create-folder.run`, passes `CREATE_FOLDER_COLUMNS` applied to
        `result.created`.
      - `.command('move-mail [id]')` → options `--to`, `--to-id`,
        `--to-parent`, `--ids-from`, `--continue-on-error`, `--stop-at`,
        `--first-match` → calls `move-mail.run`, passes `MOVE_MAIL_COLUMNS`.
    - **`list-mail` block**: add `--folder-id <id>`, `--folder-parent
      <name-or-path>` options to the existing registration.
    - **`formatErrorJson`** (line 297) — add `if (err instanceof
      CollisionError) return { error: { code: err.code, path: err.path,
      parentId: err.parentId, message: err.message } };`.
    - **`exitCodeFor`** (line 359) — add `if (err instanceof CollisionError)
      return 6;` before the `OutlookCliError` fallback.
    - **Partial-move exit-5 handling**: `move-mail.run` returns a
      `MoveMailResult` and, when `--continue-on-error` observed failures,
      sets `result.__partialFailure = true` (non-exported sentinel) — the
      action wrapper inspects it after `emitResult` and re-throws a
      synthetic `UpstreamError('UPSTREAM_PARTIAL_MOVE', ...)` routed to
      exit 5. Alternative cleaner design: `move-mail.run` itself throws
      an `UpstreamError` **after** `emitResult` was already called — but
      since commands must not touch stdout (§Patterns to preserve #2),
      the cli.ts wrapper owns the ordering.
  - **No new files.**
- **Dependencies.** P1, P2, P3, P4, P5 (all sub-phases).
- **Parallel with.** — (strictly sequential).
- **Acceptance-criteria covered.** Every folder AC end-to-end (CLI entry
  points become exercisable).
- **Verification.**
  - `npm run build` produces `dist/cli.js`.
  - `node dist/cli.js --help` lists all four new subcommands.
  - `node dist/cli.js list-folders --help` shows all four options.
  - `node dist/cli.js list-mail --help` shows `--folder-id` and
    `--folder-parent`.
  - Smoke: `node dist/cli.js find-folder DoesNotExist` against a live
    session exits 5 with `code: UPSTREAM_FOLDER_NOT_FOUND`.

---

### Phase P7 — Documentation updates

- **Goal.** Register the new features in every project-wide doc per the
  project conventions.
- **Files modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/CLAUDE.md`
    - Add four new child blocks inside `<outlook-cli>`: `<list-folders>`,
      `<find-folder>`, `<create-folder>`, `<move-mail>`.
    - Update the `<list-mail>` block description to mention the new
      `--folder-id` / `--folder-parent` flags and path-based `--folder`.
    - Append the new exit-code / code-vocabulary rows to the error table.
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/project-design.md`
    - Extend the §1 architecture diagram with the `src/folders/` module.
    - Append §2.X module-contract sections for `src/folders/resolver.ts`,
      `src/folders/types.ts`, the extended `OutlookClient` interface, and
      the new `CollisionError` class. Normative TypeScript signatures per
      §2 existing convention.
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/project-functions.MD`
    - Add FR-008 (`list-folders`), FR-009 (`find-folder`), FR-010
      (`create-folder`), FR-011 (`move-mail`), and extend FR-003 with the
      `--folder-id` / `--folder-parent` / path-based `--folder` behaviour.
    - Add FF-006 — folder resolver (path / alias / NFC / case-fold /
      ambiguity policy).
    - Add FF-007 — pagination + 50-page cap.
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/Issues - Pending Items.md`
    - If any follow-up pending items are discovered during P1-P6 (e.g.
      OQ-1..OQ-4 unresolved by user), register them per the
      "pending-on-top" convention.
- **Dependencies.** P6.
- **Parallel with.** P8 (docs and tests touch disjoint files).
- **Acceptance-criteria covered.** AC-CLAUDEMD-UPDATED-FOLDERS (grep check
  in P8).
- **Verification.**
  - `grep -c "<list-folders>\|<find-folder>\|<create-folder>\|<move-mail>"
    CLAUDE.md` returns 4.
  - `grep -n "FR-008\|FR-009\|FR-010\|FR-011" docs/design/project-functions.MD`
    returns four hits.
  - `grep -n "src/folders/" docs/design/project-design.md` returns at
    least one hit.

---

### Phase P8 — Tests (unit + AC + smoke)

- **Goal.** Cover every refined-spec AC with a dedicated script under
  `test_scripts/` and add unit tests for every new module.
- **Files created** (unit tests, each pure — mocked `OutlookClient`):
  - `test_scripts/unit/folders-resolver.spec.ts` — `parseFolderPath`,
    `buildFolderPath`, `matchesWellKnownAlias`, `resolveFolder` (all three
    kinds + ambiguity + firstMatch), `createFolderPath` (400/409
    collision + idempotent true/false + missing-parent),
    `isFolderExistsError` predicate.
  - `test_scripts/unit/outlook-client-post.spec.ts` — `client.post` writes
    the right headers + body; `listAll<T>` follows two-page nextLink +
    caps at 50 pages + rejects off-host nextLink + 401-retries.
  - `test_scripts/unit/list-folders.spec.ts`
  - `test_scripts/unit/find-folder.spec.ts`
  - `test_scripts/unit/create-folder.spec.ts`
  - `test_scripts/unit/move-mail.spec.ts`
  - `test_scripts/unit/list-mail-folder-id.spec.ts` (narrow test for the
    new path / fast-path decision branches).
- **Files created** (AC scripts — each runnable by
  `npx ts-node test_scripts/ac-folders-<name>.ts`, asserts against a live
  session via the shipped CLI. Manual-requires marked with `(manual)`):
  - `test_scripts/ac-folders-listfolders-root.ts`
  - `test_scripts/ac-folders-listfolders-children.ts`
  - `test_scripts/ac-folders-find-wellknown.ts`
  - `test_scripts/ac-folders-find-path.ts`
  - `test_scripts/ac-folders-find-id.ts`
  - `test_scripts/ac-folders-create-toplevel.ts`
  - `test_scripts/ac-folders-create-nested.ts`
  - `test_scripts/ac-folders-create-idempotent.ts`
  - `test_scripts/ac-folders-move-single.ts`
  - `test_scripts/ac-folders-move-many.ts`
  - `test_scripts/ac-folders-listmail-path.ts`
  - `test_scripts/ac-folders-listmail-id.ts`
  - `test_scripts/ac-folders-listmail-wellknown-backcompat.ts`
  - `test_scripts/ac-folders-folder-not-found.ts`
  - `test_scripts/ac-folders-folder-ambiguous.ts`
  - `test_scripts/ac-folders-create-collision.ts`
  - `test_scripts/ac-folders-create-missing-parent.ts`
  - `test_scripts/ac-folders-move-bad-dest.ts`
  - `test_scripts/ac-folders-move-bad-source.ts`
  - `test_scripts/ac-folders-move-partial.ts`
  - `test_scripts/ac-folders-move-stopat.ts`
  - `test_scripts/ac-folders-path-escape.ts`
  - `test_scripts/ac-folders-path-depth-cap.ts`
  - `test_scripts/ac-folders-wellknown-precedence.ts`
  - `test_scripts/ac-folders-401-retry.ts`
  - `test_scripts/ac-folders-no-secret-leak.ts`
  - `test_scripts/ac-folders-claudemd-updated.ts`
- **Dependencies.** P6 (CLI must run); P7 not strictly required but
  AC-CLAUDEMD-UPDATED-FOLDERS tests P7 output.
- **Parallel with.** P7 (docs in parallel with test writing once P6 lands).
  Within P8, every AC script is file-independent — parallelizable.
- **Acceptance-criteria covered.** All 27 ACs from refined §11.
- **Verification.**
  - `npm test` passes (unit + non-interactive AC scripts).
  - `grep -r 'Bearer \|cookie' <log-file-from-ac-no-secret-leak>` returns
    zero hits.

---

## 5. Parallel-Safety Matrix

Rows are implementation units (phases + sub-phases). Columns are the files
each unit writes to. A `W` means "this unit writes (creates or modifies) the
file"; blank means "read-only or untouched". Two units can run concurrently
iff the intersection of their `W` columns is empty.

| Unit | src/http/types.ts | src/folders/types.ts | src/config/errors.ts | src/http/outlook-client.ts | src/folders/resolver.ts | src/commands/list-folders.ts | src/commands/find-folder.ts | src/commands/create-folder.ts | src/commands/move-mail.ts | src/commands/list-mail.ts | src/cli.ts | CLAUDE.md | docs/design/project-design.md | docs/design/project-functions.MD | Issues - Pending Items.md | test_scripts/unit/*.spec.ts | test_scripts/ac-folders-*.ts |
|------|-------------------|---------------------|---------------------|---------------------------|------------------------|-----------------------------|----------------------------|------------------------------|---------------------------|--------------------------|-----------|-----------|------------------------------|---------------------------------|---------------------------|----------------------------|-----------------------------|
| P1   | W                 | W                   |                     |                           |                        |                             |                            |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P2   |                   |                     | W                   |                           |                        |                             |                            |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P3   |                   |                     |                     | W                         |                        |                             |                            |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P4   |                   |                     |                     |                           | W                      |                             |                            |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P5a  |                   |                     |                     |                           |                        | W                           |                            |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P5b  |                   |                     |                     |                           |                        |                             | W                          |                              |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P5c  |                   |                     |                     |                           |                        |                             |                            | W                            |                           |                          |           |           |                              |                                 |                           |                            |                             |
| P5d  |                   |                     |                     |                           |                        |                             |                            |                              | W                         |                          |           |           |                              |                                 |                           |                            |                             |
| P5e  |                   |                     |                     |                           |                        |                             |                            |                              |                           | W                        |           |           |                              |                                 |                           |                            |                             |
| P6   |                   |                     |                     |                           |                        |                             |                            |                              |                           |                          | W         |           |                              |                                 |                           |                            |                             |
| P7   |                   |                     |                     |                           |                        |                             |                            |                              |                           |                          |           | W         | W                            | W                               | W                         |                            |                             |
| P8   |                   |                     |                     |                           |                        |                             |                            |                              |                           |                          |           |           |                              |                                 |                           | W                          | W                           |

**Orchestration plan**:

- **Wave 1 (parallel)**: P1 + P2.
- **Wave 2 (single agent)**: P3 (blocked on P1+P2).
- **Wave 3 (single agent)**: P4 (blocked on P3).
- **Wave 4 (parallel, five agents)**: P5a, P5b, P5c, P5d, P5e.
- **Wave 5 (single agent)**: P6 (blocked on all of P5).
- **Wave 6 (parallel)**: P7 + P8 (write disjoint files; P7 docs can start as
  soon as P6 lands, P8 can also — they do not overlap).

The matrix guarantees that no two parallel-wave agents will ever write the
same file. `test_scripts/*.ts` files inside P8 are also file-disjoint and
can be fanned out per-AC.

---

## 6. Test Strategy

### Unit tests (`test_scripts/unit/`)

Gate: mocked `OutlookClient` (no network). Each new module has a matching
`*.spec.ts`. Focus:

- **resolver.ts** — the richest spec file. Every branch of `parseFolderPath`
  (depth cap, empty segment, escape sequences, NFC), every branch of
  `resolveFolder` (three `FolderSpec` kinds + ambiguity + firstMatch), every
  branch of `createFolderPath` (4xx+ErrorFolderExists, 9xx+other-code, happy
  path, missing parent with/without `createParents`, idempotent swallow).
- **outlook-client.ts** — `post` method: correct headers, correct body
  serialization; `listAll`: follows two nextLinks, stops on missing link,
  caps at 50 pages, rejects off-host link, 401-retries exactly once.
- **Each command file** (P5a-e) — argv validation, `UsageError` branches,
  returns typed result given a mocked client that replays canned responses.

### Integration tests (manual, with live session)

Gate: real session file at `$HOME/.outlook-cli/session.json`. Run via
`npx ts-node test_scripts/ac-folders-*.ts`. Cover every AC that depends on
real Outlook behaviour (most of §11). Bucket:

- **Passing**: AC-LISTFOLDERS-ROOT, AC-LISTFOLDERS-CHILDREN,
  AC-FIND-WELLKNOWN, AC-FIND-PATH, AC-FIND-ID, AC-CREATE-TOPLEVEL,
  AC-CREATE-NESTED, AC-CREATE-IDEMPOTENT, AC-MOVE-SINGLE, AC-MOVE-MANY,
  AC-LISTMAIL-PATH, AC-LISTMAIL-ID, AC-LISTMAIL-WELLKNOWN-BACKCOMPAT.
- **Failing / edge**: AC-FOLDER-NOT-FOUND, AC-FOLDER-AMBIGUOUS,
  AC-CREATE-COLLISION, AC-CREATE-MISSING-PARENT, AC-MOVE-BAD-DEST,
  AC-MOVE-BAD-SOURCE, AC-MOVE-PARTIAL, AC-MOVE-STOPAT, AC-PATH-ESCAPE,
  AC-PATH-DEPTH-CAP, AC-WELLKNOWN-PRECEDENCE, AC-401-RETRY-FOLDERS.

### Smoke tests (no live call)

- `AC-NO-SECRET-LEAK-FOLDERS`: run any folder command with
  `--log-file /tmp/outlook-cli-folder.log` then
  `grep -c 'Bearer \|Cookie:' /tmp/outlook-cli-folder.log` must return `0`.
- `AC-CLAUDEMD-UPDATED-FOLDERS`: grep `CLAUDE.md` for the four new child
  blocks + the documented flag list.
- `npx tsc --noEmit` must pass at every phase boundary.
- `npm run build` must produce a working `dist/cli.js` at P6 completion.
- `node dist/cli.js --help` must list all four new subcommands.

---

## 7. Risks & Mitigations

Each risk maps to a specific phase with a mitigation step. Derived from
`investigation-folders.md §5` + the three research notes.

| # | Risk | Phase | Action |
|---|---|---|---|
| R1 | `ErrorFolderExists` returned as 400 on some tenants, 409 on others | P4 | `isFolderExistsError` predicate inspects `err.body.error.code === 'ErrorFolderExists'` AND (status 400 OR 409). Per `outlook-v2-folder-duplicate-error.md §4.1`. |
| R2 | `DestinationId` alias rejection on v2.0 `/move` | P5d | **Always** resolve aliases to raw ids before `POST /move` (OQ-4). One extra `GET /MailFolders/{alias}` per move; acceptable cost. |
| R3 | `@odata.nextLink` redirected off-host (defense in depth) | P3 | `listAll<T>` validates `new URL(link).hostname === 'outlook.office.com'`; raises `ApiError('PAGINATION_OFF_HOST', ...)`. |
| R4 | Recursive `list-folders` walks blows past safety cap | P3 + P5a | Per-collection cap `MAX_FOLDER_PAGES = 50`; whole-tree cap `MAX_FOLDERS_VISITED = 5000`. Both raise `UpstreamError('UPSTREAM_PAGINATION_LIMIT', ...)` with actionable message. |
| R5 | Race window between lookup-then-create (concurrent `create-folder` runs) | P4 | Under `--idempotent`, on 400/409+ErrorFolderExists re-list the parent's children to recover the existing id — do not trust `POST` response on the collision path. |
| R6 | Move endpoint returns new id → scripts chain old id | P5d + P6 | `MoveMailResult.moved[]` explicitly surfaces `{ sourceId, newId }` pair. Documented in CLAUDE.md `<move-mail>` block. |
| R7 | Partial failure in batch move absorbed into exit 0 | P5d + P6 | `--continue-on-error` + any `failed[]` entries → exit 5. The `UpstreamError` is raised by cli.ts wrapper **after** `emitResult` (so the payload is still emitted). |
| R8 | Folder DisplayName has raw `/` — user forgets to escape | P4 | `parseFolderPath` documents the escape rule; resolver raises `UPSTREAM_FOLDER_NOT_FOUND` on the wrong segment (accurate, not misleading). |
| R9 | Ambiguous path silently resolves to one match under `--first-match` | P4 + P5b/d | `--first-match` flag documented as a foot-gun in CLAUDE.md. Default behaviour (exit 2 on ambiguity) is the safe path. |
| R10 | Well-known alias shadowed by a user folder at the root | P4 | Resolver hard-codes "well-known wins at root" (refined §6.2). User must pass `--parent MsgFolderRoot --first-match` to reach the shadowed user folder. |
| R11 | Hidden folder collides with create (invisible in pre-create list) | P4 | `createFolderPath` treats 400/409+ErrorFolderExists as authoritative even when pre-create list missed it. `--idempotent` recovers via re-list of all (including `includeHiddenFolders=true`) children. |
| R12 | Bearer expires mid-recursive-walk | P3 | 401-retry-once envelope in `doRequest` transparently recovers on the failing page request; subsequent pages use the refreshed token (mutable session). Verified by unit test in P8. |
| R13 | Token / cookie value leak in error body snippets | P3 | `throwForResponse` already runs bodies through `truncateAndRedactBody` (unchanged). Folder code does not bypass it. AC-NO-SECRET-LEAK-FOLDERS asserts. |
| R14 | Move response body empty / `Id` missing on some tenant | P5d | Treat empty `Id` in response as `UpstreamError('UPSTREAM_HTTP_200', 'move response missing new id')`; entry lands in `failed[]` under `--continue-on-error`. |
| R15 | Follow-up rename / delete requests slip in accidentally | all | NG1 / NG2 enforced by scope: no `PATCH` or `DELETE` helper added to `OutlookClient` in this plan. |

---

## 8. Verification Criteria

Claude can execute the following at each phase boundary:

### Compile-time

- `cd /Users/giorgosmarinos/aiwork/coding-platform/outlook-tool && npx tsc
  --noEmit` — must return exit 0 at every phase boundary.

### Unit tests

- `npm test` — must return exit 0 at every phase boundary after P4.
- At P3 boundary: `npm test -- outlook-client` passes (new `post` +
  `listAll` specs).
- At P4 boundary: `npm test -- folders-resolver` passes.
- At P5 boundary: `npm test -- list-folders find-folder create-folder
  move-mail list-mail-folder-id` passes.

### Build

- At P6 boundary: `npm run build` produces `dist/cli.js`.
- `node dist/cli.js --help` prints all 4 new subcommands.
- `node dist/cli.js list-mail --help` shows `--folder-id` and
  `--folder-parent`.
- `node dist/cli.js list-folders --help` shows `--parent`, `--recursive`,
  `--include-hidden`, `--top`.
- `node dist/cli.js move-mail --help` shows every option from §5.4.

### Manual checks (live session)

Each AC under `test_scripts/ac-folders-*.ts` is runnable on its own.
Recommended order for a manual smoke pass:

1. `outlook-cli list-folders --table` — prints at minimum `Inbox`,
   `SentItems`.
2. `outlook-cli find-folder Inbox --json | jq .ResolvedVia` → `"wellknown"`.
3. `outlook-cli create-folder "Outlook-CLI-Smoke-$(date +%s)"` → exit 0,
   `created[0].Id` non-empty, `PreExisting == false`.
4. Re-run step 3 → exit 6 with `code: "FOLDER_ALREADY_EXISTS"`.
5. Re-run step 3 with `--idempotent` → exit 0, `idempotent: true`.
6. `outlook-cli list-mail --folder-id <id-from-step-3> -n 5` → returns
   messages (or empty array) from that folder.
7. `outlook-cli move-mail <any-inbox-message-id> --to "<smoke-folder>"` →
   exit 0, `moved[0].newId != moved[0].sourceId`.
8. `outlook-cli --log-file /tmp/outlook-cli-smoke.log list-folders
   --recursive --table` — the log file must not contain `Bearer ` or
   `Cookie:` (AC-NO-SECRET-LEAK-FOLDERS).

### Documentation grep checks

- `grep -c "<list-folders>\|<find-folder>\|<create-folder>\|<move-mail>"
  CLAUDE.md` → `4`.
- `grep -n "FR-008\|FR-009\|FR-010\|FR-011"
  docs/design/project-functions.MD` → four hits.
- `grep -n "src/folders/resolver.ts"
  docs/design/project-design.md` → at least one hit.
- `grep -n "CollisionError\|FOLDER_ALREADY_EXISTS"
  src/config/errors.ts src/cli.ts` → hits in both files.

---

## 9. Summary

- **10 discrete implementation units** across 8 phases:
  P1 (types), P2 (errors), P3 (http client), P4 (resolver), P5a-e
  (five command files), P6 (cli.ts), P7 (docs), P8 (tests).
- **Four parallelization windows**:
  - Wave 1: `P1 || P2` (fresh types + errors).
  - Wave 4: `P5a || P5b || P5c || P5d || P5e` (five leaf command files).
  - Wave 6: `P7 || P8` (docs + tests touch disjoint files).
  - Within P8 itself, 27 AC scripts and 7 unit-spec files are all
    file-disjoint and individually parallelizable.
- **Four open questions requiring user decision before coding** —
  see §0 (OQ-1..OQ-4). Defaults have been picked and ship as written
  unless the user overrides.
- **Zero new mandatory config**, **zero new exit codes** (reuses
  0/1/2/3/4/5/6 taxonomy), **one new error class** (`CollisionError`,
  exit 6), **one new runtime dependency**: none — all additions are
  TypeScript source only.
- Full AC coverage (`refined-request-folders.md §11`) mapped 1:1 in §4
  and §6.
- Risk register (§7) pins the 15 failure modes across the three research
  notes and the investigation risk register; each has a named mitigation
  in a specific phase.

Absolute output path: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/plan-002-folders.md`
