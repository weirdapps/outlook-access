// src/commands/find-folder.ts
//
// Resolve a folder query (well-known alias, display-name path, or `id:<raw>`)
// into a single ResolvedFolder.
//
// Normative sources:
//   - docs/design/project-design.md §10.7 (find-folder CLI surface)
//   - docs/design/refined-request-folders.md §5.2
//   - docs/design/plan-002-folders.md §P5b

import type { CliConfig } from '../config/config';
import { parseFolderSpec, resolveFolder } from '../folders/resolver';
import type { FolderSpec, ResolvedFolder } from '../folders/types';
import type { OutlookClient } from '../http/outlook-client';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface FindFolderDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface FindFolderOptions {
  /**
   * Optional anchor for path-form queries. Accepts a well-known alias, a
   * display-name path, or an `id:<raw>` form — the same grammar as the
   * positional `<spec>` itself. Defaults to `MsgFolderRoot`. Only applies
   * when the positional spec is a path (well-known / id forms ignore it).
   */
  anchor?: string;
  /**
   * When true, ambiguity at any path segment is resolved deterministically
   * (ADR-14: `CreatedDateTime asc, Id asc`) instead of raising
   * `UsageError('FOLDER_AMBIGUOUS')`.
   */
  firstMatch?: boolean;
}

export async function run(
  deps: FindFolderDeps,
  spec: string,
  opts: FindFolderOptions = {},
): Promise<ResolvedFolder> {
  if (typeof spec !== 'string' || spec.length === 0) {
    throw new UsageError('find-folder: <spec> is required');
  }

  // Parse the positional query into a FolderSpec tagged union. Raw-id,
  // well-known alias, and path forms are discriminated by `parseFolderSpec`
  // — we do not replicate the grammar here (§10.5, ADR-13).
  const querySpec: FolderSpec = parseFolderSpec(spec);

  // Attach the anchor only when meaningful: the resolver ignores a `parent`
  // on wellKnown / id specs (they resolve via a single GET). Restricting
  // parse of the anchor string to the path branch keeps `find-folder Inbox
  // --anchor <garbage>` from raising a grammar error on the anchor when the
  // anchor would be silently ignored anyway.
  let finalSpec: FolderSpec = querySpec;
  if (
    querySpec.kind === 'path' &&
    typeof opts.anchor === 'string' &&
    opts.anchor.length > 0
  ) {
    const anchorSpec: FolderSpec = parseFolderSpec(opts.anchor);
    finalSpec = { ...querySpec, parent: anchorSpec };
  }

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  try {
    return await resolveFolder(client, finalSpec, {
      firstMatch: opts.firstMatch === true,
    });
  } catch (err) {
    throw mapHttpError(err);
  }
}
