// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { TfvcVersionDescriptor, TfvcVersionType, VersionControlRecursionType, TfvcChangesetSearchCriteria, VersionControlChangeType } from "azure-devops-node-api/interfaces/TfvcInterfaces.js";
import { z } from "zod";
import { getEnumKeys, safeEnumConvert, streamToString, extractAdoStreamError } from "../utils.js";
import { createExternalContentResponse } from "../shared/content-safety.js";

const TFVC_TOOLS = {
  list_changesets: "tfvc_list_changesets",
  get_changeset: "tfvc_get_changeset",
  get_changeset_changes: "tfvc_get_changeset_changes",
  get_item_content: "tfvc_get_item_content",
  get_file_diff: "tfvc_get_file_diff",
  list_items: "tfvc_list_items",
};

/** Maximum comment length returned by the TFVC API when listing changesets. Limits response size. */
const MAX_COMMENT_LENGTH = 200;

const BINARY_EXTENSIONS = new Set([
  ".png",
  ".jpg",
  ".jpeg",
  ".gif",
  ".bmp",
  ".ico",
  ".tiff",
  ".webp",
  ".svg",
  ".pdf",
  ".zip",
  ".gz",
  ".tar",
  ".rar",
  ".7z",
  ".dll",
  ".exe",
  ".so",
  ".dylib",
  ".bin",
  ".obj",
  ".lib",
  ".pdb",
  ".msi",
  ".cab",
  ".iso",
  ".img",
  ".mp3",
  ".mp4",
  ".avi",
  ".mov",
  ".wav",
  ".woff",
  ".woff2",
  ".ttf",
  ".eot",
  ".otf",
  ".pyc",
  ".class",
  ".jar",
  ".war",
  ".ear",
]);

function isBinaryPath(path: string): boolean {
  const dotIdx = path.lastIndexOf(".");
  if (dotIdx === -1) return false;
  return BINARY_EXTENSIONS.has(path.substring(dotIdx).toLowerCase());
}

function changeTypeToString(changeType?: number): string {
  if (changeType === undefined || changeType === null) return "unknown";
  // VersionControlChangeType is a flags enum — collect all matching names
  const names: string[] = [];
  for (const key of getEnumKeys(VersionControlChangeType)) {
    const val = VersionControlChangeType[key as keyof typeof VersionControlChangeType] as number;
    if (val !== 0 && (changeType & val) === val) {
      names.push(key);
    }
  }
  return names.length > 0 ? names.join(", ") : String(changeType);
}

function buildVersionDescriptor(changesetId?: number): TfvcVersionDescriptor | undefined {
  if (changesetId === undefined) return undefined;
  return {
    version: String(changesetId),
    versionType: TfvcVersionType.Changeset,
  };
}

function configureTfvcTools(server: McpServer, _tokenProvider: () => Promise<string>, connectionProvider: () => Promise<WebApi>) {
  server.tool(
    TFVC_TOOLS.list_changesets,
    "List TFVC changesets for a project. Supports filtering by author, date range, ID range, and item path. Returns changeset ID, author, date, and comment.",
    {
      project: z.string().optional().describe("Project ID or name. If not provided, lists changesets across all projects."),
      author: z.string().optional().describe("Filter changesets by author display name or unique name."),
      fromId: z.coerce.number().min(1).optional().describe("Minimum changeset ID to return (inclusive)."),
      toId: z.coerce.number().min(1).optional().describe("Maximum changeset ID to return (inclusive)."),
      fromDate: z.string().optional().describe("Return changesets created on or after this date (ISO 8601 format, e.g. '2024-01-01')."),
      toDate: z.string().optional().describe("Return changesets created on or before this date (ISO 8601 format, e.g. '2024-12-31')."),
      itemPath: z.string().optional().describe("Filter changesets that contain changes to this TFVC path (e.g. '$/MyProject/src')."),
      top: z.coerce.number().min(1).max(1000).default(50).describe("Maximum number of changesets to return. Defaults to 50, max 1000."),
      skip: z.coerce.number().min(0).default(0).describe("Number of changesets to skip for pagination. Defaults to 0."),
      orderby: z.enum(["id asc", "id desc"]).default("id desc").describe("Sort order for results. Defaults to 'id desc' (newest first)."),
    },
    async ({ project, author, fromId, toId, fromDate, toDate, itemPath, top, skip, orderby }) => {
      try {
        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const searchCriteria: TfvcChangesetSearchCriteria = {};
        if (author) searchCriteria.author = author;
        if (fromId) searchCriteria.fromId = fromId;
        if (toId) searchCriteria.toId = toId;
        if (fromDate) searchCriteria.fromDate = fromDate;
        if (toDate) searchCriteria.toDate = toDate;
        if (itemPath) searchCriteria.itemPath = itemPath;

        const changesets = await tfvcApi.getChangesets(project, MAX_COMMENT_LENGTH, skip, top, orderby, searchCriteria);

        if (!changesets || changesets.length === 0) {
          return { content: [{ type: "text", text: "No changesets found." }] };
        }

        const trimmed = changesets.map((cs) => ({
          changesetId: cs.changesetId,
          author: {
            displayName: cs.author?.displayName,
            uniqueName: cs.author?.uniqueName,
          },
          createdDate: cs.createdDate,
          comment: cs.comment,
          url: cs.url,
        }));

        return { content: [{ type: "text", text: JSON.stringify(trimmed, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error listing changesets: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    TFVC_TOOLS.get_changeset,
    "Get full details of a single TFVC changeset by its ID. Includes comment, author, creation date, policy override info, check-in notes, and optionally linked work items.",
    {
      changesetId: z.coerce.number().min(1).describe("The ID of the changeset to retrieve."),
      project: z.string().optional().describe("Project ID or name. Optional — the changeset ID uniquely identifies the changeset."),
      includeWorkItems: z.boolean().optional().default(true).describe("Whether to include linked work items. Defaults to true."),
      includeDetails: z.boolean().optional().default(true).describe("Whether to include policy details and check-in notes. Defaults to true."),
      maxChangeCount: z.coerce
        .number()
        .min(0)
        .max(100)
        .optional()
        .default(0)
        .describe("Number of file changes to include inline (0–100). Defaults to 0 (no inline changes). Use tfvc_get_changeset_changes for the full list."),
    },
    async ({ changesetId, project, includeWorkItems, includeDetails, maxChangeCount }) => {
      try {
        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const changeset = await tfvcApi.getChangeset(changesetId, project, maxChangeCount, includeDetails, includeWorkItems);

        if (!changeset) {
          return { content: [{ type: "text", text: `Changeset ${changesetId} not found.` }], isError: true };
        }

        const result = {
          changesetId: changeset.changesetId,
          author: {
            displayName: changeset.author?.displayName,
            uniqueName: changeset.author?.uniqueName,
          },
          checkedInBy: {
            displayName: changeset.checkedInBy?.displayName,
            uniqueName: changeset.checkedInBy?.uniqueName,
          },
          createdDate: changeset.createdDate,
          comment: changeset.comment,
          policyOverride: changeset.policyOverride,
          checkinNotes: changeset.checkinNotes,
          workItems: changeset.workItems?.map((wi) => ({ id: wi.id, title: wi.title, type: wi.workItemType, state: wi.state, assignedTo: wi.assignedTo })),
          changes: changeset.changes?.map((ch) => ({
            changeType: changeTypeToString(ch.changeType),
            item: { path: ch.item?.path, contentMetadata: ch.item?.contentMetadata },
          })),
          url: changeset.url,
        };

        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error getting changeset ${changesetId}: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    TFVC_TOOLS.get_changeset_changes,
    "Get all file changes within a TFVC changeset. Returns path and change type for each changed item. Binary files are automatically excluded from content retrieval.",
    {
      changesetId: z.coerce.number().min(1).describe("The ID of the changeset to get changes for."),
      top: z.coerce.number().min(1).max(1000).default(100).describe("Maximum number of changes to return. Defaults to 100, max 1000."),
      skip: z.coerce.number().min(0).default(0).describe("Number of changes to skip for pagination. Defaults to 0."),
    },
    async ({ changesetId, top, skip }) => {
      try {
        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const changesPage = await tfvcApi.getChangesetChanges(changesetId, skip, top);
        const changes = Array.isArray(changesPage) ? changesPage : [];

        if (!changes || (changes as unknown[]).length === 0) {
          return { content: [{ type: "text", text: `No changes found in changeset ${changesetId}.` }] };
        }

        const result = (changes as { changeType?: number; item?: { path?: string; contentMetadata?: { encoding?: number } } }[])
          .filter((ch) => ch.item?.path)
          .map((ch) => ({
            path: ch.item!.path,
            changeType: changeTypeToString(ch.changeType),
            isBinary: isBinaryPath(ch.item!.path!),
          }));

        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error getting changeset changes for ${changesetId}: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    TFVC_TOOLS.get_item_content,
    "Retrieve the text content of a TFVC file at a specified path and optional version (changeset ID). Returns raw file text. Binary files are not supported.",
    {
      path: z.string().describe("The TFVC server path of the file, e.g. '$/MyProject/src/MyClass.cs'."),
      project: z.string().optional().describe("Project ID or name. Optional when using a full server path starting with '$/'."),
      changesetId: z.coerce.number().min(1).optional().describe("Retrieve the file at this specific changeset version. Defaults to the latest version if not provided."),
    },
    async ({ path, project, changesetId }) => {
      try {
        if (isBinaryPath(path)) {
          return { content: [{ type: "text", text: `Binary file '${path}' cannot be returned as text.` }], isError: true };
        }

        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const versionDescriptor = buildVersionDescriptor(changesetId);
        const stream = await tfvcApi.getItemContent(path, project, undefined, false, undefined, VersionControlRecursionType.None, versionDescriptor, true);

        const content = await streamToString(stream);
        const streamError = extractAdoStreamError(content);
        if (streamError) {
          return { content: [{ type: "text", text: `Error retrieving '${path}': ${streamError}` }], isError: true };
        }

        return createExternalContentResponse(content, `TFVC file: ${path}`);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error getting content for '${path}': ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    TFVC_TOOLS.get_file_diff,
    "Get a line-by-line text diff of a specific file between two changeset versions. Fetches the before and after versions of the file and computes a unified diff. Binary files are not supported.",
    {
      path: z.string().describe("The TFVC server path of the file to diff, e.g. '$/MyProject/src/MyClass.cs'."),
      project: z.string().optional().describe("Project ID or name. Optional when the path starts with '$/'."),
      beforeChangesetId: z.coerce.number().min(1).describe("The changeset ID for the 'before' version of the file."),
      afterChangesetId: z.coerce.number().min(1).describe("The changeset ID for the 'after' version of the file."),
      contextLines: z.coerce.number().min(0).max(20).default(3).describe("Number of context lines to include around each change. Defaults to 3."),
    },
    async ({ path, project, beforeChangesetId, afterChangesetId, contextLines }) => {
      try {
        if (isBinaryPath(path)) {
          return { content: [{ type: "text", text: `Binary file '${path}' cannot be diffed as text.` }], isError: true };
        }

        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const [beforeStream, afterStream] = await Promise.all([
          tfvcApi.getItemContent(path, project, undefined, false, undefined, VersionControlRecursionType.None, buildVersionDescriptor(beforeChangesetId), true),
          tfvcApi.getItemContent(path, project, undefined, false, undefined, VersionControlRecursionType.None, buildVersionDescriptor(afterChangesetId), true),
        ]);

        const [beforeContent, afterContent] = await Promise.all([streamToString(beforeStream), streamToString(afterStream)]);

        const beforeError = extractAdoStreamError(beforeContent);
        if (beforeError) return { content: [{ type: "text", text: `Error retrieving 'before' version of '${path}': ${beforeError}` }], isError: true };

        const afterError = extractAdoStreamError(afterContent);
        if (afterError) return { content: [{ type: "text", text: `Error retrieving 'after' version of '${path}': ${afterError}` }], isError: true };

        const diff = computeUnifiedDiff(path, beforeContent, afterContent, beforeChangesetId, afterChangesetId, contextLines);

        return createExternalContentResponse(diff, `TFVC diff: ${path}`);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error computing diff for '${path}': ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    TFVC_TOOLS.list_items,
    "Browse a TFVC directory path and list its items. Equivalent to 'repo_list_directory' for Git, but for TFVC. Returns items with their type (file or folder) and server path.",
    {
      scopePath: z.string().describe("The TFVC server path to list, e.g. '$/MyProject' or '$/MyProject/src'."),
      project: z.string().optional().describe("Project ID or name. Optional when the path starts with '$/'."),
      recursionLevel: z
        .enum(getEnumKeys(VersionControlRecursionType) as [string, ...string[]])
        .default("OneLevel")
        .describe("Recursion depth: 'None' (folder metadata only), 'OneLevel' (immediate children, default), 'Full' (all descendants)."),
      changesetId: z.coerce.number().min(1).optional().describe("List items at this specific changeset version. Defaults to the latest version."),
    },
    async ({ scopePath, project, recursionLevel, changesetId }) => {
      try {
        const connection = await connectionProvider();
        const tfvcApi = await connection.getTfvcApi();

        const versionDescriptor = buildVersionDescriptor(changesetId);
        const recursion = safeEnumConvert(VersionControlRecursionType, recursionLevel);

        const items = await tfvcApi.getItems(project, scopePath, recursion, false, versionDescriptor);

        if (!items || items.length === 0) {
          return { content: [{ type: "text", text: `No items found at '${scopePath}'.` }] };
        }

        const result = items
          .filter((item) => item.path && item.path !== scopePath)
          .map((item) => ({
            path: item.path,
            isFolder: item.isFolder,
            size: item.size,
            changeDate: item.changeDate,
          }));

        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error listing items at '${scopePath}': ${errorMessage}` }], isError: true };
      }
    }
  );
}

/**
 * Compute a simple unified diff between two text strings.
 * Uses the Myers diff algorithm approximation via line-by-line LCS.
 */
function computeUnifiedDiff(path: string, before: string, after: string, beforeId: number, afterId: number, contextLines: number): string {
  const beforeLines = before.split("\n");
  const afterLines = after.split("\n");

  // Build LCS table
  const m = beforeLines.length;
  const n = afterLines.length;

  // Limit file size to avoid O(m*n) LCS memory explosion (5000 * 5000 = 25M cells × 8 bytes = ~200MB)
  const MAX_LINES = 5000;
  if (m > MAX_LINES || n > MAX_LINES) {
    return `--- ${path}\t(changeset ${beforeId})\n+++ ${path}\t(changeset ${afterId})\n@@ File too large for inline diff (>${MAX_LINES} lines). Use tfvc_get_item_content to retrieve each version separately. @@\n`;
  }

  // Build diff ops using simple O(m*n) LCS
  const dp: number[][] = Array.from({ length: m + 1 }, () => new Array<number>(n + 1).fill(0));
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      if (beforeLines[i] === afterLines[j]) {
        dp[i][j] = 1 + dp[i + 1][j + 1];
      } else {
        dp[i][j] = Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
  }

  // Generate list of ops: "=" keep, "-" remove, "+" add
  type Op = { op: "=" | "-" | "+"; line: string };
  const ops: Op[] = [];
  let i = 0,
    j = 0;
  while (i < m || j < n) {
    if (i < m && j < n && beforeLines[i] === afterLines[j]) {
      ops.push({ op: "=", line: beforeLines[i] });
      i++;
      j++;
    } else if (j < n && (i >= m || dp[i][j + 1] >= dp[i + 1][j])) {
      ops.push({ op: "+", line: afterLines[j] });
      j++;
    } else {
      ops.push({ op: "-", line: beforeLines[i] });
      i++;
    }
  }

  // Build unified diff output with context
  const header = `--- ${path}\t(changeset ${beforeId})\n+++ ${path}\t(changeset ${afterId})\n`;
  if (ops.every((o) => o.op === "=")) {
    return header + "(no differences)\n";
  }

  // Group into hunks
  const hunks: string[] = [];
  let opIdx = 0;
  while (opIdx < ops.length) {
    // Find the next change
    while (opIdx < ops.length && ops[opIdx].op === "=") opIdx++;
    if (opIdx >= ops.length) break;

    // Expand window by contextLines on each side
    const hunkStart = Math.max(0, opIdx - contextLines);
    let hunkEnd = opIdx;
    while (hunkEnd < ops.length && (ops[hunkEnd].op !== "=" || hunkEnd - opIdx < contextLines)) {
      hunkEnd++;
    }
    hunkEnd = Math.min(ops.length, hunkEnd + contextLines);

    const hunkOps = ops.slice(hunkStart, hunkEnd);
    const removedCount = hunkOps.filter((o) => o.op === "=" || o.op === "-").length;
    const addedCount = hunkOps.filter((o) => o.op === "=" || o.op === "+").length;

    // Compute line numbers
    let beforeLine = 1 + ops.slice(0, hunkStart).filter((o) => o.op === "=" || o.op === "-").length;
    let afterLine = 1 + ops.slice(0, hunkStart).filter((o) => o.op === "=" || o.op === "+").length;

    const hunkLines = [`@@ -${beforeLine},${removedCount} +${afterLine},${addedCount} @@`];
    for (const op of hunkOps) {
      if (op.op === "=") hunkLines.push(` ${op.line}`);
      else if (op.op === "-") hunkLines.push(`-${op.line}`);
      else hunkLines.push(`+${op.line}`);
    }
    hunks.push(hunkLines.join("\n"));
    opIdx = hunkEnd;
  }

  return header + hunks.join("\n");
}

export { TFVC_TOOLS, configureTfvcTools };
