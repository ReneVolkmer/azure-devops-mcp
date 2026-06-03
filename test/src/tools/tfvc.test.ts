// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it, beforeEach } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { configureTfvcTools, TFVC_TOOLS } from "../../../src/tools/tfvc";
import { Readable } from "stream";

function makeReadable(text: string): Readable {
  const r = new Readable();
  r.push(text);
  r.push(null);
  return r;
}

describe("configureTfvcTools", () => {
  let server: McpServer;
  let tokenProvider: jest.MockedFunction<() => Promise<string>>;
  let connectionProvider: jest.MockedFunction<() => Promise<WebApi>>;
  let mockTfvcApi: {
    getChangesets: jest.Mock;
    getChangeset: jest.Mock;
    getChangesetChanges: jest.Mock;
    getItemContent: jest.Mock;
    getItems: jest.Mock;
  };

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn();

    mockTfvcApi = {
      getChangesets: jest.fn(),
      getChangeset: jest.fn(),
      getChangesetChanges: jest.fn(),
      getItemContent: jest.fn(),
      getItems: jest.fn(),
    };

    connectionProvider = jest.fn().mockResolvedValue({
      getTfvcApi: jest.fn().mockResolvedValue(mockTfvcApi),
    } as unknown as WebApi);
  });

  describe("tool registration", () => {
    it("registers all TFVC tools", () => {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const registeredNames = (server.tool as jest.Mock).mock.calls.map(([name]) => name);
      expect(registeredNames).toContain(TFVC_TOOLS.list_changesets);
      expect(registeredNames).toContain(TFVC_TOOLS.get_changeset);
      expect(registeredNames).toContain(TFVC_TOOLS.get_changeset_changes);
      expect(registeredNames).toContain(TFVC_TOOLS.get_item_content);
      expect(registeredNames).toContain(TFVC_TOOLS.get_file_diff);
      expect(registeredNames).toContain(TFVC_TOOLS.list_items);
    });
  });

  describe("tfvc_list_changesets", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.list_changesets);
      if (!call) throw new Error("tfvc_list_changesets not registered");
      return call[3];
    }

    it("returns trimmed changeset list", async () => {
      mockTfvcApi.getChangesets.mockResolvedValue([
        {
          changesetId: 100,
          author: { displayName: "Jane Dev", uniqueName: "jane@example.com" },
          createdDate: "2024-01-15T10:00:00Z",
          comment: "Fix compilation error",
          url: "https://dev.azure.com/org/proj/_apis/tfvc/changesets/100",
        },
        {
          changesetId: 99,
          author: { displayName: "John Dev", uniqueName: "john@example.com" },
          createdDate: "2024-01-14T09:00:00Z",
          comment: "Add new feature",
          url: "https://dev.azure.com/org/proj/_apis/tfvc/changesets/99",
        },
      ]);

      const result = await getHandler()({ project: "MyProject", top: 50, skip: 0, orderby: "id desc" });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].changesetId).toBe(100);
      expect(parsed[0].author.displayName).toBe("Jane Dev");
      expect(parsed[0].comment).toBe("Fix compilation error");
      expect(result.isError).toBeUndefined();
    });

    it("returns 'no changesets found' when list is empty", async () => {
      mockTfvcApi.getChangesets.mockResolvedValue([]);
      const result = await getHandler()({ project: "MyProject", top: 50, skip: 0, orderby: "id desc" });
      expect(result.content[0].text).toContain("No changesets found");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getChangesets.mockRejectedValue(new Error("Network timeout"));
      const result = await getHandler()({ project: "MyProject", top: 50, skip: 0, orderby: "id desc" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Network timeout");
    });

    it("passes author and date filters to the API", async () => {
      mockTfvcApi.getChangesets.mockResolvedValue([]);
      await getHandler()({ project: "MyProject", author: "jane@example.com", fromDate: "2024-01-01", toDate: "2024-12-31", top: 10, skip: 0, orderby: "id asc" });
      expect(mockTfvcApi.getChangesets).toHaveBeenCalledWith("MyProject", 200, 0, 10, "id asc", expect.objectContaining({ author: "jane@example.com", fromDate: "2024-01-01", toDate: "2024-12-31" }));
    });
  });

  describe("tfvc_get_changeset", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.get_changeset);
      if (!call) throw new Error("tfvc_get_changeset not registered");
      return call[3];
    }

    it("returns full changeset details with work items", async () => {
      mockTfvcApi.getChangeset.mockResolvedValue({
        changesetId: 100,
        author: { displayName: "Jane Dev", uniqueName: "jane@example.com" },
        checkedInBy: { displayName: "Jane Dev", uniqueName: "jane@example.com" },
        createdDate: "2024-01-15T10:00:00Z",
        comment: "Fix compilation error",
        workItems: [{ id: 12345, title: "Bug: compile error", workItemType: "Bug", state: "Active", assignedTo: "Jane Dev" }],
        changes: [],
        checkinNotes: [],
        url: "https://dev.azure.com/org/proj/_apis/tfvc/changesets/100",
      });

      const result = await getHandler()({ changesetId: 100, includeWorkItems: true, includeDetails: true, maxChangeCount: 0 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.changesetId).toBe(100);
      expect(parsed.workItems).toHaveLength(1);
      expect(parsed.workItems[0].id).toBe(12345);
      expect(result.isError).toBeUndefined();
    });

    it("returns error when changeset not found", async () => {
      mockTfvcApi.getChangeset.mockResolvedValue(null);
      const result = await getHandler()({ changesetId: 9999, includeWorkItems: true, includeDetails: true, maxChangeCount: 0 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("9999");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getChangeset.mockRejectedValue(new Error("Unauthorized"));
      const result = await getHandler()({ changesetId: 100, includeWorkItems: true, includeDetails: true, maxChangeCount: 0 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Unauthorized");
    });
  });

  describe("tfvc_get_changeset_changes", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.get_changeset_changes);
      if (!call) throw new Error("tfvc_get_changeset_changes not registered");
      return call[3];
    }

    it("returns file changes with type and binary flag", async () => {
      mockTfvcApi.getChangesetChanges.mockResolvedValue([
        { changeType: 2, item: { path: "$/MyProject/src/MyClass.cs" } },
        { changeType: 1, item: { path: "$/MyProject/assets/logo.png" } },
      ]);

      const result = await getHandler()({ changesetId: 100, top: 100, skip: 0 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].path).toBe("$/MyProject/src/MyClass.cs");
      expect(parsed[0].isBinary).toBe(false);
      expect(parsed[1].path).toBe("$/MyProject/assets/logo.png");
      expect(parsed[1].isBinary).toBe(true);
    });

    it("returns 'no changes found' when list is empty", async () => {
      mockTfvcApi.getChangesetChanges.mockResolvedValue([]);
      const result = await getHandler()({ changesetId: 100, top: 100, skip: 0 });
      expect(result.content[0].text).toContain("No changes found");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getChangesetChanges.mockRejectedValue(new Error("Server error"));
      const result = await getHandler()({ changesetId: 100, top: 100, skip: 0 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Server error");
    });
  });

  describe("tfvc_get_item_content", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.get_item_content);
      if (!call) throw new Error("tfvc_get_item_content not registered");
      return call[3];
    }

    it("returns text file content", async () => {
      const fileContent = "public class MyClass { }";
      mockTfvcApi.getItemContent.mockResolvedValue(makeReadable(fileContent));

      const result = await getHandler()({ path: "$/MyProject/src/MyClass.cs", changesetId: 100 });
      expect(result.content[0].text).toContain(fileContent);
      expect(result.isError).toBeUndefined();
    });

    it("rejects binary files without calling API", async () => {
      const result = await getHandler()({ path: "$/MyProject/assets/logo.png" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Binary file");
      expect(mockTfvcApi.getItemContent).not.toHaveBeenCalled();
    });

    it("handles ADO stream error response", async () => {
      const errorJson = JSON.stringify({ typeName: "GitItemNotFoundException", message: "Item not found at path" });
      mockTfvcApi.getItemContent.mockResolvedValue(makeReadable(errorJson));

      const result = await getHandler()({ path: "$/MyProject/src/Missing.cs" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Item not found at path");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getItemContent.mockRejectedValue(new Error("Path not found"));
      const result = await getHandler()({ path: "$/MyProject/src/MyClass.cs" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Path not found");
    });
  });

  describe("tfvc_get_file_diff", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.get_file_diff);
      if (!call) throw new Error("tfvc_get_file_diff not registered");
      return call[3];
    }

    it("rejects binary files without calling API", async () => {
      const result = await getHandler()({ path: "$/MyProject/assets/logo.png", beforeChangesetId: 99, afterChangesetId: 100, contextLines: 3 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Binary file");
      expect(mockTfvcApi.getItemContent).not.toHaveBeenCalled();
    });

    it("returns a unified diff for text files", async () => {
      const before = "line1\noriginal\nline3\n";
      const after = "line1\nmodified\nline3\n";
      mockTfvcApi.getItemContent.mockResolvedValueOnce(makeReadable(before)).mockResolvedValueOnce(makeReadable(after));

      const result = await getHandler()({ path: "$/MyProject/src/MyClass.cs", beforeChangesetId: 99, afterChangesetId: 100, contextLines: 3 });
      expect(result.content[0].text).toContain("---");
      expect(result.content[0].text).toContain("+++");
      expect(result.isError).toBeUndefined();
    });

    it("returns '(no differences)' when files are identical", async () => {
      const content = "line1\nline2\nline3\n";
      mockTfvcApi.getItemContent.mockResolvedValueOnce(makeReadable(content)).mockResolvedValueOnce(makeReadable(content));

      const result = await getHandler()({ path: "$/MyProject/src/MyClass.cs", beforeChangesetId: 99, afterChangesetId: 100, contextLines: 3 });
      expect(result.content[0].text).toContain("(no differences)");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getItemContent.mockRejectedValue(new Error("Auth error"));
      const result = await getHandler()({ path: "$/MyProject/src/MyClass.cs", beforeChangesetId: 99, afterChangesetId: 100, contextLines: 3 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Auth error");
    });
  });

  describe("tfvc_list_items", () => {
    function getHandler() {
      configureTfvcTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === TFVC_TOOLS.list_items);
      if (!call) throw new Error("tfvc_list_items not registered");
      return call[3];
    }

    it("returns items excluding the scope root", async () => {
      mockTfvcApi.getItems.mockResolvedValue([
        { path: "$/MyProject/src", isFolder: true, size: 0 },
        { path: "$/MyProject/src/MyClass.cs", isFolder: false, size: 1024 },
        { path: "$/MyProject/src/Utils.cs", isFolder: false, size: 512 },
      ]);

      const result = await getHandler()({ scopePath: "$/MyProject/src", recursionLevel: "OneLevel" });
      const parsed = JSON.parse(result.content[0].text);
      // The scope root itself should be excluded
      expect(parsed.every((item: { path: string }) => item.path !== "$/MyProject/src")).toBe(true);
      expect(parsed).toHaveLength(2);
    });

    it("returns 'no items found' when list is empty", async () => {
      mockTfvcApi.getItems.mockResolvedValue([]);
      const result = await getHandler()({ scopePath: "$/MyProject/empty", recursionLevel: "OneLevel" });
      expect(result.content[0].text).toContain("No items found");
    });

    it("returns error on API failure", async () => {
      mockTfvcApi.getItems.mockRejectedValue(new Error("TFVC not enabled"));
      const result = await getHandler()({ scopePath: "$/MyProject", recursionLevel: "OneLevel" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("TFVC not enabled");
    });
  });
});
