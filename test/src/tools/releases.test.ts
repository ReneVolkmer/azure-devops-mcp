// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it, beforeEach } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { configureReleaseTools, RELEASE_TOOLS } from "../../../src/tools/releases";
import { Readable } from "stream";

function makeReadable(text: string): Readable {
  const r = new Readable();
  r.push(text);
  r.push(null);
  return r;
}

describe("configureReleaseTools", () => {
  let server: McpServer;
  let tokenProvider: jest.MockedFunction<() => Promise<string>>;
  let connectionProvider: jest.MockedFunction<() => Promise<WebApi>>;
  let mockReleaseApi: {
    getReleaseDefinitions: jest.Mock;
    getReleases: jest.Mock;
    getRelease: jest.Mock;
    getReleaseEnvironment: jest.Mock;
    getDeployments: jest.Mock;
    getTaskLog: jest.Mock;
    getLogs: jest.Mock;
  };

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn();

    mockReleaseApi = {
      getReleaseDefinitions: jest.fn(),
      getReleases: jest.fn(),
      getRelease: jest.fn(),
      getReleaseEnvironment: jest.fn(),
      getDeployments: jest.fn(),
      getTaskLog: jest.fn(),
      getLogs: jest.fn(),
    };

    connectionProvider = jest.fn().mockResolvedValue({
      getReleaseApi: jest.fn().mockResolvedValue(mockReleaseApi),
    } as unknown as WebApi);
  });

  describe("tool registration", () => {
    it("registers all release tools", () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const registeredNames = (server.tool as jest.Mock).mock.calls.map(([name]) => name);
      expect(registeredNames).toContain(RELEASE_TOOLS.list_definitions);
      expect(registeredNames).toContain(RELEASE_TOOLS.list_releases);
      expect(registeredNames).toContain(RELEASE_TOOLS.get_release);
      expect(registeredNames).toContain(RELEASE_TOOLS.get_environment);
      expect(registeredNames).toContain(RELEASE_TOOLS.list_deployments);
      expect(registeredNames).toContain(RELEASE_TOOLS.get_deployment_logs);
    });
  });

  describe("releases_list_definitions", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.list_definitions);
      if (!call) throw new Error("releases_list_definitions not registered");
      return call[3];
    }

    it("returns trimmed definition list", async () => {
      mockReleaseApi.getReleaseDefinitions.mockResolvedValue([
        {
          id: 1,
          name: "Production Release",
          description: "Deploy to production",
          path: "\\",
          lastRelease: { id: 42, name: "Release-42", createdOn: "2024-01-15T10:00:00Z" },
        },
        { id: 2, name: "Staging Release", description: "Deploy to staging", path: "\\", lastRelease: undefined },
      ]);

      const result = await getHandler()({ project: "MyProject", top: 50 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].id).toBe(1);
      expect(parsed[0].name).toBe("Production Release");
      expect(parsed[0].lastRelease.id).toBe(42);
      expect(result.isError).toBeUndefined();
    });

    it("returns 'no definitions found' when list is empty", async () => {
      mockReleaseApi.getReleaseDefinitions.mockResolvedValue([]);
      const result = await getHandler()({ project: "MyProject", top: 50 });
      expect(result.content[0].text).toContain("No release definitions found");
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getReleaseDefinitions.mockRejectedValue(new Error("Service unavailable"));
      const result = await getHandler()({ project: "MyProject", top: 50 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Service unavailable");
    });
  });

  describe("releases_list_releases", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.list_releases);
      if (!call) throw new Error("releases_list_releases not registered");
      return call[3];
    }

    it("returns trimmed release list", async () => {
      mockReleaseApi.getReleases.mockResolvedValue([
        {
          id: 42,
          name: "Release-42",
          status: 1, // Active
          createdOn: "2024-01-15T10:00:00Z",
          createdBy: { displayName: "Jane Dev", uniqueName: "jane@example.com" },
          releaseDefinition: { id: 1, name: "Production Release" },
          environments: [{ id: 10, name: "Production", status: 4 }],
        },
      ]);

      const result = await getHandler()({ project: "MyProject", top: 50 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].id).toBe(42);
      expect(parsed[0].name).toBe("Release-42");
      expect(parsed[0].releaseDefinition.id).toBe(1);
      expect(parsed[0].environments).toHaveLength(1);
      expect(result.isError).toBeUndefined();
    });

    it("returns 'no releases found' when list is empty", async () => {
      mockReleaseApi.getReleases.mockResolvedValue([]);
      const result = await getHandler()({ project: "MyProject", top: 50 });
      expect(result.content[0].text).toContain("No releases found");
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getReleases.mockRejectedValue(new Error("Unauthorized"));
      const result = await getHandler()({ project: "MyProject", top: 50 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Unauthorized");
    });
  });

  describe("releases_get_release", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.get_release);
      if (!call) throw new Error("releases_get_release not registered");
      return call[3];
    }

    it("returns full release details", async () => {
      const mockRelease = {
        id: 42,
        name: "Release-42",
        status: 1,
        environments: [{ id: 10, name: "Production", status: 4 }],
      };
      mockReleaseApi.getRelease.mockResolvedValue(mockRelease);

      const result = await getHandler()({ project: "MyProject", releaseId: 42 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe(42);
      expect(parsed.environments).toHaveLength(1);
      expect(result.isError).toBeUndefined();
    });

    it("returns error when release not found", async () => {
      mockReleaseApi.getRelease.mockResolvedValue(null);
      const result = await getHandler()({ project: "MyProject", releaseId: 9999 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("9999");
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getRelease.mockRejectedValue(new Error("Not found"));
      const result = await getHandler()({ project: "MyProject", releaseId: 42 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Not found");
    });
  });

  describe("releases_get_environment", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.get_environment);
      if (!call) throw new Error("releases_get_environment not registered");
      return call[3];
    }

    it("returns environment details", async () => {
      const mockEnv = { id: 10, name: "Production", status: 4, deploySteps: [] };
      mockReleaseApi.getReleaseEnvironment.mockResolvedValue(mockEnv);

      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 10 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe(10);
      expect(parsed.name).toBe("Production");
      expect(result.isError).toBeUndefined();
    });

    it("returns error when environment not found", async () => {
      mockReleaseApi.getReleaseEnvironment.mockResolvedValue(null);
      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 99 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("99");
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getReleaseEnvironment.mockRejectedValue(new Error("Permission denied"));
      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 10 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Permission denied");
    });
  });

  describe("releases_list_deployments", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.list_deployments);
      if (!call) throw new Error("releases_list_deployments not registered");
      return call[3];
    }

    it("returns trimmed deployment list", async () => {
      mockReleaseApi.getDeployments.mockResolvedValue([
        {
          id: 1,
          release: { id: 42, name: "Release-42" },
          releaseDefinition: { id: 1, name: "Production Release" },
          releaseEnvironment: { id: 10, name: "Production" },
          deploymentStatus: 4, // Succeeded
          operationStatus: 128, // Approved
          requestedBy: { displayName: "Jane Dev", uniqueName: "jane@example.com" },
          queuedOn: "2024-01-15T10:00:00Z",
          startedOn: "2024-01-15T10:01:00Z",
          completedOn: "2024-01-15T10:15:00Z",
        },
      ]);

      const result = await getHandler()({ project: "MyProject", top: 50, latestAttemptsOnly: true });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].id).toBe(1);
      expect(parsed[0].release.name).toBe("Release-42");
      expect(parsed[0].requestedBy.displayName).toBe("Jane Dev");
      expect(result.isError).toBeUndefined();
    });

    it("returns 'no deployments found' when list is empty", async () => {
      mockReleaseApi.getDeployments.mockResolvedValue([]);
      const result = await getHandler()({ project: "MyProject", top: 50, latestAttemptsOnly: true });
      expect(result.content[0].text).toContain("No deployments found");
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getDeployments.mockRejectedValue(new Error("Timeout"));
      const result = await getHandler()({ project: "MyProject", top: 50, latestAttemptsOnly: true });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Timeout");
    });
  });

  describe("releases_get_deployment_logs", () => {
    function getHandler() {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === RELEASE_TOOLS.get_deployment_logs);
      if (!call) throw new Error("releases_get_deployment_logs not registered");
      return call[3];
    }

    it("returns task log when phaseId and taskId provided", async () => {
      const logContent = "Task started\nTask completed successfully\n";
      mockReleaseApi.getTaskLog.mockResolvedValue(makeReadable(logContent));

      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 10, releaseDeployPhaseId: 1, taskId: 2 });
      expect(result.content[0].text).toContain(logContent);
      expect(mockReleaseApi.getTaskLog).toHaveBeenCalledWith("MyProject", 42, 10, 1, 2);
      expect(mockReleaseApi.getLogs).not.toHaveBeenCalled();
    });

    it("returns all logs when only releaseId and environmentId provided", async () => {
      const logContent = "Full release log content\n";
      mockReleaseApi.getLogs.mockResolvedValue(makeReadable(logContent));

      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 10 });
      expect(result.content[0].text).toContain(logContent);
      expect(mockReleaseApi.getLogs).toHaveBeenCalledWith("MyProject", 42);
      expect(mockReleaseApi.getTaskLog).not.toHaveBeenCalled();
    });

    it("returns error on API failure", async () => {
      mockReleaseApi.getLogs.mockRejectedValue(new Error("Logs not available"));
      const result = await getHandler()({ project: "MyProject", releaseId: 42, environmentId: 10 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Logs not available");
    });
  });
});
