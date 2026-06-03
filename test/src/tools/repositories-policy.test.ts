// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it, beforeEach } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { configureRepoTools, REPO_TOOLS } from "../../../src/tools/repositories";
import { getCurrentUserDetails, getUserIdFromEmail } from "../../../src/tools/auth";

// Mock the auth module
jest.mock("../../../src/tools/auth", () => ({
  getCurrentUserDetails: jest.fn(),
  getUserIdFromEmail: jest.fn(),
}));

const mockGetCurrentUserDetails = getCurrentUserDetails as jest.MockedFunction<typeof getCurrentUserDetails>;

describe("repos policy tools", () => {
  let server: McpServer;
  let tokenProvider: jest.MockedFunction<() => Promise<string>>;
  let connectionProvider: jest.MockedFunction<() => Promise<WebApi>>;
  let userAgentProvider: () => string;
  let mockPolicyApi: {
    getPolicyConfigurations: jest.Mock;
    getPolicyEvaluations: jest.Mock;
  };
  let mockCoreApi: {
    getProjects: jest.Mock;
  };
  let mockGitApi: Record<string, jest.Mock>;

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn();
    userAgentProvider = () => "Jest";

    mockPolicyApi = {
      getPolicyConfigurations: jest.fn(),
      getPolicyEvaluations: jest.fn(),
    };

    mockCoreApi = {
      getProjects: jest.fn(),
    };

    // Minimal git API mock to satisfy configureRepoTools
    mockGitApi = {};

    connectionProvider = jest.fn().mockResolvedValue({
      getGitApi: jest.fn().mockResolvedValue(mockGitApi),
      getPolicyApi: jest.fn().mockResolvedValue(mockPolicyApi),
      getCoreApi: jest.fn().mockResolvedValue(mockCoreApi),
    } as unknown as WebApi);

    mockGetCurrentUserDetails.mockResolvedValue({
      authenticatedUser: { id: "user123", uniqueName: "testuser@example.com", displayName: "Test User" },
    } as any);
  });

  describe("repo_list_branch_policies", () => {
    function getHandler() {
      configureRepoTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === REPO_TOOLS.list_branch_policies);
      if (!call) throw new Error("repo_list_branch_policies not registered");
      return call[3];
    }

    const mockConfigurations = [
      {
        id: 1,
        isEnabled: true,
        isBlocking: true,
        type: { id: "fa4e907d-c16b-452d-8106-7efa0cb84489", displayName: "Minimum number of reviewers" },
        settings: {
          minimumApproverCount: 2,
          scope: [{ repositoryId: "repo-guid-123", refName: "refs/heads/main" }],
        },
        createdDate: "2024-01-01T00:00:00Z",
        revision: 1,
      },
      {
        id: 2,
        isEnabled: true,
        isBlocking: false,
        type: { id: "0609b952-1397-4640-95ec-e00a01b2f659", displayName: "Work item linking" },
        settings: {
          scope: [{ repositoryId: "other-repo-guid", refName: "refs/heads/develop" }],
        },
        createdDate: "2024-01-01T00:00:00Z",
        revision: 1,
      },
    ];

    it("returns all policy configurations for a project", async () => {
      mockPolicyApi.getPolicyConfigurations.mockResolvedValue(mockConfigurations);

      const result = await getHandler()({ project: "MyProject" });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].id).toBe(1);
      expect(parsed[0].type.displayName).toBe("Minimum number of reviewers");
      expect(parsed[0].isBlocking).toBe(true);
      expect(result.isError).toBeUndefined();
    });

    it("filters by repositoryId", async () => {
      mockPolicyApi.getPolicyConfigurations.mockResolvedValue(mockConfigurations);

      const result = await getHandler()({ project: "MyProject", repositoryId: "repo-guid-123" });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].id).toBe(1);
    });

    it("filters by branch name", async () => {
      mockPolicyApi.getPolicyConfigurations.mockResolvedValue(mockConfigurations);

      const result = await getHandler()({ project: "MyProject", branch: "develop" });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].id).toBe(2);
    });

    it("filters by branch with full refs/heads/ prefix", async () => {
      mockPolicyApi.getPolicyConfigurations.mockResolvedValue(mockConfigurations);

      const result = await getHandler()({ project: "MyProject", branch: "refs/heads/main" });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].id).toBe(1);
    });

    it("returns 'no policies found' when filtered result is empty", async () => {
      mockPolicyApi.getPolicyConfigurations.mockResolvedValue(mockConfigurations);

      const result = await getHandler()({ project: "MyProject", repositoryId: "nonexistent-repo" });
      expect(result.content[0].text).toContain("No branch policies found");
    });

    it("returns error on API failure", async () => {
      mockPolicyApi.getPolicyConfigurations.mockRejectedValue(new Error("Unauthorized"));

      const result = await getHandler()({ project: "MyProject" });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Unauthorized");
    });
  });

  describe("repo_get_pull_request_policy_evaluations", () => {
    function getHandler() {
      configureRepoTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([name]) => name === REPO_TOOLS.get_pull_request_policy_evaluations);
      if (!call) throw new Error("repo_get_pull_request_policy_evaluations not registered");
      return call[3];
    }

    const mockProjects = [{ id: "project-guid-abc", name: "MyProject" }];

    const mockEvaluations = [
      {
        evaluationId: "eval-1",
        status: 2, // Approved
        configuration: {
          id: 1,
          isEnabled: true,
          isBlocking: true,
          type: { id: "fa4e907d-c16b-452d-8106-7efa0cb84489", displayName: "Minimum number of reviewers" },
        },
        startedDate: "2024-01-15T10:00:00Z",
        completedDate: "2024-01-15T10:01:00Z",
        context: {},
      },
      {
        evaluationId: "eval-2",
        status: 3, // Rejected
        configuration: {
          id: 2,
          isEnabled: true,
          isBlocking: false,
          type: { id: "0609b952-1397-4640-95ec-e00a01b2f659", displayName: "Work item linking" },
        },
        startedDate: "2024-01-15T10:00:00Z",
        completedDate: "2024-01-15T10:00:30Z",
        context: {},
      },
    ];

    it("returns policy evaluations for a pull request", async () => {
      mockCoreApi.getProjects.mockResolvedValue(mockProjects);
      mockPolicyApi.getPolicyEvaluations.mockResolvedValue(mockEvaluations);

      const result = await getHandler()({ project: "MyProject", pullRequestId: 123, includeNotApplicable: false });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].evaluationId).toBe("eval-1");
      expect(parsed[0].configuration.type.displayName).toBe("Minimum number of reviewers");
      expect(result.isError).toBeUndefined();
    });

    it("passes correct artifactId to policy API", async () => {
      mockCoreApi.getProjects.mockResolvedValue(mockProjects);
      mockPolicyApi.getPolicyEvaluations.mockResolvedValue(mockEvaluations);

      await getHandler()({ project: "MyProject", pullRequestId: 123, includeNotApplicable: false });

      expect(mockPolicyApi.getPolicyEvaluations).toHaveBeenCalledWith("MyProject", "vstfs:///CodeReview/CodeReviewId/project-guid-abc/123", false);
    });

    it("returns error when project not found", async () => {
      mockCoreApi.getProjects.mockResolvedValue([]);

      const result = await getHandler()({ project: "NonExistentProject", pullRequestId: 123, includeNotApplicable: false });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("NonExistentProject");
    });

    it("returns 'no evaluations found' when list is empty", async () => {
      mockCoreApi.getProjects.mockResolvedValue(mockProjects);
      mockPolicyApi.getPolicyEvaluations.mockResolvedValue([]);

      const result = await getHandler()({ project: "MyProject", pullRequestId: 123, includeNotApplicable: false });
      expect(result.content[0].text).toContain("No policy evaluations found");
    });

    it("returns error on API failure", async () => {
      mockCoreApi.getProjects.mockResolvedValue(mockProjects);
      mockPolicyApi.getPolicyEvaluations.mockRejectedValue(new Error("Policy API error"));

      const result = await getHandler()({ project: "MyProject", pullRequestId: 123, includeNotApplicable: false });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Policy API error");
    });
  });
});
