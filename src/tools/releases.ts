// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import {
  ReleaseStatus,
  DeploymentStatus,
  DeploymentOperationStatus,
  ReleaseDefinitionQueryOrder,
  ReleaseQueryOrder,
  ReleaseExpands,
  ReleaseDefinitionExpands,
} from "azure-devops-node-api/interfaces/ReleaseInterfaces.js";
import { z } from "zod";
import { getEnumKeys, safeEnumConvert, streamToString } from "../utils.js";
import { createExternalContentResponse } from "../shared/content-safety.js";

const RELEASE_TOOLS = {
  list_definitions: "releases_list_definitions",
  list_releases: "releases_list_releases",
  get_release: "releases_get_release",
  get_environment: "releases_get_environment",
  list_deployments: "releases_list_deployments",
  get_deployment_logs: "releases_get_deployment_logs",
};

function configureReleaseTools(server: McpServer, _tokenProvider: () => Promise<string>, connectionProvider: () => Promise<WebApi>) {
  server.tool(
    RELEASE_TOOLS.list_definitions,
    "List release pipeline definitions for a project. Returns the ID, name, last release information, and description of each definition.",
    {
      project: z.string().describe("Project ID or name to list release definitions for."),
      searchText: z.string().optional().describe("Filter definitions by name (partial match)."),
      top: z.coerce.number().min(1).max(1000).default(50).describe("Maximum number of definitions to return. Defaults to 50."),
      continuationToken: z.string().optional().describe("Continuation token for pagination from a previous response."),
      queryOrder: z
        .enum(getEnumKeys(ReleaseDefinitionQueryOrder) as [string, ...string[]])
        .optional()
        .describe("Sort order for the results."),
      expand: z
        .enum(getEnumKeys(ReleaseDefinitionExpands) as [string, ...string[]])
        .optional()
        .describe("Properties to expand in the response (e.g. 'lastRelease', 'artifacts')."),
    },
    async ({ project, searchText, top, continuationToken, queryOrder, expand }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const definitions = await releaseApi.getReleaseDefinitions(
          project,
          searchText,
          safeEnumConvert(ReleaseDefinitionExpands, expand),
          undefined,
          undefined,
          top,
          continuationToken,
          safeEnumConvert(ReleaseDefinitionQueryOrder, queryOrder)
        );

        if (!definitions || (definitions as unknown[]).length === 0) {
          return { content: [{ type: "text", text: "No release definitions found." }] };
        }

        const trimmed = (definitions as { id?: number; name?: string; description?: string; path?: string; lastRelease?: { id?: number; name?: string; createdOn?: string } }[]).map((def) => ({
          id: def.id,
          name: def.name,
          description: def.description,
          path: def.path,
          lastRelease: def.lastRelease
            ? {
                id: def.lastRelease.id,
                name: def.lastRelease.name,
                createdOn: def.lastRelease.createdOn,
              }
            : undefined,
        }));

        return { content: [{ type: "text", text: JSON.stringify(trimmed, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error listing release definitions: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.list_releases,
    "List releases for a project. Supports filtering by definition, status, creation date, and search text. Returns ID, name, status, and environment summary for each release.",
    {
      project: z.string().describe("Project ID or name to list releases for."),
      definitionId: z.coerce.number().min(1).optional().describe("Filter releases by release definition ID."),
      statusFilter: z
        .enum(getEnumKeys(ReleaseStatus) as [string, ...string[]])
        .optional()
        .describe("Filter releases by status (e.g. 'Active', 'Abandoned', 'Draft', 'Undefined')."),
      searchText: z.string().optional().describe("Filter releases by name (partial match)."),
      createdBy: z.string().optional().describe("Filter releases created by this user (unique name or display name)."),
      minCreatedTime: z.coerce.date().optional().describe("Return releases created on or after this date."),
      maxCreatedTime: z.coerce.date().optional().describe("Return releases created on or before this date."),
      top: z.coerce.number().min(1).max(1000).default(50).describe("Maximum number of releases to return. Defaults to 50."),
      queryOrder: z
        .enum(getEnumKeys(ReleaseQueryOrder) as [string, ...string[]])
        .optional()
        .describe("Sort order for the results."),
      expand: z
        .enum(getEnumKeys(ReleaseExpands) as [string, ...string[]])
        .optional()
        .describe("Properties to expand in the response (e.g. 'environments', 'artifacts')."),
    },
    async ({ project, definitionId, statusFilter, searchText, createdBy, minCreatedTime, maxCreatedTime, top, queryOrder, expand }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const releases = await releaseApi.getReleases(
          project,
          definitionId,
          undefined,
          searchText,
          createdBy,
          safeEnumConvert(ReleaseStatus, statusFilter),
          undefined,
          minCreatedTime,
          maxCreatedTime,
          safeEnumConvert(ReleaseQueryOrder, queryOrder),
          top,
          undefined,
          safeEnumConvert(ReleaseExpands, expand)
        );

        if (!releases || (releases as unknown[]).length === 0) {
          return { content: [{ type: "text", text: "No releases found." }] };
        }

        const trimmed = (
          releases as {
            id?: number;
            name?: string;
            status?: number;
            createdOn?: string;
            createdBy?: { displayName?: string; uniqueName?: string };
            releaseDefinition?: { id?: number; name?: string };
            environments?: { id?: number; name?: string; status?: number }[];
          }[]
        ).map((rel) => ({
          id: rel.id,
          name: rel.name,
          status: rel.status !== undefined ? ReleaseStatus[rel.status] : undefined,
          createdOn: rel.createdOn,
          createdBy: {
            displayName: rel.createdBy?.displayName,
            uniqueName: rel.createdBy?.uniqueName,
          },
          releaseDefinition: rel.releaseDefinition
            ? {
                id: rel.releaseDefinition.id,
                name: rel.releaseDefinition.name,
              }
            : undefined,
          environments: rel.environments?.map((env) => ({
            id: env.id,
            name: env.name,
            status: env.status,
          })),
        }));

        return { content: [{ type: "text", text: JSON.stringify(trimmed, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error listing releases: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.get_release,
    "Get full details of a specific release including all environment stages, artifacts, triggers, approvals, and deployment attempts.",
    {
      project: z.string().describe("Project ID or name where the release is located."),
      releaseId: z.coerce.number().min(1).describe("The ID of the release to retrieve."),
    },
    async ({ project, releaseId }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const release = await releaseApi.getRelease(project, releaseId);

        if (!release) {
          return { content: [{ type: "text", text: `Release ${releaseId} not found.` }], isError: true };
        }

        return { content: [{ type: "text", text: JSON.stringify(release, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error getting release ${releaseId}: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.get_environment,
    "Get details of a specific environment (stage) within a release, including its current deployment status, approval state, conditions, and deployment attempts.",
    {
      project: z.string().describe("Project ID or name where the release is located."),
      releaseId: z.coerce.number().min(1).describe("The ID of the release."),
      environmentId: z.coerce.number().min(1).describe("The ID of the release environment (stage) to retrieve."),
    },
    async ({ project, releaseId, environmentId }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const environment = await releaseApi.getReleaseEnvironment(project, releaseId, environmentId);

        if (!environment) {
          return { content: [{ type: "text", text: `Environment ${environmentId} not found in release ${releaseId}.` }], isError: true };
        }

        return { content: [{ type: "text", text: JSON.stringify(environment, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error getting environment ${environmentId} in release ${releaseId}: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.list_deployments,
    "List deployments across release definitions in a project. Supports filtering by definition, environment, status, and date range. Useful for tracking deployment health across all pipelines.",
    {
      project: z.string().describe("Project ID or name to list deployments for."),
      definitionId: z.coerce.number().min(1).optional().describe("Filter deployments by release definition ID."),
      definitionEnvironmentId: z.coerce.number().min(1).optional().describe("Filter deployments by release definition environment ID."),
      createdBy: z.string().optional().describe("Filter deployments triggered by this user (unique name)."),
      deploymentStatus: z
        .enum(getEnumKeys(DeploymentStatus) as [string, ...string[]])
        .optional()
        .describe("Filter by deployment status (e.g. 'Succeeded', 'Failed', 'InProgress', 'NotDeployed')."),
      operationStatus: z
        .enum(getEnumKeys(DeploymentOperationStatus) as [string, ...string[]])
        .optional()
        .describe("Filter by deployment operation status (e.g. 'Approved', 'Rejected', 'Pending')."),
      minModifiedTime: z.coerce.date().optional().describe("Return deployments modified on or after this date."),
      maxModifiedTime: z.coerce.date().optional().describe("Return deployments modified on or before this date."),
      top: z.coerce.number().min(1).max(1000).default(50).describe("Maximum number of deployments to return. Defaults to 50."),
      queryOrder: z
        .enum(getEnumKeys(ReleaseQueryOrder) as [string, ...string[]])
        .optional()
        .describe("Sort order for the results."),
      latestAttemptsOnly: z.boolean().optional().default(true).describe("Return only the latest deployment attempt for each environment. Defaults to true."),
    },
    async ({ project, definitionId, definitionEnvironmentId, createdBy, deploymentStatus, operationStatus, minModifiedTime, maxModifiedTime, top, queryOrder, latestAttemptsOnly }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const deployments = await releaseApi.getDeployments(
          project,
          definitionId,
          definitionEnvironmentId,
          createdBy,
          minModifiedTime,
          maxModifiedTime,
          safeEnumConvert(DeploymentStatus, deploymentStatus),
          safeEnumConvert(DeploymentOperationStatus, operationStatus),
          latestAttemptsOnly,
          safeEnumConvert(ReleaseQueryOrder, queryOrder),
          top
        );

        if (!deployments || (deployments as unknown[]).length === 0) {
          return { content: [{ type: "text", text: "No deployments found." }] };
        }

        const trimmed = (
          deployments as {
            id?: number;
            release?: { id?: number; name?: string };
            releaseDefinition?: { id?: number; name?: string };
            releaseEnvironment?: { id?: number; name?: string };
            deploymentStatus?: number;
            operationStatus?: number;
            requestedBy?: { displayName?: string; uniqueName?: string };
            queuedOn?: string;
            startedOn?: string;
            completedOn?: string;
          }[]
        ).map((dep) => ({
          id: dep.id,
          release: dep.release ? { id: dep.release.id, name: dep.release.name } : undefined,
          releaseDefinition: dep.releaseDefinition ? { id: dep.releaseDefinition.id, name: dep.releaseDefinition.name } : undefined,
          releaseEnvironment: dep.releaseEnvironment ? { id: dep.releaseEnvironment.id, name: dep.releaseEnvironment.name } : undefined,
          deploymentStatus: dep.deploymentStatus !== undefined ? DeploymentStatus[dep.deploymentStatus] : undefined,
          operationStatus: dep.operationStatus !== undefined ? DeploymentOperationStatus[dep.operationStatus] : undefined,
          requestedBy: {
            displayName: dep.requestedBy?.displayName,
            uniqueName: dep.requestedBy?.uniqueName,
          },
          queuedOn: dep.queuedOn,
          startedOn: dep.startedOn,
          completedOn: dep.completedOn,
        }));

        return { content: [{ type: "text", text: JSON.stringify(trimmed, null, 2) }] };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return { content: [{ type: "text", text: `Error listing deployments: ${errorMessage}` }], isError: true };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.get_deployment_logs,
    "Get task-level logs for a specific release environment (stage). Returns the log content for each task that ran during deployment, including failed tasks. Use this to diagnose failed deployments.",
    {
      project: z.string().describe("Project ID or name where the release is located."),
      releaseId: z.coerce.number().min(1).describe("The ID of the release."),
      environmentId: z.coerce.number().min(1).describe("The ID of the release environment (stage) to get logs for."),
      releaseDeployPhaseId: z.coerce.number().min(1).optional().describe("Optional deploy phase ID to scope logs to a specific deploy phase. If not provided, returns logs for all phases."),
      taskId: z.coerce.number().min(1).optional().describe("Optional task ID to scope logs to a specific task within a phase."),
    },
    async ({ project, releaseId, environmentId, releaseDeployPhaseId, taskId }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        // If specific phase and task are provided, get the targeted log
        if (releaseDeployPhaseId !== undefined && taskId !== undefined) {
          const stream = await releaseApi.getTaskLog(project, releaseId, environmentId, releaseDeployPhaseId, taskId);
          const content = await streamToString(stream);
          return createExternalContentResponse(content, `release ${releaseId} environment ${environmentId} task ${taskId} log`);
        }

        // Otherwise return the full environment logs as a zip stream text
        const stream = await releaseApi.getLogs(project, releaseId);
        const content = await streamToString(stream);
        return createExternalContentResponse(content, `release ${releaseId} logs`);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error getting deployment logs for release ${releaseId} environment ${environmentId}: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );
}

export { RELEASE_TOOLS, configureReleaseTools };
