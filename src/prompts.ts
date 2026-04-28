// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { CORE_TOOLS } from "./tools/core.js";
import { TFVC_TOOLS } from "./tools/tfvc.js";
import { PIPELINE_TOOLS } from "./tools/pipelines.js";

function configurePrompts(server: McpServer) {
  server.prompt("Projects", "Lists all projects in the Azure DevOps organization.", {}, () => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: String.raw`
# Task
Use the '${CORE_TOOLS.list_projects}' tool to retrieve all 'wellFormed' projects in the current Azure DevOps organization.
Present the results in alphabetical order in a table with the following columns: Name and ID.`,
        },
      },
    ],
  }));

  server.prompt(
    "ReviewChangeset",
    "Perform a D365 F&O code review on a TFVC changeset. Retrieves changeset details and all file changes, then analyzes code quality, architecture compliance, security, and performance against Microsoft D365 F&O best practices.",
    {
      changesetId: z.string().describe("The TFVC changeset ID to review."),
      project: z.string().optional().describe("Optional Azure DevOps project name to scope the review."),
    },
    ({ changesetId, project }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: String.raw`
# D365 F&O Code Review — Changeset ${changesetId}

You are an expert Microsoft Dynamics 365 Finance & Operations (D365 F&O) architect and code reviewer.

## Step 1: Retrieve changeset details
Use the '${TFVC_TOOLS.get_changeset}' tool with changesetId=${changesetId}${project ? `, project="${project}"` : ""}, includeWorkItems=true, includeDetails=true.

## Step 2: Retrieve all file changes
Use the '${TFVC_TOOLS.get_changeset_changes}' tool with changesetId=${changesetId} to list all modified files.

## Step 3: Review file contents
For each non-binary file in the changeset, use the '${TFVC_TOOLS.get_item_content}' tool to retrieve the file content at this changeset version.

## Step 4: Produce a structured code review

Analyze the code against D365 F&O best practices and produce a structured review covering:

### Code Quality & Standards
- X++ language standards, naming conventions, coding patterns
- Error handling and exception management
- Proper transaction handling
- Memory leak prevention and resource disposal

### Architecture Compliance
- Extension patterns vs overlayering (Chain of Command, event handlers, post-event handlers)
- Separation of concerns — business logic placement (table vs form vs class)
- Proper use of D365 F&O frameworks (SysOperation, Workflow, Number Sequences, Dimension Framework)

### Security Review
- Role-based access control validation
- Data protection and sensitive data handling
- Record-level security (RLS)
- Audit trail implementations

### Performance Assessment
- Database query optimization (avoiding full table scans)
- Batch processing patterns
- Caching strategies
- RecordSortedList usage for large data sets

### Integration Patterns
- OData/REST API design (if applicable)
- Data Management Framework (DMF) entity patterns
- Integration with external systems

## Feedback Format
Use the following severity categories:
- 🔴 **Critical**: Security vulnerabilities, performance bottlenecks, compliance violations, data integrity risks
- 🟡 **Important**: Best practice violations, maintainability concerns, architecture deviations
- 🟢 **Suggestion**: Optimization opportunities, code improvements, alternative patterns
- 📚 **Educational**: Learning opportunities and D365 F&O best practice explanations

Be specific — reference exact file paths and provide concrete suggestions for each finding.`,
          },
        },
      ],
    })
  );

  server.prompt(
    "BuildHealthCheck",
    "Get an overview of recent build health for a project. Retrieves the last N builds and summarizes pass/fail trends, average duration, and common failure patterns.",
    {
      project: z.string().describe("Azure DevOps project name to check build health for."),
      top: z.string().optional().describe("Number of recent builds to analyze. Defaults to 20."),
    },
    ({ project, top }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: String.raw`
# Build Health Check — Project: ${project}

## Step 1: Retrieve recent builds
Use the '${PIPELINE_TOOLS.pipelines_get_builds}' tool with project="${project}", top=${top ?? 20}, queryOrder="QueueTimeDescending".

## Step 2: Analyze build health

Produce a structured health report covering:

### Summary Table
Present builds in a table with columns: Build ID, Build Number, Definition, Status, Result, Queued Time, Duration.

### Health Metrics
- Total builds analyzed
- Success rate (succeeded / total)
- Failure rate and partial-success rate
- Average build duration
- Trend: is the success rate improving or declining over the period?

### Failure Analysis
- List all failed builds with their build numbers and definitions
- Identify the most frequently failing definitions
- Note any patterns (e.g., failures only on certain branches or times of day)

### Recommendations
- Flag any definitions with a failure rate above 20%
- Suggest investigation areas based on failure patterns
- Recommend if a deeper analysis with '${PIPELINE_TOOLS.pipelines_get_failed_tasks_with_logs}' is warranted`,
          },
        },
      ],
    })
  );

  server.prompt(
    "AnalyzeBuildFailure",
    "Perform a root-cause analysis of a specific failed build. Retrieves all failed tasks with their log output and produces a structured diagnosis with actionable recommendations.",
    {
      project: z.string().describe("Azure DevOps project name containing the build."),
      buildId: z.string().describe("The ID of the failed build to analyze."),
    },
    ({ project, buildId }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: String.raw`
# Build Failure Analysis — Build ${buildId} in Project: ${project}

You are an expert Azure DevOps and D365 F&O build engineer.

## Step 1: Get build details
Use the '${PIPELINE_TOOLS.pipelines_get_build_status}' tool with project="${project}", buildId=${buildId} to get overall build metadata.

## Step 2: Retrieve failed tasks with logs
Use the '${PIPELINE_TOOLS.pipelines_get_failed_tasks_with_logs}' tool with project="${project}", buildId=${buildId}, tailLines=100.

## Step 3: Produce root-cause analysis

### Failed Tasks Summary
For each failed task, provide:
- Task name and type
- Result status
- Issues/errors reported by the task
- Key lines from the log that indicate the root cause

### Root Cause Assessment
- Identify the primary failure cause (compilation error, test failure, deployment error, infrastructure issue, etc.)
- Distinguish between the root cause and downstream failures caused by it
- For D365 F&O builds: identify if the failure relates to X++ compilation, package deployment, database sync, or test execution

### Recommended Actions
- Specific steps to resolve each root cause
- Commands or configuration changes needed
- Links to relevant documentation if applicable
- Indicate if the failure appears to be a flaky/infrastructure issue vs a code problem`,
          },
        },
      ],
    })
  );

  server.prompt(
    "CompareChangesets",
    "Compare two TFVC changesets side-by-side. Shows what files changed in each, highlights overlapping changes, and summarizes the overall impact difference between the two changesets.",
    {
      changesetId1: z.string().describe("The first TFVC changeset ID to compare."),
      changesetId2: z.string().describe("The second TFVC changeset ID to compare."),
      project: z.string().optional().describe("Optional Azure DevOps project name."),
    },
    ({ changesetId1, changesetId2, project }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: String.raw`
# Changeset Comparison — ${changesetId1} vs ${changesetId2}${project ? ` (Project: ${project})` : ""}

## Step 1: Retrieve both changesets
Use the '${TFVC_TOOLS.get_changeset}' tool for each:
- changesetId=${changesetId1}${project ? `, project="${project}"` : ""}, includeWorkItems=true
- changesetId=${changesetId2}${project ? `, project="${project}"` : ""}, includeWorkItems=true

## Step 2: Retrieve file changes for both changesets
Use the '${TFVC_TOOLS.get_changeset_changes}' tool for each:
- changesetId=${changesetId1}
- changesetId=${changesetId2}

## Step 3: For files changed in both changesets, get the diff
For any file that appears in both change lists, use the '${TFVC_TOOLS.get_file_diff}' tool to compare the two versions.

## Step 4: Produce a comparison report

### Changeset Metadata Comparison
| Field | Changeset ${changesetId1} | Changeset ${changesetId2} |
|---|---|---|
| Author | | |
| Date | | |
| Comment | | |
| Linked Work Items | | |

### File Changes Comparison
- Files only in changeset ${changesetId1}
- Files only in changeset ${changesetId2}
- Files changed in both changesets (with diff summary)

### Impact Analysis
- Which changeset has broader impact (more files, more modules)?
- Are there conflicting changes to the same files?
- Are there dependencies between the two changesets?

### D365 F&O Specific Notes
- Do the changes follow extension patterns?
- Are there any breaking changes to public APIs?
- Are data entity or table changes that require data migration?`,
          },
        },
      ],
    })
  );
}

export { configurePrompts };
