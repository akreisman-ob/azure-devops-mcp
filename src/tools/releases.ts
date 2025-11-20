// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { AuthToken } from "./authToken.js";
import { z } from "zod";
import { getEnumKeys, safeEnumConvert } from "../utils.js";
import {
  ReleaseStatus,
  ReleaseExpands,
  ReleaseQueryOrder,
  ReleaseDefinitionExpands,
  ReleaseDefinitionQueryOrder,
  SingleReleaseExpands,
  ApprovalFilters,
} from "azure-devops-node-api/interfaces/ReleaseInterfaces.js";

const RELEASE_TOOLS = {
  releases_list_releases: "releases_list_releases",
  releases_get_release: "releases_get_release",
  releases_list_definitions: "releases_list_definitions",
  releases_get_definition: "releases_get_definition",
  releases_add_tag: "releases_add_tag",
  releases_remove_tag: "releases_remove_tag",
  releases_list_tags: "releases_list_tags",
  releases_get_release_tags: "releases_get_release_tags",
  releases_create_release: "releases_create_release",
};

function configureReleaseTools(
  server: McpServer,
  _: () => Promise<AuthToken>,
  connectionProvider: () => Promise<WebApi>
) {
  server.tool(
    RELEASE_TOOLS.releases_list_releases,
    "Retrieves a list of releases for a given project. Supports filtering by definition, environment, status, tags, and more.",
    {
      project: z.string().describe("Project ID or name to get releases for"),
      definitionId: z.number().optional().describe("Release definition ID to filter releases"),
      definitionEnvironmentId: z
        .number()
        .optional()
        .describe("Definition environment ID to filter releases"),
      searchText: z.string().optional().describe("Search text to filter releases"),
      createdBy: z.string().optional().describe("User ID or name who created the release"),
      statusFilter: z
        .enum(getEnumKeys(ReleaseStatus) as [string, ...string[]])
        .optional()
        .describe("Status filter for releases (e.g., Active, Abandoned)"),
      environmentStatusFilter: z
        .number()
        .optional()
        .describe("Environment status filter for releases"),
      minCreatedTime: z.coerce
        .date()
        .optional()
        .describe("Minimum created time to filter releases"),
      maxCreatedTime: z.coerce
        .date()
        .optional()
        .describe("Maximum created time to filter releases"),
      queryOrder: z
        .enum(getEnumKeys(ReleaseQueryOrder) as [string, ...string[]])
        .optional()
        .describe("Order in which releases are returned (e.g., Descending, Ascending)"),
      top: z.number().optional().describe("Maximum number of releases to return"),
      continuationToken: z
        .number()
        .optional()
        .describe("Token for continuing paged results"),
      expand: z
        .enum(getEnumKeys(ReleaseExpands) as [string, ...string[]])
        .optional()
        .describe(
          "Expand options for releases (e.g., Environments, Artifacts, Approvals)"
        ),
      artifactTypeId: z.string().optional().describe("Artifact type ID to filter releases"),
      sourceId: z.string().optional().describe("Source ID to filter releases"),
      artifactVersionId: z
        .string()
        .optional()
        .describe("Artifact version ID to filter releases"),
      sourceBranchFilter: z
        .string()
        .optional()
        .describe("Source branch to filter releases"),
      isDeleted: z
        .boolean()
        .optional()
        .describe("Whether to include deleted releases, defaults to false"),
      tagFilter: z
        .array(z.string())
        .optional()
        .describe("Array of tags to filter releases"),
      propertyFilters: z
        .array(z.string())
        .optional()
        .describe("Array of property filters"),
      releaseIdFilter: z
        .array(z.number())
        .optional()
        .describe("Array of release IDs to filter"),
      path: z.string().optional().describe("Path to filter releases"),
    },
    async ({
      project,
      definitionId,
      definitionEnvironmentId,
      searchText,
      createdBy,
      statusFilter,
      environmentStatusFilter,
      minCreatedTime,
      maxCreatedTime,
      queryOrder,
      top,
      continuationToken,
      expand,
      artifactTypeId,
      sourceId,
      artifactVersionId,
      sourceBranchFilter,
      isDeleted,
      tagFilter,
      propertyFilters,
      releaseIdFilter,
      path,
    }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const releases = await releaseApi.getReleases(
          project,
          definitionId,
          definitionEnvironmentId,
          searchText,
          createdBy,
          safeEnumConvert(ReleaseStatus, statusFilter),
          environmentStatusFilter,
          minCreatedTime,
          maxCreatedTime,
          safeEnumConvert(ReleaseQueryOrder, queryOrder),
          top,
          continuationToken,
          safeEnumConvert(ReleaseExpands, expand),
          artifactTypeId,
          sourceId,
          artifactVersionId,
          sourceBranchFilter,
          isDeleted,
          tagFilter,
          propertyFilters,
          releaseIdFilter,
          path
        );

        return {
          content: [{ type: "text", text: JSON.stringify(releases, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error fetching releases: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_get_release,
    "Retrieves a specific release by ID with optional filters and expansions.",
    {
      project: z.string().describe("Project ID or name"),
      releaseId: z.number().describe("ID of the release to retrieve"),
      approvalFilters: z
        .enum(getEnumKeys(ApprovalFilters) as [string, ...string[]])
        .optional()
        .describe("Approval filters to apply"),
      propertyFilters: z
        .array(z.string())
        .optional()
        .describe("Array of property filters"),
      expand: z
        .enum(getEnumKeys(SingleReleaseExpands) as [string, ...string[]])
        .optional()
        .describe("Expand options for the release"),
      topGateRecords: z
        .number()
        .optional()
        .describe("Number of top gate records to include"),
    },
    async ({
      project,
      releaseId,
      approvalFilters,
      propertyFilters,
      expand,
      topGateRecords,
    }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const release = await releaseApi.getRelease(
          project,
          releaseId,
          safeEnumConvert(ApprovalFilters, approvalFilters),
          propertyFilters,
          safeEnumConvert(SingleReleaseExpands, expand),
          topGateRecords
        );

        return {
          content: [{ type: "text", text: JSON.stringify(release, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error fetching release: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_list_definitions,
    "Retrieves a list of release definitions for a given project.",
    {
      project: z.string().describe("Project ID or name to get release definitions for"),
      searchText: z
        .string()
        .optional()
        .describe("Search text to filter release definitions"),
      expand: z
        .enum(getEnumKeys(ReleaseDefinitionExpands) as [string, ...string[]])
        .optional()
        .describe("Expand options for release definitions"),
      artifactType: z
        .string()
        .optional()
        .describe("Artifact type to filter release definitions"),
      artifactSourceId: z
        .string()
        .optional()
        .describe("Artifact source ID to filter release definitions"),
      top: z
        .number()
        .optional()
        .describe("Maximum number of release definitions to return"),
      continuationToken: z
        .string()
        .optional()
        .describe("Token for continuing paged results"),
      queryOrder: z
        .enum(getEnumKeys(ReleaseDefinitionQueryOrder) as [string, ...string[]])
        .optional()
        .describe("Order in which release definitions are returned"),
      path: z.string().optional().describe("Path to filter release definitions"),
      isExactNameMatch: z
        .boolean()
        .optional()
        .describe("Whether to perform exact name match, defaults to false"),
      tagFilter: z
        .array(z.string())
        .optional()
        .describe("Array of tags to filter release definitions"),
      propertyFilters: z
        .array(z.string())
        .optional()
        .describe("Array of property filters"),
      definitionIdFilter: z
        .array(z.string())
        .optional()
        .describe("Array of definition IDs to filter"),
      isDeleted: z
        .boolean()
        .optional()
        .describe("Whether to include deleted definitions, defaults to false"),
      searchTextContainsFolderName: z
        .boolean()
        .optional()
        .describe("Whether search text contains folder name, defaults to false"),
    },
    async ({
      project,
      searchText,
      expand,
      artifactType,
      artifactSourceId,
      top,
      continuationToken,
      queryOrder,
      path,
      isExactNameMatch,
      tagFilter,
      propertyFilters,
      definitionIdFilter,
      isDeleted,
      searchTextContainsFolderName,
    }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const definitions = await releaseApi.getReleaseDefinitions(
          project,
          searchText,
          safeEnumConvert(ReleaseDefinitionExpands, expand),
          artifactType,
          artifactSourceId,
          top,
          continuationToken,
          safeEnumConvert(ReleaseDefinitionQueryOrder, queryOrder),
          path,
          isExactNameMatch,
          tagFilter,
          propertyFilters,
          definitionIdFilter,
          isDeleted,
          searchTextContainsFolderName
        );

        return {
          content: [{ type: "text", text: JSON.stringify(definitions, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            { type: "text", text: `Error fetching release definitions: ${errorMessage}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_get_definition,
    "Retrieves a specific release definition by ID.",
    {
      project: z.string().describe("Project ID or name"),
      definitionId: z.number().describe("ID of the release definition to retrieve"),
      propertyFilters: z
        .array(z.string())
        .optional()
        .describe("Array of property filters"),
    },
    async ({ project, definitionId, propertyFilters }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const definition = await releaseApi.getReleaseDefinition(
          project,
          definitionId,
          propertyFilters
        );

        return {
          content: [{ type: "text", text: JSON.stringify(definition, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            { type: "text", text: `Error fetching release definition: ${errorMessage}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_add_tag,
    "Adds a tag to a specific release.",
    {
      project: z.string().describe("Project ID or name"),
      releaseId: z.number().describe("ID of the release to add tag to"),
      tag: z.string().describe("Tag to add to the release"),
    },
    async ({ project, releaseId, tag }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const tags = await releaseApi.addReleaseTag(project, releaseId, tag);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ releaseId, addedTag: tag, allTags: tags }, null, 2),
            },
          ],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error adding tag to release: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_remove_tag,
    "Removes a tag from a specific release.",
    {
      project: z.string().describe("Project ID or name"),
      releaseId: z.number().describe("ID of the release to remove tag from"),
      tag: z.string().describe("Tag to remove from the release"),
    },
    async ({ project, releaseId, tag }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const tags = await releaseApi.deleteReleaseTag(project, releaseId, tag);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { releaseId, removedTag: tag, remainingTags: tags },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            { type: "text", text: `Error removing tag from release: ${errorMessage}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_list_tags,
    "Retrieves all tags for a given project.",
    {
      project: z.string().describe("Project ID or name to get release tags for"),
    },
    async ({ project }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const tags = await releaseApi.getTags(project);

        return {
          content: [{ type: "text", text: JSON.stringify(tags, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error fetching tags: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_get_release_tags,
    "Retrieves tags for a specific release.",
    {
      project: z.string().describe("Project ID or name"),
      releaseId: z.number().describe("ID of the release to get tags for"),
    },
    async ({ project, releaseId }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const tags = await releaseApi.getReleaseTags(project, releaseId);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ releaseId, tags }, null, 2),
            },
          ],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            { type: "text", text: `Error fetching release tags: ${errorMessage}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    RELEASE_TOOLS.releases_create_release,
    "Creates a new release from a release definition.",
    {
      project: z.string().describe("Project ID or name"),
      definitionId: z.number().describe("ID of the release definition to create release from"),
      description: z.string().optional().describe("Description for the release"),
      artifacts: z
        .array(
          z.object({
            alias: z.string().describe("Alias of the artifact"),
            instanceReference: z.object({
              id: z.string().describe("ID of the artifact version"),
              name: z.string().optional().describe("Name of the artifact version"),
            }),
          })
        )
        .optional()
        .describe("Array of artifacts to include in the release"),
      isDraft: z.boolean().optional().describe("Whether the release is a draft"),
      reason: z.string().optional().describe("Reason for creating the release"),
      manualEnvironments: z
        .array(z.string())
        .optional()
        .describe("Array of environment names to deploy manually"),
    },
    async ({
      project,
      definitionId,
      description,
      artifacts,
      isDraft,
      reason,
      manualEnvironments,
    }) => {
      try {
        const connection = await connectionProvider();
        const releaseApi = await connection.getReleaseApi();

        const releaseStartMetadata: any = {
          definitionId,
          description,
          isDraft,
          reason,
          manualEnvironments,
        };

        if (artifacts) {
          releaseStartMetadata.artifacts = artifacts;
        }

        const release = await releaseApi.createRelease(releaseStartMetadata, project);

        return {
          content: [{ type: "text", text: JSON.stringify(release, null, 2) }],
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error creating release: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );
}

export { configureReleaseTools };

