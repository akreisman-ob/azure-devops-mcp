// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { AuthToken } from "../../../src/tools/authToken";
import { configureReleaseTools } from "../../../src/tools/releases";

type TokenProviderMock = () => Promise<AuthToken>;
type ConnectionProviderMock = () => Promise<WebApi>;

describe("configureReleaseTools", () => {
  let server: McpServer;
  let tokenProvider: TokenProviderMock;
  let connectionProvider: ConnectionProviderMock;
  let mockConnection: { getReleaseApi: jest.Mock };

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn();
    mockConnection = {
      getReleaseApi: jest.fn(),
    };
    connectionProvider = jest.fn().mockResolvedValue(mockConnection);
  });

  describe("tool registration", () => {
    it("registers release tools on the server", () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      expect(server.tool as jest.Mock).toHaveBeenCalled();

      const toolNames = (server.tool as jest.Mock).mock.calls.map(
        (call) => call[0]
      );

      expect(toolNames).toContain("releases_list_releases");
      expect(toolNames).toContain("releases_get_release");
      expect(toolNames).toContain("releases_list_definitions");
      expect(toolNames).toContain("releases_get_definition");
      expect(toolNames).toContain("releases_add_tag");
      expect(toolNames).toContain("releases_remove_tag");
      expect(toolNames).toContain("releases_list_tags");
      expect(toolNames).toContain("releases_get_release_tags");
      expect(toolNames).toContain("releases_create_release");
    });
  });

  describe("releases_list_releases tool", () => {
    it("should list releases with minimal parameters", async () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(
        ([toolName]) => toolName === "releases_list_releases"
      );
      if (!call) throw new Error("releases_list_releases tool not registered");
      const [, , , handler] = call;

      const mockReleases = [
        { id: 1, name: "Release-1" },
        { id: 2, name: "Release-2" },
      ];

      const mockReleaseApi = {
        getReleases: jest.fn().mockResolvedValue(mockReleases),
      };
      mockConnection.getReleaseApi.mockResolvedValue(mockReleaseApi);

      const params = {
        project: "test-project",
      };

      const result = await handler(params);

      expect(mockConnection.getReleaseApi).toHaveBeenCalled();
      expect(mockReleaseApi.getReleases).toHaveBeenCalled();
      expect(result.content[0].text).toBe(JSON.stringify(mockReleases, null, 2));
      expect(result.isError).toBeUndefined();
    });

    it("should handle errors when listing releases", async () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(
        ([toolName]) => toolName === "releases_list_releases"
      );
      if (!call) throw new Error("releases_list_releases tool not registered");
      const [, , , handler] = call;

      const mockReleaseApi = {
        getReleases: jest.fn().mockRejectedValue(new Error("API error")),
      };
      mockConnection.getReleaseApi.mockResolvedValue(mockReleaseApi);

      const params = {
        project: "test-project",
      };

      const result = await handler(params);

      expect(result.content[0].text).toBe("Error fetching releases: API error");
      expect(result.isError).toBe(true);
    });
  });

  describe("releases_add_tag tool", () => {
    it("should add a tag to a release", async () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(
        ([toolName]) => toolName === "releases_add_tag"
      );
      if (!call) throw new Error("releases_add_tag tool not registered");
      const [, , , handler] = call;

      const mockTags = ["tag1", "tag2", "new-tag"];

      const mockReleaseApi = {
        addReleaseTag: jest.fn().mockResolvedValue(mockTags),
      };
      mockConnection.getReleaseApi.mockResolvedValue(mockReleaseApi);

      const params = {
        project: "test-project",
        releaseId: 123,
        tag: "new-tag",
      };

      const result = await handler(params);

      expect(mockConnection.getReleaseApi).toHaveBeenCalled();
      expect(mockReleaseApi.addReleaseTag).toHaveBeenCalledWith(
        "test-project",
        123,
        "new-tag"
      );
      expect(result.content[0].text).toContain("new-tag");
      expect(result.isError).toBeUndefined();
    });
  });

  describe("releases_list_tags tool", () => {
    it("should list all tags for a project", async () => {
      configureReleaseTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(
        ([toolName]) => toolName === "releases_list_tags"
      );
      if (!call) throw new Error("releases_list_tags tool not registered");
      const [, , , handler] = call;

      const mockTags = ["tag1", "tag2", "tag3"];

      const mockReleaseApi = {
        getTags: jest.fn().mockResolvedValue(mockTags),
      };
      mockConnection.getReleaseApi.mockResolvedValue(mockReleaseApi);

      const params = {
        project: "test-project",
      };

      const result = await handler(params);

      expect(mockConnection.getReleaseApi).toHaveBeenCalled();
      expect(mockReleaseApi.getTags).toHaveBeenCalledWith("test-project");
      expect(result.content[0].text).toBe(JSON.stringify(mockTags, null, 2));
      expect(result.isError).toBeUndefined();
    });
  });
});

