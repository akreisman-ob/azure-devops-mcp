// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AzureCliCredential, ChainedTokenCredential, DefaultAzureCredential, TokenCredential } from "@azure/identity";
import { AccountInfo, AuthenticationResult, PublicClientApplication } from "@azure/msal-node";
import open from "open";
import { AuthToken } from "./tools/authToken.js";

const scopes = ["499b84ac-1321-427f-aa17-267ca6975798/.default"];


class OAuthAuthenticator {
  static clientId = "0d50963b-7bb9-4fe7-94c7-a99af00b5136";
  static defaultAuthority = "https://login.microsoftonline.com/common";
  static zeroTenantId = "00000000-0000-0000-0000-000000000000";

  private accountId: AccountInfo | null;
  private publicClientApp: PublicClientApplication;

  constructor(tenantId?: string) {
    this.accountId = null;

    let authority = OAuthAuthenticator.defaultAuthority;
    if (tenantId && tenantId !== OAuthAuthenticator.zeroTenantId) {
      authority = `https://login.microsoftonline.com/${tenantId}`;
    }

    this.publicClientApp = new PublicClientApplication({
      auth: {
        clientId: OAuthAuthenticator.clientId,
        authority,
      },
    });
  }

  public async getToken(): Promise<string> {
    let authResult: AuthenticationResult | null = null;
    if (this.accountId) {
      try {
        authResult = await this.publicClientApp.acquireTokenSilent({
          scopes,
          account: this.accountId,
        });
      } catch {
        authResult = null;
      }
    }
    if (!authResult) {
      authResult = await this.publicClientApp.acquireTokenInteractive({
        scopes,
        openBrowser: async (url) => {
          open(url);
        },
      });
      this.accountId = authResult.account;
    }

    if (!authResult.accessToken) {
      throw new Error("Failed to obtain Azure DevOps OAuth token.");
    }
    return authResult.accessToken;
  }
}

function createAuthenticator(type: string, tenantId?: string): () => Promise<AuthToken> {
  switch (type) {
    case "pat":
      // Read token from fixed environment variable
      return async () => {
        const pat = process.env["ADO_MCP_PAT"];
        if (!pat) {
          throw new Error("Environment variable 'ADO_MCP_PAT' is not set or empty. Please set it with a valid Azure DevOps Personal Access Token.");
        }
        return AuthToken.withPAT(pat);
      };
    case "envvar":
      // Read token from fixed environment variable
      return async () => {
        const token = process.env["ADO_MCP_AUTH_TOKEN"];
        if (!token) {
          throw new Error("Environment variable 'ADO_MCP_AUTH_TOKEN' is not set or empty. Please set it with a valid Azure DevOps Personal Access Token.");
        }
        return AuthToken.withApiToken(token);
      };

    case "azcli":
    case "env":
      if (type !== "env") {
        process.env.AZURE_TOKEN_CREDENTIALS = "dev";
      }
      let credential: TokenCredential = new DefaultAzureCredential(); // CodeQL [SM05138] resolved by explicitly setting AZURE_TOKEN_CREDENTIALS
      if (tenantId) {
        // Use Azure CLI credential if tenantId is provided for multi-tenant scenarios
        const azureCliCredential = new AzureCliCredential({ tenantId });
        credential = new ChainedTokenCredential(azureCliCredential, credential);
      }
      return async () => {
        const result = await credential.getToken(scopes);
        if (!result) {
          throw new Error("Failed to obtain Azure DevOps token. Ensure you have Azure CLI logged or use interactive type of authentication.");
        }
        return AuthToken.withApiToken(result.token);
      };

    default:
      const authenticator = new OAuthAuthenticator(tenantId);
      return async () => {
        const token = await authenticator.getToken();
        return AuthToken.withApiToken(token);
      };
  }
}
export { createAuthenticator, AuthToken };
