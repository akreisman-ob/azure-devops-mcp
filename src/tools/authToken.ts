// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getBasicHandler, getBearerHandler, getPersonalAccessTokenHandler } from "azure-devops-node-api";

class AuthToken {
  token: string;
  type: string;
  private _pat?: string;

  constructor(token: string, type: string, _pat?: string) {
    this.token = token;
    this.type = type;
    this._pat = _pat;
  }

  static withPAT(pat: string): AuthToken {
    return new AuthToken(Buffer.from(`:${pat}`).toString('base64'), "Basic", pat);
  }

  static withApiToken(apiToken: string): AuthToken {
    return new AuthToken(apiToken, "Bearer");
  }

  getHandler() {
    switch (this.type) {
      case "Basic":
        return getPersonalAccessTokenHandler(this._pat ?? "");
      default:
        throw new Error("Method not implemented.");
        //return getBearerHandler(this.token);

    }
  }
}

export { AuthToken };

