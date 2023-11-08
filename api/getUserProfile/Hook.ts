import { OnBehalfOfUserCredential } from "@microsoft/teamsfx";
import { AccessToken } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { HttpsProxyAgent } from "hpagent";
import fetch from "node-fetch";

function shouldHook(): boolean {
  return process.env.HOOK_GRAPH === "true";
}

if (shouldHook()) {
  // hook getToken to return a mocked access token. The token can not be used to call real graph api.
  OnBehalfOfUserCredential.prototype.getToken = async (scopes: string | string[], options?: any) => {
    const accessToken: AccessToken = {
      token: "mocked token",
      expiresOnTimestamp: 2147483647,
    };
    return accessToken;
  };

  global["fetch"] = fetch;
  const agent = new HttpsProxyAgent({
    proxy: 'http://127.0.0.1:8000',
    rejectUnauthorized: false,
  });

  const oldInitWithMiddleware = Client.initWithMiddleware;
  Client.initWithMiddleware = (options: any) => {
    options.fetchOptions = options.fetchOptions ?? {}
    options.fetchOptions.agent = agent;
    return oldInitWithMiddleware(options);
  };

  const oldInit = Client.init;
  Client.init = (options: any) => {
    options.fetchOptions = options.fetchOptions ?? {}
    options.fetchOptions.agent = agent;
    return oldInit(options);
  };
}
