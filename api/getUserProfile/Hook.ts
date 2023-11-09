import { AuthenticationProvider, Client } from "@microsoft/microsoft-graph-client";
import { HttpsProxyAgent } from "hpagent";
import fetch from "node-fetch";

function shouldHook(): boolean {
  return process.env.HOOK_GRAPH === "true";
}

if (shouldHook()) {
  global["fetch"] = fetch;
  const agent = new HttpsProxyAgent({
    proxy: 'http://127.0.0.1:8000',
    rejectUnauthorized: false,
  });

  // replace the default auth provider with a custom mocked one in graph client, 
  // so that we can handle application auth as well as obo auth
  class AuthenticationProviderImpl implements AuthenticationProvider {
    private accessToken = "mocked token";
    public async getAccessToken(): Promise<string> {
      return this.accessToken;
    }
  }
  const oldInitWithMiddleware = Client.initWithMiddleware;
  Client.initWithMiddleware = (options: any) => {
    options.authProvider = new AuthenticationProviderImpl();
    options.fetchOptions = options.fetchOptions ?? {}
    options.fetchOptions.agent = agent;
    return oldInitWithMiddleware(options);
  };

  const oldInit = Client.init;
  Client.init = (options: any) => {
    options.authProvider = new AuthenticationProviderImpl();
    options.fetchOptions = options.fetchOptions ?? {}
    options.fetchOptions.agent = agent;
    return oldInit(options);
  };
}
