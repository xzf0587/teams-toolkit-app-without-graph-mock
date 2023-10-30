import { OnBehalfOfUserCredential } from "@microsoft/teamsfx";
import { AccessToken } from "@azure/identity";
import { AuthenticationProvider, Client, Context, Middleware, MiddlewareFactory } from "@microsoft/microsoft-graph-client";
import { HttpsProxyAgent } from "hpagent";
import fetch from "node-fetch";

function shouldHook(): boolean {
  return process.env.HOOK_GRAPH === "true";
}

if (shouldHook()) {
  // hook getToken to return a mocked access token. The token can not be used to call real graph api. To go through the flow of graph api call, please handle th.
  OnBehalfOfUserCredential.prototype.getToken = async (scopes: string | string[], options?: any) => {
    const accessToken: AccessToken = {
      token: "mocked token",
      expiresOnTimestamp: 2147483647,
    };
    return accessToken;
  };

  class ProxyMiddleware implements Middleware {
    private url: string;

    private nextMiddleware!: Middleware;

    public constructor(url: string) {
      this.url = url;
    }

    public async execute(context: Context): Promise<void> {
      if (context.options) {
        context.options.agent = new HttpsProxyAgent({
          proxy: this.url,
          rejectUnauthorized: false,
        });
      }

      return await this.nextMiddleware.execute(context);
    }

    public setNext(next: Middleware): void {
      this.nextMiddleware = next;
    }
  }

  class AuthenticationProviderImpl implements AuthenticationProvider {
    private accessToken = "mocked token";
    public async getAccessToken(): Promise<string> {
      return this.accessToken;
    }
  }

  const oldInitWithMiddleware = Client.initWithMiddleware;
  Client.initWithMiddleware = (options: any) => {
    // const middleware = MiddlewareFactory.getDefaultMiddlewareChain(authProvider);
    const middleware = MiddlewareFactory.getDefaultMiddlewareChain(new AuthenticationProviderImpl());
    // can use process.env.HTTPS_PROXY or process.env.HTTP_PROXY to create proxy
    middleware.splice(-1, 0, new ProxyMiddleware("http://127.0.0.1:8000"));
    middleware[middleware.length - 1].execute = async (context: Context) => {
      context.response = await fetch(context.request, context.options);
    };
    // options.middleware = middleware;
    // options.baseUrl = "https://graph1.microsoft.com/";
    // return oldInitWithMiddleware(options);
    return oldInitWithMiddleware({ middleware: middleware });
  };
}
