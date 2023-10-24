import { AuthenticationProvider, Client, Context, Middleware, MiddlewareFactory } from "@microsoft/microsoft-graph-client";
import { HttpsProxyAgent } from "hpagent";
import fetch from "node-fetch";
// import "isomorphic-fetch";

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

export class AuthenticationProviderImpl implements AuthenticationProvider {
	private accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlJTSmZsWkNBMHg5cWZBdW9LdGV1WUpLRVdnOFN2eGt3S1lYYmdpaEtVNVkiLCJhbGciOiJSUzI1NiIsIng1dCI6IjlHbW55RlBraGMzaE91UjIybXZTdmduTG83WSIsImtpZCI6IjlHbW55RlBraGMzaE91UjIybXZTdmduTG83WSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81ZWU4ZjBiNS04YmRlLTQzMWItOWNkNS0wZDI3MTE0YmYwNmQvIiwiaWF0IjoxNjk2ODk5ODM5LCJuYmYiOjE2OTY4OTk4MzksImV4cCI6MTY5NjkwMzgyNCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFUUUF5LzhVQUFBQU0vWWtWSzcxMmd4Y091T1FSc1B0NGcvTlB1cm9lSXlpWmhCUEpjVGRiTGIrNDdTR3A1dUVaKzJBdmNua0NyM0kiLCJhbXIiOlsicHdkIiwicnNhIl0sImFwcF9kaXNwbGF5bmFtZSI6ImhlbGxvLXdvcmxkLXRhYi13aXRoLWJhY2tlbmQtYWFkIiwiYXBwaWQiOiI1M2NkYzhlYy01ZmY5LTQ3MWEtYmRmOS1hNTI0ZjEzZjY1NTIiLCJhcHBpZGFjciI6IjAiLCJkZXZpY2VpZCI6IjQwYzcwNTQ4LTMyOWYtNGM1NC04MmU3LTllZjA0MmE3MmQ2NSIsImZhbWlseV9uYW1lIjoieHYiLCJnaXZlbl9uYW1lIjoia25pZmUiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIyNDA0OmY4MDE6OTAwMDoxYTo4Njg3OmZkODA6NjllOTpmNTIiLCJuYW1lIjoia25pZmUgeHYiLCJvaWQiOiJiZDFlNWJkMS00YWUwLTQ5NzEtYTQzZS00YjlhYmFjZTc3YjMiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDBFQ0EwNEQxMyIsInJoIjoiMC5BWFlBdGZEb1h0NkxHME9jMVEwbkVVdndiUU1BQUFBQUFBQUF3QUFBQUFBQUFBQjJBTHMuIiwic2NwIjoib3BlbmlkIHByb2ZpbGUgVXNlci5SZWFkIFVzZXIuUmVhZC5BbGwgZW1haWwiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiI5UVJUYkdxYV9CVWR3OEJvUWdmLXFhSzBJOURBRFFDcDkzMFNCM2pVU0tVIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiNWVlOGYwYjUtOGJkZS00MzFiLTljZDUtMGQyNzExNGJmMDZkIiwidW5pcXVlX25hbWUiOiJ4emZAemhhb2Zlbmdvcmcub25taWNyb3NvZnQuY29tIiwidXBuIjoieHpmQHpoYW9mZW5nb3JnLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IlFwRmI4ajRrMzB1VEFZS2s2OWhkQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoid3BaLWxIbWQzaHY2YV9aUHRJLWxYN25NcUZrSzZnUHRNckxHLU9uY1lubyJ9LCJ4bXNfdGNkdCI6MTYwMjQ5MjgwMX0.WfvaonXh7IWHWmeXhyldhkfD4iyS5CPiVE6vpLtTgm6RF1W2VU_Zvp3yjeoPOdz33_QuHP2-X129BaTDSlSH1wWQyYMpyUsBZCKcMheE0nv2BO0fwjOkjud2stlopi_oyoX9ZXp6mCAzJl3GTUCyLE6qUzJXlMYwrPXqc6xCAk6VVerjGnhfac_6EswKnfeHGpRFbmVZUyLRv-_YV61-K_6g_ps9aMiNZWP-zligwyBGDyZCFXe-DVVT9h80lSbh6zdad0RwXYs6YsKvgo_IEST7Lw4xqUuVSCGnYu6gXjJFTXFoaCkIxChLaLg0Q2L60s99wZEUh0b8z6inHDVpFw";
	public async getAccessToken(): Promise<string> {
		return this.accessToken;
	}
}
export function getCustomGraphClient(authProvider: AuthenticationProvider): Client {
	// const middleware = MiddlewareFactory.getDefaultMiddlewareChain(authProvider);
	const middleware = MiddlewareFactory.getDefaultMiddlewareChain(new AuthenticationProviderImpl());
	// can use process.env.HTTPS_PROXY or process.env.HTTP_PROXY to create proxy
	middleware.splice(-1, 0, new ProxyMiddleware("http://127.0.0.1:8000"));
	middleware[middleware.length - 1].execute = async (context: Context) => {
		context.response = await fetch(context.request, context.options);
	};
	return Client.initWithMiddleware({ middleware: middleware });
}
