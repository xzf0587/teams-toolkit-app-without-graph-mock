/* This code sample provides a starter kit to implement server side logic for your Teams App in TypeScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

import "isomorphic-fetch";
import { Context, HttpRequest } from "@azure/functions";
import {
  OnBehalfOfCredentialAuthConfig,
  OnBehalfOfUserCredential,
} from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import config from "../config";
import "./Hook";
import { ClientSecretCredential } from "@azure/identity";

interface Response {
  status: number;
  body: { [key: string]: any; };
}

type TeamsfxContext = { [key: string]: any; };

export default async function run(
  context: Context,
  req: HttpRequest,
  teamsfxContext: TeamsfxContext
): Promise<Response> {
  context.log("HTTP trigger function processed a request.");

  // Initialize response.
  const res: Response = {
    status: 200,
    body: {},
  };

  // Prepare access token.
  const accessToken: string = teamsfxContext["AccessToken"];
  if (!accessToken) {
    return {
      status: 400,
      body: {
        error: "No access token was found in request header.",
      },
    };
  }

  const oboAuthConfig: OnBehalfOfCredentialAuthConfig = {
    authorityHost: config.authorityHost,
    clientId: config.clientId,
    tenantId: config.tenantId,
    clientSecret: config.clientSecret,
  };

  let oboCredential: OnBehalfOfUserCredential;
  try {
    oboCredential = new OnBehalfOfUserCredential(accessToken, oboAuthConfig);
  } catch (e) {
    context.log.error(e);
    return {
      status: 500,
      body: {
        error:
          "Failed to construct OnBehalfOfUserCredential using your accessToken. " +
          "Ensure your function app is configured with the right Azure AD App registration.",
      },
    };
  }
  const scopes = [
    "https://graph.microsoft.com/TeamSettings.Read.All",
    "https://graph.microsoft.com/TeamMember.Read.All"
  ];
  try {
    const authProvider = new TokenCredentialAuthenticationProvider(
      oboCredential,
      {
        scopes,
      }
    );
    // Initialize Graph client instance with obo authProvider
    const graphClient = Client.initWithMiddleware({
      // in mock mode, use a mocked authProvider to replace the real authProvider. No user code is required to change.
      authProvider: authProvider,
    });
    let joinedTeams: any;
    try {
      joinedTeams = await graphClient.api("/me/joinedTeams").get();
    } catch (error) {
      joinedTeams = `get graph api me/joinedTeams failed for statusCode: ${(error as any).statusCode}`;
    }
    res.body.joinedTeams = joinedTeams.value.map((team: any) => team.displayName );

    // Initialize Graph client instance with application authProvider
    const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
    const applicationAuthProvider = new TokenCredentialAuthenticationProvider(credential, { scopes });
    const applicationGraphClient = Client.initWithMiddleware({
      // in mock mode, use a mocked authProvider to replace the real authProvider. No user code is required to change.
      authProvider: applicationAuthProvider,
    });
    let teamMembers: any;
    try {
      teamMembers = await applicationGraphClient.api("teams/{{mocked teams}}/members").get();
    } catch (error) {
      teamMembers = `get graph api teams/{{teams}}/members failed for statusCode: ${(error as any).statusCode}`;
    }
    res.body.teamMembers = teamMembers.value.map((member: any) => member.displayName );
  } catch (e) {
    context.log.error(e);
    return {
      status: 500,
      body: {
        error:
          "Failed to retrieve information from Microsoft Graph. The application credential is invalid.",
      },
    };
  }

  return res;
}
