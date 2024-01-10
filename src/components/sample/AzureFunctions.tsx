import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import * as axios from "axios";
import { BearerTokenAuthProvider, createApiClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";
import { useData } from "@microsoft/teamsfx-react";
import { shouldHook } from "../Hook";

const functionName = config.apiName || "myFunc";
async function callFunction(teamsUserCredential: TeamsUserCredential) {
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(async () => (await teamsUserCredential.getToken(""))!.token)
    );
    const response = await apiClient.get(functionName);
    return response.data;
  } catch (err: unknown) {
    if (axios.default.isAxiosError(err)) {
      let funcErrorMsg = "";
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
      throw new Error(funcErrorMsg);
    }
    throw err;
  }
}

export function AzureFunctions(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    const calendarView = await callFunction(teamsUserCredential!);
    return calendarView;
  }, { autoLoad: false });

  return (
    <div>
      <h2>Call GraphApi from Azure Function</h2>
      <pre>
        {`Grpah API called in backend: \n`}
        {`https://graph.microsoft.com/v1.0/me/joinedTeams (using obo auth)\n`}
        {shouldHook() && (`https://graph.microsoft.com/v1.0/teams/*/members (using application auth)\n`)}
      </pre>
      {shouldHook() && (<div>
        <h2>Hook Mode</h2>
        <pre>
          {`1. Graph request will always acquire mocked token in hook mode.\n`}
          {`2. Code change is needed for sending the Graph request to the proxy.\n`}
        </pre>
        {/* <pre>
          {`Hook the grapClient creation in backend by setting fetchOptions.agent \n    HttpsProxyAgent([M365 prxoy listening address]):\nImport the hook file in index.ts.\n`}
        </pre> */}
      </div>)}
      <div className="call backend api">
        {!loading && (
          <Button appearance="primary" disabled={loading} onClick={reload}>
            Call Azure Function
          </Button>
        )}
        {loading && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loading && !!data && !error && <pre className="fixed">{JSON.stringify(data, null, 2)}</pre>}
        {!loading && !!error && <div className="error fixed">{`failed for error: ${(error as any).toString()}`}</div>}
      </div>
    </div>
  );
}
