import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import * as axios from "axios";
import { BearerTokenAuthProvider, createApiClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";
import { useData } from "@microsoft/teamsfx-react";

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
        {`Call Backend API from frontend using: const apiClient = createApiClient().\n`}
        {`Hook the grapClient creation by adding proxy middleware:\n`}
        <code>  ProxyMiddleware("http://LOCAL_PROXY_ADDRESS")</code>
        {`\nGrpah API called in backend: \n`}
        {`https://graph.microsoft.com/v1.0/me/joinedTeams (using obo auth)\n`}
        {`https://graph.microsoft.com/v1.0/teams/{{mocked teams}}/members (using application auth)\n`}
      </pre>
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
