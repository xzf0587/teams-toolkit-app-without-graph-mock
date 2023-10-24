import { useContext, useState } from "react";
import { Button } from "@fluentui/react-components";
import * as axios from "axios";
import { BearerTokenAuthProvider, createApiClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";

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

      if (err?.response?.status === 404) {
        funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
      } else if (err.message === "Network Error") {
        funcErrorMsg =
          "Cannot call Azure Function due to network error, please check your network connection status and ";
        if (err.config?.url && err.config.url.indexOf("localhost") >= 0) {
          funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
        } else {
          funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
        }
      } else {
        funcErrorMsg = err.message;
        if (err.response?.data?.error) {
          funcErrorMsg += ": " + err.response.data.error;
        }
      }

      throw new Error(funcErrorMsg);
    }
    throw err;
  }
}

export function AzureFunctions(props: { codePath?: string; docsUrl?: string; }) {
  const { codePath, docsUrl } = {
    codePath: `api/${functionName}/index.ts`,
    docsUrl: "https://aka.ms/teamsfx-azure-functions",
    ...props,
  };
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const [res, setRes] = useState({
    data: "Click to Call GraphApi from Azure Function" as any,
    error: undefined,
  });
  const clickHandler = async () => {
    try {
      if (!teamsUserCredential) {
        throw new Error("TeamsFx SDK is not initialized.");
      }
      const functionRes = await callFunction(teamsUserCredential);
      setRes({
        data: functionRes,
        error: undefined,
      });
    } catch (error: any) {
      setRes({
        data: undefined,
        error: error.message,
      });
    }
  };

  return (
    <div>
      <h2>Call GraphApi from Azure Function</h2>
      <pre>
      {`Call Backend API from frontend using: const apiClient = createApiClient().\n`}
      {`Hook the grapClient creation by adding proxy middleware:\n`}
      <code>  ProxyMiddleware("http://LOCAL_PROXY_ADDRESS")</code>
      </pre>
      {(
        <Button appearance="primary" onClick={clickHandler}>
          Call Azure Function
        </Button>
      )}
      {res.data && <pre className="fixed">{JSON.stringify(res.data, null, 2)}</pre>}
      {!!res.error && <div className="error fixed">{(res.error as any).toString()}</div>}
    </div>
  );
}
