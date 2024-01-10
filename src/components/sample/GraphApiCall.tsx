import { useContext } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { useData } from "@microsoft/teamsfx-react";
import { shouldHook } from "../Hook";
export function GraphApiCall(props: { codePath?: string; docsUrl?: string; }) {
  let teamsId: string;
  let channelId: string;
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const authProvider = new TokenCredentialAuthenticationProvider(
    teamsUserCredential!,
    {
      // as the TeamsUserCredential getToken is mocked, any scope is ok.
      scopes: ["https://graph.microsoft.com/User.Read.All"],
    }
  );
  const graphClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
  const { loading: loadingMessages, data: messagesData, error: messagesError, reload: reloadMessages } = useData(async () => {
    if (!teamsId) {
      const teamsResponse = await graphClient.api("me/joinedTeams").get();
      teamsId = teamsResponse.value[0].id;
    }
    if (!channelId) {
      const channelResponse = await graphClient.api(`teams/${teamsId}/channels`).get();
      channelId = channelResponse.value[0].id;
    }
    let response = await graphClient.api(`teams/${teamsId}/channels/${channelId}/messages`).version("beta").get();
    const res = response.value.map((item: any) => {
      return item.body.content;
    });
    return res.slice(0, 3);
  }, { autoLoad: false });

  const { loading: loadingTeams, data: teamsData, error: teamsError, reload: reloadTeams } = useData(async () => {
    const response = await graphClient.api("me/joinedTeams").get();
    const res = response.value.map((item: any) => {
      return `${item.displayName}`;
    });
    return res;
  }, { autoLoad: false });

  const { loading: loadingQueryString, data: queryStringData, error: queryStringError, reload: reloadQueryString } = useData(async () => {
    const groups = await graphClient.api("groups").query({
      $count: "true",
      $filter: "startswith(displayName, 'zhao')",
      $select: "id,displayName",
    }).get();
    const res = groups.value.map((item: any) => {
      return `${item.displayName}`;
    });
    return res;
  }, { autoLoad: false });

  const { loading: loadingPhoto, data: photoData, error: photoError, reload: reloadPhoto } = useData(async () => {
    const photo = await graphClient.api(`me/photo/$value`).get();
    const url = window.URL || window.webkitURL;
    return url.createObjectURL(photo);
  }, { autoLoad: false });
  return (
    <div>
      <h2>Call Graph API in frontend</h2>
      <pre>
        {`Call Graph API in frontend by graph client.\n`}
        {`Use teamsUserCredential as authProvider of graph client.\n`}
        <code>{`  const graphClient = Client.initWithMiddleware({ authProvider });`}</code>
        {`Call Graph API: \n`}
        <code>{`  const result: any = await graphClient.api([URI]).get();`}</code>
        {`\nGrpah API used: \n`}
        {`https://graph.microsoft.com/beta/teams/*/channels/*/messages\n`}
        {`https://graph.microsoft.com/v1.0/me/joinedTeams"\n`}
        {`https://graph.microsoft.com/v1.0/groups\n`}
        {`https://graph.microsoft.com/v1.0/me/photo/$value`}
      </pre>
      {!shouldHook() && (<div>
        <h2>Hook Mode</h2>
        <pre>
          {`1. Graph request will always acquire mocked token in hook mode.\n`}
          {`2. Graph request will be sent to system https proxy(if existing) in browser by default.\n   No code change is needed.\n`}
        </pre>
      </div>)}
      <div className="profile">
        {!loadingMessages && (
          <Button appearance="primary" disabled={loadingMessages} onClick={reloadMessages}>
            Get graph.microsoft.com/beta/teams/[teams id]/channels/[channel id]/messages
          </Button>
        )}
        {loadingMessages && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingMessages && !!messagesData && !messagesError && <pre className="fixed">{`channel messages:\n${JSON.stringify(messagesData, null, 2)}`}</pre>}
        {!loadingMessages && !!messagesError && <div className="error fixed">{`failed for statusCode: ${(messagesError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="teams">
        {!loadingTeams && (
            <Button appearance="primary" disabled={loadingTeams} onClick={reloadTeams}>
              Get graph.microsoft.com/v1.0/me/joinedTeams
            </Button>
        )}
        {!loadingTeams && shouldHook() && (
           <span>&nbsp;&nbsp;(The second request is different as the proxy configured)</span>
        )}
        {loadingTeams && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingTeams && !!teamsData && !teamsError && <pre className="fixed">{`joinedTeams list:\n${JSON.stringify(teamsData, null, 2)}`}</pre>}
        {!loadingTeams && !!teamsError && <div className="error fixed">{`failed for statusCode: ${(teamsError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="query string">
        {!loadingQueryString && (
          <div>
            <Button appearance="primary" disabled={loadingQueryString} onClick={reloadQueryString}>
              Get graph.microsoft.com/v1.0/groups with query string $filter and $select
            </Button>
          </div>
        )}
        {loadingQueryString && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingQueryString && !!queryStringData && !queryStringError && <pre className="fixed">{`group list for query string:\n${JSON.stringify(queryStringData, null, 2)}`}</pre>}
        {!loadingQueryString && !!queryStringError && <div className="error fixed">{`failed for statusCode: ${(queryStringError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="photo">
        {!loadingPhoto && (
          <Button appearance="primary" disabled={loadingPhoto} onClick={reloadPhoto}>
            Get graph.microsoft.com/v1.0/me/photo/$value
          </Button>
        )}
        {loadingPhoto && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        <p></p>
        {!loadingPhoto && !!photoData && !photoError && <img src={photoData} loading="lazy" />}
        {!loadingPhoto && !!photoError && <div className="error fixed">{`failed for statusCode: ${(photoError as any).statusCode}`}</div>}
      </div>

    </div>
  );
}

