import { useContext } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { useData } from "@microsoft/teamsfx-react";
export function GraphApiCall(props: { codePath?: string; docsUrl?: string; }) {
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
  const { loading: loadingInstalledApp, data: installedAppData, error: installedAppError, reload: reloadInstalledApp } = useData(async () => {
    let response = await graphClient.api("/users/{{mocked user id}}/teamwork/installedApps").get();
    const res = response.value.map((item: any) => {
      return `${item.teamsAppDefinition.displayName} v${item.teamsAppDefinition.version}`;
    });
    return res;
  }, { autoLoad: false });

  const { loading: loadingGroups, data: groupsData, error: groupsError, reload: reloadGroups } = useData(async () => {
    const calendarView = await graphClient.api("groups").get();
    const res = calendarView.value.map((item: any) => {
      return `${item.displayName}`;
    });
    return res;
  }, { autoLoad: false });

  const { loading: loadingQueryString, data: queryStringData, error: queryStringError, reload: reloadQueryString } = useData(async () => {
    const calendarView = await graphClient.api("groups").query({
      $count: "true",
      $filter: "hasMembersWithLicenseErrors+eq+true",
      $select: "id,displayName",
    }).get();
    const res = calendarView.value.map((item: any) => {
      return `${item.displayName}`;
    });
    return res;
  }, { autoLoad: false });

  const { loading: loadingPhoto, data: photoData, error: photoError, reload: reloadPhoto } = useData(async () => {
    const photo = await graphClient.api("/users/{{mocked user id}}/photo/$value").get();
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
        {`\nAs the getToken method of teamsUserCredential has been hooked, there will be no permission error.\n`}
        {`Call Graph API: \n`}
        <code>{`  const result: any = await graphClient.api([URI]).get();`}</code>
        {`\nGrpah API used: \n`}
        {`https://graph.microsoft.com/v1.0/users/*/teamwork/installedApps\n`}
        {`https://graph.microsoft.com/v1.0/groups\n`}
        {`https://graph.microsoft.com/v1.0/users/*/photo/$value`}
      </pre>
      <div className="profile">
        {!loadingInstalledApp && (
          <Button appearance="primary" disabled={loadingInstalledApp} onClick={reloadInstalledApp}>
            Send Request graph.microsoft.com/v1.0/users/[user id]/teamwork/installedApps
          </Button>
        )}
        {loadingInstalledApp && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingInstalledApp && !!installedAppData && !installedAppError && <pre className="fixed">{`installed app list:\n${JSON.stringify(installedAppData, null, 2)}`}</pre>}
        {!loadingInstalledApp && !!installedAppError && <div className="error fixed">{`failed for statusCode: ${(installedAppError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="group">
        {!loadingGroups && (
          <div>
            <Button appearance="primary" disabled={loadingGroups} onClick={reloadGroups}>
              Send Request graph.microsoft.com/v1.0/groups
            </Button>
            &nbsp;&nbsp;<span>(The second request will has a different response)</span>
          </div>
        )}
        {loadingGroups && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingGroups && !!groupsData && !groupsError && <pre className="fixed">{`group list in organization:\n${JSON.stringify(groupsData, null, 2)}`}</pre>}
        {!loadingGroups && !!groupsError && <div className="error fixed">{`failed for statusCode: ${(groupsError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="query string">
        {!loadingQueryString && (
          <div>
            <Button appearance="primary" disabled={loadingQueryString} onClick={reloadQueryString}>
              Send Request graph.microsoft.com/v1.0/groups with query string $filter and $select
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
            Send Request graph.microsoft.com/v1.0/users/[user id]/photo/$value
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

