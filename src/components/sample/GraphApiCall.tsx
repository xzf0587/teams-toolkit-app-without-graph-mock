import { useContext, useState } from "react";
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
  const { loading: loadingProfile, data: profileData, error: profileError, reload: reloadProfile } = useData(async () => {
    const profile = await graphClient.api("/me").get();
    return profile;
  }, { autoLoad: false });

  const { loading: loadingCalendar, data: calendarData, error: calendarError, reload: reloadCalendar } = useData(async () => {
    const calendarView = await graphClient.api("me/calendarview").get();
    return calendarView;
  }, { autoLoad: false });

  const { loading: loadingPhoto, data: photoData, error: photoError, reload: reloadPhoto } = useData(async () => {
    const photo = await graphClient.api("/users/me/photo/$value").get();
    const url = window.URL || window.webkitURL;
    return url.createObjectURL(photo);
  }, { autoLoad: false });
  return (
    <div>
      <h2>Call Graph API in frontend</h2>
      <pre>
        {`Call Graph API in frontend by graph client.\n`}
        {`Use teamsUserCredential as authProvider of graph client.\n`}
        {`As the getToken method of teamsUserCredential has been hooked, there is no permission error.\n`}
        {`Call Graph API: \n`}
        <code>{`  const result: any = await graphClient.api({URI}).get();`}</code>
        {`\nGrpah API used: \n`}
        {`https://graph.microsoft.com/v1.0/me\n`}
        {`https://graph.microsoft.com/v1.0/me/calendarview\n`}
        {`https://graph.microsoft.com/v1.0/users/{userId}/photo/$value`}
      </pre>
      <div className="profile">
        {!loadingProfile && (
          <Button appearance="primary" disabled={loadingProfile} onClick={reloadProfile}>
            Send Request graph.microsoft.com/v1.0/me
          </Button>
        )}
        {loadingProfile && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingProfile && !!profileData && !profileError && <pre className="fixed">{JSON.stringify(profileData, null, 2)}</pre>}
        {!loadingProfile && !!profileError && <div className="error fixed">{`failed for statusCode: ${(profileError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="calendarview">
        {!loadingCalendar && (
          <Button appearance="primary" disabled={loadingCalendar} onClick={reloadCalendar}>
            Send Request graph.microsoft.com/v1.0/me/calendarview
          </Button>
        )}
        {loadingCalendar && (
          <pre className="fixed">
            <Spinner />
          </pre>
        )}
        {!loadingCalendar && !!calendarData && !calendarError && <pre className="fixed">{JSON.stringify(calendarData, null, 2)}</pre>}
        {!loadingCalendar && !!calendarError && <div className="error fixed">{`failed for statusCode: ${(calendarError as any).statusCode}`}</div>}
      </div>
      <p></p>
      <div className="photo">
        {!loadingPhoto && (
          <Button appearance="primary" disabled={loadingPhoto} onClick={reloadPhoto}>
            Send Request graph.microsoft.com/v1.0/me/photo/$value
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

