import { useContext, useState } from "react";
import { Button } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
export function GraphApiCall(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const [data, setData] = useState({
    profile: "Click to Call GraphApi from browser" as any,
    calendarView: {} as any,
    imgUrl: "",
    error: undefined as any,
  });
  const clickHandler = async () => {
    try {
      if (!teamsUserCredential) {
        throw new Error("TeamsFx SDK is not initialized.");
      }
      // await teamsUserCredential!.getToken(["User.Write.All"]);
      // await teamsUserCredential!.getToken(["User.Read"]);
      {
        const authProvider = new TokenCredentialAuthenticationProvider(
          teamsUserCredential,
          {
            scopes: ["https://graph.microsoft.com/User.Read"],
          }
        );
        // Initialize Graph client instance with authProvider
        const graphClient = Client.initWithMiddleware({
          authProvider: authProvider,
        });
        const profile: any = await graphClient.api("/me").get();
        const calendarView: any = await graphClient.api("me/calendarview").get();
        const image: any = await graphClient.api(`/users/${profile.id}/photo/$value`).get();
        const url = window.URL || window.webkitURL;
        const imgUrl = url.createObjectURL(image);
        // return {
        //   profile,
        //   imgUrl,
        // };
        setData({
          profile,
          imgUrl,
          calendarView,
          error: "",
        });
      }
    } catch (error: any) {
      setData({
        profile: undefined,
        imgUrl: "",
        calendarView: {},
        error: error.message,
      });
    }
  };
  // const { loading, data, error, reload } = useData(async () => {
  //   if (!teamsUserCredential) {
  //     throw new Error("TeamsFx SDK is not initialized.");
  //   }
  //   {
  //     const authProvider = new TokenCredentialAuthenticationProvider(
  //       teamsUserCredential,
  //       {
  //         scopes: ["https://graph.microsoft.com/User.Read"],
  //       }
  //     );
  //     // Initialize Graph client instance with authProvider
  //     const graphClient = Client.initWithMiddleware({
  //       authProvider: authProvider,
  //     });
  //     const profile: any = await graphClient.api("/me").get();
  //     const image: any = await graphClient.api(`/users/${profile.id}/photo/$value`).get();
  //     const url = window.URL || window.webkitURL;
  //     const imgUrl = url.createObjectURL(image);
  //     return {
  //       profile,
  //       imgUrl,
  //     };
  //   }
  // });
  return (
    <div>
      <h2>Call Graph API in frontend</h2>
      <pre>
        {`Call Graph API in frontend by graph client.\n`}
        {`Use teamsUserCredential as authProvider of graph client.\n`}
        {`As the getToken method of teamsUserCredential has been hooked, there is no permission error.\n`}
        {`Call Graph API: \n`}
        <code>{`  const profile: any = await graphClient.api("/me").get();`}</code>
        {`\nUsed API: \n`}
        {`https://graph.microsoft.com/v1.0/me\n`}
        {`https://graph.microsoft.com/v1.0/me/calendarview\n`}
        {`https://graph.microsoft.com/v1.0/users/{userId}/photo/$value`}
      </pre>
      {(
        <Button appearance="primary" onClick={clickHandler}>
          Call Graph API in frontend
        </Button>
      )}
      {!!data.profile && !data.error && <pre className="fixed">{JSON.stringify(data.profile, null, 2)}</pre>}
      {!!data.imgUrl && !data.error && <img src={data.imgUrl} loading="lazy" />}
      {!!data.calendarView && !data.error && <pre className="fixed">{JSON.stringify(data.calendarView, null, 2)}</pre>}
      {!!data.error && <div className="error fixed">{data.error}</div>}
    </div>
  );
}
