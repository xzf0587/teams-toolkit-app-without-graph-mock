import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";
import { shouldHook } from "../Hook";

export function Login(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    await teamsUserCredential.getUserInfo();
    const scopes = [
      "User.Read.All",
      "Team.ReadBasic.All",
      "Group.Read.All",
      "TeamMember.Read.All",
      "TeamSettings.Read.All",
    ];
    await teamsUserCredential!.login(scopes);
    const loginRes = "login success";
    return loginRes;
  }, { autoLoad: false });

  return (
    <div>
      {shouldHook() && (<div>
        <h2>Hook Login</h2>
        <pre>
          {`Hook the teamsUserCredential login and getToken method. Import the hook file in App.tsx\n`}
          <code>TeamsUserCredential.prototype.login</code><br />
          <code>TeamsUserCredential.prototype.getToken</code><br />
          {`Login will always success and getToken will return a mocked token.\n`}
        </pre>
      </div>)}
      {!loading && !shouldHook() &&(
        <pre>Requested scopes: <br/> "User.Read.All",  <br/> "Team.ReadBasic.All",  <br/> "Group.Read.All",  <br/> "TeamMember.Read.All",  <br/> "TeamSettings.Read.All"</pre>
      )}
      {!loading && (
        <Button appearance="primary" disabled={loading} onClick={reload}>
          Login
        </Button>
      )}
      
      {loading && (
        <pre className="fixed">
          <Spinner />
        </pre>
      )}
      {!loading && !!data && !error && <p className="fixed">{data}</p>}
      {!loading && !!error && <div className="error fixed">{(error as any).toString()}</div>}
    </div>
  );
}
