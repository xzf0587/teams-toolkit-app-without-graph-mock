import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";

export function Login(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    await teamsUserCredential.getUserInfo();
    const scopes = [
      "User.Read.All",
      "Calendars.Read.All"
    ];
    await teamsUserCredential!.login(scopes);
    const loginRes = "login success";
    return loginRes;
  }, { autoLoad: false });

  return (
    <div>
      <h2>Hook Login</h2>
      <pre>
        {`Hook the teamsUserCredential login method. It will always return success.\n`}
        {`Login for scopes: ["User.Read.All", "Calendars.Read.All"]`}
      </pre>
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
