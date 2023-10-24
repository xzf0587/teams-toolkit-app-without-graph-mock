import { useContext, useState } from "react";
import { Button } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";

export function Login(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const [message, setMessage] = useState("click login button to login first");
  const clickHandler = async () => {
    try {
      if (!teamsUserCredential) {
        throw new Error("TeamsFx SDK is not initialized.");
      }
      await teamsUserCredential!.login(["User.Read"]);
      // The first time to get token will set msal.account.keys in session storage. 
      // If there is no account info in session storage, it will not execute acquireTokenSilent to get access token by refresh token.
      // It will execute ssoSilent login instead. Currently, the mocked response of the oauth token api can not support auth code 
      // await teamsUserCredential!.getToken(["User.Read"]);
      return setMessage("login success");
    } catch (error: any) {
      setMessage(error.message);
    }
  };
  return (
    <div>
      <h2>Login</h2>
      <pre>
        {`Use hook to handle the teamsUserCredential login method. It will always return success.`}
      </pre>
      {(
        <Button appearance="primary" onClick={clickHandler}>
          Login
        </Button>
      )}
      {<p className="fixed">{message}</p>}
    </div>
  );
}
