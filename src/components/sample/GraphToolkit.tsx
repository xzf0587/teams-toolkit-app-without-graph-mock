import { useContext } from "react";
import { Button } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import {
  Person,
  PersonViewType,
} from "@microsoft/mgt-react";
import { CacheService } from "@microsoft/mgt";

export function GraphToolkit(props: { codePath?: string; docsUrl?: string; }) {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const scopes = ["https://graph.microsoft.com/User.Read.All"];
  CacheService.config.isEnabled = false;
  const provider = new TeamsFxProvider(teamsUserCredential!, scopes);
  Providers.globalProvider = provider;
  const reloadInstalledApp = async () => {
    CacheService.clearCaches();
    Providers.globalProvider.setState(ProviderState.SignedIn);
  };

  return (
    <div>
      <h2>Graph Toolkit</h2>
      <div className="my-account-area">
        <pre>
          {`Person Component Code of Graph Toolkit.\n`}
          <code>{`<Person`}</code><br />
          <code>{`  userId={userId}`}</code><br />
          <code>{`  view={PersonViewType.threelines}`}</code><br />
          <code>{`></Person>`}</code><br />
        </pre>
        <Button appearance="primary" onClick={reloadInstalledApp}>
          Load Person Information
        </Button>
        <br />
        <br />
        <Person
          // userId={"00000000-0000-0000-0000-000000000000"}
          userId={"3ea99f0d-bdf5-4fdc-b9c5-57646a283974"}
          view={PersonViewType.threelines}
        ></Person>
        {/* <pre>
          {`Load Person for me.\n`}
          <code>{`<Person`}</code><br />
          <code>{`  personQuery="me"`}</code><br />
          <code>{`  view={PersonViewType.threelines}`}</code><br />
          <code>{`></Person>`}</code><br />
        </pre>
        <Person
          personQuery="me"
          view={PersonViewType.threelines}
        ></Person> */}
      </div>
    </div>
  );
}

