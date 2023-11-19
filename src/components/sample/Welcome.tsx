import { useContext, useState } from "react";
import {
  Image,
  TabList,
  Tab,
  SelectTabEvent,
  SelectTabData,
  TabValue,
} from "@fluentui/react-components";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { AzureFunctions } from "./AzureFunctions";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";
import { GraphApiCall } from "./GraphApiCall";
import { Login } from "./Login";
import { GraphToolkit } from "./GraphToolkit";

export function Welcome(props: { showFunction?: boolean; environment?: string; }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const [selectedValue, setSelectedValue] = useState<TabValue>("local");

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value);
  };
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>

        <div className="tabList">
          <TabList selectedValue={selectedValue} onTabSelect={onTabSelect}>
            <Tab id="Local" value="local">
              1. Login
            </Tab>
            <Tab id="Azure" value="azure">
              2. Call Graph API from browser
            </Tab>
            <Tab id="Publish" value="publish">
              3. Call Graph API from Azure Function
            </Tab>
            <Tab id="Toolkit" value="toolkit">
              4. Use Graph Toolkit
            </Tab>
          </TabList>
          <div>
            {selectedValue === "local" && (
              <div>
                {/* <CurrentUser userName={userName} /> */}
                {<Login />}
              </div>
            )}
            {selectedValue === "azure" && (
              <div>
                {<GraphApiCall />}
              </div>
            )}
            {selectedValue === "publish" && (
              <div>
                {<AzureFunctions />}
              </div>
            )}
             {selectedValue === "toolkit" && (
              <div>
                {<GraphToolkit />}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
