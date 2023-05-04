import { useContext, useState } from "react";
import { Image, Menu, tabListBehavior } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { pages } from "@microsoft/teams-js";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";

export function Welcome(props: {
  showFunction?: boolean;
  environment?: string;
}) {
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

  //const steps = ["local", "azure", "publish"];
  const friendlyStepsName: { [key: string]: string } = {
    local: "1. Build your app locally",
    azure: "2. Provision and Deploy to the Cloud",
    publish: "3. Publish to Teams",
    search: "Search",
    provision: "Provision Guest Users",
  };

  const [selectedMenuItem, setSelectedMenuItem] = useState<string | undefined>(
    "search"
  );
  const menus = ["search", "provision"].map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const setSelectedMenu = (step: string) => {
    //setSelectedMenuItem(step);
    if (pages.isSupported()) {
      pages.getConfig().then((config) => {
        console.log(config);
      });
      const navPromise = pages.navigateToApp({
        appId: "a5b8cacf-17f9-41cf-b6b2-ed293b31665f",
        pageId: "index1",
      });
      navPromise.then((result) => {
        console.log(result);
      });
    }
  };
  // const items = steps.map((step) => {
  //   return {
  //     key: step,
  //     content: friendlyStepsName[step] || "",
  //     onClick: () => setSelectedMenuItem(step),
  //   };
  // });

  //const { teamsUserCredential } = useContext(TeamsFxContext);
  // const { loading, data, error } = useData(async () => {
  //   if (teamsUserCredential) {
  //     const userInfo = await teamsUserCredential.getUserInfo();
  //     return userInfo;
  //   }
  // });
  //const userName = loading || error ? "" : data!.displayName;
  const hubName = useData(async () => {
    await microsoftTeams.app.initialize();
    const context = await microsoftTeams.app.getContext();
    return context.app.host.name;
  })?.data;
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Welcome to Guest Search!</h1>
        <div className="fixed-position">
          {hubName && <p>Your app is running in {hubName}</p>}
          <p>Your app is running in your {friendlyEnvironmentName}</p>
        </div>
        {/* <Menu
          activeIndex={selectedMenuItem === "search" ? 0 : 1}
          items={menus}
          underlined
          primary
          accessibility={tabListBehavior}
        /> */}
        <div className="sections">
          {selectedMenuItem === "search" && (
            <Graph changeMenu={setSelectedMenu} />
          )}
          {/* {selectedMenuItem === "provision" && (
            <div>
              <Deploy />
            </div>
          )} */}

          {/* {selectedMenuItem === "local" && (
            <div>
              <EditCode showFunction={showFunction} />
              <CurrentUser userName={userName} />
              <Graph />
              {showFunction && <AzureFunctions />}
            </div>
          )}
          {selectedMenuItem === "azure" && (
            <div>
              <Deploy />
            </div>
          )}
          {selectedMenuItem === "publish" && (
            <div>
              <Publish />
            </div>
          )} */}
        </div>
      </div>
    </div>
  );
}
