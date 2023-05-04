import "./Graph.css";
import { useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import {
  Button,
  Input,
  RadioGroup,
  Flex,
  Label,
} from "@fluentui/react-northstar";
import { Design } from "./Design";
import { PersonCardFluentUI } from "./PersonCardFluentUI";
import { PersonCardGraphToolkit } from "./PersonCardGraphToolkit";
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import { PersonCardGrid } from "./PersonCardGrid";

export function Graph(props: { changeMenu?: Function }) {
  const [query, setQuery] = useState<string>();
  const [queryType, setQueryType] = useState<string | number | undefined>(
    "mail"
  );
  const [users, setUsers] = useState<Array<String>>([]);
  const [queryState, setQueryState] = useState<number | undefined>(0);
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, error, data, reload } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      setQueryState(1);
      const profile = await graph.api("/me").get();

      let searchUsers: any = undefined;
      //let photoUrl = "";
      let resultUsers = [];
      try {
        // const photo = await graph.api("/me/photo/$value").get();
        // photoUrl = URL.createObjectURL(photo);
        if (query !== "") {
          let url =
            "/users?$filter=userType eq 'Guest' &$filter=mail eq '" +
            query +
            "'";
          searchUsers = await graph.api(url).get();
        }

        setQueryState(2);
        if (searchUsers && searchUsers.value) {
          for (let user of searchUsers.value) {
            resultUsers.push(user.id);
          }
          //setUsers(resultUsers);
        }

        // Initialize Graph Toolkit TeamsFx provider
        const provider = new TeamsFxProvider(teamsUserCredential, scope);
        Providers.globalProvider = provider;
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, resultUsers };
    },
    { scope: ["User.Read", "User.Read.All"], credential: teamsUserCredential }
  );

  return (
    <div>
      {/* <Design /> */}
      <div className="center">
        <div>Enter your search term and Click Search</div>
        <Flex hAlign="center" gap="gap.small">
          <Label content="Email" />
          <Input
            fluid
            type="text"
            placeholder={"user@example.com"}
            onChange={async (e, v) => {
              await setQuery(v?.value);
              await setQueryState(0);
            }}
          />

          <Button
            primary
            content="Search"
            disabled={loading}
            onClick={reload}
          />
        </Flex>
      </div>
      <div>
        <h4>Search Result(s)</h4>
        <PersonCardGrid
          loading={loading}
          data={data}
          error={error}
          query={query}
          queryState={queryState}
          changeMenu={props.changeMenu}
        />
      </div>
    </div>
  );
}
