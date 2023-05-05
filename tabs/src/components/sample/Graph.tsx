import "./Graph.css";
import { useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { Button, Input, Flex } from "@fluentui/react-northstar";
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import { PersonCardGrid } from "./PersonCardGrid";

export function Graph(props: { changeMenu?: Function }) {
  const [query, setQuery] = useState<string>();
  const [isEmailValid, setIsEmailValid] = useState<number>(0); // 0: not checked, 1: valid, 2: invalid
  const [queryState, setQueryState] = useState<number | undefined>(0);
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, error, data, reload } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      if (query !== undefined && query.length > 0 && !validateEmail(query)) {
        setIsEmailValid(2);
        return;
      } else {
        setIsEmailValid(1);
      }
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
            resultUsers.push(user);
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
    {
      scope: ["User.Read", "User.Read.All"],
      credential: teamsUserCredential,
    }
  );

  const validateEmail = (email: string) => {
    return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,10}$/i.test(email);
  };

  return (
    <div>
      <div className="center">
        <div>Enter your search term and Click Search</div>
        <Flex hAlign="center" gap="gap.small">
          <p>Email</p>
          <Input
            error={isEmailValid === 2}
            fluid
            type="email"
            placeholder={"user@example.com"}
            onChange={async (e, v) => {
              await setQuery(v?.value);
              await setQueryState(0);
            }}
            onKeyPress={(e) => {
              if (e.key === "Enter") {
                reload();
              }
            }}
            onBlur={(e) => {
              if (e.target.value.length > 0 && !validateEmail(e.target.value)) {
                setIsEmailValid(2);
              } else {
                setIsEmailValid(1);
              }
            }}
          />

          <Button
            style={{ marginTop: "8px" }}
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
