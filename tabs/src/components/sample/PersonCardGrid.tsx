// import Team from "../../models/Team";
// import { TeamCard } from "./TeamCard";
import { PersonCard } from "@microsoft/mgt-react";
import {
  Flex,
  FlexItem,
  Grid,
  Loader,
  Image,
  Button,
} from "@fluentui/react-northstar";
import { PersonCardGraphToolkit } from "./PersonCardGraphToolkit";

export function PersonCardGrid(props: {
  loading?: boolean;
  error?: any;
  query?: string;
  queryState?: number;
  changeMenu?: Function;
  data?:
    | {
        profile: any;
        resultUsers: string[];
      }
    | undefined;
}) {
  let users: JSX.Element[] = [];
  if (!props.loading && props.data && props.data.resultUsers) {
    users = props.data.resultUsers.map((user) => {
      return (
        <PersonCard
          className="custom-card"
          style={{ margin: "0.5em" }}
          key={user}
          userId={user}
          isExpanded={false}
        />
        // <Flex.Item shrink={false} size="340" styles={{ width: "340" }}>
        //   <PersonCard key={user} userId={user} isExpanded={false} />
        // </Flex.Item>
      );
    });
  }
  return (
    <div className="section">
      {props.loading && props.queryState === 2 && (
        <>
          <Loader label="Loading..." />
          {/* {<PersonCard loading={true} data={undefined} />} */}
        </>
      )}
      {!props.loading && props.error && (
        <div className="error">
          Failed to read your profile. Please try again later. <br /> Details:{" "}
          {props.error.toString()}
        </div>
      )}
      {!props.loading && props.data && props.data.resultUsers && (
        // <Flex gap="gap.smaller" fill={false}>
        //   {users}
        // </Flex>
        <Grid
          columns="repeat(2, 1fr)"
          styles={{ width: "100%" }}
          content={users}
        />
      )}
      {!props.loading &&
        props.data &&
        props.data.resultUsers &&
        props.data.resultUsers.length === 0 &&
        props.query &&
        props.queryState === 2 && (
          // <Flex gap="gap.smaller" fill={false}>
          //   {users}
          // </Flex>
          <Flex column hAlign="center">
            <Image src="not-found.svg" style={{ maxWidth: "20.4rem" }} />
            <h3>We couldn't find any results for '{props.query}'</h3>
            <span>
              Would you like to submit request for provisioning of your Guest
              User?{" "}
              <Button
                tinted
                content="Submit"
                onClick={(e, v) =>
                  props.changeMenu && props.changeMenu("provision")
                }
              />
            </span>
          </Flex>
        )}
    </div>
  );
}
