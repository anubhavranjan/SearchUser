import { Flex, Grid, Loader, Image, Button } from "@fluentui/react-northstar";
import { ProfileCard } from "./ProfileCard";

export function PersonCardGrid(props: {
  loading?: boolean;
  error?: any;
  query?: string;
  queryState?: number;
  changeMenu?: Function;
  data?:
    | {
        profile: any;
        resultUsers: any[];
      }
    | undefined;
}) {
  let users: JSX.Element[] = [];
  if (!props.loading && props.data && props.data.resultUsers) {
    users = props.data.resultUsers.map((user) => {
      return (
        <>{user !== null && ProfileCard(false, props.data?.profile, user)}</>
      );
    });
  }
  return (
    <div className="section">
      {props.loading && props.queryState === 2 && (
        <>
          <Loader label="Loading..." />
        </>
      )}
      {!props.loading && props.error && (
        <div className="error">
          Failed to read your profile. Please try again later. <br /> Details:{" "}
          {props.error.toString()}
        </div>
      )}
      {!props.loading && props.data && props.data.resultUsers && (
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
