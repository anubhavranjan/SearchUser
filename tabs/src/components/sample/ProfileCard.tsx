import React from "react";
import {
  Avatar,
  Card,
  CardBody,
  Flex,
  Skeleton,
  Text,
  Button,
  Divider,
  Accordion,
} from "@fluentui/react-northstar";
import { EmailIcon, ChatIcon } from "@fluentui/react-icons-northstar";
import { chat } from "@microsoft/teams-js";

export const ProfileCard = (
  loading: boolean,
  profile?: any,
  data?: any
  //chatFn?: Function
) => (
  <Card
    aria-roledescription="card avatar"
    elevated
    ghost
    styles={{
      height: "max-content",
      margin: "0.5em 0",
      width: "340px",
      background: "#faf9f8",
      borderRadius: "8px",
    }}
  >
    <Card.Header styles={{ "margin-bottom": "0" }}>
      {loading && (
        <Skeleton animation="wave">
          <Flex gap="gap.medium">
            <Skeleton.Avatar size="larger" />
            <div>
              <Skeleton.Line width="100px" />
              <Skeleton.Line width="150px" />
            </div>
          </Flex>
        </Skeleton>
      )}
      {!loading && data && (
        <>
          <Flex gap="gap.medium">
            <Avatar
              size="larger"
              image={data.photoUrl}
              name={data.displayName}
            />{" "}
            <Text content={data.displayName} size="large" weight="semibold" />
          </Flex>
          <div className="base-icons">
            <a href={"mailto:" + data.mail}>
              <Button
                icon={<EmailIcon />}
                content="Send email"
                text
                title="Send email"
              />
            </a>
            <Button
              icon={<ChatIcon />}
              text
              content="Start chat"
              onClick={async () => {
                if (chat.isSupported()) {
                  const chatPromise = chat.openChat({
                    user: data.userPrincipalName,
                  });
                  chatPromise
                    .then((result) => console.log("then: ", result))
                    .catch((error) => console.log(error));
                }
              }}
              title={"Start a chat with " + data.displayName}
            />
          </div>
        </>
      )}
    </Card.Header>
    <Card.Body>
      <Accordion
        panels={[
          {
            title: <Divider content="Contact" />,
            content: (
              <>
                <div style={{ marginLeft: "-26px" }}>
                  <Button
                    icon={<EmailIcon />}
                    content="Email"
                    text
                    size="small"
                    title="Email"
                    disabled
                    style={{ color: "#484644" }}
                  />
                  <br />
                  <a href={"mailto:" + data.mail}>
                    <Button
                      content={data.mail}
                      text
                      title={data.mail}
                      size="small"
                    />
                  </a>
                  <br />
                  <br />
                  <Button
                    icon={<ChatIcon />}
                    content="Teams"
                    text
                    size="small"
                    title="Teams"
                    disabled
                    style={{ color: "#484644" }}
                  />
                  <br />
                  <Button
                    text
                    content={data.mail}
                    title={data.mail}
                    size="small"
                    onClick={async () => {
                      console.log(
                        "https://teams.microsoft.com/l/chat/0/0?users=" +
                          data.userPrincipalName
                      );
                      if (chat.isSupported()) {
                        const chatPromise = chat.openChat({
                          user: data.userPrincipalName,
                        });
                        chatPromise
                          .then((result) => console.log("then: ", result))
                          .catch((error) => console.log(error));
                      }
                    }}
                  />
                </div>
              </>
            ),
          },
        ]}
      />
    </Card.Body>
  </Card>
);
