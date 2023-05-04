import React from "react";
import "./Deploy.css";
import { Embed, Flex, Image } from "@fluentui/react-northstar";

export function Deploy(props: { docsUrl?: string }) {
  const { docsUrl } = {
    docsUrl: "https://aka.ms/teamsfx-docs",
    ...props,
  };
  return (
    <Embed
      active
      iframe={{
        src: "https://www.bing.com",
        height: "100%",
        width: "100%",
      }}
      title="Microsoft Teams"
    />
  );
}
