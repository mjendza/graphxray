import { Stack } from "@fluentui/react/lib/Stack";
import { DefaultPalette } from "@fluentui/react";
import { FontSizes } from "@fluentui/theme";
import { CommandMenu } from "./CommandMenu.js";

// Set at build time by the dedicated `npm run build:dev` script (via DefinePlugin).
// Undefined in production builds, so the DEV badge never renders there.
const isDevBuild = process.env.REACT_APP_DEV_BUILD === "true";

const devBadgeStyle = {
  marginLeft: 8,
  padding: "1px 6px",
  borderRadius: 4,
  background: DefaultPalette.orangeLight,
  color: DefaultPalette.white,
  fontSize: FontSizes.size12,
  fontWeight: 600,
  textTransform: "uppercase",
  letterSpacing: 0.5,
};

// Non-mutating styles definition
const stackItemStyles = {
  root: {
    alignItems: "center",
    background: DefaultPalette.themeDarker,
    color: DefaultPalette.white,
    display: "flex",
    height: 50,
    justifyContent: "center",
    overflow: "hidden",
  },
};
const nonShrinkingStackItemStyles = {
  root: {
    alignItems: "center",
    background: DefaultPalette.themeDarker,
    color: DefaultPalette.white,
    height: 50,
    display: "flex",
    justifyContent: "start",
    overflow: "hidden",
    paddingLeft: "10px",
  },
};

export const AppHeader = ({ hideSettings }) => {
  return (
    <Stack horizontal>
      <Stack.Item grow styles={stackItemStyles}>
        <img alt="logo" src="./img/icon-16.svg" height="22px"></img>
      </Stack.Item>
      <Stack.Item grow={10} styles={nonShrinkingStackItemStyles} disableShrink>
        <div
          style={{
            fontSize: FontSizes.size16,
            fontWeight: 600,
            display: "flex",
            alignItems: "center",
          }}
        >
          Microsoft Graph X-Ray
          {isDevBuild && <span style={devBadgeStyle}>DEV Build</span>}
        </div>
      </Stack.Item>
      <Stack.Item grow styles={stackItemStyles}>
        {!hideSettings && <CommandMenu></CommandMenu>}
      </Stack.Item>
    </Stack>
  );
};
