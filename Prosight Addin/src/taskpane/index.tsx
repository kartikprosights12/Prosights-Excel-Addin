import "./index.css";
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";


const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;


/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title="ProSights Excel Assistant" />
    </FluentProvider>
  );
});
// Define a type for module with hot property
interface HotModule {
  hot?: {
    accept: (path: string, callback: () => void) => void;
  };
}

declare const module: HotModule;

if (module.hot) {
  module.hot.accept("./components/App", () => {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const NextApp = require("./components/App").default;
    root?.render(<NextApp />);
  });
}
