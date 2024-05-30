import React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./index.css";
import { FluentProvider, teamsLightTheme } from "@fluentui/react-components";

const container = document.getElementById("root");
const root = createRoot(container!);
root.render(
    <FluentProvider theme={teamsLightTheme}>
        <App />
    </FluentProvider>,
);
