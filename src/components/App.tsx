// https://fluentsite.z22.web.core.windows.net/quick-start
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  Spinner,
  tokens,
} from "@fluentui/react-components";
import { HashRouter as Router, Navigate, Route, Routes } from "react-router-dom";
import { useTeamsUserCredential } from "@microsoft/teamsfx-react";
import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import Tab from "./Tab";
import { TeamsFxContext } from "./Context";
import config from "./sample/lib/config";
import RaiseRequest from "./RaiseRequest";
import { InteractiveBrowserCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { MyRequests } from "./MyRequests";
import { Dummy } from "./Dummy";
import { API_SCOPES } from "../Constants";
import { Login } from "@microsoft/mgt-react";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { loading, theme, themeString, teamsUserCredential } = useTeamsUserCredential({
    initiateLoginEndpoint: config.initiateLoginEndpoint!,
    clientId: config.clientId!,
  });

  const credential = new InteractiveBrowserCredential({
    tenantId: "b67fae7e-909b-42b3-a571-9fc7262d1439",
    clientId: "0be5fed3-eaaa-4ce8-a04a-957d94c18b47",
    redirectUri: "https://localhost:53000/auth-end.html?clientId=0be5fed3-eaaa-4ce8-a04a-957d94c18b47"
  });

  // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
  const authProvider = new TokenCredentialAuthenticationProvider(
    credential,
    {
      scopes: API_SCOPES,
    });

  const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

  return (
    <TeamsFxContext.Provider value={{ theme, themeString, teamsUserCredential }}>
      <FluentProvider
        theme={
          themeString === "dark"
            ? teamsDarkTheme
            : themeString === "contrast"
              ? teamsHighContrastTheme
              : {
                ...teamsLightTheme,
                colorNeutralBackground3: "#eeeeee",
              }
        }
        style={{ background: tokens.colorNeutralBackground3 }}
      >
        <Router>
          {loading ? (
            <Spinner style={{ margin: 100 }} />
          ) : (
            <Routes>
              <Route path="/privacy" element={<Privacy />} />
              <Route path="/termsofuse" element={<TermsOfUse />} />
              <Route path="/raiseRequest" element={<RaiseRequest graphClient={graphClient} />} />
              <Route path="/myRequests" element={<MyRequests graphClient={graphClient} />} />
              <Route path="/dummy" element={<Dummy graphClient={graphClient} />} />
              <Route path="/tab" element={<Tab />} />
              <Route path="*" element={<Navigate to={"/tab"} />}></Route>
            </Routes>
          )}
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}
