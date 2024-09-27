import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.tsx";
import "./index.css";

const config = {
  clientId: "81f42388-549a-4f97-b7e4-a2e9b9d83c11",
  tenantId: "0cb87174-c7ed-4063-a036-cc8a3c4ee938",
  tokenExchangeURL:
    "https://default0cb87174c7ed4063a036cc8a3c4ee9.38.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cr756_ssoBot/directline/token?api-version=2022-03-01-preview",
};

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App {...config} />
  </StrictMode>
);
