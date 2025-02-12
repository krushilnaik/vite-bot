import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.tsx";
import "./index.css";

const config = {
  botName: "React SSO Bot",
  clientId: "81f42388-549a-4f97-b7e4-a2e9b9d83c11",
  tenantId: "0cb87174-c7ed-4063-a036-cc8a3c4ee938",
  tokenExchangeURL: "https://directline.botframework.com/v3/directline/tokens/generate",
};

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App {...config} />
  </StrictMode>
);
