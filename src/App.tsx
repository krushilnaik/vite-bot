import { PublicClientApplication } from "@azure/msal-browser";
import { createDirectLine, createStore, StyleOptions } from "botframework-webchat";
import { useEffect, useMemo, useState } from "react";
import { exchangeTokenAsync, fetchJSON, getOAuthCardResourceUri } from "./lib";
import ReactWebChat from "botframework-webchat";

interface Props {
  clientId: string;
  tenantId: string;
  tokenExchangeURL: string;
}

function App(props: Props) {
  const { clientId, tenantId, tokenExchangeURL } = props;
  const [token, setToken] = useState("");

  let msalConfig = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: true,
    },
  };

  const client = useMemo(() => {
    return new PublicClientApplication(msalConfig);
  }, []);

  const directLine = useMemo(() => createDirectLine({ token }), [token]);

  const store = useMemo(() => {
    return createStore({}, ({ dispatch }) => (next) => (action) => {
      let userdetails = client.getAllAccounts()[0];
      let userId = userdetails?.localAccountId
        ? ("sso-chatbot" + userdetails.localAccountId).substring(0, 64)
        : (Math.random().toString() + Date.now().toString()).substring(0, 64);

      if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
        dispatch({
          type: "WEB_CHAT/SEND_EVENT",
          payload: {
            name: "startConversation",
            type: "event",
            value: { text: userdetails.name },
          },
        });
        return next(action);
      }

      if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
        const activity = action.payload.activity;
        let resourceUri;
        if (activity.from && activity.from.role === "bot" && (resourceUri = getOAuthCardResourceUri(activity))) {
          exchangeTokenAsync(client, resourceUri).then(function (token) {
            if (token) {
              directLine
                .postActivity({
                  // @ts-ignore
                  type: "invoke",
                  name: "signin/tokenExchange",
                  value: {
                    id: activity.attachments[0].content.tokenExchangeResource.id,
                    connectionName: activity.attachments[0].content.connectionName,
                    token,
                  },
                  from: {
                    id: userId,
                    name: userdetails.name,
                    role: "user",
                  },
                })
                .subscribe(
                  (id) => {
                    if (id === "retry") {
                      // bot was not able to handle the invoke, so display the oauthCard
                      return next(action);
                    }
                    // else: tokenexchange successful and we do not display the oauthCard
                  },
                  (_error) => {
                    // an error occurred to display the oauthCard
                    return next(action);
                  }
                );
              return;
            } else return next(action);
          });
        } else return next(action);
      } else return next(action);
    });
  }, [token]);

  const init = async () => {
    await client.initialize();

    const account = client.getAllAccounts()[0];
    if (!account) {
      await client
        .loginPopup({
          scopes: ["user.read", "openid", "profile"],
        })
        .catch(console.error);
    }

    fetchJSON(tokenExchangeURL).then(({ token }) => setToken(token));
  };

  useEffect(() => {
    init();
  }, []);

  const styleOptions: StyleOptions = {
    hideUploadButton: true,
  };

  // return <Chatbot token={token} store={store} />;
  return token ? <ReactWebChat directLine={directLine} store={store} styleOptions={styleOptions} /> : <div />;
}

export default App;
