import { PublicClientApplication } from "@azure/msal-browser";
import { createDirectLine, createStore, StyleOptions } from "botframework-webchat";
import { useEffect, useMemo, useState } from "react";
import { exchangeTokenAsync, fetchJSON, getOAuthCardResourceUri } from "./lib";
import ReactWebChat from "botframework-webchat";
import { IconChevronUp, IconChevronDown, IconCpu } from "@tabler/icons-react";

interface Props {
  clientId: string;
  tenantId: string;
  tokenExchangeURL: string;
  botName: string;
}

const sizes = {
  none: 3.5,
  half: 20,
  full: 50,
};

function App(props: Props) {
  const { clientId, tenantId, tokenExchangeURL, botName } = props;
  const [token, setToken] = useState("");
  const [size, setSize] = useState<"none" | "half" | "full">("full");

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

  const directLine = useMemo(() => {
    if (token) {
      return createDirectLine({ token });
    }

    return null;
  }, [token]);

  const store = useMemo(() => {
    if (token) {
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
                  ?.postActivity({
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
              } else {
                const temp = next(action);
                console.log(temp);

                return temp;
              }
            });
          } else {
            const temp = next(action);
            console.log(temp);

            return temp;
          }
        } else {
          const temp = next(action);
          console.log(temp);

          return temp;
        }
      });
    }

    return null;
  }, [token]);

  const styleOptions: StyleOptions = useMemo(
    () => ({
      hideUploadButton: true,
      backgroundColor: "white",
      bubbleBackground: "#f3f3f3",
      bubbleBorderWidth: 0,
      rootWidth: 500,
      rootHeight: "100%",
      bubbleMaxWidth: 400,
      bubbleBorderRadius: 8,
      botAvatarBackgroundColor: "red",
      botAvatarInitials: "CSC",
      bubbleFromUserBackground: "#d2d7e2",
      bubbleFromUserBorderRadius: 8,
      userAvatarBackgroundColor: "#415385",
      sendBoxButtonColor: "white",
      sendBoxButtonShadeColor: "#d93954",
      sendBoxButtonShadeColorOnHover: "#f04864",
      suggestedActionLayout: "flow",
      suggestedActionBackgroundColor: "transparent",
      suggestedActionBorderRadius: 4,
      suggestedActionBorderWidth: 1,
      suggestedActionHeight: "2em",
      suggestedActionTextColor: "black",
      sendBoxBorderTop: "solid 1px lightgray",
      sendBoxBorderBottom: "solid 1px lightgray",
      sendBoxBorderLeft: "solid 1px lightgray",
      sendBoxBorderRight: "solid 1px lightgray",
      sendBoxButtonShadeBorderRadius: 8,
    }),
    []
  );

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

  return (
    <div
      style={{
        display: "grid",
        gridTemplateRows: `${sizes["none"]}rem 1fr`,
        transition: "height linear 200ms",
        height: `${sizes[size]}rem`,
        maxHeight: "60vh",
        borderRadius: "8px 8px 0 0",
        overflowX: "hidden",
        overflowY: size === "none" ? "hidden" : "auto",
      }}
    >
      <div
        style={{
          display: "flex",
          position: "sticky",
          top: 0,
          zIndex: 9999,
          backgroundColor: "#d93954",
          color: "white",
          justifyContent: "space-between",
          paddingInline: "16px",
        }}
      >
        {/* header */}
        <div style={{ height: "100%", display: "grid", gridTemplateColumns: "auto 1fr", gap: 8, alignItems: "center" }}>
          <div
            style={{
              backgroundColor: "white",
              display: "grid",
              placeContent: "center",
              padding: 6,
              borderRadius: 9999,
            }}
          >
            <IconCpu stroke={2} color="#d93954" />
          </div>
          <h1>{botName}</h1>
        </div>
        <ul style={{ display: "flex", listStyle: "none", alignItems: "center", gap: 8 }}>
          {/* <li>
            <button onClick={() => setSize("full")} className="nav_button">
              <IconArrowsMaximize stroke={2} />
            </button>
          </li> */}
          <li>
            <button onClick={() => setSize(size === "none" ? "full" : "none")} className="nav_button">
              {size === "none" ? <IconChevronUp stroke={2} /> : <IconChevronDown stroke={2} />}
            </button>
          </li>
        </ul>
      </div>

      {/* chat interface */}
      {token && <ReactWebChat directLine={directLine} store={store} styleOptions={styleOptions} />}

      {/* <span>AI generated content may be inaccurate and requires human review.</span> */}

      {/* add bot name to timestamp of responses */}
      <style>
        {`.webchat__activity-status:not(.webchat__activity-status--self) span::after {
          content: " • ${botName}";
        }`}
      </style>
    </div>
  );
}

export default App;
