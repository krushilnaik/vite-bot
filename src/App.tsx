import { PublicClientApplication } from "@azure/msal-browser";
import ReactWebChat, { createDirectLine, createStore, StyleOptions } from "botframework-webchat";
import { useEffect, useMemo, useRef, useState } from "react";
import { exchangeTokenAsync, fetchJSON, getOAuthCardResourceUri } from "./lib";

interface Props {
  clientId: string;
  tenantId: string;
  tokenExchangeURL: string;
  botName: string;
}

const sizes = {
  none: 3.5,
  half: 27,
  full: 50,
};

function App(props: Props) {
  const { clientId, tenantId, tokenExchangeURL, botName } = props;

  const [token, setToken] = useState("");
  const [size, setSize] = useState<"none" | "half" | "full">("full");
  const ref = useRef<HTMLDivElement>(null);

  // this tracks the secondary button shown on the header
  const [secondary, setSecondary] = useState<"min" | "max">("min");

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
              } else return next(action);
            });
          } else return next(action);
        } else return next(action);
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
      accent: "#d93954",
      bubbleMaxWidth: 400,
      bubbleBorderRadius: 8,
      botAvatarInitials: "CSC",
      bubbleFromUserBackground: "#d2d7e2",
      bubbleFromUserBorderRadius: 8,
      userAvatarBackgroundColor: "#415385",

      // input box and send button
      sendBoxButtonColor: "white",
      sendBoxButtonShadeColor: "#d93954",
      sendBoxButtonShadeColorOnHover: "#f04864",
      sendBoxBorderTop: "solid 1px lightgray",
      sendBoxBorderBottom: "solid 1px lightgray",
      sendBoxBorderLeft: "solid 1px lightgray",
      sendBoxBorderRight: "solid 1px lightgray",
      sendBoxButtonShadeBorderRadius: 8,

      // quick replies
      suggestedActionLayout: "flow",
      suggestedActionBackgroundColor: "transparent",
      suggestedActionBorderColor: "#a0a0a0",
      suggestedActionBorderRadius: 4,
      suggestedActionBorderWidth: 1,
      suggestedActionHeight: "2em",
      suggestedActionTextColor: "black",
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

  // fix jittering caused by scrollbar render on minimize
  useEffect(() => {
    if (size !== "none") {
      setTimeout(() => {
        if (ref.current) {
          ref.current.style.overflowY = "auto";
        }
      }, 200);
    } else {
      if (ref.current) {
        ref.current.style.overflowY = "hidden";
      }
    }
  }, [size]);

  // SSO into to chatbot
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
        filter: "drop-shadow(0 0 1px gray)",
        maxHeight: "60vh",
      }}
    >
      {/* header */}
      <div
        style={{
          display: "flex",
          position: "sticky",
          top: 0,
          zIndex: 9999,
          backgroundColor: "#d93954",
          borderRadius: "8px 8px 0 0",
          color: "white",
          justifyContent: "space-between",
          paddingInline: "16px",
        }}
      >
        <div style={{ height: "100%", display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" }}>
          <div
            style={{
              backgroundColor: "white",
              display: "grid",
              placeContent: "center",
              padding: 6,
              borderRadius: 9999,
            }}
          >
            ✨
          </div>
          <h1>{botName}</h1>
        </div>
        <ul style={{ display: "flex", listStyle: "none", alignItems: "center", gap: 8 }}>
          <li>
            <button
              onClick={() => {
                setSize(secondary === "max" ? "full" : "half");
                setSecondary(secondary === "max" ? "min" : "max");
              }}
              className="nav_button"
            >
              {secondary}
            </button>
          </li>
          <li>
            <button onClick={() => setSize(size === "none" ? "full" : "none")} className="nav_button">
              {size === "none" ? "^" : "v"}
            </button>
          </li>
        </ul>
      </div>

      {/* chat interface */}
      <div
        ref={ref}
        style={{
          backgroundColor: "white",
          overflowX: "hidden",
        }}
      >
        <div
          style={{
            paddingInline: "8px",
            minHeight: "calc(100% - 32px)",
            display: "flex",
            flexDirection: "column",
            justifyContent: "end",
          }}
        >
          {token && (
            <ReactWebChat
              directLine={directLine}
              store={store}
              styleOptions={styleOptions}
              // customize bot "typing" indicator
              sendTypingIndicator={true}
              typingIndicatorMiddleware={() =>
                (_next) =>
                ({ activeTyping }) => {
                  // @ts-ignore
                  activeTyping = Object.values(activeTyping);

                  if (activeTyping.length) {
                    const { role } = activeTyping[0];

                    if (role === "bot") {
                      return <span className="webchat__typing-indicator"></span>;
                    }
                  } else {
                    if (ref.current) {
                      ref.current.scrollTop = ref.current.scrollHeight;
                    }
                  }
                }}
            />
          )}
        </div>
        <div
          style={{
            backgroundColor: "white",
            position: "sticky",
            bottom: 0,
            paddingInline: "16px",
            paddingBlock: "0 16px",
            fontSize: "12px",
          }}
        >
          <span style={{ fontFamily: "Segoe UI" }}>
            AI-generated content may be inaccurate and requires human review.
          </span>
        </div>
      </div>

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
