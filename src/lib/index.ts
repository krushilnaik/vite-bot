import { PublicClientApplication } from "@azure/msal-browser";
import { Activity } from "botframework-schema";

export function getOAuthCardResourceUri({ attachments }: Activity) {
  if (
    attachments &&
    attachments[0]?.contentType === "application/vnd.microsoft.card.oauth" &&
    attachments[0].content.tokenExchangeResource
  ) {
    // asking for token exchange with AAD
    return attachments[0].content.tokenExchangeResource.uri;
  }
}

export async function fetchJSON(url: string, options: RequestInit = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      ...options.headers,
      accept: "application/json",
    },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch JSON due to ${res.status}`);
  }

  return await res.json();
}

export function exchangeTokenAsync(client: PublicClientApplication, resourceUri: string) {
  let user = client.getAllAccounts()[0];

  if (user) {
    let requestObj = {
      scopes: [resourceUri],
    };

    client.setActiveAccount(user);

    return client
      .acquireTokenSilent(requestObj)
      .then(({ accessToken }) => accessToken)
      .catch(console.error);
  } else {
    return Promise.resolve(null);
  }
}
