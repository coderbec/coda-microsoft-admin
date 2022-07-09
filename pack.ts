/*
Start making Packs!
Try out the hello world sample below to create your first build.
*/

// This import statement gives you access to all parts of the Coda Packs SDK.
import * as coda from "@codahq/packs-sdk";
import { User, Message } from '@microsoft/microsoft-graph-types';
// This line creates your new Pack.
export const pack = coda.newPack();

// Per-user authentication to the Microsoft Graph API, using an OAuth2 flow.
// See https://developer.todoist.com/guides/#oauth
pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    authorizationUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    tokenUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    scopes:  [
      'user.read',
      'mail.read',
      'mail.send'
    ],
  
     // Additional parameters to ensure a refresh_token is returned.
    additionalParams: {
      prompt: "consent",
    },

    // Enable PKCE (optional but recommended).
    useProofKeyForCodeExchange: true,

    // Determines the display name of the connected account.
    getConnectionName: async function (context) {
      let response = await context.fetcher.fetch({
        method: "GET",
        url: "https://graph.microsoft.com/v1.0/me",
      });
      let user = response.body;
      return user.displayName;
    },
  });
  
// Allow the pack to make requests to Microsoft.
pack.addNetworkDomain("microsoft.com");

// Here, we add a new formula to this Pack.
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "User",
  description: "Print out current user information.",

  // If your formula requires one or more inputs, you’ll define them here.
  // Here, we're creating a string input called “name”.
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "name",
      description: "The name you would like to say hello to.",
    }),
  ],

  // The resultType defines what will be returned in your Coda doc. Here, we're
  // returning a simple text string.
  resultType: coda.ValueType.String,

  // Everything inside this execute statement will happen anytime your Coda
  // formula is called in a doc. An array of all user inputs is always the 1st
  // parameter.
  execute: async ([], context) => {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: "https://graph.microsoft.com/v1.0/me",
    });
    // Any following code won't run until the response is received.
    let user: User = response.body;
    return user.displayName;
  },

});
