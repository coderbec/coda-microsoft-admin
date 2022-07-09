/*
Start making Packs!
Try out the hello world sample below to create your first build.
*/

// This import statement gives you access to all parts of the Coda Packs SDK.
import * as coda from "@codahq/packs-sdk";
import { User, Message, Group } from '@microsoft/microsoft-graph-types';
// This line creates your new Pack.
export const pack = coda.newPack();

const GroupSchema = coda.makeObjectSchema({
  type: coda.ValueType.Object,
  properties: {
    id: { description: "id", type: coda.ValueType.String },
    displayName: { description: "id", type: coda.ValueType.String },
    mailEnabled: { description: "id", type: coda.ValueType.Boolean },
    mail: { description: "id", type: coda.ValueType.String },
    description: { description: "id", type: coda.ValueType.String },

    // Add more properties here.
  },
  idProperty: "id",
  displayProperty: "displayName", // Which property above to display by default.
});
// Per-user authentication to the Microsoft Graph API, using an OAuth2 flow.
// See https://developer.todoist.com/guides/#oauth
pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    authorizationUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    tokenUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    scopes:  [
      '.default'
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
    // coda.makeParameter({
    //   type: coda.ParameterType.String,
    //   name: "name",
    //   description: "The name you would like to say hello to.",
    // }),
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

// Here, we add a new formula to this Pack.
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "Groups",
  description: "Print out all the organisational groups.",

  // If your formula requires one or more inputs, you’ll define them here.
  // Here, we're creating a string input called “name”.
  parameters: [
    // coda.makeParameter({
    //   type: coda.ParameterType.String,
    //   name: "name",
    //   description: "The name you would like to say hello to.",
    // }),
  ],

  // The resultType defines what will be returned in your Coda doc. Here, we're
  // returning a simple text string.
  resultType: coda.ValueType.Array,
  items: GroupSchema,
  // Everything inside this execute statement will happen anytime your Coda
  // formula is called in a doc. An array of all user inputs is always the 1st
  // parameter.
  execute: async ([], context) => {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: "https://graph.microsoft.com/v1.0/groups",
    });
    // Any following code won't run until the response is received.
    let groups = response.body.value;
    return groups;
  },

});

pack.addSyncTable({
  name: "Groups",
  schema: GroupSchema,
  identityName: "Group",
  formula: {
    name: "SyncGroups", 
    description: "Sync groups from Microsoft Graph",
    parameters: [],
    execute: async ([], context) => {
      let response = await context.fetcher.fetch({
        method: "GET",
        url: "https://graph.microsoft.com/v1.0/groups",
      });
      // Any following code won't run until the response is received.
      let groups = response.body.value;
      return {
        result: groups
      };
    },
  },
});