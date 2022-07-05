/*
Start making Packs!
Try out the hello world sample below to create your first build.
*/

// This import statement gives you access to all parts of the Coda Packs SDK.
import * as coda from "@codahq/packs-sdk";

// This line creates your new Pack.
export const pack = coda.newPack();

// Per-user authentication to the Microsoft Graph API, using an OAuth2 flow.
// See https://developer.todoist.com/guides/#oauth
pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    authorizationUrl: "https://todoist.com/oauth/authorize",
    tokenUrl: "https://todoist.com/oauth/access_token",
    scopes: ["data:read_write"],
  
    // Determines the display name of the connected account.
    getConnectionName: async function (context) {
      let url = coda.withQueryParams("https://api.todoist.com/sync/v8/sync", {
        resource_types: JSON.stringify(["user"]),
      });
      let response = await context.fetcher.fetch({
        method: "GET",
        url: url,
      });
      return response.body.user?.full_name;
    },
  });
  

// Here, we add a new formula to this Pack.
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "Hello",
  description: "A Hello World example.",

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
  execute: async function ([name], context) {
    return "Hello " + name + "!";
  },
});
