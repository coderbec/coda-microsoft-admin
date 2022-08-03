/*
Start making Packs!
Try out the hello world sample below to create your first build.
*/

// This import statement gives you access to all parts of the Coda Packs SDK.
import * as coda from "@codahq/packs-sdk";
import { User, Message, Group, GroupType } from '@microsoft/microsoft-graph-types';
// This line creates your new Pack.
export const pack = coda.newPack();

const GroupSchema = coda.makeObjectSchema({
  type: coda.ValueType.Object,
  properties: {
    id: { description: "ID", type: coda.ValueType.String },
    displayName: { description: "Display Name", type: coda.ValueType.String },
    mailEnabled: { description: "Mail Enabled", type: coda.ValueType.Boolean },
    mail: { description: "Mail", type: coda.ValueType.String },
    description: { description: "Description", type: coda.ValueType.String },

    // Add more properties here.
  },
  idProperty: "id",
  displayProperty: "displayName", // Which property above to display by default.
});

const UserSchema = coda.makeObjectSchema({
  type: coda.ValueType.Object,
  properties: {
    id: { description: "id", type: coda.ValueType.String },
    displayName: { description: "id", type: coda.ValueType.String },
    mail: { description: "id", type: coda.ValueType.String },
    givenName: { description: "id", type: coda.ValueType.String },
    surname: { description: "id", type: coda.ValueType.String },
    jobTitle: { description: "id", type: coda.ValueType.String },
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

//add a column format tor Group names in order to facilitate an RBAC
// pack.addColumnFormat({
//   name: "groupName",
//   instructions: "Use this to display group names from Microsoft 365. If you enter a Group Name and a user the column will show if the user has access to that group.",
//   formulaName: "groupAccessForUser",
//   matchers: [
//     // If formatting a URL, a regular expression that matches that URL.
//   ],
// });

// Does the User have access to the supplied group
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "groupAccessForUser",
  description: "Is a specified user a member of a supplied group?",

  // If your formula requires one or more inputs, you’ll define them here.
  // Here, we're creating a string input called “name”.
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "groupId",
      description: "The UUID of the group.",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "userId",
      description: "The UUID of the user.",
    }),
  ],

  // The resultType defines what will be returned in your Coda doc. Here, we're
  // returning a simple text string.
  resultType: coda.ValueType.Boolean,
  // Everything inside this execute statement will happen anytime your Coda
  // formula is called in a doc. An array of all user inputs is always the 1st
  // parameter.
  execute: async ([groupId, userId], context) => {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: `https://graph.microsoft.com/v1.0/groups/${groupId}/members`,
    });
    // Any following code won't run until the response is received.
    let users: [Object] = response.body.value;
    var ids = users.map(function(a: User) {return a.id;});
    if(ids.includes(userId)){
      return true;
    }
    return false;
  },
});


// Here, we add a new formula to this Pack.
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "Group",
  description: "Get properties and relationship of group.",

  // If your formula requires one or more inputs, you’ll define them here.
  // Here, we're creating a string input called “name”.
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "id",
      description: "The UUID of the group.",
    }),
  ],

  // The resultType defines what will be returned in your Coda doc. Here, we're
  // returning a simple text string.
  resultType: coda.ValueType.Object,
  schema: GroupSchema,
  // Everything inside this execute statement will happen anytime your Coda
  // formula is called in a doc. An array of all user inputs is always the 1st
  // parameter.
  execute: async ([id], context) => {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: "https://graph.microsoft.com/v1.0/groups/" + id,
    });
    // Any following code won't run until the response is received.
    let group: GroupType = response.body;
    return group;
  },

});


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

// Get all the groups the current user belongs to
pack.addFormula({
  // This is the name that will be called in the formula builder.
  // Remember, your formula name cannot have spaces in it.
  name: "currentUserGroups",
  description: "Print out all the groups the current user is enrolled in.",

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
      url: "https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group?$count=true",
    });
    // Any following code won't run until the response is received.
    let groups: GroupType = response.body.value;
    return groups;
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

pack.addSyncTable({
  name: "Users",
  schema: UserSchema,
  identityName: "User",
  formula: {
    name: "SyncUsers", 
    description: "Sync user from Microsoft Graph",
    parameters: [],
    execute: async ([], context) => {
      let response = await context.fetcher.fetch({
        method: "GET",
        url: "https://graph.microsoft.com/v1.0/users",
      });
      // Any following code won't run until the response is received.
      let users = response.body.value;
      return {
        result: users
      };
    },
  },
});

// Schema for a Channels
const ChannelSchema = coda.makeObjectSchema({
  properties: {
    displayName: {
      type: coda.ValueType.String,
    },
    description: {
      type: coda.ValueType.String,
    },
    webUrl: {
      type: coda.ValueType.String,
      codaType: coda.ValueHintType.Url,
    },
    id: { type: coda.ValueType.String },
  },
  displayProperty: "displayName",
  idProperty: "id",
  featuredProperties: ["webUrl"],
});

//https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')


// list all channels that are a part of a team, uses a dynamic autocompete to accept the team name (with id as parameter)
pack.addSyncTable({
  name: "Channels",
  identityName: "Channel",
  schema: ChannelSchema,
  connectionRequirement: coda.ConnectionRequirement.None,
  formula: {
    name: "SyncChannels",
    description: "Display Channels that are a part of a team.",
    parameters: [
      coda.makeParameter({
        type: coda.ParameterType.String,
        name: "team",
        description: "Only Channels from this team will be shown - UUID",
        optional: false,
        // Pull the list of tags to use for autocomplete from the API.
        // autocomplete: async function (context, search) {
        //   let response = await context.fetcher.fetch({
        //     method: "GET",
        //     url: "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')",
        //   });
        //   let tags = response.body.value;
        //   // Convert the tags into a list of autocomplete options.
        //   return coda.autocompleteSearchObjects(search, tags, "displayName", "id");
        // },
      }),
    ],
    execute: async function ([team], context) {
      let url = `https://graph.microsoft.com/v1.0/teams/${team}/channels`
      let response = await context.fetcher.fetch({
        method: "GET",
        url: url,
      });
      let channels = response.body.value;
      return {
        result: channels,
      };
    },
  },
});

// Schema for a Channels
const TeamsSchema = coda.makeObjectSchema({
  properties: {
    displayName: {
      type: coda.ValueType.String,
    },
    description: {
      type: coda.ValueType.String,
    },
    // webUrl: {
    //   type: coda.ValueType.String,
    //   codaType: coda.ValueHintType.Url,
    // },
    id: { type: coda.ValueType.String },
  },
  displayProperty: "displayName",
  idProperty: "id"
});


pack.addSyncTable({
  name: "Teams",
  schema: TeamsSchema,
  identityName: "Team",
  formula: {
    name: "SyncTeams", 
    description: "Get all Teams for the org",
    parameters: [],
    execute: async ([], context) => {
      let response = await context.fetcher.fetch({
        method: "GET",
        url: "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')",
      });
      // Any following code won't run until the response is received.
      let teams = response.body.value;
      return {
        result: teams
      };
    },
  },
});


// "accountEnabled": true,
//     "city": "Seattle",
//     "country": "United States",
//     "department": "Sales & Marketing",
//     "mailNickname": "MelissaD",
//     "passwordPolicies": "DisablePasswordExpiration",
//     "passwordProfile": {
//         "password": "94d1ebe9-aa89-c2a7-8fc8-deb326dc8df2",
//         "forceChangePasswordNextSignIn": false
//     },
//     "officeLocation": "131/1105",
//     "postalCode": "98052",
//     "preferredLanguage": "en-US",
//     "state": "WA",
//     "streetAddress": "9256 Towne Center Dr., Suite 400",
//     "mobilePhone": "+1 206 555 0110",
//     "usageLocation": "US",
//     "userPrincipalName": "MelissaD@{domain}"
// Action formula (for buttons and automations) that adds a new task in Todoist.
pack.addFormula({
  name: "AddUser",
  description: "Add a new User.",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "givenName",
      description: "First Name",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "surname",
      description: "Last Name",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "displayName",
      description: "displayName",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "jobTitle",
      description: "jobTitle",
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,

  execute: async function ([name], context) {
    let response = await context.fetcher.fetch({
      url: "https://graph.microsoft.com/v1.0/users",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name
      }),
    });
    // Return values are optional but recommended. Returning a URL or other
    // unique identifier is recommended when creating a new entity.
    return response.body.url;
  },
});
