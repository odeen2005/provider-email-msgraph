# provider-email-msgraph
Microsoft Graph email provider plugin for Strapi 4

## Prerequisites

An app registration for the tenant with Mail.Send permission is required. You'll need:

- Tenant ID
- Client App ID
- Client App Secret

## Installation

This package is scoped so you'll need to add an alias to your package.json. Replace `<version>` with the version number of your choice (eg 1.0.0).

```json
  "dependencies": {
    ...
    "@strapi/provider-email-msgraph": "npm:@odeen2005/provider-email-msgraph@<version>"
    ...
  }
```

Then run either `yarn` or `npm install` (depending on which package manager you're using).

## Configuration

To use this provider setup your config/plugins.js file:

```javascript
module.exports = ({ env }) => ({
  email: {
    provider: "msgraph",
    providerOptions: {
      clientId: env("GRAPH_MAIL_CLIENT_ID"),
      clientSecret: env("GRAPH_MAIL_CLIENT_SECRET"),
      tenantId: env("GRAPH_MAIL_TENANT_ID"),
    },
    settings: {
      defaultFrom: "hello@example.com",
    },
  },
});
```