This package is externally forked off: [mondago](https://github.com/mondago/strapi-provider-email-msgraph) - I do not take any credits

# strapi-provider-email-msgraph

Microsoft Graph email provider plugin for Strapi 4.x.

## Prerequisites

An app registration for the tenant with Mail.Send permission is required. You'll need:

- Tenant ID
- Client App ID
- Client App Secret

## Installation

This package is scoped so you'll need to add an alias to your package.json.

```json
  "dependencies": {
    ...
    "strapi-provider-email-msgraph": "npm:strapi-email-msgraph"
    ...
  }
```

Then run either `yarn` or `npm install` (depending on which package manager you're using).

## Configuration

To use this provider setup your config/plugins.ts file:

```typescript
export default ({ env }) => ({
  email: {
    config: {
      provider: "strapi-provider-email-msgraph",
      providerOptions: {
        clientId: env("GRAPH_MAIL_CLIENT_ID"),
        clientSecret: env("GRAPH_MAIL_CLIENT_SECRET"),
        tenantId: env("GRAPH_MAIL_TENANT_ID"),
      },
      settings: {
        defaultFrom: "hello@example.com",
      },
    }
  },
});
```

## Support Matrix

| Our Version | Strapi Version |
|-------------|----------------|
| 1.x.x       | 4.x            |
