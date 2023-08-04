"use strict";

/**
 * Module dependencies
 */

/* eslint-disable import/no-unresolved */
/* eslint-disable prefer-template */
// Public node modules.
require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");

module.exports = {
  provider: "msgraph",
  name: "Microsoft Graph Email Plugin",

  init: (providerOptions = {}, settings = {}) => {
    const authProvider = new TokenCredentialAuthenticationProvider(
      new ClientSecretCredential(
        providerOptions.tenantId,
        providerOptions.clientId,
        providerOptions.clientSecret
      ),
      { scopes: ["https://graph.microsoft.com/.default"] }
    );

    return {
      send: (options) => {
        const getEmailFromAddress = () => {
          if (!options.from) {
            return settings.defaultFrom;
          }

          const regex = /[^< ]+(?=>)/g;
          const matches = options.from.match(regex);
          return matches.length ? matches[0] : settings.defaultFrom;
        };

        return new Promise((resolve, reject) => {
          const client = Client.initWithMiddleware({
            debugLogging: false,
            authProvider: authProvider,
          });

          const from = getEmailFromAddress();
          const mail = {
            subject: options.subject,
            from: {
              emailAddress: { address: from },
            },
            toRecipients: [
              {
                emailAddress: {
                  address: options.to,
                },
              },
            ],
            attachments: options.attachments,
            body: options.html
              ? {
                  content: options.html,
                  contentType: "html",
                }
              : {
                  content: options.text,
                  contentType: "text",
                },
          };

          client
            .api(`/users/${from}/sendMail`)
            .post({ message: mail })
            .then(resolve)
            .catch(reject);
        });
      },
    };
  },
};
