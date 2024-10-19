'use strict';
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// src/index.ts
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const identity_1 = require("@azure/identity");
const azureTokenCredentials_1 = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
module.exports = {
    provider: 'msgraph',
    name: 'Microsoft Graph Email Plugin',
    init(providerOptions, settings) {
        const authProvider = new azureTokenCredentials_1.TokenCredentialAuthenticationProvider(new identity_1.ClientSecretCredential(providerOptions.tenantId, providerOptions.clientId, providerOptions.clientSecret), { scopes: ['https://graph.microsoft.com/.default'] });
        return {
            send: (options) => __awaiter(this, void 0, void 0, function* () {
                const getEmailFromAddress = () => {
                    if (!options.from) {
                        return settings.defaultFrom;
                    }
                    const regex = /[^< ]+(?=>)|([^< ]+@[^< ]+)/; // Adjusted regex to capture plain email format
                    const matches = options.from.match(regex);
                    return (matches === null || matches === void 0 ? void 0 : matches.length) ? matches[0].trim() : settings.defaultFrom;
                };
                const client = microsoft_graph_client_1.Client.initWithMiddleware({
                    debugLogging: false,
                    authProvider: authProvider,
                });
                const from = getEmailFromAddress();
                const mail = {
                    subject: options.subject,
                    from: {
                        emailAddress: { address: from },
                    },
                    toRecipients: options.to.map(email => ({
                        emailAddress: {
                            address: email,
                        },
                    })),
                    ccRecipients: options.cc ? options.cc.map(email => ({
                        emailAddress: {
                            address: email,
                        },
                    })) : [],
                    bccRecipients: options.bcc ? options.bcc.map(email => ({
                        emailAddress: {
                            address: email,
                        },
                    })) : [],
                    body: options.html
                        ? {
                            content: options.html,
                            contentType: 'html',
                        }
                        : {
                            content: options.text,
                            contentType: 'text',
                        },
                };
                yield client.api(`/users/${from}/sendMail`).post({ message: mail });
            })
        };
    }
};
//# sourceMappingURL=index.js.map