'use strict';

// src/index.ts

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

interface Settings {
    defaultFrom: string;
    defaultReplyTo: string;
}

interface SendOptions {
    from?: string;
    to: string[];
    cc?: string[];
    bcc?: string[];
    replyTo?: string;
    subject: string;
    text: string;
    html?: string;
    [key: string]: unknown;
}

interface ProviderOptions {
    tenantId: string;
    clientId: string;
    clientSecret: string;
}

export = {
    provider: 'msgraph',
    name: 'Microsoft Graph Email Plugin',
    init(providerOptions: ProviderOptions, settings: Settings) {
        const authProvider = new TokenCredentialAuthenticationProvider(
            new ClientSecretCredential(providerOptions.tenantId, providerOptions.clientId, providerOptions.clientSecret),
            { scopes: ['https://graph.microsoft.com/.default'] },
        );
        
        return {
            send: async (options: SendOptions) => {
                const getEmailFromAddress = () => {
                    if (!options.from) {
                        return settings.defaultFrom;
                    }
                    const regex = /[^< ]+(?=>)|([^< ]+@[^< ]+)/; // Adjusted regex to capture plain email format
                    const matches = options.from.match(regex);
                    return matches?.length ? matches[0].trim() : settings.defaultFrom;
                };
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
                }
                await client.api(`/users/${from}/sendMail`).post({ message: mail });
            }
        }
    }
}