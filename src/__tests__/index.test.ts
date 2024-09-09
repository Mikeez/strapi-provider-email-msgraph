// src/index.test.ts

import { Client } from '@microsoft/microsoft-graph-client';
import emailProviderModule from '../index'

jest.mock('@microsoft/microsoft-graph-client');

describe('Email Provider', () => {
    let emailProvider: any;
    const mockSendMail = jest.fn();

    beforeAll(() => {
        // Mock the Client's initWithMiddleware method
        (Client.initWithMiddleware as jest.Mock).mockReturnValue({
            api: jest.fn().mockReturnValue({
                post: mockSendMail,
            }),
        });

        // Initialize provider with mock options
        const providerOptions = {
            tenantId: 'tenantId',
            clientId: 'clientId',
            clientSecret: 'clientSecret',
        };

        const settings = {
            defaultFrom: 'default@example.com',
            defaultReplyTo: 'reply@example.com',
        };

        emailProvider = emailProviderModule.init(providerOptions, settings);
    });

    beforeEach(() => {
        mockSendMail.mockClear();
    });

    it('should send an email with the correct parameters', async () => {
        const emailOptions = {
            from: 'sender@example.com',
            to: ['recipient1@example.com', 'recipient2@example.com'],
            subject: 'Test Subject',
            text: 'Test email body',
            html: '<h1>Test email body</h1>',
        };

        await emailProvider.send(emailOptions);

        // Check that the post method was called correctly
        expect(mockSendMail).toHaveBeenCalledWith({
            message: {
                subject: emailOptions.subject,
                from: {
                    emailAddress: { address: emailOptions.from },
                },
                toRecipients: [
                    { emailAddress: { address: emailOptions.to[0] } },
                    { emailAddress: { address: emailOptions.to[1] } },
                ],
                body: {
                    content: emailOptions.html,
                    contentType: 'html',
                },
            },
        });
    });

    it('should use default from email if not provided', async () => {
        const emailOptions = {
            to: ['recipient@example.com'],
            subject: 'Test Subject',
            text: 'Test email body',
        };

        await emailProvider.send(emailOptions);

        // Check that the post method was called with default from email
        expect(mockSendMail).toHaveBeenCalledWith({
            message: {
                subject: emailOptions.subject,
                from: {
                    emailAddress: { address: 'default@example.com' }, // default from
                },
                toRecipients: [
                    { emailAddress: { address: emailOptions.to[0] } },
                ],
                body: {
                    content: emailOptions.text,
                    contentType: 'text',
                },
            },
        });
    });
});