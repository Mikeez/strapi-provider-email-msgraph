interface Settings {
    defaultFrom: string;
    defaultReplyTo: string;
}
interface Attachment {
    name: string;
    contentBytes: string;
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
    attachments?: Attachment[];
    [key: string]: unknown;
}
interface ProviderOptions {
    tenantId: string;
    clientId: string;
    clientSecret: string;
}
declare const _default: {
    provider: string;
    name: string;
    init(providerOptions: ProviderOptions, settings: Settings): {
        send: (options: SendOptions) => Promise<void>;
    };
};
export = _default;
