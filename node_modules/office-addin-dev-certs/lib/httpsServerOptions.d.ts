/// <reference types="node" />
interface IHttpsServerOptions {
    cert: Buffer;
    key: Buffer;
}
export declare function getHttpsServerOptions(): Promise<IHttpsServerOptions>;
export {};
