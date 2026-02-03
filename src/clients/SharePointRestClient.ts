import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { ISharePointRestClient } from './ISharePointRestClient';

export class SharePointRestClient implements ISharePointRestClient {
  constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async get<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`: HTTP ${response.status} for ${url}`);
    }

    return response.json() as Promise<T>;
  }
}
