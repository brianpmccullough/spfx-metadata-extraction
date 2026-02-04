import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import type { ISharePointRestClient } from './ISharePointRestClient';

export class SharePointRestClient implements ISharePointRestClient {
  constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async get<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} for ${url}`);
    }

    return response.json() as Promise<T>;
  }

  public async post<T>(
    url: string,
    body: unknown,
    headers?: Record<string, string>
  ): Promise<T> {
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(body),
      headers: headers,
    };

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} for ${url}`);
    }

    return response.json() as Promise<T>;
  }
}
