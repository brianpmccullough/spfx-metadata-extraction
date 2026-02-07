import { AadHttpClient, AadHttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

export interface IAadHttpClient {
  post<T>(url: string, body: unknown): Promise<T>;
}

export class AadHttpClientWrapper implements IAadHttpClient {
  constructor(private readonly _aadHttpClient: AadHttpClient) {}

  public async post<T>(url: string, body: unknown): Promise<T> {
    const options: IHttpClientOptions = {
      body: JSON.stringify(body),
      headers: {
        'Content-Type': 'application/json',
      },
    };

    const response: AadHttpClientResponse = await this._aadHttpClient.post(
      url,
      AadHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    return response.json() as Promise<T>;
  }
}
