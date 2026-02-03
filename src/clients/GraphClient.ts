import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { IGraphClient } from './IGraphClient';

export class GraphClient implements IGraphClient {
  constructor(private readonly _graphClient: MSGraphClientV3) {}

  public async get<T>(path: string): Promise<T> {
    return this._graphClient.api(path).get() as Promise<T>;
  }
}
