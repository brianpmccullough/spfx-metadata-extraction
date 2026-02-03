export interface IGraphClient {
  get<T>(path: string): Promise<T>;
}
