export interface ISharePointRestClient {
  get<T>(url: string): Promise<T>;
}
