export interface ISharePointRestClient {
  get<T>(url: string): Promise<T>;
  post<T>(url: string, body: unknown, headers?: Record<string, string>): Promise<T>;
}
