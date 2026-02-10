import type { IAadHttpClient } from '../clients/AadHttpClientWrapper';
import type {
  ILlmExtractionService,
  IExtractionRequest,
  IExtractionResponse,
} from './ILlmExtractionService';

export class LlmExtractionService implements ILlmExtractionService {
  constructor(
    private readonly _aadHttpClient: IAadHttpClient,
    private readonly _baseUrl: string
  ) {}

  public async extract(request: IExtractionRequest): Promise<IExtractionResponse> {
    return this._aadHttpClient.post<IExtractionResponse>(
      `${this._baseUrl}/api/extract/metadata`,
      request
    );
  }
}
