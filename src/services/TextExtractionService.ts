import type { IAadHttpClient } from '../clients';
import type { ITextExtractionService, ITextExtractionResponse } from './ITextExtractionService';

export class TextExtractionService implements ITextExtractionService {
  constructor(
    private readonly _aadHttpClient: IAadHttpClient,
    private readonly _baseUrl: string
  ) {}

  public async extractText(documentUrl: string): Promise<ITextExtractionResponse> {
    return this._aadHttpClient.post<ITextExtractionResponse>(
      `${this._baseUrl}/api/extract/text`,
      { document: { path: documentUrl } }
    );
  }
}
