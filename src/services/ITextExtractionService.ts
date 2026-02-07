export interface ITextExtractionResponse {
  content: string;
  format: 'markdown' | 'plain';
}

export interface ITextExtractionService {
  extractText(documentUrl: string): Promise<ITextExtractionResponse>;
}
