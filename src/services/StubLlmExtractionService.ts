import type {
  ILlmExtractionService,
  IExtractionRequest,
  IExtractionResponse,
  IExtractionResult,
  IDocumentDriveLocation,
} from './ILlmExtractionService';
import type { ExtractionConfidence } from '../models/extraction';
import { MetadataExtractionFieldType } from '../models/extraction';

/**
 * Stubbed LLM extraction service that simulates extraction with fake values.
 * Replace with real API implementation when the extraction API is available.
 */
export class StubLlmExtractionService implements ILlmExtractionService {
  private readonly _simulatedDelay: number;

  constructor(simulatedDelayMs: number = 1500) {
    this._simulatedDelay = simulatedDelayMs;
  }

  public async extract(request: IExtractionRequest): Promise<IExtractionResponse> {
    // Simulate network delay
    await this._delay(this._simulatedDelay);

    // Log request for debugging
    console.log('[StubLlmExtractionService] Extraction request:', {
      document: request.document,
      fields: request.fields.map((f) => f.title),
    });

    // Determine document identifier for simulation
    const docId = this._isDriveLocation(request.document)
      ? request.document.driveItemId
      : request.document.path;

    // Generate simulated extraction results
    const results: IExtractionResult[] = request.fields.map((field) => ({
      fieldName: field.title,
      confidence: this._generateConfidence(),
      value: this._generateSimulatedValue(field.dataType, field.title, docId),
    }));

    const document = this._isDriveLocation(request.document)
      ? { driveId: request.document.driveId, driveItemId: request.document.driveItemId }
      : { driveId: '', driveItemId: '' };

    return { document, results };
  }

  private _isDriveLocation(doc: IExtractionRequest['document']): doc is IDocumentDriveLocation {
    return 'driveId' in doc;
  }

  private _generateSimulatedValue(
    type: MetadataExtractionFieldType,
    fieldTitle: string,
    fileName: string
  ): string | number | boolean | null {
    // Simulate some fields not being extractable
    if (Math.random() < 0.15) {
      return null;
    }

    switch (type) {
      case MetadataExtractionFieldType.Number:
        return this._generateNumericValue(fieldTitle);
      case MetadataExtractionFieldType.Boolean:
        return Math.random() > 0.5;
      case MetadataExtractionFieldType.String:
      default:
        return this._generateStringValue(fieldTitle, fileName);
    }
  }

  private _generateNumericValue(fieldTitle: string): number {
    const lowerTitle = fieldTitle.toLowerCase();
    if (lowerTitle.includes('year')) {
      return 2020 + Math.floor(Math.random() * 6);
    }
    if (lowerTitle.includes('count') || lowerTitle.includes('number')) {
      return Math.floor(Math.random() * 100);
    }
    if (lowerTitle.includes('amount') || lowerTitle.includes('price') || lowerTitle.includes('cost')) {
      return Math.round(Math.random() * 10000 * 100) / 100;
    }
    if (lowerTitle.includes('percent') || lowerTitle.includes('rate')) {
      return Math.round(Math.random() * 1000) / 10;
    }
    return Math.floor(Math.random() * 1000);
  }

  private _generateStringValue(fieldTitle: string, fileName: string): string {
    const lowerTitle = fieldTitle.toLowerCase();

    if (lowerTitle.includes('title') || lowerTitle.includes('name')) {
      // Extract a meaningful title from filename
      const baseName = fileName.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ');
      return baseName.charAt(0).toUpperCase() + baseName.slice(1);
    }
    if (lowerTitle.includes('author') || lowerTitle.includes('owner')) {
      const authors = ['John Smith', 'Jane Doe', 'Alex Johnson', 'Sarah Williams', 'Mike Brown'];
      return authors[Math.floor(Math.random() * authors.length)];
    }
    if (lowerTitle.includes('department') || lowerTitle.includes('team')) {
      const depts = ['Engineering', 'Marketing', 'Sales', 'Finance', 'HR', 'Legal'];
      return depts[Math.floor(Math.random() * depts.length)];
    }
    if (lowerTitle.includes('status')) {
      const statuses = ['Draft', 'In Review', 'Approved', 'Published', 'Archived'];
      return statuses[Math.floor(Math.random() * statuses.length)];
    }
    if (lowerTitle.includes('category') || lowerTitle.includes('type')) {
      const categories = ['Report', 'Proposal', 'Contract', 'Memo', 'Policy', 'Procedure'];
      return categories[Math.floor(Math.random() * categories.length)];
    }
    if (lowerTitle.includes('date')) {
      const date = new Date();
      date.setDate(date.getDate() - Math.floor(Math.random() * 365));
      return date.toISOString().split('T')[0];
    }
    if (lowerTitle.includes('description') || lowerTitle.includes('summary')) {
      return `Extracted summary for ${fileName}`;
    }

    // Default: generate a generic extracted value
    return `Extracted: ${fieldTitle}`;
  }

  private _generateConfidence(): ExtractionConfidence {
    const rand = Math.random();
    if (rand < 0.15) return 'red';
    if (rand < 0.4) return 'yellow';
    return 'green';
  }

  private _delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
