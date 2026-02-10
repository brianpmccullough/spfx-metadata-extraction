import type { IFieldSchema, ExtractionConfidence } from '../models/extraction';

/**
 * Document location specified by path.
 */
export interface IDocumentPathLocation {
  /** Absolute or server-relative path to the document */
  path: string;
}

/**
 * Document location specified by Graph drive identifiers.
 */
export interface IDocumentDriveLocation {
  /** The Graph API drive ID */
  driveId: string;
  /** The Graph API drive item ID */
  driveItemId: string;
}

/**
 * Request payload for LLM metadata extraction.
 */
export interface IExtractionRequest {
  /** Document location - specify either path OR driveId/driveItemId */
  document: IDocumentPathLocation | IDocumentDriveLocation;
  /** Array of metadata fields to extract */
  fields: IFieldSchema[];
}

/**
 * Single field extraction result from LLM.
 */
export interface IExtractionResult {
  /** The field name (corresponds to the title sent in the request schema) */
  fieldName: string;
  /** Confidence level of the extraction */
  confidence: ExtractionConfidence;
  /** The extracted value (or null if not found) */
  value: string | number | boolean | null;
}

/**
 * Response from LLM metadata extraction.
 */
export interface IExtractionResponse {
  /** Echo of the document identifiers from the request */
  document: IDocumentDriveLocation;
  results: IExtractionResult[];
}

/**
 * Service interface for LLM-based metadata extraction.
 */
export interface ILlmExtractionService {
  /**
   * Extracts metadata values from a document using LLM.
   * @param request The extraction request with document details and schema
   * @returns Promise resolving to extraction results
   */
  extract(request: IExtractionRequest): Promise<IExtractionResponse>;
}
