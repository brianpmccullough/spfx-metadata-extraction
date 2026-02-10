import type { ITerm } from '../models/fields/TaxonomyField';

/**
 * Service interface for fetching taxonomy terms from SharePoint.
 * Abstracted for testability.
 */
export interface ITaxonomyService {
  /**
   * Fetches all terms from a term set.
   * @param termSetId - The GUID of the term set
   * @param siteUrl - The SharePoint site URL for API calls
   * @returns Array of terms with their GUIDs and labels
   */
  getTerms(termSetId: string, siteUrl: string): Promise<ITerm[]>;
}
