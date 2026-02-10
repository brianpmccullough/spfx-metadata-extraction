import type { ISharePointRestClient } from '../clients/ISharePointRestClient';
import type { ITerm } from '../models/fields/TaxonomyField';
import type { ITaxonomyService } from './ITaxonomyService';

interface ITermStoreTermResponse {
  id: string;
  labels: Array<{
    name: string;
    isDefault: boolean;
    languageTag: string;
  }>;
}

interface ITermStoreResponse {
  value: ITermStoreTermResponse[];
}

/**
 * Service for fetching taxonomy terms from SharePoint term store.
 */
export class TaxonomyService implements ITaxonomyService {
  constructor(private readonly _spoClient: ISharePointRestClient) {}

  public async getTerms(
    termSetId: string,
    siteUrl: string
  ): Promise<ITerm[]> {
    const url = `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms`;

    try {
      const response = await this._spoClient.get<ITermStoreResponse>(url);

      return response.value.map((term) => ({
        termGuid: term.id,
        label: this.getDefaultLabel(term.labels),
      }));
    } catch (error) {
      // If the v2.1 API fails, return empty array
      // This allows the field to still be created without terms
      console.warn(`Failed to load terms for term set ${termSetId}:`, error);
      return [];
    }
  }

  private getDefaultLabel(
    labels: Array<{ name: string; isDefault: boolean; languageTag: string }>
  ): string {
    const defaultLabel = labels.find((l) => l.isDefault);
    return defaultLabel?.name ?? labels[0]?.name ?? '';
  }
}
