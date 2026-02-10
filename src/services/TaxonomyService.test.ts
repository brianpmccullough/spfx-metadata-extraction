import { TaxonomyService } from './TaxonomyService';
import type { ISharePointRestClient } from '../clients';

describe('TaxonomyService', () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/TestSite';

  const makeMockSPOClient = (response?: unknown): ISharePointRestClient => ({
    get: jest.fn().mockResolvedValue(response ?? { value: [] }),
    post: jest.fn().mockResolvedValue({}),
  });

  describe('getTerms', () => {
    it('constructs the correct REST URL', async () => {
      const spoClient = makeMockSPOClient({ value: [] });
      const service = new TaxonomyService(spoClient);

      await service.getTerms('termset-guid-123', siteUrl);

      expect(spoClient.get).toHaveBeenCalledWith(
        'https://contoso.sharepoint.com/sites/TestSite/_api/v2.1/termStore/sets/termset-guid-123/terms'
      );
    });

    it('maps term response to ITerm array', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          {
            id: 'term-guid-1',
            labels: [{ name: 'Finance', isDefault: true, languageTag: 'en-US' }],
          },
          {
            id: 'term-guid-2',
            labels: [{ name: 'HR', isDefault: true, languageTag: 'en-US' }],
          },
        ],
      });
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms).toHaveLength(2);
      expect(terms[0]).toEqual({ termGuid: 'term-guid-1', label: 'Finance' });
      expect(terms[1]).toEqual({ termGuid: 'term-guid-2', label: 'HR' });
    });

    it('uses default label when multiple labels exist', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          {
            id: 'term-guid-1',
            labels: [
              { name: 'Alternate', isDefault: false, languageTag: 'en-US' },
              { name: 'Default Label', isDefault: true, languageTag: 'en-US' },
            ],
          },
        ],
      });
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms[0].label).toBe('Default Label');
    });

    it('uses first label when no default is marked', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          {
            id: 'term-guid-1',
            labels: [
              { name: 'First Label', isDefault: false, languageTag: 'en-US' },
              { name: 'Second Label', isDefault: false, languageTag: 'en-US' },
            ],
          },
        ],
      });
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms[0].label).toBe('First Label');
    });

    it('returns empty array when no terms exist', async () => {
      const spoClient = makeMockSPOClient({ value: [] });
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms).toEqual([]);
    });

    it('returns empty array on API error', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn().mockRejectedValue(new Error('HTTP 403: Forbidden')),
        post: jest.fn().mockResolvedValue({}),
      };
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms).toEqual([]);
    });

    it('handles empty labels array gracefully', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          {
            id: 'term-guid-1',
            labels: [],
          },
        ],
      });
      const service = new TaxonomyService(spoClient);

      const terms = await service.getTerms('termset-guid', siteUrl);

      expect(terms[0].label).toBe('');
    });
  });
});
