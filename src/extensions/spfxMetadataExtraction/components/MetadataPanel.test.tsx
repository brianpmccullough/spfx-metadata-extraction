import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MetadataPanel, IMetadataPanelProps } from './MetadataPanel';
import { FieldBase, StringField, NumericField, BooleanField } from '../../../models/fields';
import type { IDocumentContext } from '../../../models/IDocumentContext';
import type { ILlmExtractionService, IExtractionResponse } from '../../../services';

function makeFields(): FieldBase[] {
  return [
    new StringField('f1', 'TitleField', 'Title Field', 'desc 1', false, 'Sample Title'),
    new NumericField('f2', 'CountField', 'Count Field', 'desc 2', false, 42),
    new BooleanField('f3', 'ActiveField', 'Active Field', '', false, true),
  ];
}

function makeMockDocumentContext(): IDocumentContext {
  return {
    contentTypeId: '0x0101',
    fileExtension: 'docx',
    fileName: 'test-document.docx',
    fileLeafRef: 'test-document.docx',
    fileRef: '/sites/test/Documents/test-document.docx',
    fileSize: 1024,
    fileSizeInBytes: 1024,
    fileSizeInKiloBytes: 1,
    fileSizeInMegaBytes: 0.001,
    fileType: 'docx',
    serverRelativeUrl: '/sites/test/Documents/test-document.docx',
    itemId: 1,
    uniqueId: 'guid-1234',
    driveId: 'drive-id-1',
    driveItemId: 'item-id-1',
    spItemUrl: 'https://tenant.sharepoint.com/sites/test/_api/items/1',
    siteUrl: 'https://tenant.sharepoint.com/sites/test',
    webUrl: 'https://tenant.sharepoint.com/sites/test',
    siteId: 'site-guid',
    webId: 'web-guid',
    listId: 'list-guid',
  };
}

function makeMockLlmService(): ILlmExtractionService {
  return {
    extract: jest.fn().mockResolvedValue({
      document: { driveId: 'drive-id-1', driveItemId: 'item-id-1' },
      results: [
        { fieldName: 'Title Field', confidence: 'green', value: 'Extracted Title' },
        { fieldName: 'Count Field', confidence: 'green', value: 100 },
        { fieldName: 'Active Field', confidence: 'yellow', value: true },
      ],
    } as IExtractionResponse),
  };
}

async function flushPromises(): Promise<void> {
  // Multiple microtask flushes to handle chained promises
  await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

async function renderPanel(
  container: HTMLDivElement,
  props: Partial<IMetadataPanelProps> = {}
): Promise<void> {
  const defaultProps: IMetadataPanelProps = {
    loadFields: jest.fn().mockResolvedValue(makeFields()),
    documentContext: makeMockDocumentContext(),
    llmService: makeMockLlmService(),
    onDismiss: jest.fn(),
    onSave: jest.fn(),
    ...props,
  };

  await act(async () => {
    ReactDOM.render(<MetadataPanel {...defaultProps} />, container);
    await flushPromises();
  });
}

describe('MetadataPanel', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  });

  it('shows a spinner while loading', () => {
    act(() => {
      ReactDOM.render(
        <MetadataPanel
          loadFields={() => new Promise(() => { /* never resolves */ })}
          documentContext={makeMockDocumentContext()}
          llmService={makeMockLlmService()}
          onDismiss={jest.fn()}
          onSave={jest.fn()}
        />,
        container
      );
    });

    const spinner = container.querySelector('.ms-Spinner');
    expect(spinner).not.toBeNull();
  });

  it('renders a heading "Field Metadata" after loading', async () => {
    await renderPanel(container);

    const heading = container.querySelector('h2');
    expect(heading).not.toBeNull();
    expect(heading!.textContent).toBe('Field Metadata');
  });

  it('renders a row for each field', async () => {
    await renderPanel(container);

    const labels = container.querySelectorAll('.ms-Label');
    const labelTexts = Array.from(labels).map((l) => l.textContent);
    expect(labelTexts).toContain('Title Field');
    expect(labelTexts).toContain('Count Field');
    expect(labelTexts).toContain('Active Field');
  });

  it('renders Extract and Cancel buttons', async () => {
    await renderPanel(container);

    const buttons = container.querySelectorAll('.ms-Button');
    const buttonTexts = Array.from(buttons).map((b) => b.textContent);
    expect(buttonTexts).toContain('Extract');
    expect(buttonTexts).toContain('Cancel');
  });

  it('calls onDismiss when Cancel is clicked', async () => {
    const onDismiss = jest.fn();
    await renderPanel(container, { onDismiss });

    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const cancelButton = buttons.find((b) => b.textContent === 'Cancel') as HTMLElement;
    expect(cancelButton).toBeDefined();

    act(() => {
      cancelButton.click();
    });

    expect(onDismiss).toHaveBeenCalled();
  });

  it('calls llmService.extract when Extract is clicked', async () => {
    const llmService = makeMockLlmService();
    await renderPanel(container, { llmService });

    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const extractButton = buttons.find((b) => b.textContent === 'Extract') as HTMLElement;
    expect(extractButton).toBeDefined();

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    expect(llmService.extract).toHaveBeenCalledTimes(1);
  });

  it('renders field values using formatForDisplay', async () => {
    await renderPanel(container);

    expect(container.textContent).toContain('Sample Title');
    expect(container.textContent).toContain('42');
    expect(container.textContent).toContain('Yes');
  });

  it('shows an error message when loadFields rejects', async () => {
    const loadFields = jest.fn().mockRejectedValue(new Error('Network error'));
    await renderPanel(container, { loadFields });

    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('Network error');
  });

  it('shows a Close button on error', async () => {
    const loadFields = jest.fn().mockRejectedValue(new Error('fail'));
    const onDismiss = jest.fn();
    await renderPanel(container, { loadFields, onDismiss });

    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const closeButton = buttons.find((b) => b.textContent === 'Close') as HTMLElement;
    expect(closeButton).toBeDefined();

    act(() => {
      closeButton.click();
    });

    expect(onDismiss).toHaveBeenCalled();
  });

  it('renders extraction type dropdown for each field', async () => {
    await renderPanel(container);

    const dropdowns = container.querySelectorAll('.ms-Dropdown');
    expect(dropdowns.length).toBeGreaterThanOrEqual(3);
  });

  it('renders description textarea for each field', async () => {
    await renderPanel(container);

    const textareas = container.querySelectorAll('textarea');
    expect(textareas.length).toBe(3);
  });

  it('populates description from field description', async () => {
    await renderPanel(container);

    const textareas = container.querySelectorAll('textarea');
    const descriptions = Array.from(textareas).map((t) => t.value);
    expect(descriptions.some((d) => d.includes('desc 1'))).toBe(true);
    expect(descriptions.some((d) => d.includes('desc 2'))).toBe(true);
  });

  it('shows a friendly message when there are no fields', async () => {
    const loadFields = jest.fn().mockResolvedValue([]);
    await renderPanel(container, { loadFields });

    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('No fields available for metadata extraction');
  });
});
