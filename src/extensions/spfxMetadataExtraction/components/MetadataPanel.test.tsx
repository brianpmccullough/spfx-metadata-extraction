import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MetadataPanel, IMetadataPanelProps } from './MetadataPanel';
import { FieldBase, StringField, NumericField, BooleanField, TaxonomyField } from '../../../models/fields';
import type { ILlmExtractionService, IExtractionResponse } from '../../../services';
import { makeMockDocumentContext } from '../../../__test-utils__/fixtures';

function makeFields(): FieldBase[] {
  return [
    new StringField('f1', 'TitleField', 'Title Field', 'desc 1', false, 'Sample Title'),
    new NumericField('f2', 'CountField', 'Count Field', 'desc 2', false, 42),
    new BooleanField('f3', 'ActiveField', 'Active Field', '', false, true),
  ];
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
): Promise<() => void> {
  const defaultProps: IMetadataPanelProps = {
    loadFields: jest.fn().mockResolvedValue(makeFields()),
    documentContext: makeMockDocumentContext(),
    llmService: makeMockLlmService(),
    onDismiss: jest.fn(),
    onApply: jest.fn().mockResolvedValue(undefined),
    ...props,
  };

  await act(async () => {
    ReactDOM.render(<MetadataPanel {...defaultProps} />, container);
    await flushPromises();
  });

  return () => {
    ReactDOM.unmountComponentAtNode(container);
  };
}

describe('MetadataPanel', () => {
  let container: HTMLDivElement;
  let cleanup: (() => void) | undefined;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    if (cleanup) {
      act(() => {
        cleanup!();
      });
      cleanup = undefined;
    }
    container.remove();
  });

  it('shows a spinner while loading', async () => {
    cleanup = await renderPanel(container, {
      loadFields: () => new Promise(() => { /* never resolves */ }),
    });

    const spinner = container.querySelector('.ms-Spinner');
    expect(spinner).not.toBeNull();
  });

  it('renders a heading "Field Metadata" after loading', async () => {
    cleanup = await renderPanel(container);

    const heading = container.querySelector('h2');
    expect(heading).not.toBeNull();
    expect(heading!.textContent).toBe('Field Metadata');
  });

  it('renders a row for each field', async () => {
    cleanup = await renderPanel(container);

    const labels = container.querySelectorAll('.ms-Label');
    const labelTexts = Array.from(labels).map((l) => l.textContent);
    expect(labelTexts).toContain('Title Field');
    expect(labelTexts).toContain('Count Field');
    expect(labelTexts).toContain('Active Field');
  });

  it('renders Extract and Cancel buttons', async () => {
    cleanup = await renderPanel(container);

    const buttons = container.querySelectorAll('.ms-Button');
    const buttonTexts = Array.from(buttons).map((b) => b.textContent);
    expect(buttonTexts).toContain('Extract');
    expect(buttonTexts).toContain('Cancel');
  });

  it('calls onDismiss when Cancel is clicked', async () => {
    const onDismiss = jest.fn();
    cleanup = await renderPanel(container, { onDismiss });

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
    cleanup = await renderPanel(container, { llmService });

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
    cleanup = await renderPanel(container);

    expect(container.textContent).toContain('Sample Title');
    expect(container.textContent).toContain('42');
    expect(container.textContent).toContain('Yes');
  });

  it('shows an error message when loadFields rejects', async () => {
    const loadFields = jest.fn().mockRejectedValue(new Error('Network error'));
    cleanup = await renderPanel(container, { loadFields });

    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('Network error');
  });

  it('shows a Close button on error', async () => {
    const loadFields = jest.fn().mockRejectedValue(new Error('fail'));
    const onDismiss = jest.fn();
    cleanup = await renderPanel(container, { loadFields, onDismiss });

    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const closeButton = buttons.find((b) => b.textContent === 'Close') as HTMLElement;
    expect(closeButton).toBeDefined();

    act(() => {
      closeButton.click();
    });

    expect(onDismiss).toHaveBeenCalled();
  });

  it('renders extraction type dropdown for each field', async () => {
    cleanup = await renderPanel(container);

    const dropdowns = container.querySelectorAll('.ms-Dropdown');
    expect(dropdowns.length).toBeGreaterThanOrEqual(3);
  });

  it('renders description textarea for each field', async () => {
    cleanup = await renderPanel(container);

    const textareas = container.querySelectorAll('textarea');
    expect(textareas.length).toBe(3);
  });

  it('populates description from field description', async () => {
    cleanup = await renderPanel(container);

    const textareas = container.querySelectorAll('textarea');
    const descriptions = Array.from(textareas).map((t) => t.value);
    expect(descriptions.some((d) => d.includes('desc 1'))).toBe(true);
    expect(descriptions.some((d) => d.includes('desc 2'))).toBe(true);
  });

  it('shows a friendly message when there are no fields', async () => {
    const loadFields = jest.fn().mockResolvedValue([]);
    cleanup = await renderPanel(container, { loadFields });

    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('No fields available for metadata extraction');
  });

  it('shows extraction error inline while preserving field state', async () => {
    const llmService: ILlmExtractionService = {
      extract: jest.fn().mockRejectedValue(new Error('LLM service unavailable')),
    };
    cleanup = await renderPanel(container, { llmService });

    // Fields should be rendered
    const labels = container.querySelectorAll('.ms-Label');
    expect(Array.from(labels).map((l) => l.textContent)).toContain('Title Field');

    // Click Extract to trigger the error
    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const extractButton = buttons.find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    // Error should be shown inline
    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('LLM service unavailable');

    // Fields should still be visible
    const labelsAfter = container.querySelectorAll('.ms-Label');
    expect(Array.from(labelsAfter).map((l) => l.textContent)).toContain('Title Field');

    // Extract button should still be available for retry
    const buttonsAfter = Array.from(container.querySelectorAll('.ms-Button'));
    expect(buttonsAfter.map((b) => b.textContent)).toContain('Extract');
  });

  it('shows success message after apply succeeds', async () => {
    const onApply = jest.fn().mockResolvedValue(undefined);
    const llmService = makeMockLlmService();
    cleanup = await renderPanel(container, { llmService, onApply });

    // Click Extract first to populate results
    const extractButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    // Click Apply
    const applyButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Apply') as HTMLElement;

    await act(async () => {
      applyButton.click();
      await flushPromises();
    });

    // Success message should be shown
    const messageBars = container.querySelectorAll('.ms-MessageBar');
    const successBar = Array.from(messageBars).find((bar) =>
      bar.textContent?.includes('Document properties updated successfully')
    );
    expect(successBar).toBeDefined();
  });

  it('clears success message when Extract is clicked again', async () => {
    const onApply = jest.fn().mockResolvedValue(undefined);
    const llmService = makeMockLlmService();
    cleanup = await renderPanel(container, { llmService, onApply });

    // Extract -> Apply to get success state
    const extractButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    const applyButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Apply') as HTMLElement;

    await act(async () => {
      applyButton.click();
      await flushPromises();
    });

    // Success message should be shown
    expect(container.textContent).toContain('Document properties updated successfully');

    // Click Extract again
    const extractButton2 = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton2.click();
      await flushPromises();
    });

    // Success message should be cleared
    expect(container.textContent).not.toContain('Document properties updated successfully');
  });

  it('shows apply error inline while preserving field state', async () => {
    const onApply = jest.fn().mockRejectedValue(new Error('Permission denied'));
    const llmService = makeMockLlmService();
    cleanup = await renderPanel(container, { llmService, onApply });

    // Click Extract first to populate results
    const extractButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    // Click Apply to trigger the error
    const applyButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Apply') as HTMLElement;

    await act(async () => {
      applyButton.click();
      await flushPromises();
    });

    // Error should be shown inline
    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.textContent).toContain('Permission denied');

    // Fields should still be visible
    const labels = container.querySelectorAll('.ms-Label');
    expect(Array.from(labels).map((l) => l.textContent)).toContain('Title Field');
  });

  it('applies taxonomy field with label|termGuid format', async () => {
    const taxField = new TaxonomyField(
      'f-tax', 'DeptField', 'Department', '', false, null, 'ts1',
      [
        { termGuid: 'guid-exec', label: 'Executive Management' },
        { termGuid: 'guid-hr', label: 'HR' },
      ]
    );

    const loadFields = jest.fn().mockResolvedValue([taxField]);
    const llmService: ILlmExtractionService = {
      extract: jest.fn().mockResolvedValue({
        document: { driveId: 'drive-id-1', driveItemId: 'item-id-1' },
        results: [
          { fieldName: 'Department', confidence: 'green', value: 'Executive Management' },
        ],
      } as IExtractionResponse),
    };
    const onApply = jest.fn().mockResolvedValue(undefined);
    cleanup = await renderPanel(container, { loadFields, llmService, onApply });

    // Extract
    const extractButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    // Apply
    const applyButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Apply') as HTMLElement;

    await act(async () => {
      applyButton.click();
      await flushPromises();
    });

    expect(onApply).toHaveBeenCalledWith([
      { internalName: 'DeptField', value: 'Executive Management|guid-exec' },
    ]);
  });

  it('disables apply for taxonomy field with unmatched term', async () => {
    const taxField = new TaxonomyField(
      'f-tax', 'DeptField', 'Department', '', false, null, 'ts1',
      [
        { termGuid: 'guid-exec', label: 'Executive Management' },
        { termGuid: 'guid-hr', label: 'HR' },
      ]
    );

    const loadFields = jest.fn().mockResolvedValue([taxField]);
    const llmService: ILlmExtractionService = {
      extract: jest.fn().mockResolvedValue({
        document: { driveId: 'drive-id-1', driveItemId: 'item-id-1' },
        results: [
          { fieldName: 'Department', confidence: 'green', value: 'Unknown Department' },
        ],
      } as IExtractionResponse),
    };
    cleanup = await renderPanel(container, { loadFields, llmService });

    // Extract
    const extractButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Extract') as HTMLElement;

    await act(async () => {
      extractButton.click();
      await flushPromises();
    });

    // The Apply checkbox should be disabled (confidence is red)
    // The Apply button should be disabled because no checkboxes are checked
    const applyButton = Array.from(container.querySelectorAll('.ms-Button'))
      .find((b) => b.textContent === 'Apply') as HTMLElement;
    expect(applyButton.getAttribute('aria-disabled')).toBe('true');
  });
});
