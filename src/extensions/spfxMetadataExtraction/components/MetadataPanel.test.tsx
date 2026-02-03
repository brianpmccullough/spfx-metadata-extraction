import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MetadataPanel, IMetadataPanelProps } from './MetadataPanel';
import type { IFieldMetadata } from '../../../models/IFieldMetadata';

function makeFields(): IFieldMetadata[] {
  return [
    { id: 'f1', title: 'Title Field', description: 'desc 1', type: 'string' },
    { id: 'f2', title: 'Count Field', description: 'desc 2', type: 'number' },
    { id: 'f3', title: 'Active Field', description: '', type: 'boolean' },
  ];
}

function flushPromises(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function renderPanel(
  container: HTMLDivElement,
  props: Partial<IMetadataPanelProps> = {}
): Promise<void> {
  const defaultProps: IMetadataPanelProps = {
    loadFields: jest.fn().mockResolvedValue(makeFields()),
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

  it('renders Save and Cancel buttons', async () => {
    await renderPanel(container);

    const buttons = container.querySelectorAll('.ms-Button');
    const buttonTexts = Array.from(buttons).map((b) => b.textContent);
    expect(buttonTexts).toContain('Save');
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

  it('calls onSave with field data when Save is clicked', async () => {
    const onSave = jest.fn();
    await renderPanel(container, { onSave });

    const buttons = Array.from(container.querySelectorAll('.ms-Button'));
    const saveButton = buttons.find((b) => b.textContent === 'Save') as HTMLElement;
    expect(saveButton).toBeDefined();

    act(() => {
      saveButton.click();
    });

    expect(onSave).toHaveBeenCalledWith(makeFields());
  });

  it('renders correct number of text inputs for descriptions', async () => {
    await renderPanel(container);

    const inputs = container.querySelectorAll('input');
    expect(inputs.length).toBeGreaterThanOrEqual(3);
  });

  it('shows an error message when loadFields rejects', async () => {
    const loadFields = jest.fn().mockRejectedValue(new Error('Network error'));
    await renderPanel(container, { loadFields });

    const messageBar = container.querySelector('.ms-MessageBar');
    expect(messageBar).not.toBeNull();
    expect(messageBar!.innerHTML).toContain('Network error');
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
});
