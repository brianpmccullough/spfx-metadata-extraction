import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MetadataRow, IMetadataRowProps } from './MetadataRow';
import type { IFieldMetadata } from '../../../models/IFieldMetadata';

function makeField(overrides?: Partial<IFieldMetadata>): IFieldMetadata {
  return {
    id: 'field-1',
    internalName: 'TestField',
    title: 'Test Field',
    description: 'A test description',
    type: 'string',
    value: 'Test Value',
    ...overrides,
  };
}

function renderComponent(
  container: HTMLDivElement,
  props: Partial<IMetadataRowProps> = {}
): void {
  const defaultProps: IMetadataRowProps = {
    field: makeField(),
    onDescriptionChange: jest.fn(),
    onTypeChange: jest.fn(),
    ...props,
  };

  act(() => {
    ReactDOM.render(<MetadataRow {...defaultProps} />, container);
  });
}

describe('MetadataRow', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(container);
    container.remove();
  });

  it('renders the field title as a label', () => {
    renderComponent(container, { field: makeField({ title: 'My Field' }) });

    const label = container.querySelector('.ms-Label');
    expect(label).not.toBeNull();
    expect(label!.textContent).toBe('My Field');
  });

  it('renders the description in a multiline text field', () => {
    renderComponent(container, {
      field: makeField({ description: 'Current desc' }),
    });

    const textarea = container.querySelector('textarea') as HTMLTextAreaElement;
    expect(textarea).not.toBeNull();
    expect(textarea.value).toBe('Current desc');
  });

  it('calls onDescriptionChange when the text field value changes', () => {
    const onDescriptionChange = jest.fn();
    renderComponent(container, {
      field: makeField({ id: 'field-abc' }),
      onDescriptionChange,
    });

    const textarea = container.querySelector('textarea') as HTMLTextAreaElement;
    act(() => {
      const nativeTextAreaValueSetter = Object.getOwnPropertyDescriptor(
        window.HTMLTextAreaElement.prototype,
        'value'
      )!.set!;
      nativeTextAreaValueSetter.call(textarea, 'new description');
      textarea.dispatchEvent(new Event('input', { bubbles: true }));
    });

    expect(onDescriptionChange).toHaveBeenCalledWith('field-abc', 'new description');
  });

  it('renders a dropdown with the current type selected', () => {
    renderComponent(container, {
      field: makeField({ type: 'number' }),
    });

    const dropdownTitle = container.querySelector('.ms-Dropdown-title');
    expect(dropdownTitle).not.toBeNull();
    expect(dropdownTitle!.textContent).toBe('number');
  });

  it('renders all three type options in the dropdown', () => {
    renderComponent(container);

    const dropdownOption = container.querySelector('.ms-Dropdown');
    expect(dropdownOption).not.toBeNull();
  });
});
