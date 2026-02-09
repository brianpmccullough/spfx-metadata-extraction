import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MetadataRow, IMetadataRowProps } from './MetadataRow';
import { StringField, BooleanField, NumericField } from '../../../models/fields';
import { MetadataExtractionField, MetadataExtractionFieldType } from '../../../models/extraction';

function renderComponent(
  container: HTMLDivElement,
  props: Omit<IMetadataRowProps, 'applyChecked' | 'onApplyCheckedChange' | 'isApplyEnabled'> & Partial<Pick<IMetadataRowProps, 'applyChecked' | 'onApplyCheckedChange' | 'isApplyEnabled'>>
): void {
  const fullProps: IMetadataRowProps = {
    applyChecked: false,
    onApplyCheckedChange: jest.fn(),
    isApplyEnabled: false,
    ...props,
  };
  act(() => {
    ReactDOM.render(<MetadataRow {...fullProps} />, container);
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
    const field = new StringField('id', 'TestField', 'My Field', '', false, 'test');
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    const label = container.querySelector('.ms-Label');
    expect(label).not.toBeNull();
    expect(label!.textContent).toBe('My Field');
  });

  it('renders field value using formatForDisplay', () => {
    const field = new StringField('id', 'Notes', 'Notes', '', false, 'Hello World');
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    expect(container.textContent).toContain('Hello World');
  });

  it('renders (empty) for null value', () => {
    const field = new StringField('id', 'Notes', 'Notes', '', false, null);
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    expect(container.textContent).toContain('(empty)');
  });

  it('renders Yes for true boolean value', () => {
    const field = new BooleanField('id', 'Active', 'Active', '', false, true);
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    expect(container.textContent).toContain('Yes');
  });

  it('renders No for false boolean value', () => {
    const field = new BooleanField('id', 'Active', 'Active', '', false, false);
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    expect(container.textContent).toContain('No');
  });

  it('renders a dropdown for extraction type', () => {
    const field = new StringField('id', 'Notes', 'Notes', '', false, 'test');
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    const dropdown = container.querySelector('.ms-Dropdown');
    expect(dropdown).not.toBeNull();
  });

  it('renders a multiline text field for description', () => {
    const field = new StringField('id', 'Notes', 'Notes', 'Field description', false, 'test');
    const extractionField = new MetadataExtractionField(field);
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange: jest.fn(),
    });

    const textarea = container.querySelector('textarea');
    expect(textarea).not.toBeNull();
    expect(textarea!.value).toBe('Field description');
  });

  it('calls onDescriptionChange when description is edited', () => {
    const field = new StringField('id', 'Notes', 'Notes', '', false, 'test');
    const extractionField = new MetadataExtractionField(field);
    const onDescriptionChange = jest.fn();
    renderComponent(container, {
      extractionField,
      onExtractionTypeChange: jest.fn(),
      onDescriptionChange,
    });

    const textarea = container.querySelector('textarea') as HTMLTextAreaElement;
    act(() => {
      textarea.value = 'New description';
      textarea.dispatchEvent(new Event('input', { bubbles: true }));
    });

    expect(onDescriptionChange).toHaveBeenCalledWith('New description');
  });

  it('infers extraction type as Number for NumericField', () => {
    const field = new NumericField('id', 'Count', 'Count', '', false, 42);
    const extractionField = new MetadataExtractionField(field);

    expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.Number);
  });

  it('infers extraction type as Boolean for BooleanField', () => {
    const field = new BooleanField('id', 'Active', 'Active', '', false, true);
    const extractionField = new MetadataExtractionField(field);

    expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.Boolean);
  });

  it('infers extraction type as String for StringField', () => {
    const field = new StringField('id', 'Notes', 'Notes', '', false, 'test');
    const extractionField = new MetadataExtractionField(field);

    expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.String);
  });
});
