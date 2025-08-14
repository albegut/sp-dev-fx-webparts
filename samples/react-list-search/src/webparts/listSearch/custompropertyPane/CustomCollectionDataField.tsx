import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { IBaseFieldData, IListData, IMappingFieldData } from '../model/IListConfigProps';

export const GetCustomCollectionDropDown = (
  options: IPropertyPaneDropdownOption[],
  field: ICustomCollectionField,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  row: any,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  updateFunction: (fieldId: string, value: any) => void,
  errorFunction?: (fieldId: string, value: string) => void,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  customOnchangeFunction?: (row: any, fieldId: string, option: IDropdownOption<any>, updateFunction: (fieldId: string, value: any) => void, errorFunction: (fieldId: string, value: string) => void) => void
): JSX.Element => {
  return (<Dropdown placeholder={field.placeholder || field.title}
    options={options.sort((a, b) => { return a.text.localeCompare(b.text); })}
    selectedKey={row[field.id] || null}
    required={field.required}
    onChange={(evt, option) => customOnchangeFunction ? customOnchangeFunction(row, field.id, option, updateFunction, errorFunction) : updateFunction(field.id, option.key)}
    onRenderOption={field.onRenderOption}
    className="PropertyFieldCollectionData__panel__dropdown-field" />);
}

export const GetPickerByStringOptions = (
  possibleOptions: Array<string>,
  field: ICustomCollectionField,
  row: IMappingFieldData | IBaseFieldData | IListData,
  updateFunction: (fieldId: string, value: string | number) => void,
  customOnChange: (row: IMappingFieldData | IBaseFieldData | IListData, fieldId: string, option: IDropdownOption<string>, updateFunction: (fieldId: string, value: string | number) => void, errorFunction: (fieldId: string, value: string) => void) => void,
  customError?: (fieldId: string, value: string) => void
): JSX.Element => {
  let options: IPropertyPaneDropdownOption[] = [];
  if (possibleOptions) {
    options = possibleOptions.map(option => { return { key: option, text: option }; });
  }
  return GetCustomCollectionDropDown(options, field, row, updateFunction, customError, customOnChange);
}
