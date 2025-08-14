import { ICustomCollectionField } from "@pnp/spfx-property-controls";
import { IBaseFieldData, ICustomOption, ListField } from "../model/IListConfigProps";
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { IDropdownOption } from "@fluentui/react";
import { GetCustomCollectionDropDown } from "./CustomCollectionDataField";

export const GetFieldPickerByList = (
    possibleOptions: Array<ListField>,
    field: ICustomCollectionField,
    row: IBaseFieldData,
    updateFunction: (fieldId: string, value: string) => void,
    customOnChange: (row: IBaseFieldData, fieldId: string, option: IDropdownOption<string>, updateFunction: (fieldId: string, value: string) => void, errorFunction: (fieldId: string, value: string) => void) => void,
    customOptions?: Array<ICustomOption>
  ): JSX.Element => {
    let options: IPropertyPaneDropdownOption[] = [];
    if (possibleOptions) {
      options = possibleOptions.map(option => { return { key: option.InternalName, text: option.Title, title: option.InternalName, FieldType: option.TypeAsString }; });
    }
    if (customOptions) {
      customOptions.map(option => {
        options.push({
          key: option.Key,
          text: option.Option,
        });
      });
    }
    return GetCustomCollectionDropDown(options, field, row, updateFunction, null, customOnChange);
}