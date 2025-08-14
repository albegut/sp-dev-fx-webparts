import { ICustomCollectionField } from "@pnp/spfx-property-controls";
import { IBaseFieldData, IListData, IRedirectData, SiteList } from "../model/IListConfigProps";
import { IDropdownOption } from "@fluentui/react";
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { GetCustomCollectionDropDown } from "./CustomCollectionDataField";


export const GetListPickerBySiteOptions = (
    possibleOptions: Array<IListData>,
    field: ICustomCollectionField,
    row: IRedirectData | IBaseFieldData,
    updateFunction: (fieldId: string, value: string) => void,
    customOnChange?: (row: IRedirectData | IBaseFieldData, fieldId: string, option: IDropdownOption<string>, updateFunction: (fieldId: string, value: string) => void, errorFunction: (fieldId: string, value: string) => void) => void,
): JSX.Element => {
    const currentOptions: IPropertyPaneDropdownOption[] = [];
    possibleOptions.filter(option => {
        if (row.SiteCollectionSource && option.SiteCollectionSource == row.SiteCollectionSource) {
            currentOptions.push({
                key: option.ListSourceField,
                text: option.ListSourceFieldName
            });
        }
    });
    return GetCustomCollectionDropDown(currentOptions, field, row, updateFunction, null, customOnChange);
}

export const GetListPicker = (
    possibleOptions: Array<SiteList>,
    field: ICustomCollectionField,
    row: IListData,
    updateFunction: (fieldId: string, value: string | number) => void,
    customOnChange: (row: IListData, fieldId: string, option: IDropdownOption<string>, updateFunction: (fieldId: string, value: string | number) => void, errorFunction: (fieldId: string, value: string) => void) => void,
    customError?: (fieldId: string, value: string) => string
): JSX.Element => {
    let options: IPropertyPaneDropdownOption[] = [];
    if (possibleOptions) {
        options = possibleOptions.map(option => { return { key: option.Id, text: option.Title }; });
    }
    return GetCustomCollectionDropDown(options, field, row, updateFunction, customError, customOnChange);
}