import * as React from "react";
import { Checkbox } from "@fluentui/react";
import { ICustomCollectionField } from "@pnp/spfx-property-controls";
import { IDetailListFieldData } from "../model/IListConfigProps";
import { getKeyValue } from "../utils/ObjectUtils";

export const getDisabledCheckBox = (
    field: ICustomCollectionField,
    item: IDetailListFieldData,
    updateFunction: (fieldId: string, value: boolean) => void
): JSX.Element => {
    const itemValue = getKeyValue(item, field.id as keyof IDetailListFieldData) as boolean;
    return <Checkbox checked={itemValue ? itemValue : false}
        onChange={(ev, value) => updateFunction(field.id, value)}
        disabled={true}
        className="PropertyFieldCollectionData__panel__boolean-field" />;
}