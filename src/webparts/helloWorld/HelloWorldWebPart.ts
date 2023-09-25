import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "HelloWorldWebPartStrings";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

export interface IHelloWorldWebPartProps {
  description: string;
  Remisiones: any;
  DatosAI: any;
  ListCheck:any;
  Tablareglas:any;
  ListaValidaciom:any;
  Documentos:any;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        Tablareglas:this.properties.Tablareglas,
        description: this.properties.description,
        context: this.context,
        Remisiones: this.properties.Remisiones,
        DatosAI: this.properties.DatosAI,
        ListCheck:this.properties.ListCheck,
        ListaValidaciom:this.properties.ListaValidaciom,
        Documentos:this.properties.Documentos
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      spfi().using(SPFx(this.context));
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker("Remisiones", {
                  label: "Selecciona la lista de Remisiones",
                  selectedList: this.properties.Remisiones,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("DatosAI", {
                  label: "Selecciona la lista de DatosAI",
                  selectedList: this.properties.DatosAI,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId1",
                }),
                PropertyFieldListPicker("ListCheck", {
                  label: "Selecciona la lista de ListCheck",
                  selectedList: this.properties.ListCheck,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId2",
                }),
                PropertyFieldListPicker("Tablareglas", {
                  label: "Selecciona la lista de Tabla de reglas",
                  selectedList: this.properties.Tablareglas,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId3",
                }),
                PropertyFieldListPicker("ListaValidaciom", {
                  label: "Selecciona la lista de validación",
                  selectedList: this.properties.ListaValidaciom,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId4",
                }),
                PropertyFieldListPicker("Documentos", {
                  label: "Selecciona la lista de Documentos",
                  selectedList: this.properties.Documentos,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId5",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
