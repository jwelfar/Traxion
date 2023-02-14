import * as React from "react";
import { IHelloWorldProps } from "./IHelloWorldProps";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
} from "office-ui-fabric-react";

let _sp: SPFI = null;

export interface IDetailsTableItem {
  NO_ORDEN_REPOSICION_UNOPS: string;
  NO_REMISION: string;
  NO_LICITACION: string;
  NO_CONTRATO: string;
  PROCEDENCIA: string;
  Registro_Sanitario: string;
  MARCA: string;
  TIPO_MONEDA: string;
  CLAVE: string;
  Fecha_Caducidad: string;
  Lote: string;
  Cantidad: string;
  Fecha_Fabircada: string;
  Presion_sin_iva: string;
  IVA: string;
  ID_x002d_Remision: string;
}

export interface ITableState {
  columns: IColumn[];
  DatosAI: IDetailsTableItem[];
  Remisiones: IDetailsTableItem[];
  DefTable: IColumn[];
}

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  ITableState
> {
  constructor(props: IHelloWorldProps) {
    super(props);

    const columnas: IColumn[] = [
      {
        key: "column1",
        name: "No Orden",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "NO_ORDEN_REPOSICION_UNOPS",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.NO_ORDEN_REPOSICION_UNOPS}</span>;
        },
      },
      {
        key: "column2",
        name: "No Remision",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "NO_REMISION",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.NO_REMISION}</span>;
        },
      },
      {
        key: "column3",
        name: "No Licitación",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "NO_LICITACION",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.NO_LICITACION}</span>;
        },
      },
      {
        key: "column4",
        name: "No Contrato",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "NO_CONTRATO",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.NO_CONTRATO}</span>;
        },
      },
      {
        key: "column5",
        name: "Procedencia",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "PROCEDENCIA",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.PROCEDENCIA}</span>;
        },
      },
      {
        key: "column6",
        name: "Registro Sanitario",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Registro_Sanitario",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Registro_Sanitario}</span>;
        },
      },
      {
        key: "column7",
        name: "Marca",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "MARCA",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.MARCA}</span>;
        },
      },
      {
        key: "column8",
        name: "Tipo Moneda",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "TIPO_MONEDA",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.TIPO_MONEDA}</span>;
        },
      },
      {
        key: "column9",
        name: "Clave",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "CLAVE",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.CLAVE}</span>;
        },
      },
      {
        key: "column10",
        name: "Fecha Caducidad",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Fecha_Caducidad",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Fecha_Caducidad}</span>;
        },
      },
      {
        key: "column11",
        name: "Lote",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Lote",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Lote}</span>;
        },
      },
      {
        key: "column12",
        name: "Cantidad",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Cantidad",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Cantidad}</span>;
        },
      },
      {
        key: "column13",
        name: "Fecha Fabricación",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Fecha_Fabircada",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Fecha_Fabircada}</span>;
        },
      },
      {
        key: "column14",
        name: "Precio",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "Presion_sin_iva",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.Presion_sin_iva}</span>;
        },
      },
      {
        key: "column15",
        name: "IVA",
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        fieldName: "IVA",
        isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {
          return <span>{global.IVA}</span>;
        },
      },
    ];

    this.state = {
      columns: columnas,
      DatosAI: [],
      Remisiones: [],
      DefTable: [],
    };
  }

  private async getAIDataTable(): Promise<void> {
    let items: [];
    if (this.props.DatosAI) {
      try {
        items = await getSP(this.props.context)
          .web.lists.getById(this.props.DatosAI.id)
          .items.select(
            "NO_ORDEN_REPOSICION_UNOPS",
            "NO_REMISION",
            "NO_LICITACION",
            "NO_CONTRATO",
            "PROCEDENCIA",
            "MARCA",
            "TIPO_MONEDA",
            "CLAVE",
            "IVA"
          )();

        this.setState({
          DatosAI: items,
        });
      } catch (err) {
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  private async getRemisionDataTable(): Promise<void> {
    let items: [];
    if (this.props.Remisiones) {
      try {
        items = await getSP(this.props.context)
          .web.lists.getById(this.props.Remisiones.id)
          .items.select(
            "Lote",
            "Registro_Sanitario",
            "Presion_sin_iva",
            "Cantidad",
            "Fecha_Fabircada",
            "Fecha_Caducidad",
            "ID_x002d_Remision"
          )();

        this.setState({
          Remisiones: items,
        });
      } catch (err) {
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  finalDataTable = async (): Promise<void> => {
    const result: any = [];
    if (this.state.DatosAI && this.state.Remisiones) {
      this.state.DatosAI.forEach((datoAI) => {
        const remFilter = this.state.Remisiones.filter((remision) => {
          return (
            datoAI.NO_ORDEN_REPOSICION_UNOPS === remision.ID_x002d_Remision
          );
        });

        const results = remFilter.reduce((x: any, y: any) => {
          (x[y.Lote] = x[y.Lote] || []).push(y);
          return x;
        }, {});

        const datos = Object.keys(results);
        const dato: any = [];
        datos.forEach((ele) => {
          dato.push(results[ele]);
        });

        const joinObject = (dataJson: any) => {
          let resultObj = {};
          const resultArray = [];

          const finalObj = (currentObj: any = {}, nextObj: any = {}) => {
            let resObj = { ...currentObj };
            for (const k in nextObj) {
              if (nextObj[k] === null) {
                resObj = { ...resObj };
              } else {
                resObj = { ...resObj, [k]: nextObj[k] };
              }
            }
            return resObj;
          };

          for (let i = 0; i < dataJson.length; i++) {
            for (let j = 0; j < dataJson[i].length; j++) {
              resultObj = finalObj(resultObj, dataJson[i][j]);
            }
            resultArray.push(resultObj);
            resultObj = {};
          }

          return resultArray;
        };

        result.push(
          joinObject(dato).map((item) => {
            return {
              ...datoAI,
              ...item,
            };
          })
        );
        // console.log("resukt", result);
      });
      this.setState({
        DefTable: result.flat(),
      });
    }
  };

  async componentDidMount(): Promise<void> {
    await this.getAIDataTable();
    await this.getRemisionDataTable();
    await this.finalDataTable();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <section>
        <DetailsList
          layoutMode={DetailsListLayoutMode.justified}
          items={this.state.DefTable}
          columns={this.state.columns}
          compact={true}
          setKey="set"
        />
      </section>
    );
  }
}
