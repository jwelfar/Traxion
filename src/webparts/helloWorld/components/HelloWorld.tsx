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
  DefaultButton,
  TextField,
  ITextFieldStyles,
} from "office-ui-fabric-react";
import * as XLSX from "xlsx";
import DataTable from "react-data-table-component";

let _sp: SPFI = null;

export interface IDetailsTableItem {
  Title: string;
  NO_ORDEN_REPOSICION_UNOPS: string;
  ID_x002d_Remision: string;
  NO_REMISION: string;
  NO_LICITACION: string;
  NO_CONTRATO: string;
  FECHA_SELLO_RECEPCION: string;
  PROCEDENCIA: string;
  RFC_LABORATORIO: string;
  Registro_Sanitario: string;
  REGISTRO_SANITARIO: string;
  MARCA: string;
  TIPO_MONEDA: string;
  CLAVE: string;
  Fecha_Caducidad: string;
  Lote: string;
  Cantidad: string;
  CANTIDAD_RECIBIDA: string;
  Fecha_Fabircada: string;
  Presion_sin_iva: string;
  PRECIO_SIN_IVA: string;
  IVA: string;
  ENTIDAD_FEDERATIVA: string;
}

export interface ITableState {
  columns: any[];
  DatosAI: IDetailsTableItem[];
  Remisiones: IDetailsTableItem[];
  DefTable: any[];
  filterText: string;
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

    const columnas = [
      {
        id: "column1",
        center: true,
        name: "No Orden",
        minWidth: "250px",
        maxWidth: "350px",
        selector: (row: any) => {
          return <span>{row.RAZON_SOCIAL}</span>;
        },
      },
      {
        id: "column2",
        grow: 2,
        center: true,
        name: "No Orden",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.NO_ORDEN_REPOSICION_UNOPS}</span>;
        },
      },
      {
        id: "column3",
        center: true,
        name: "OR",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          const numOrden =
            row.NO_ORDEN_REPOSICION_UNOPS || row.ID_x002d_Remision;
          const or = numOrden.substring(numOrden.lastIndexOf("/") + 1);
          return <span>{or}</span>;
        },
      },
      {
        id: "column4",
        center: true,
        name: "No Remisión",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.NO_REMISION}</span>;
        },
      },
      {
        id: "column5",
        center: true,
        name: "No Licitación",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.NO_LICITACION}</span>;
        },
      },
      {
        id: "column6",
        center: true,
        name: "No Contrato",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.NO_CONTRATO}</span>;
        },
      },
      {
        id: "column7",
        center: true,
        name: "Procedencia",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.PROCEDENCIA}</span>;
        },
      },
      {
        id: "column8",
        center: true,
        name: "Registro Sanitario",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return (
            <span>{row.Registro_Sanitario || row.REGISTRO_SANITARIO}</span>
          );
        },
      },
      {
        id: "column9",
        center: true,
        name: "Marca",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.MARCA}</span>;
        },
      },
      {
        id: "column10",
        center: true,
        name: "Tipo Moneda",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.TIPO_MONEDA}</span>;
        },
      },
      {
        id: "column11",
        center: true,
        name: "Clave",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.CLAVE}</span>;
        },
      },
      {
        id: "column12",
        center: true,
        name: "Fecha Caducidad",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.Fecha_Caducidad}</span>;
        },
      },
      {
        id: "column13",
        center: true,
        name: "Lote",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.Lote}</span>;
        },
      },
      {
        id: "column14",
        center: true,
        name: "Cantidad",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.Cantidad || row.CANTIDAD_RECIBIDA}</span>;
        },
      },
      {
        id: "column15",
        center: true,
        name: "Fecha Fabricación",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.Fecha_Fabircada}</span>;
        },
      },
      {
        id: "column16",
        center: true,
        name: "Precio",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.Presion_sin_iva || row.PRECIO_SIN_IVA}</span>;
        },
      },
      {
        id: "column17",
        center: true,
        name: "IVA",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.IVA}</span>;
        },
      },
    ];

    this.state = {
      columns: columnas,
      DatosAI: [],
      Remisiones: [],
      DefTable: [],
      filterText: "",
    };
  }

  private async getAIDataTable(): Promise<void> {
    let items: any = [];
    let response: any = [];
    if (this.props.DatosAI) {
      try {
        let next = true;
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
            "IVA",
            "REGISTRO_SANITARIO",
            "CANTIDAD_RECIBIDA",
            "PRECIO_SIN_IVA",
            "Title",
            "ENTIDAD_FEDERATIVA",
            "RFC_LABORATORIO",
            "FECHA_SELLO_RECEPCION",
            "RAZON_SOCIAL"
          )
          .top(50)
          .getPaged();

        const data = items.results;
        response = response.concat(data);

        while (next) {
          if (items.hasNext) {
            items = await items.getNext();
            response = response.concat(items.results);
          } else {
            next = false;
          }
        }

        this.setState({
          DatosAI: response,
        });

        return response;
      } catch (err) {
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  private async getRemisionDataTable(): Promise<void> {
    let items: any = [];
    let response: any = [];
    if (this.props.Remisiones) {
      try {
        let next = true;
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
          )
          .top(50)
          .getPaged();

        const data = items.results;
        response = response.concat(data);

        while (next) {
          if (items.hasNext) {
            items = await items.getNext();
            response = response.concat(items.results);
          } else {
            next = false;
          }
        }
        this.setState({
          Remisiones: response,
        });

        return response;
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
          return datoAI.Title === remision.ID_x002d_Remision;
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
      });
      this.setState({
        DefTable: result.flat(),
      });
      // console.log(this.state.DefTable);
    }
  };

  async componentDidMount(): Promise<void> {
    await this.getAIDataTable();
    await this.getRemisionDataTable();
    await this.finalDataTable();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const handleOnExport = () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(
        this.state.DefTable.map((item) => {
          const OR = item.NO_ORDEN_REPOSICION_UNOPS;
          const or = OR.substring(OR.lastIndexOf("/") + 1);
          return {
            NO_ORDEN_REPOSICION_UNOPS: item.NO_ORDEN_REPOSICION_UNOPS,
            OR: or,
            NO_REMISION: item.NO_REMISION || item.ID_x002d_Remisio,
            NO_LICITACION: item.NO_LICITACION,
            NO_CONTRATO: item.NO_CONTRATO,
            Procedencia: item.PROCEDENCIA,
            REGISTRO_SANITARIO:
              item.REGISTRO_SANITARIO || item.Registro_Sanitario,
            MARCA: item.MARCA,
            TIPO_MONEDA: item.TIPO_MONEDA,
            CLAVE: item.CLAVE,
            FECHA_CADUCIDAD: item.Fecha_Caducidad,
            LOTE: item.Lote,
            CANTIDAD_RECIBIDA: item.CANTIDAD_RECIBIDA || item.Cantidad,
            "Fecha Fabricación": item.Fecha_Fabircada,
            PRECIO_SIN_IVA: item.PRECIO_SIN_IVA || item.Presion_sin_iva,
            IVA: item.IVA,
          };
        })
      );
      const ws2 = XLSX.utils.json_to_sheet(
        this.state.DefTable.map((item) => {
          return {
            CLAS_PTAL_OL: "098316150905",
            NO_LICITACION: item.NO_LICITACION,
            NO_CONTRATO: item.NO_CONTRATO,
            RFC_LABORATORIO: item.RFC_LABORATORIO,
            NO_ORDEN_REPOSICION_UNOPS: item.NO_ORDEN_REPOSICION_UNOPS,
            FECHA_SELLO_RECEPCION: item.FECHA_SELLO_RECEPCION,
            CLAVE: item.CLAVE,
            PROCEDENCIA: item.PROCEDENCIA,
            REGISTRO_SANITARIO:
              item.REGISTRO_SANITARIO || item.Registro_Sanitario,
            MARCA: item.MARCA,
            FECHA_FABRICACION: item.Fecha_Fabircada,
            FECHA_CADUCIDAD: item.Fecha_Caducidad,
            LOTE: item.Lote,
            CANTIDAD_RECIBIDA: item.CANTIDAD_RECIBIDA || item.Cantidad,
            PRECIO_SIN_IVA: item.PRECIO_SIN_IVA || item.Presion_sin_iva,
            TIPO_MONEDA: item.TIPO_MONEDA,
          };
        })
      );
      const ws3 = XLSX.utils.json_to_sheet(
        this.state.DefTable.map((item) => {
          return {
            CLAS_PTAL_OL: "098316150905",
            NO_ORDEN_REPOSICION_UNOPS: item.NO_ORDEN_REPOSICION_UNOPS,
            ENTIDAD_FEDERATIVA: item.ENTIDAD_FEDERATIVA,
            CLAVE: item.CLAVE,
            CANTIDAD_RECIBIDA: item.CANTIDAD_RECIBIDA || item.Cantidad,
            NO_REMISION: item.NO_REMISION,
          };
        })
      );
      XLSX.utils.book_append_sheet(wb, ws, "WMS IZEL");
      XLSX.utils.book_append_sheet(wb, ws2, "PCCA");
      XLSX.utils.book_append_sheet(wb, ws3, "PCC2");

      XLSX.writeFile(wb, "Factura.xlsx");
    };

    const handleChange = (e: any) => {
      this.setState({
        filterText: e.target.value,
      });
    };

    const filteredList = this.state.DefTable.filter(
      (item) =>
        item.RAZON_SOCIAL &&
        item.RAZON_SOCIAL.toLowerCase().includes(
          this.state.filterText.toLowerCase()
        )
    );
    // console.log("filtro", filteredList);

    const textFieldStyles: Partial<ITextFieldStyles> = {
      fieldGroup: { width: 300 },
    };

    return (
      <section>
        <div
          style={{
            padding: "8px",
            display: "flex",
            alignItems: "end",
            justifyContent: "space-between",
          }}
        >
          <TextField
            label="Buscar por Razón Social"
            type="search"
            value={this.state.filterText}
            onChange={(e) => handleChange(e)}
            styles={textFieldStyles}
          />

          <DefaultButton
            text="Exportar"
            allowDisabledFocus
            onClick={() => handleOnExport()}
          />
        </div>

        <br />
        <DataTable
          columns={this.state.columns}
          data={filteredList}
          pagination
        />
      </section>
    );
  }
}
