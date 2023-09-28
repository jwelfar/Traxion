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
import "@pnp/sp/batching";
import "@pnp/sp/items/get-all";
import styless from "./HelloWorld.module.scss";
import Sendata from "./Sendata/Sendata"; // Import the Sendata component

//import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  DefaultButton,
  TextField,
  ITextFieldStyles,
  Link,
  Spinner,
  SpinnerSize,
  IconButton,
  Modal,
} from "office-ui-fabric-react";
import * as XLSX from "xlsx";
import DataTable from "react-data-table-component";
import * as moment from "moment";
import Swal from "sweetalert2";

let _sp: SPFI = null;
export interface IFederatibeitem {
  RFC: string;
  Title: string;
}
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
  Created: any;
  UrlArchivo: any;
  stylored: any;
  FechaRegistroSanitario: any;
  Tablacanje: any;
}

export interface ITableState {
  titleId: any;
  isModalOpen: boolean;
  showModal: boolean;
  hideModal: boolean;
  cuerpo: any;
  Entidadfederativatabla: any;
  Archivoeleminar: any;
  Cartaviciostabla: any;
  Cartagarantiatabla: any;
  cartacanjetabla: any;
  cartacanjeclave: any;
  cartacanjefecha: any;
  Cartacertificado: any;
  FechaRegistro: any;
  ordenreposicionsuma: any;
  errornoai: any;
  columns: any[];
  columns2: any[];
  ListCheck: any[];
  Tablereglas: any[];
  listafederal: IFederatibeitem[];
  Listacheckd: IDetailsTableItem[];
  DatosAI: IDetailsTableItem[];
  Remisiones: IDetailsTableItem[];
  DefTable: any[];
  NumOrderSearch: string;
  vinculo: string;
  ClaveSearch: string;
  LoteSearch: string;
  datefrom: string;
  dateto: string;
  pending: boolean;
  filteredData: any[];
  filteredDataf: any[];
  filteredDatalistache: any[];
  loading: any;
  cfnFilter: any;
  selectedRows: any[];
  selectedRowsData: any[];
}

interface MyItem {
  Errordocumento: string;
  Fechaerror: string;
  UrlArchivo: string;
  // add more properties as needed
}
let rowdataselet:any=[];
const conditionalRowStyles = [
  {
    when: (row: any) => row.selected, // Apply style to selected rows
    style: {
      backgroundColor: "lightgreen", // Set the background color for selected rows
    },
  },
];
export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
const modalStyles = {
  main: {
    padding: "2em",
    maxWidth: "100%",
    maxHeight: "100%",
    height: 400,
    width: 400,
    overflowY: "auto",
  },
};

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  ITableState
> {
  constructor(props: IHelloWorldProps) {
    super(props);

    const columnas2 = [
      {
        id: "column1",
        center: true,
        name: "Documento Faltante",
        minWidth: "250px",
        maxWidth: "350px",
        selector: (row: any) => {
          return <span>{row.Errordocumento}</span>;
        },
      },
      {
        id: "column2",
        grow: 2,
        center: true,
        name: "Fechas erroneas",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.FechasError}</span>;
        },
      },

      {
        id: "column3",
        center: true,
        name: "Archivo",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          const file = row.UrlArchivo.substring(
            row.UrlArchivo.lastIndexOf("/") + 1
          );

          return (
            <Link href={row.UrlArchivo} target="_blank">
              {file}
            </Link>
          );
        },
      },
    ];

    const columnas = [
      {
        id: "column1",
        center: true,
        name: "Razón social",
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
          if (row.stylored === "red") {
            return (
              <>
                <IconButton
                  onClick={() => {
                    this._showModal(row);
                  }}
                  iconProps={{ iconName: "AlertSolid" }}
                  title="AlertSolid"
                  ariaLabel="AlertSolid"
                  color="red"
                />
                <span className={styless.redalert}>
                  {row.NO_ORDEN_REPOSICION_UNOPS}
                </span>
              </>
            );
          } else {
            return <span>{row.NO_ORDEN_REPOSICION_UNOPS}</span>;
          }
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
          const or = numOrden?.substring(numOrden?.lastIndexOf("/") + 1);
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
          return <span>{row.IVA === null ? "0" : row.IVA}</span>;
        },
      },
      {
        id: "column18",
        center: true,
        name: "Fecha digitalización",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{moment(row.Created).format("DD-MM-YYYY")}</span>;
        },
      },
      {
        id: "column19",
        center: true,
        name: "Archivo",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          const file = row.LinkTitle.substring(
            row.LinkTitle.lastIndexOf("/") + 1
          );

          return (
            <Link
              data-interception="off"
              rel="noopener noreferrer"
              href={row.LinkTitle}
              target="_blank"
            >
              {file}
            </Link>
          );
        },
      },
    ];

    this.state = {
      Tablereglas: [],
      titleId: "",
      isModalOpen: false,
      showModal: false,
      hideModal: false,
      cuerpo: "",
      Entidadfederativatabla: "",
      Archivoeleminar: "",
      Cartaviciostabla: "",
      Cartagarantiatabla: "",
      cartacanjetabla: "",
      Cartacertificado: "",
      cartacanjefecha: "",
      cartacanjeclave: "",
      FechaRegistro: "",
      ordenreposicionsuma: "",
      errornoai: "",
      listafederal: [],
      columns2: columnas2,
      columns: columnas,
      Listacheckd: [],
      DatosAI: [],
      Remisiones: [],
      DefTable: [],
      NumOrderSearch: "",
      vinculo: "",
      ClaveSearch: "",
      LoteSearch: "",
      datefrom: "",
      dateto: "",
      pending: true,
      filteredData: [],
      filteredDataf: [],
      loading: false,
      ListCheck: [],
      filteredDatalistache: [],
      cfnFilter: [],
      selectedRows: [],
      selectedRowsData: [],
    };
  }

  private _showModal = (row: any): void => {
    this.setState({
      titleId: row.LinkTitle,
      cuerpo: row.texto,
      Entidadfederativatabla:
        row?.ErrorTabla === null || row?.ErrorTabla === undefined
          ? ""
          : row?.ErrorTabla,
      FechaRegistro:
        row?.ErrorfechaRegistro === null ||
        row?.ErrorfechaRegistro === undefined
          ? ""
          : row?.ErrorfechaRegistro,
      cartacanjetabla:
        row?.Tablacanje === null || row?.Tablacanje === undefined
          ? ""
          : row?.Tablacanje,
      cartacanjeclave:
        row?.ErrorTableClave === null || row?.ErrorTableClave === undefined
          ? ""
          : row?.ErrorTableClave,
      cartacanjefecha:
        row?.ErrorTableClaveFecha === null ||
        row?.ErrorTableClaveFecha === undefined
          ? ""
          : row?.ErrorTableClaveFecha,
      ordenreposicionsuma:
        (row?.ErrorSUma === null || row?.ErrorSUma === undefined)||  row.LinkTitle.indexOf('CROSS') === -1
          ? ""
          : row?.ErrorSUma,
      errornoai:
        row?.nolectura === null || row?.nolectura === undefined
          ? ""
          : row?.nolectura,
      /* Cartaviciostabla:any;
       Cartagarantiatabla:any;
       cartacanjetabla:any;
       Cartacertificado:any;*/
    });
    this.setState({ showModal: true });
  };

  private _hideModal = (): void => {
    this.setState({
      titleId: "",
      cuerpo: "",
    });
    this.setState({ showModal: false });
  };

  private async getAIDataTable(): Promise<void> {
    this.setState({
      loading: true,
    });
    let query = "";
    if (this.state.ClaveSearch.length >= 5) {
      if (query.length === 0) {
        query = "substringof('" + this.state.ClaveSearch + "', CLAVE)";
      }
    }
    if (this.state.NumOrderSearch.length >= 5) {
      if (query.length === 0) {
        query =
          "substringof('" +
          this.state.NumOrderSearch +
          "', NO_ORDEN_REPOSICION_UNOPS)";
      } else {
        query +=
          " and substringof('" +
          this.state.NumOrderSearch +
          "', NO_ORDEN_REPOSICION_UNOPS)";
      }
    }
    let items: any = [];
    let response: any = [];

    if (this.state.datefrom) {
      if (query.length === 0) {
        query = "Created ge datetime'" + this.state.datefrom + "T00:00:00'";
        if (this.state.dateto) {
          query +=
            " and Created le datetime'" + this.state.dateto + "T23:59:59'";
        }
      } else {
        query +=
          " and Created ge datetime'" + this.state.datefrom + "T00:00:00'";
        if (this.state.dateto) {
          query +=
            " and Created le datetime'" + this.state.dateto + "T23:59:59'";
        }
      }
    }

    if (this.props.DatosAI) {
      try {
        let next = true;
        if (query.length > 0) {
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
              "RAZON_SOCIAL",
              "Created",
              "LinkTitle",
              "FechaRegistroSanitario",
              "CLUES_x002f_CPTALDESTINO",
              "Id"
            )
            .top(2000)
            .filter(query)
            .getPaged();
        } else {
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
              "RAZON_SOCIAL",
              "Created",
              "LinkTitle",
              "FechaRegistroSanitario",
              "CLUES_x002f_CPTALDESTINO",
              "Id"
            )
            .top(2000)
            .getPaged();
        }
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
        this.setState({
          loading: false,
        });
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  private async getAIDataTablefiltedate(): Promise<void> {
    if (this.state.Remisiones.length > 0) {
      const maxDate = new Date(
        Math.max(
          ...this.state.Remisiones.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      // ✅ Get Min date
      const minDate = new Date(
        Math.min(
          ...this.state.Remisiones.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      this.setState({
        loading: true,
      });
      let query = "";
      if (this.state.ClaveSearch.length >= 5) {
        if (query.length === 0) {
          query = "substringof('" + this.state.ClaveSearch + "', CLAVE)";
        }
      }
      if (this.state.NumOrderSearch.length >= 5) {
        if (query.length === 0) {
          query =
            "substringof('" +
            this.state.NumOrderSearch +
            "', NO_ORDEN_REPOSICION_UNOPS)";
        } else {
          query +=
            " and substringof('" +
            this.state.NumOrderSearch +
            "', NO_ORDEN_REPOSICION_UNOPS)";
        }
      }
      let items: any = [];
      let response: any = [];

      if (minDate) {
        if (query.length === 0) {
          query =
            "Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:59'";
          }
        } else {
          query +=
            " and Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:59'";
          }
        }
      }

      if (this.props.DatosAI) {
        try {
          let next = true;
          if (query.length > 0) {
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
                "RAZON_SOCIAL",
                "Created",
                "LinkTitle",
                "FechaRegistroSanitario",
                "Id"
              )
              .top(2000)
              .filter(query)
              .getPaged();
          } else {
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
                "RAZON_SOCIAL",
                "Created",
                "LinkTitle",
                "FechaRegistroSanitario",
                "Id"
              )
              .top(2000)
              .getPaged();
          }
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
          this.setState({
            loading: false,
          });
          console.log("Error", err);
          err.res.json().then(() => {
            console.log("Failed to get list items!", err);
          });
        }
      }
    }
  }

  private formatDate(date: string): string {
    const datef = new Date(date);
    const formattedDate = moment(datef).format("YYYY-MM-DD");
    return formattedDate;
  }
  private async getListacheckDataTable(): Promise<void> {
    if (this.state.DatosAI.length > 0) {
      const maxDate = new Date(
        Math.max(
          ...this.state.DatosAI.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      // ✅ Get Min date
      const minDate = new Date(
        Math.min(
          ...this.state.DatosAI.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      this.setState({
        loading: true,
      });
      let query = "";

      let items: any = [];
      let response: any = [];

      if (minDate) {
        if (query.length === 0) {
          query =
            "Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:00'";
          }
        } else {
          query =
            "  Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:00'";
          }
        }
      }

      if (this.props.ListCheck) {
        try {
          let next = true;
          if (query.length > 0) {
            items = await getSP(this.props.context)
              .web.lists.getById(this.props.ListCheck.id)
              .items.select(
                "Title",
                "Certificado_Calidad",
                "Prorroga_Sanitario",
                "Registro_Sanitario",
                "Carta_Vicios",
                "Carta_Garantia",
                "Manifiesto",
                "OrdenReposicion",
                "UrlArchivo",
                "Id"
              )
              .top(2000)
              .filter(query)
              .getPaged();
          } else {
            items = await getSP(this.props.context)
              .web.lists.getById(this.props.ListCheck.id)
              .items.select(
                "Title",
                "Certificado_Calidad",
                "Prorroga_Sanitario",
                "Registro_Sanitario",
                "Carta_Vicios",
                "Carta_Garantia",
                "Manifiesto",
                "OrdenReposicion",
                "UrlArchivo",
                "Id"
              )
              .top(2000)
              .getPaged();
          }
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
            Listacheckd: response,
          });

          return response;
        } catch (err) {
          this.setState({
            loading: false,
          });
          console.log("Error", err);
          err.res.json().then(() => {
            console.log("Failed to get list items!", err);
          });
        }
      }
    }
  }

  private async geTablaReglas(): Promise<void> {
    if (this.state.DatosAI.length > 0) {
      const maxDate = new Date(
        Math.max(
          ...this.state.DatosAI.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      // ✅ Get Min date
      const minDate = new Date(
        Math.min(
          ...this.state.DatosAI.map((remision) => {
            return new Date(remision.Created).getTime();
          })
        )
      );
      this.setState({
        loading: true,
      });
      let query = "";

      // let items: any = [];
      let response: any = [];

      if (minDate) {
        if (query.length === 0) {
          query =
            "Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:00'";
          }
        } else {
          query =
            "Created ge datetime'" +
            this.formatDate(minDate.toString()) +
            "T00:00:00'";
          if (maxDate) {
            query +=
              " and Created le datetime'" +
              this.formatDate(maxDate.toString()) +
              "T23:59:00'";
          }
        }
      }

      if (this.props.Tablareglas) {
        try {
          if (query.length > 0) {
            const allItems: any[] = await _sp.web.lists
              .getById(this.props.Tablareglas.id)
              .items.select(
                "Title",
                "ENTIDAD_FEDERATIVA",
                "CANTIDAD_RECIBIDA",
                "LOTE",
                "FECHA_FABRICACION",
                "FECHA_CADUCIDAD",
                "UrlArchivo",
                "TipoTabla",
                "Id"
              )
              .getAll(4000);
            //  const result: any = [];
            const datosai = this.state.DatosAI;
            datosai.forEach(async (element) => {
              const remFilter = allItems.filter((remision) => {
                return remision.UrlArchivo === element.Title;
              });

              if (remFilter?.length > 0) {
                const data = remFilter;
                response = response.concat(data);
              }
            });
            this.setState({
              Tablereglas: response,
            });
          }

          return response;
        } catch (err) {
          this.setState({
            loading: false,
          });
          console.log("Error", err);
          err.res.json().then(() => {
            console.log("Failed to get list items!", err);
          });
        }
      }
    }
  }

  private async getRemisionDataTable(): Promise<void> {
    let query = "";
    if (this.state.LoteSearch.length >= 3) {
      query = "substringof('" + this.state.LoteSearch + "',Lote)";
    }
    if (this.state.datefrom) {
      if (query.length === 0) {
        query = "Created ge datetime'" + this.state.datefrom + "T00:00:00'";
        if (this.state.dateto) {
          query +=
            " and Created le datetime'" + this.state.dateto + "T23:59:59'";
        }
      } else {
        query +=
          " and Created ge datetime'" + this.state.datefrom + "T00:00:00'";
        if (this.state.dateto) {
          query +=
            " and Created le datetime'" + this.state.dateto + "T23:59:59'";
        }
      }
    }

    let items: any = [];
    let response: any = [];
    if (this.props.Remisiones) {
      try {
        let next = true;
        if (this.state.LoteSearch) {
          items = await getSP(this.props.context)
            .web.lists.getById(this.props.Remisiones.id)
            .items.select(
              "Lote",
              "Registro_Sanitario",
              "Presion_sin_iva",
              "Cantidad",
              "CFN",
              "Fecha_Fabircada",
              "Fecha_Caducidad",
              "ID_x002d_Remision",
              "Created",
              "Id"
            )
            .top(2000)
            .filter(query)
            .getPaged();
        } else {
          items = await getSP(this.props.context)
            .web.lists.getById(this.props.Remisiones.id)
            .items.select(
              "Lote",
              "Registro_Sanitario",
              "Presion_sin_iva",
              "Cantidad",
              "CFN",
              "Fecha_Fabircada",
              "Fecha_Caducidad",
              "ID_x002d_Remision",
              "Created",
              "Id"
            )
            .top(2000)
            .getPaged();
        }
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
        this.setState({
          loading: false,
        });
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  private async getreglafederal(): Promise<void> {
    let items: any = [];
    let response: any = [];
    try {
      let next = true;
      items = await getSP(this.props.context)
        .web.lists.getById(this.props.ListaValidaciom.id)
        .items.select("Title", "RFC", "Id")
        .top(2000)
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
        listafederal: response,
      });

      return response;
    } catch (err) {
      this.setState({
        loading: false,
      });
      console.log("Error", err);
      err.res.json().then(() => {
        console.log("Failed to get list items!", err);
      });
    }
  }

  finalistcheckDataTable = async (): Promise<void> => {
    const result: any = [];
    if (this.state.DatosAI.length > 0 && this.state.Listacheckd.length > 0) {
      this.state.DatosAI.forEach((datoAI) => {
        const remFilter = this.state.Listacheckd.filter((remision) => {
          return remision.UrlArchivo === datoAI.Title;
        });
        result.push(remFilter?.[0]);
      });
      const dataultimo: MyItem[] = [];
      let textvalidation: any = "";
      try {
        if (result !== undefined || result.length > 0) {
          result.forEach(
            (element: {
              Title: string;
              Certificado_Calidad: string;
              Prorroga_Sanitario: string;
              Registro_Sanitario: string;
              Manifiesto: string;
              Carta_Vicios: string;
              Carta_Garantia: string;
              Carta_Canje: string;
              OrdenReposicion: string;
              UrlArchivo: string;
            }) => {
              textvalidation = "";
              if (
                element?.Title === "" ||
                element?.Title === undefined ||
                element?.Title === null
              ) {
                textvalidation += "Remision, ";
              }
              if (
                element?.Certificado_Calidad === "" ||
                element?.Certificado_Calidad === undefined ||
                element?.Certificado_Calidad === null
              ) {
                textvalidation += "Certificado Calidad, ";
              }
              if (
                element?.Prorroga_Sanitario === "" ||
                element?.Prorroga_Sanitario === undefined ||
                element?.Prorroga_Sanitario === null
              ) {
                if (
                  element?.Registro_Sanitario === "" ||
                  element?.Registro_Sanitario === undefined ||
                  element?.Registro_Sanitario === null
                ) {
                  textvalidation += "Registro Sanitario, ";
                }
              }

              if (
                element?.Manifiesto === "" ||
                element?.Manifiesto === undefined ||
                element?.Manifiesto === null
              ) {
                textvalidation += "Manifiesto, ";
              }
              if (
                element?.Carta_Vicios === "" ||
                element?.Carta_Vicios === undefined ||
                element?.Carta_Vicios === null
              ) {
                textvalidation += "Carta Vicios, ";
              }
              if (
                element?.Carta_Garantia === "" ||
                element?.Carta_Garantia === undefined ||
                element?.Carta_Garantia === null
              ) {
                textvalidation += "Carta Garantia, ";
              }
              if (
                element?.Carta_Canje === "" ||
                element?.Carta_Canje === undefined ||
                element?.Carta_Canje === null
              ) {
                const canj: any = this.state.filteredData.filter((canje) => {
                  return canje.LinkTitle === element.UrlArchivo;
                });

                if (moment().subtract(12, "months") > canj.Fecha_Caducidad) {
                  textvalidation += "Carta Canje, ";
                }
              }
              if (
               ( element.OrdenReposicion === "" ||
                element.OrdenReposicion === undefined ||
                element.OrdenReposicion === null) &&  element.UrlArchivo.indexOf('CROSS') !== -1
              ) {
                textvalidation += "Orden Reposicióm";
              }

              const daultimo: MyItem = {
                Errordocumento: textvalidation,
                Fechaerror: "",
                UrlArchivo: element?.UrlArchivo,
              };
              if (daultimo.Errordocumento !== "") {
                dataultimo.push(daultimo);
              }
            }
          );
        }
      } catch (error) {
        console.log(error);
      }
      this.setState({
        filteredDatalistache: dataultimo,
      });

      this.setState({
        loading: false,
      });
    }
  };

  finalDataTable = async (): Promise<void> => {
    const result: any = [];
    let resutllote: any = [];
    if (
      this.state.LoteSearch.length > 0 &&
      (this.state.ClaveSearch.length === 0 ||
        this.state.NumOrderSearch.length === 0)
    ) {
      await this.getAIDataTablefiltedate();
      this.state.Remisiones.forEach((remision: { ID_x002d_Remision: any }) => {
        this.state.DatosAI.filter((datoAI: { Title: any }) => {
          if (datoAI.Title === remision.ID_x002d_Remision) {
            resutllote.push(datoAI);
          }
        });
      });
    } else {
      resutllote = this.state.DatosAI;
    }
    if (this.state.DatosAI.length > 0) {
      if (this.state.Remisiones.length > 0) {
        resutllote.forEach(
          (datoAI: { NO_ORDEN_REPOSICION_UNOPS: any; Title: any }) => {
            const filename = datoAI?.Title.toString().substring(
              datoAI?.Title.toString().lastIndexOf("/") + 1
            );
            const dato =
              filename.indexOf("PO-") > -1 || filename.indexOf("PO/") > -1;
            if (dato === false) {
              const remFilter = this.state.Remisiones.filter(
                (remision: { ID_x002d_Remision: any }) => {
                  return datoAI.Title === remision.ID_x002d_Remision;
                }
              );
              if (remFilter.length === 0) {
                result.push(datoAI);
              } else {
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

                  const finalObj = (
                    currentObj: any = {},
                    nextObj: any = {}
                  ) => {
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
              }
            }
          }
        );

        this.setState({
          filteredDataf: result.flat(),
        });
        this.setState({
          cfnFilter: this.state.filteredDataf.filter((obj) => {
            if (obj?.RAZON_SOCIAL?.includes("MEDTRONIC") && obj?.CFN) {
              return obj.CFN.includes("SP");
            }
          }),
        });
        this.setState({
          loading: false,
        });
      } else {
        this.state.DatosAI.forEach(
          (datoAI: { NO_ORDEN_REPOSICION_UNOPS: any; Title: any }) => {
            const filename = datoAI?.Title.toString().substring(
              datoAI?.Title.toString().lastIndexOf("/") + 1
            );
            const datos =
              filename.indexOf("PO-") > -1 || filename.indexOf("PO/") > -1;
            if (datos === false) {
              const dato: any = [];

              dato.push(datoAI);
            }
          }
        );
        this.setState({
          filteredDataf: result.flat(),
        });
        this.setState({
          loading: false,
        });
      }
    }
  };

  finalDataTableval = async (): Promise<void> => {
    if (
      this.state.filteredDataf.length > 0 &&
      this.state.filteredDatalistache.length > 0
    ) {
      this.state.filteredDatalistache.forEach((datoAI): void => {
        this.state.filteredDataf.forEach((item) => {
          if (item.Title === datoAI.UrlArchivo) {
            item.stylored = "red";
            item.texto = datoAI.Errordocumento;
          }
        });
      });
    }
    this.setState({ filteredData: this.state.filteredDataf });
    this.setState({
      loading: false,
    });
  };

  CheckOrdenSalida = async (): Promise<void> => {
    if (this.state.filteredDataf.length > 0) {
      const cartacanjelotes: any = this.state.Tablereglas?.reduce(
        (acc, curr) => {
          if (!acc[curr.UrlArchivo]) {
            acc[curr.UrlArchivo] = [];
          }
          acc[curr.UrlArchivo].push(curr);
          return acc;
        },
        {}
      );
      let validadornumero = 0;
      this.state.filteredDataf.forEach((datoAI): void => {
        const ordenes = this.state.Tablereglas?.filter((item) => {
          return (
            item.TipoTabla === "Tabla-Ordenes" &&
            datoAI.Title === item.UrlArchivo
          );
        });
        if (ordenes.length > 0) {
          const ordenefina = this.findDuplicates(ordenes);
          if (ordenefina?.length > 0) {
            datoAI.stylored = "red";
            datoAI.ErrorTabla =
              "Entidad federativa: " +
              ordenefina[0].Title +
              " - " +
              ordenefina[0].ENTIDAD_FEDERATIVA;
          }
        }
        if (cartacanjelotes.length > 0) {
          const found = cartacanjelotes[datoAI.Title].filter(
            (a: any) =>
              datoAI?.Lote === a.LOTE &&
              datoAI.Title === a.UrlArchivo &&
              a.TipoTabla === "Tabla-CartaCanje"
          );
          if (found.length <= 0) {
            datoAI.stylored = "red";
            datoAI.Tablacanje = "Error en lote: " + datoAI.Lote;
            validadornumero = validadornumero + 1;
          }
        }
        const tablacanje = this.state.Tablereglas.filter((item) => {
          return (
            item.TipoTabla === "Tabla-CartaCanje" &&
            datoAI.Title === item.UrlArchivo
          );
        });

        if (tablacanje.length > 0) {
          tablacanje.forEach((tcanje): void => {
            if (tcanje.Title === datoAI.Clave) {
              datoAI.stylored = "red";
              datoAI.ErrorTableClave = "Error en Clave: " + datoAI.Clave;
              validadornumero = validadornumero + 1;
            }
            if (tcanje.Fecha_Caducidad) {
              const fechacad = tcanje.Fecha_Caducidad;
              if (fechacad.include("/")) {
                if (fechacad.length === 10) {
                  const date = moment(fechacad, "DD-MM-YYYY").format(
                    "DD-MM-YYYY"
                  );
                  if (date < moment().format("DD-MM-YYYY")) {
                    datoAI.stylored = "red";
                    datoAI.ErrorTableClaveFecha =
                      "Error en Fecha carta canje: " + fechacad;
                    validadornumero = validadornumero + 1;
                  }
                } else {
                  const dateString = fechacad;
                  const [month, year] = dateString.split("/");
                  const lastDayOfMonth = moment(`${year}-${month}`, "YYYY/MM")
                    .endOf("month")
                    .date();
                  const date = moment(
                    `${year}-${month}-${lastDayOfMonth}`,
                    "YYYY-MM-DD"
                  ).format("DD-MM-YYYY");
                  if (date < moment().format("DD-MM-YYYY")) {
                    datoAI.stylored = "red";
                    datoAI.ErrorTableClaveFecha =
                      "Error en Fecha carta canje: " + fechacad;
                    validadornumero = validadornumero + 1;
                  }
                }
              } else {
                if (fechacad.length === 10) {
                  const date = moment(fechacad, "DD-MM-YYYY").format(
                    "DD-MM-YYYY"
                  );
                  if (date < moment().format("DD-MM-YYYY")) {
                    datoAI.stylored = "red";
                    datoAI.ErrorTableClaveFecha =
                      "Error en Fecha carta canje: " + fechacad;
                    validadornumero = validadornumero + 1;
                  }
                } else {
                  const dateString = fechacad;
                  const [month, year] = dateString.split("-");
                  const lastDayOfMonth = moment(`${year}-${month}`, "YYYY-MM")
                    .endOf("month")
                    .date();
                  const date = moment(
                    `${year}-${month}-${lastDayOfMonth}`,
                    "YYYY-MM-DD"
                  ).format("DD-MM-YYYY");
                  if (date < moment().format("DD-MM-YYYY")) {
                    datoAI.stylored = "red";
                    datoAI.ErrorTableClaveFecha =
                      "Error en Fecha carta canje: " + fechacad;
                    validadornumero = validadornumero + 1;
                  }
                }
              }
            }
          });

          if (datoAI?.FechaRegistroSanitario) {
            if (
              moment(datoAI.FechaRegistroSanitario).format("DD-MM-YYYY") <=
              moment().subtract(150, "days").format("DD-MM-YYYY")
            ) {
              datoAI.stylored = "red";
              datoAI.ErrorfechaRegistro =
                "Error en fecha de registro sanitario: " +
                datoAI?.FechaRegistroSanitario;
              validadornumero = validadornumero + 1;
            }
          }
          const ordenesclave = this.findsumm(
            this.state.filteredDataf,
            datoAI?.Title
          );
          const findsum =
            datoAI.Cantidad === ordenesclave.toString() ? true : false;
          if (findsum === false) {
            const lotes = datoAI?.Lote === undefined ? "" : datoAI?.Lote;
            datoAI.stylored = "red";
            datoAI.ErrorSUma =
              "Error en lote: " + lotes + " cantidad:" + datoAI?.Cantidad;
            validadornumero = validadornumero + 1;
          }
        }

        const ordenesval = this.state.Tablereglas?.filter((item) => {
          return (
            item.TipoTabla === "Tabla-Ordenes" &&
            datoAI.Title === item.UrlArchivo
          );
        });
        const remFilter = this.state.Remisiones.filter(
          (remision: { ID_x002d_Remision: any }) => {
            return datoAI.Title === remision.ID_x002d_Remision;
          }
        );
        const categoriaFilter = this.state.Listacheckd.filter((remision) => {
          return remision.UrlArchivo === datoAI.Title;
        });
        let messageAi = "";
        if (ordenesval?.length === 0 || ordenesval === undefined) {
          messageAi = "Error no tiene lectura por la AI en reglas de negocio";
        }
        if (remFilter === undefined || remFilter.length === 0) {
          if (messageAi.length > 0)
            messageAi += ", lectura por categoría documentos";
          else {
            datoAI.stylored = "red";
            messageAi =
              "Error no tiene lectura por la AI en categoría documentos";
          }
        }
        if (categoriaFilter === undefined || categoriaFilter.length === 0) {
          if (messageAi.length > 0) messageAi += ", lectura por remisiones";
          else {
            datoAI.stylored = "red";
            messageAi =
              "Error no tiene lectura por la AI en lecturas por remisiones";
          }
        }
        if (messageAi) {
          datoAI.stylored = "red";
          datoAI.nolectura = messageAi;
        }
      });
    }
  };

  findDuplicates = (arr: any[]): any[] => {
    const duplicates: any = [];
    for (let i = 0; i < arr.length; i++) {
      for (let j = i + 1; j < arr.length; j++) {
        if (
          arr[i].ENTIDAD_FEDERATIVA === arr[j].ENTIDAD_FEDERATIVA &&
          arr[i].Title === arr[j].Title
        ) {
          if (!duplicates.includes(arr[i])) {
            duplicates.push(arr[i]);
          }
          if (!duplicates.includes(arr[j])) {
            duplicates.push(arr[j]);
          }
        }
      }
    }
    return duplicates;
  };

  findsumm = (arr: any[], Title: any): any => {
    let cantidad: any = 0;
    const filterlote = arr.filter((obj) => obj.Title === Title);
    filterlote.forEach((element) => {
      cantidad = Number(element.Cantidad) + Number(cantidad);
    });
    return cantidad;
  };

  findDuplicateslotes = (arr: any[]): any[] => {
    const duplicates: any = [];
    for (let i = 0; i < arr.length - 1; i++) {
      for (let j = i + 1; j < arr.length; j++) {
        if (arr[i].Lote === arr[j].Lote && arr[i].Title === arr[j].Title) {
          if (!duplicates.includes(arr[i])) {
            duplicates.push(arr[i]);
          }
          if (!duplicates.includes(arr[j])) {
            duplicates.push(arr[j]);
          }
        }
      }
    }
    return duplicates;
  };

  executeFunctionWithSelectedRows = () => {
    
    this.setState({ selectedRows: this.state.selectedRowsData });
    // Do something with selectedRows in DataProcess component
    console.log("datosselecionado", this.state.selectedRowsData);
  };

  handleRowSelection = (rowData: any) => {
    this.setState((prevState) => {
      const { selectedRowsData } = prevState;

      // Check if the row is already selected
      const isRowSelected = selectedRowsData.some(
        (row: any) => row.ID === rowData.ID
      );

      if (isRowSelected) {
        // Remove the row from selectedRowsData if it's already selected
        const updatedSelectedRowsData = selectedRowsData.filter(
          (row: any) => row.ID !== rowData.ID
        );
        rowdataselet=rowData;
        
        return {
          selectedRowsData: updatedSelectedRowsData,
        };
      } else {
        // Clone the rowData and toggle the selected property
        const newRowData = { ...rowData.selectedRows, selected: true };
        rowdataselet=rowData;
        return {
          selectedRowsData: [...selectedRowsData, newRowData],
        };
      }
    });
  };

  handleFilter = async (): Promise<void> => {
    this.setState({
      filteredDataf: [],
    });

    setTimeout(async () => {
      if (
        this.state.LoteSearch.length > 0 &&
        (this.state.ClaveSearch.length === 0 ||
          this.state.NumOrderSearch.length === 0)
      ) {
        await this.getRemisionDataTable();
        await this.geTablaReglas();
        await this.getreglafederal();
        await this.getListacheckDataTable();
        await this.finalDataTable();
        await this.CheckOrdenSalida();
        await this.finalistcheckDataTable();
        await this.finalDataTableval();
      } else {
        await this.getAIDataTable();
        await this.getRemisionDataTable();
        await this.geTablaReglas();
        await this.getreglafederal();
        await this.getListacheckDataTable();
        await this.finalDataTable();
        await this.CheckOrdenSalida();
        await this.finalistcheckDataTable();
        await this.finalDataTableval();
      }
    }, 3000);
  };

  handledelete = async (archivo: any): Promise<void> => {
    this._hideModal();
    Swal.fire({
      title: "¿Está seguro?",
      text: "el archivo se borrara!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Sí",
    }).then(async (result: any) => {
      if (result.isConfirmed) {
        this.setState({
          filteredDataf: [],
        });
        Swal.fire({
          title: "Eliminando",
          text: "Si los archivos son muy grandes, esta carga tomara unos minutos",
          allowOutsideClick: false,
          didOpen: () => {
            Swal.showLoading();
          },
        });
        //elimina de tabla reglas
        const columnValue = archivo;
        let myList = spfi()
          .using(SPFx(this.props.context))
          .web.lists.getById(this.props.Tablareglas.id);
        let itemsToDelete = await myList.items.filter(
          `UrlArchivo eq '${columnValue}'`
        )();

        // Delete each item
        for (const item of itemsToDelete) {
          await spfi()
            .using(SPFx(this.props.context))
            .web.lists.getById(this.props.Tablareglas.id)
            .items.getById(item.Id).delete;
        }

        //elimina de tabla lista checkeo
        myList = spfi()
          .using(SPFx(this.props.context))
          .web.lists.getById(this.props.ListCheck.id);
        itemsToDelete = await myList.items.filter(
          `UrlArchivo eq '${columnValue}'`
        )();

        // Delete each item
        for (const item of itemsToDelete) {
          await myList.items.getById(item.Id).delete();
        }
        //elimina de tabla lista remisiones
        myList = spfi()
          .using(SPFx(this.props.context))
          .web.lists.getById(this.props.Remisiones.id);
        itemsToDelete = await myList.items.filter(
          `ID_x002d_Remision eq '${columnValue}'`
        )();

        // Delete each item
        for (const item of itemsToDelete) {
          await myList.items.getById(item.Id).delete();
        }

        //elimina de tabla lista DatosAI
        myList = spfi()
          .using(SPFx(this.props.context))
          .web.lists.getById(this.props.DatosAI.id);
        itemsToDelete = await myList.items.filter(
          `Title eq '${columnValue}'`
        )();

        // Delete each item
        for (const item of itemsToDelete) {
          await myList.items.getById(item.Id).delete();
        }

        myList = spfi()
          .using(SPFx(this.props.context))
          .web.lists.getById(this.props.Documentos.id);
        itemsToDelete = await myList.items.filter(
          `Title eq '${columnValue}'`
        )();

        for (const item of itemsToDelete) {
          await myList.items.getById(item.Id).delete();
        }

        Swal.fire({
          icon: "success",
          title: "Documento removido",
          showConfirmButton: false,
          timer: 1500,
        });
        this.setState({
          filteredData: [],
          filteredDataf: [],
          filteredDatalistache: [],
        });
        setTimeout(async () => {
          if (
            this.state.LoteSearch.length > 0 &&
            (this.state.ClaveSearch.length === 0 ||
              this.state.NumOrderSearch.length === 0)
          ) {
            await this.getRemisionDataTable();
            await this.geTablaReglas();
            await this.getreglafederal();
            await this.getListacheckDataTable();
            await this.finalDataTable();
            await this.CheckOrdenSalida();
            await this.finalistcheckDataTable();
            await this.finalDataTableval();
          } else {
            await this.getAIDataTable();
            await this.getRemisionDataTable();
            await this.geTablaReglas();
            await this.getreglafederal();
            await this.getListacheckDataTable();
            await this.finalDataTable();
            await this.CheckOrdenSalida();
            await this.finalistcheckDataTable();
            await this.finalDataTableval();
          }
        }, 3000);
      }
    });
  };

  
  

  async componentDidMount(): Promise<void> {
    this.setState({
      pending: false,
    });
  }
  public render(): React.ReactElement<IHelloWorldProps> {
    const handleOnExport = (): void => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(
        this.state.filteredDataf.map(
          (item: {
            NO_ORDEN_REPOSICION_UNOPS: any;
            TIPO_MONEDA: string;
            NO_REMISION: any;
            ID_x002d_Remisio: any;
            NO_LICITACION: any;
            NO_CONTRATO: any;
            PROCEDENCIA: any;
            REGISTRO_SANITARIO: any;
            Registro_Sanitario: any;
            MARCA: any;
            CLAVE: any;
            Fecha_Caducidad: any;
            Lote: any;
            CANTIDAD_RECIBIDA: any;
            Cantidad: any;
            Fecha_Fabircada: any;
            PRECIO_SIN_IVA: string;
            Presion_sin_iva: string;
            IVA: any;
          }) => {
            const OR = item.NO_ORDEN_REPOSICION_UNOPS;
            const or = OR?.substring(OR.lastIndexOf("/") + 1);
            let tipomoneda = item.TIPO_MONEDA?.replace("(", "");
            tipomoneda = tipomoneda?.replace(")", "");
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
              TIPO_MONEDA: tipomoneda,
              CLAVE: item.CLAVE,
              FECHA_CADUCIDAD: item.Fecha_Caducidad,
              LOTE: item.Lote,
              CANTIDAD_RECIBIDA: item.CANTIDAD_RECIBIDA || item.Cantidad,
              "Fecha Fabricación": item.Fecha_Fabircada,
              PRECIO_SIN_IVA:
                item.PRECIO_SIN_IVA?.replace("$", "") ||
                item.Presion_sin_iva?.replace("$", ""),
              IVA: item.IVA === null ? "0" : item.IVA,
            };
          }
        )
      );
      const ws2 = XLSX.utils.json_to_sheet(
        this.state.filteredDataf.map(
          (item: {
            RAZON_SOCIAL: string;
            TIPO_MONEDA: string;
            NO_LICITACION: any;
            NO_CONTRATO: any;
            RFC_LABORATORIO: any;
            NO_ORDEN_REPOSICION_UNOPS: any;
            FECHA_SELLO_RECEPCION: any;
            CLAVE: any;
            PROCEDENCIA: any;
            REGISTRO_SANITARIO: any;
            Registro_Sanitario: any;
            MARCA: any;
            Fecha_Fabircada: any;
            Fecha_Caducidad: any;
            Lote: any;
            CANTIDAD_RECIBIDA: any;
            Cantidad: any;
            PRECIO_SIN_IVA: string;
            Presion_sin_iva: string;
          }) => {
            let tipomoneda = item.TIPO_MONEDA?.replace("(", "");
            tipomoneda = item.TIPO_MONEDA?.replace(")", "");
            const filtefedar = this.state.listafederal?.filter(
              (a) =>
                item?.RAZON_SOCIAL?.toLowerCase().indexOf(
                  a.Title.toLowerCase()
                ) >= 0
            );
            const rfcvalue =
              filtefedar.length > 0 ? filtefedar[0].RFC : item.RFC_LABORATORIO;
            return {
              CLAS_PTAL_OL: "098316150905",
              NO_LICITACION: item.NO_LICITACION,
              NO_CONTRATO: item.NO_CONTRATO,
              RFC_LABORATORIO: rfcvalue,
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
              PRECIO_SIN_IVA:
                item.PRECIO_SIN_IVA?.replace("$", "") ||
                item.Presion_sin_iva?.replace("$", ""),
              TIPO_MONEDA: tipomoneda,
            };
          }
        )
      );
      const ws3 = XLSX.utils.json_to_sheet(
        this.state.filteredDataf.map(
          (item: {
            NO_ORDEN_REPOSICION_UNOPS: any;
            ENTIDAD_FEDERATIVA: any;
            CLAVE: any;
            CANTIDAD_RECIBIDA: any;
            Cantidad: any;
            NO_REMISION: any;
          }) => {
            return {
              CLAS_PTAL_OL: "098316150905",
              NO_ORDEN_REPOSICION_UNOPS: item.NO_ORDEN_REPOSICION_UNOPS,
              ENTIDAD_FEDERATIVA: item.ENTIDAD_FEDERATIVA,
              CLAVE: item.CLAVE,
              CANTIDAD_RECIBIDA: item.CANTIDAD_RECIBIDA || item.Cantidad,
              NO_REMISION: item.NO_REMISION,
            };
          }
        )
      );
      XLSX.utils.book_append_sheet(wb, ws, "WMS IZEL");
      XLSX.utils.book_append_sheet(wb, ws2, "PCCA");
      XLSX.utils.book_append_sheet(wb, ws3, "PCC2");
      XLSX.writeFile(wb, "Factura.xlsx");
    };

    const textFieldStyles: Partial<ITextFieldStyles> = {
      fieldGroup: { width: 300 },
    };
    return (
      <>
        <section>
          <div
            style={{
              padding: "8px",
              display: "flex",
              alignItems: "flex-end",
              justifyContent: "space-between",
              flexWrap: "wrap",
            }}
          >
            <TextField
              label="Buscar por Número de Orden"
              type="search"
              value={this.state.NumOrderSearch}
              onChange={(e) => {
                this.setState({
                  NumOrderSearch: (e.target as HTMLInputElement).value,
                });
              }}
              styles={textFieldStyles}
            />

            <TextField
              label="Buscar por Clave"
              type="search"
              value={this.state.ClaveSearch}
              onChange={(e) => {
                this.setState({
                  ClaveSearch: (e.target as HTMLInputElement).value,
                });
              }}
              styles={textFieldStyles}
            />

            <TextField
              label="Buscar por Lote"
              type="search"
              value={this.state.LoteSearch}
              onChange={(e) => {
                this.setState({
                  LoteSearch: (e.target as HTMLInputElement).value,
                });
              }}
              styles={textFieldStyles}
            />

            <div
              style={{
                width: "300px",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-end",
              }}
            >
              <TextField
                label="Buscar por Fecha Desde"
                type="date"
                value={this.state.datefrom}
                onChange={(e) => {
                  this.setState({
                    datefrom: (e.target as HTMLInputElement).value,
                  });
                }}
              />

              <TextField
                label="Hasta"
                type="date"
                value={this.state.dateto}
                onChange={(e) => {
                  this.setState({
                    dateto: (e.target as HTMLInputElement).value,
                  });
                }}
                min={this.state.datefrom}
                disabled={!this.state.datefrom}
              />
            </div>
            </div>
            <div
              style={{
                paddingTop:"10px",
                width: "100%",
                display: "flex",
                justifyContent: "flex-end",
                alignItems: "flex-end",
              }}>
            <DefaultButton
              text="Buscar"
              allowDisabledFocus
              onClick={() => this.handleFilter()}
            />
  </div>
  <div
              style={{
                paddingTop:"10px",
                width: "100%",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-end",
              }}>
                
            <Sendata
              selectedRows={rowdataselet}
              description={""}
              Remisiones={this.props.Remisiones}
              DatosAI={this.props.DatosAI}
              ListCheck={undefined}
              Tablareglas={undefined}
              ListaValidaciom={undefined}
              Documentos={undefined}
              context={this.props.context}
              
            />
            <DefaultButton
              text="Exportar"
              allowDisabledFocus
              onClick={() => handleOnExport()}
            />
 
            </div>
          <br />
          {this.state.loading && (
            <Spinner label="Loading items..." size={SpinnerSize.large} />
          )}
          {!this.state.loading && (
            <>
              <DataTable
                columns={this.state.columns}
                data={this.state.filteredData}
                pagination
                onRowClicked={(rowData) => {
                  this.handleRowSelection([
                    ...this.state.selectedRows,
                    {
                      ...rowData,
                      selected: !rowData.selected,
                    },
                  ]);
                }}
                selectableRows // Enable row selection
                selectableRowsHighlight // Highlight selected rows
                onSelectedRowsChange={(rowData) => {
                  this.handleRowSelection(rowData);
                }}
                conditionalRowStyles={conditionalRowStyles}
                progressPending={this.state.pending}
              />
            </>
          )}
        </section>
        <Modal
          isOpen={this.state.showModal}
          onDismiss={this._hideModal}
          isBlocking={false}
          styles={modalStyles}
        >
          <div>
            <h1>Errores en Documento</h1>
            {this.state.cuerpo && (
              <>
                <h2>Documentos faltantes</h2>
                <p>{this.state.cuerpo}</p>
              </>
            )}
            {this.state.Entidadfederativatabla && (
              <>
                {" "}
                <h2>Errores en Entidad federativa</h2>
                <p>{this.state.Entidadfederativatabla}</p>
              </>
            )}
            {this.state.FechaRegistro && (
              <>
                {" "}
                <h2>Errores en Fecha registro sanitario</h2>
                <p>{this.state.FechaRegistro}</p>
              </>
            )}

            {this.state.cartacanjetabla && (
              <>
                {" "}
                <h2>Errores en Carta canje lote</h2>
                <p>{this.state.cartacanjetabla}</p>
              </>
            )}

            {this.state.cartacanjeclave && (
              <>
                {" "}
                <h2>Errores en Carta canje Clave</h2>
                <p>{this.state.cartacanjeclave}</p>
              </>
            )}
            {this.state.cartacanjefecha && (
              <>
                {" "}
                <h2>Errores en Carta canje fecha</h2>
                <p>{this.state.cartacanjefecha}</p>
              </>
            )}
            {this.state.ordenreposicionsuma && (
              <>
                {" "}
                <h2>Errores en Orden reposición</h2>
                <p>{this.state.ordenreposicionsuma}</p>
              </>
            )}
            {this.state.errornoai && (
              <>
                {" "}
                <h2>Errores en lectura AI</h2>
                <p>{this.state.errornoai}</p>
              </>
            )}
            <div
              style={{
                padding: "8px",
                display: "flex",
                alignItems: "flex-end",
                justifyContent: "space-between",
                flexWrap: "wrap",
              }}
            >
              <DefaultButton
                text="Eliminar"
                allowDisabledFocus
                onClick={() => this.handledelete(this.state.titleId)}
                styles={{
                  root: {
                    right: "50",
                    textalign: "right",
                    top: "0",
                    backgroundColor: "#f00",
                    color: "#fff",
                  },
                }}
              />
              <DefaultButton
                onClick={this._hideModal}
                text="Close"
                styles={{
                  root: {
                    right: "0",
                    textalign: "center",
                    top: "0",
                    backgroundColor: "#f00",
                    color: "#fff",
                  },
                }}
              />
            </div>
          </div>
        </Modal>
      </>
    );
  }
}
