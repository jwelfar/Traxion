import * as React from "react";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/batching";
import "@pnp/sp/items/get-all";
import { IHelloWorldProps } from "../IHelloWorldProps";
import { SPFI, SPFx, spfi } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as moment from "moment";
import { IDetailsTableItem } from "../HelloWorld";
import { SID_API_KEY, SID_USER_TOKEN } from "../../config/apiconfig";

export interface Root {
  iZELShippingNotification: IZelshippingNotification;
}

export interface IZelshippingNotification {
  plantCode: string;
  documentNumber: string;
  documentDate?: string;
  documentType: string;
  billOfLading: string;
  containerSeal: string;
  storageLocation?: string;
  provider?: string;
  providerName: string;
  auxiliarydate01?: string;
  auxiliarydate02?: string;
  auxiliarydate03?: string;
  auxiliarydate04?: string;
  auxiliarydate05?: string;
  auxiliary01: string;
  auxiliary02: string;
  auxiliary03: string;
  auxiliary04: string;
  auxiliary05: string;
  auxiliary06?: string;
  auxiliary07?: string;
  auxiliary08: string;
  auxiliary09: string;
  auxiliary10?: string;
  auxiliary11?: string;
  auxiliary12?: string;
  auxiliary13?: string;
  auxiliary14?: string;
  auxiliary15?: string;
  items: Item[];
}

export interface Item {
  itemNumber: string;
  sscc?: string;
  material: string;
  batch: string;
  uom: string;
  stockCategory?: string;
  serialNumber?: string;
  quantity: number;
  productionDate: string;
  expireDate: string;
  stockType?: string;
  cost: number;
  auxiliary01: string;
  auxiliary02: string;
  auxiliary03: string;
  auxiliary04?: string;
  auxiliary05?: string;
}

interface SendataProps {
  context: WebPartContext; // Replace with the appropriate context type for SharePoint
  selectedRows: any[]; // Define the type of selectedRows
}

export interface ITableState {
  RemisionesSelected: IDetailsTableItem[];
}
let _sp: SPFI = null;
export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
//let _sp: SPFI = null;
export default class Sendata extends React.Component<
  SendataProps & IHelloWorldProps,
  any
> {
  constructor(props: SendataProps & IHelloWorldProps) {
    super(props);
    this.state = {
      dataToSend: {
        key1: "value1",
        key2: "value2",
      },
      apiResponse: null,
    };
  }

  executeAction = async () => {
    const { selectedRows } = this.props;

    if (selectedRows.length > 0) {
      const serverResponse = await this.getRemisionDataTable();

      const shippingNotifications: IZelshippingNotification[] = [];

      selectedRows.forEach((row) => {
        row.selectedRows.forEach((selectedRow: any) => {
          const { ID_x002d_Remision } = selectedRow;
          const matchingResponse = serverResponse.find(
            (responseObj: any) =>
              responseObj.ID_x002d_Remision === ID_x002d_Remision
          );
          if (matchingResponse) {
            const iZELShippingNotification: IZelshippingNotification = {
              plantCode: "AVIORBR",
              documentNumber: "",
              documentType: "5",
              billOfLading: selectedRow.NO_CONTRATO,
              containerSeal: selectedRow.NO_LICITACION,
              provider: "",
              providerName: selectedRow.RAZON_SOCIAL,
              auxiliary01: "TEM_AMB",
              auxiliary02: "MXN",
              auxiliary03: "ESTANDAR",
              auxiliary04: selectedRow.NO_REMISION,
              auxiliary05: "",
              auxiliary08: "",
              auxiliary09: selectedRow.CLUES_x002f_CPTALDESTINO,
              items: [],
            };

            const newItem: Item = {
              itemNumber: "1",
              material: selectedRow.CLAVE,
              batch: matchingResponse.Lote,
              uom: "PZA",
              quantity: matchingResponse.Cantidad,
              productionDate: matchingResponse.Fecha_Fabircada,
              expireDate: matchingResponse.Fecha_Caducidad,
              cost: 0,
              auxiliary01: selectedRow.REGISTRO_SANITARIO,
              auxiliary02: selectedRow.MARCA,
              auxiliary03: selectedRow.PROCEDENCIA,
            };

            iZELShippingNotification.items.push(newItem);
            shippingNotifications.push(iZELShippingNotification);
          }
        });
      });

      console.log("iZELShippingNotification:", shippingNotifications);

      // const filteredAndMappedData = selectedRows.map((row) => {
      //   return row.selectedRows.map((selectedRow: any) => {
      //     const { ID_x002d_Remision } = selectedRow;
      //     const matchingResponse = serverResponse.find(
      //       (responseObj: any) =>
      //         responseObj.ID_x002d_Remision === ID_x002d_Remision
      //     );
      //     if (matchingResponse) {
      // console.log("Filtered remision:", matchingResponse);
      //       return { ...selectedRow, ...matchingResponse };
      //     }
      //     return null;
      //   });
      // });

      // const flattenedData = filteredAndMappedData.reduce((acc, curr) => {
      //   return acc.concat(curr);
      // }, []);

      // console.log("Filtered data:", flattenedData);
    }
    console.log("selected", selectedRows);
  };

  private formatDate(date: string): string {
    const datef = new Date(date);
    const formattedDate = moment(datef).format("YYYY-MM-DD");
    return formattedDate;
  }

  private async getRemisionDataTable(): Promise<any> {
    const { selectedRows } = this.props;
    const selectedArray = (selectedRows[0] as any)?.selectedRows;
    const validDates = selectedArray
      .map((remision: any) => {
        const date = new Date(remision.Created);
        return date.toString() !== "Invalid Date" ? date.getTime() : null;
      })
      .filter((timestamp: any) => timestamp !== null);
    const maxDate = new Date(Math.max(...validDates));
    const minDate = new Date(Math.min(...validDates));

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
    let items: any = [];
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
            "CFN",
            "Fecha_Fabircada",
            "Fecha_Caducidad",
            "ID_x002d_Remision",
            "Created",
            "Id"
          )
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
          RemisionesSelected: response,
        });
        // console.log("response", this.state.RemisionesSelected);
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

  sendDataToAPI = () => {
    const apiUrl =
      "https://wapps02.gruposid.com.mx:4443/gas401/ws/r/izelwms/856/v1/add";

    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "SID-API-KEY": SID_API_KEY,
        "SID-USER-TOKEN": SID_USER_TOKEN,
      },
      body: JSON.stringify(this.state.dataToSend),
    };

    fetch(apiUrl, requestOptions)
      .then((response) => response.json())
      .then((data) => {
        // Handle the API response here
        this.setState({ apiResponse: data });
      })
      .catch((error) => {
        // Handle errors, e.g., network issues
        console.error("API request error:", error);
      });
  };

  render() {
    context: this.context;
    return (
      <div>
        <button onClick={this.executeAction}>Execute Action</button>
      </div>
    );
  }
}
