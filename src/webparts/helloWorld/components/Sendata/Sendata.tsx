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
import { DefaultButton, Modal } from "office-ui-fabric-react";

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
  cost: string;
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
    super(props); this.state = {
       // Initialize an array to store error responses
    };
    this.state = {
      dataToSend: {
        key1: "value1",
        key2: "value2",
      },
      apiErrorResponses: [],
      apiResponse: null,
      showModal:false 
    };
  }

  executeAction = async () => {
    const { selectedRows } = this.props;

    if (Object.keys(selectedRows).length>0) {
      const serverResponse = await this.getRemisionDataTable();

     
const usuario=this.props.context.pageContext.user.displayName;

      const selectedArray = (selectedRows as any)?.selectedRows;
      
      selectedArray.forEach(async (selectedRow: any) => {
        const parts =selectedRow.CLUES_x002f_CPTALDESTINO.split(/[-\s]+/); // Split by either hyphen or space
const firstPart = parts[0].trim();
const currentDateTime = moment();

// Format the date and time as "YYYY-MM-DD HH:MM:SS"
const formattedDateTime = currentDateTime.format('YYYY-MM-DD HH:mm:ss');
const plancode= selectedRow.Title.indexOf('LOTIS') === -1? "IMSS":"LOTIS";
        const iZELShippingNotification: IZelshippingNotification = {
          plantCode: plancode,
          documentNumber: "NA",
          documentType: "5",
          documentDate:formattedDateTime,
          billOfLading: selectedRow.NO_CONTRATO|| "NA",
          containerSeal: selectedRow.NO_LICITACION|| "NA",
          provider: selectedRow.RFC_LABORATORIO|| "NA",
          providerName: selectedRow.RAZON_SOCIAL|| "NA",
          auxiliary01: "TEM_AMB",
          auxiliary02: "MXN",
          auxiliary03: "ESTANDAR",
          auxiliary04: selectedRow.NO_REMISION|| "NA",
          auxiliary05: selectedRow.NO_REMISION|| "NA",
          auxiliary08: usuario,
          auxiliary09: firstPart|| "NA",
          auxiliary10:"NA",
          auxiliary11:"NA",
          items: [],
        };
          const { ID_x002d_Remision } = selectedRow;
          const matchingResponse = serverResponse.find(
            (responseObj: any) =>
              responseObj.ID_x002d_Remision === ID_x002d_Remision
          );

          if (matchingResponse) {
            let contador = 0;
            // Wrap matchingResponse in an array to ensure it's treated as an array
            const matchingResponseArray = Array.isArray(matchingResponse) ? matchingResponse : [matchingResponse];
            const extractedValue = selectedRow.CLAVE.match(/\d{3}\.\d{3}\.\d{4}/);

            const result = extractedValue ? extractedValue[0] : "NA";
            matchingResponseArray.forEach((matchingResponseItem: any) => {
              const newItem: Item = {
                itemNumber: (contador + 1).toString(),
                material: result,
                batch: matchingResponseItem.Lote,
                uom: "NA",
                quantity: matchingResponseItem.Cantidad|| "NA",
                productionDate: this.formatDateString(matchingResponseItem.Fecha_Fabircada),
                expireDate:this.formatDateString(matchingResponseItem.Fecha_Caducidad),
                cost: "0",
                stockType: "UN",
                auxiliary01: selectedRow.REGISTRO_SANITARIO || matchingResponseItem.Registro_Sanitario || "NA",
                auxiliary02: selectedRow.MARCA || matchingResponseItem.Marca|| "NA",
                auxiliary03: selectedRow.PROCEDENCIA || matchingResponseItem.Procedendia|| "NA",
              };
              iZELShippingNotification.items.push(newItem);
              contador=contador+1;

            });

            
          
          const resutlapi =  await this.sendDataToAPI2(iZELShippingNotification);
           if (resutlapi) {
            console.log("api", resutlapi);
          } else {
            console.log("API request did not return a valid result.");
          }
          }
        });

     if(  this.state.apiErrorResponses.length>0){
      this.setState({showModal:true});
     }
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
    const selectedArray = (selectedRows as any)?.selectedRows;
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
          .filter(query)
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
  padStart = (input: string, targetLength: number, padString: string): string => {
    while (input.length < targetLength) {
      input = padString + input;
    }
    return input;
  }
  
  convertDate(match: string, p1: string, p2: string | undefined): string {
    let year: string;
  if (p2 !== undefined && p2.length === 4) {
    year = p2; // Assuming 4-digit year
  } else if (p1.length >= 6) {
    year = `20${p1.slice(6)}`;
  } else {
    // Handle invalid cases where year is missing or not in the expected format
    return "Invalid date format";
  }
    const month = p1.slice(3, 5); // Extract characters at indices 3 and 4
const day = p1.slice(0, 2);   // Extract characters at indices 0 and 1
return `${year}-${month}-${day}`;
  }
  formatDateString = (input: string): string => {
    // Define regular expressions to match different date formats
    const regex1 = /^(\d{2}\/\d{2}\/\d{2})$/; // "06/03/23"
    const regex2 = /^:\s*(\d{2}\/\d{2}\/\d{2})$/; // ": 06/03/23"
    const regex3 = /^(\d{2}\/\d{2}\/\d{2})\.$/; // "06/03/23."
    const regex4 = /^(\d{2}\/\d{2}\/\d{2})\s+(\d{4})$/; // "06/03/23 1233"
  
    if (regex1.test(input)) {
      return input.replace(regex1, this.convertDate);
    } else if (regex2.test(input)) {
      return input.replace(regex2, this.convertDate);
    } else if (regex3.test(input)) {
      return input.replace(regex3, this.convertDate);
    } else if (regex4.test(input)) {
      return input.replace(regex4, this.convertDate);
    } else {
      return "Invalid date format";
    }
  }
  sendDataToAPI2 = async (iZELShippingNotification: any) => {
    const myHeaders = new Headers();
myHeaders.append("SID-API-KEY", SID_API_KEY);
myHeaders.append("SID-USER-TOKEN",SID_USER_TOKEN);
myHeaders.append("Content-Type", "application/json");

const raw = JSON.stringify({
  iZELShippingNotification
});

const requestOptions: RequestInit = {
  method: 'POST',
  headers: myHeaders,
  body: raw,
  redirect: 'follow' as RequestRedirect, // Explicitly specify the type
};
const response = await fetch("https://wapps02.gruposid.com.mx:4443/gas401/ws/r/izelwms/856/v1/add", requestOptions);
const result = await response.text();
return result; // Return the result

// fetch("https://wapps02.gruposid.com.mx:4443/gas401/ws/r/izelwms/856/v1/add", requestOptions)
//   .then(response => response.text())
//   .then(result => console.log(result))
//   .catch(error => console.log('error', error));
//   }

  
  // sendDataToAPI = async (datesend: any) => {
    
  //   const apiUrl =
  //     "https://wapps02.gruposid.com.mx:4443/gas401/ws/r/izelwms/856/v1/add";
      
  //   const requestOptions = {
  //     method: "POST",
  //     headers: {
  //       "Accept":"*/*",
  //       "Content-Type": "application/json",
  //       "SID-API-KEY": SID_API_KEY,
  //       "SID-USER-TOKEN": SID_USER_TOKEN,
  //     },
  //     body: JSON.stringify(datesend),
  //   };
   
  //   const response = await fetch("https://wapps02.gruposid.com.mx:4443/gas401/ws/r/izelwms/856/v1/add", requestOptions);
  //   const result = await response.text();
  //   return result; // Return the result
    // fetch(apiUrl, requestOptions)
    // .then((response) => {
    //   console.log("Response status:", response.status);
    //   return response.json();
    // })
    // .then((data) => {
    //   console.log("API response:", data);
    //   if (data.error) {
    //     // If the response has an error, add it to the error responses array
    //     this.setState((prevState: { apiErrorResponses: any; }) => ({
    //       apiErrorResponses: [...prevState.apiErrorResponses, data],
    //     }));
    //   }
    //   // Handle other cases or successful responses as needed
    //   // ...
    // })
    // .catch((error) => {
    //   // Handle network errors or other exceptions
    //   console.error("API request error:", error);
    // });
  };
  private _hideModal = (): void => {

    this.setState({ showModal: false });
  };

  render() {
    return (
      <><div>
        <DefaultButton
          text="Enviar WMS"
          allowDisabledFocus
          onClick={this.executeAction} />

      </div><Modal
        isOpen={this.state.showModal}
        onDismiss={this._hideModal}
        isBlocking={false}
        styles={modalStyles}
      >
          <div>
            <h1>Errores en Documento</h1>
            {this.state.apiErrorResponses.map((error:any, index:any) => (
            <li key={index}>{error}</li>
    ))}
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
                }} />
            </div>
          </div>
        </Modal></>
    );
  }
}
