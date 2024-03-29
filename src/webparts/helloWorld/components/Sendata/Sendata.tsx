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
//import { SID_API_KEY, SID_USER_TOKEN } from "../../config/apiconfig";
import { DefaultButton } from "office-ui-fabric-react";
import Swal from "sweetalert2";


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
  quantity: string;
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
  selectedRows: any[];

 // Define the type of selectedRows
}
interface erorrarray{
  Razonsocial:any,
  Remision:any,
  Errorrow:any
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
    super(props); this.state = {
       // Initialize an array to store error responses
    };
 

    this.state = {
      dataToSend: {
        key1: "value1",
        key2: "value2",
      },
      apiError:[],
      apiResponse: null,
      showModalSend:false ,
      loadingbut:false,
      renisiones:[],
      filteredDatafsend:[]
    };
  }
   formatoNumerico= async (numero:string) => {
    // Primero, eliminamos cualquier carácter que no sea un dígito
    numero = numero.replace(/\D/g, '');

    // Definimos los patrones de formato
    const patrones = ['xx.xxx.xx', 'xx.xxx.xxxx.xx'];

    // Elegimos el patrón de formato en función de la longitud del número
    let patron = patrones[0];
    if (numero.length >= 9) {
        patron = patrones[1];
    }

    // Aplicamos el formato
    let resultado = '';
    let indice = 0;
    for (let i = 0; i < patron.length; i++) {
        if (patron[i] === 'x') {
            resultado += numero[indice] || 'x'; // Usamos 'x' si no hay más dígitos
            indice++;
        } else {
            resultado += patron[i];
        }
    }

    return resultado;
}


 formatString= async (inputString:any) => {
  const formattedString = inputString.replace(/[^0-9]/g, "");
  let dotting = formattedString.substring(0,3) + "." + formattedString.substring(3,6) + "." + formattedString.substring(6,10) + "." + formattedString.substring(10,12)+ "."+formattedString.substring(12,14);
  while (dotting[dotting.length-1] === ".")
    dotting = dotting.slice(0,-1);

  return dotting
}

  executeAction = async () => {
    const { selectedRows } = this.props;
    const erorrarrayInstance: erorrarray []= [];
    if (Object.keys(selectedRows).length>0) {
    Swal.fire({
      title: "Enviando ",
      text: `Esto puede tardar uno segundo o minutos depende de la cantidad de archivos a enviar `,
      allowOutsideClick: false,
      footer:"por favor espere",
      didOpen: () => {
        Swal.showLoading()
      }
    });
    this.setState({
      loadingbut: true,
    });
    const apiError:any = [];
    
      const serverResponse = await this.getRemisionDataTable();
      this.setState({remisione:serverResponse});
    
      const usuario=this.props.context.pageContext.user.displayName;
      const selectedArray = (selectedRows as any)?.selectedRows;
      
      for (const selectedRow of selectedArray) {
        
        let firstPart:any;
        if (selectedRow.CLUES_x002f_CPTALDESTINO) {
          // Format the numerical value
          const numeross = this.formatoNumerico(selectedRow.CLUES_x002f_CPTALDESTINO);
            console.log(numeross);
        
          // Split by either hyphen or space
          const parts = selectedRow.CLUES_x002f_CPTALDESTINO.split(/[-\s]+/);
        
          // Check if the first part exists and trim it
          if (parts.length > 0) {
             firstPart = parts[0].trim();
            console.log(firstPart);
          } else {
            console.log('No parts found.');
          }
        } 
        let rfcLaboratorio:any;
        if (selectedRow.RFC_LABORATORIO) {
          rfcLaboratorio = selectedRow.RFC_LABORATORIO.replace(/[-\s]/g, '');
        }
        let calvematerial:any;
        if (selectedRow.CLAVE) {
          calvematerial =await  this.formatString(selectedRow.CLAVE);
        }else{
          calvematerial="NA";
        }
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
          provider: rfcLaboratorio|| "NA",
          providerName: selectedRow.RAZON_SOCIAL|| "NA",
          auxiliary01: "TEM_AMB",
          auxiliary02: "MXN",
          auxiliary03: "ESTANDAR",
          auxiliary04: selectedRow.NO_REMISION|| "NA",
          auxiliary05: selectedRow.NO_REMISION|| "NA",
          auxiliary08: usuario,
          auxiliary09: selectedRow.CLUES_x002f_CPTALDESTINO|| "NA",
          auxiliary10:"NA",
          auxiliary11:"NA",
          items: [],
        };
          const { ID_x002d_Remision } = selectedRow;
          const matchingResponse = serverResponse.filter(
            (responseObj: any) =>
              responseObj.ID_x002d_Remision === ID_x002d_Remision
          );


          if (matchingResponse) {
            let contador = 0;
            // Wrap matchingResponse in an array to ensure it's treated as an array
            const matchingResponseArray = Array.isArray(matchingResponse) ? matchingResponse : [matchingResponse];
            //const extractedValue = selectedRow.CLAVE.match(/\d{1,3}\.\d{3}\.\d{4}\.\d{2}/);
          //  const result = extractedValue ? extractedValue[0] : "NA";
        
            matchingResponseArray.forEach((matchingResponseItem: any) => {
              const newItem: Item = {
                itemNumber: (contador + 1).toString(),
                material: calvematerial,
                batch: matchingResponseItem.Lote || "NA",
                uom: "NA",
                quantity: matchingResponseItem.Cantidad,
                productionDate: (() => {
                  const formattedDate = this.formatDateString(matchingResponseItem.Fecha_Fabircada);
                  return formattedDate !== "Invalid date format" ? formattedDate : "1900-01-01";
                })(),
                expireDate: (() => {
                  const formattedDate = this.formatDateString(matchingResponseItem.Fecha_Caducidad);
                  return formattedDate !== "Invalid date format" ? formattedDate : "1900-01-01";
                })(),
                cost: "0",
                stockType: "UN",
                auxiliary01: selectedRow.REGISTRO_SANITARIO || matchingResponseItem.Registro_Sanitario || "NA",
                auxiliary02: selectedRow.MARCA || matchingResponseItem.Marca|| "NA",
                auxiliary03: selectedRow.PROCEDENCIA || matchingResponseItem.Procedendia|| "NA",
              };
              iZELShippingNotification.items.push(newItem);
              contador=contador+1;

            });
          try{
          const resutlapi =  await this.sendDataToAPI(iZELShippingNotification);
          if (!resutlapi.ok) {
            const responseData = await resutlapi.json();
            console.log("responseerrro",responseData);
            const errorMessagesText = responseData.Errors
            .map((x:any) => `${x} errores}`)
            .join("\n");
            const data1: erorrarray = {
              Remision: selectedRow.NO_REMISION,
              Errorrow: errorMessagesText,
              Razonsocial: selectedRow.RAZON_SOCIAL,
            };
            
            // Push the object to the array
            erorrarrayInstance.push(data1);
                       
            apiError.push(selectedRow);
            
            this.setState({
              loadloadingbuting: false,
            });
            this.setState((prevState:any) => ({
              apiError: [...prevState.apiError, apiError],
            }));
            // Handle API errors here and add them to the apiError array
          
            console.log(resutlapi);
               this.setState({
                loadingbut: false,
      });
      Swal.close();
          }
        } catch (error) {
          this.setState({
            loadingbut: false,
          });
          Swal.close();
          console.log("Error", error);
          // Handle network or other errors here
          apiError.push(selectedRow);
          this.setState((prevState:any) => ({
            apiError: [...prevState.apiError, "Network error or other issue occurred"],
          }));
        }     
          }
        }
        const updatedSelectedArray = selectedArray.filter((x:any) => {
          // Check if selectedRow is present in apiErrorResponses based on a unique identifier,
          // for example, an ID or some other property that uniquely identifies each item
          const existsInApiErrors = apiError.some((ap:any) => {
            // Replace "uniqueIdentifierProperty" with the actual property that uniquely identifies each item
            return x.uniqueIdentifierProperty === ap.uniqueIdentifierProperty;
          });
        
          // If existsInApiErrors is true, keep the item; otherwise, remove it
          return existsInApiErrors;
        });
        this.setState({apiError:apiError})
        if( apiError.length>0){
          selectedArray.length = 0;
         // Clear the original selectedArray
        Array.prototype.push.apply(selectedArray, updatedSelectedArray);
        }
     if( apiError.length>0){
      // const errorMessagesText = selectedArray
      // .map((x:any) => `${x.NO_REMISION} con proveedor: ${x.RAZON_SOCIAL}`)
      // .join("\n");
      const errorMessagesText = erorrarrayInstance.map((x: erorrarray) => `${x.Remision} con proveedor: ${x.Razonsocial} con los siguientes campos:\n${x.Errorrow}\n`);
      Swal.fire({
        title: "Erro al enviar ",
        html: `<div style="white-space: pre-line;">Las siguientes Remisiones no se enviaron:\n${errorMessagesText.join("\n")}</div>`,
        allowOutsideClick: false,
        footer:"por favor verificar"
      });
    }
    else{
      Swal.fire({
        title: "Envío exitoso ",
        text: `Se ha enviado exitosamente los datos.`,
        allowOutsideClick: false,
      });
     
        // Call the callback function to update rowdataselet in the parent component
        this.setState({selectedRows:selectedArray});
      
    }
    this.setState({
      loadingbut: false,
    });
    console.log("selected", selectedRows);
  }
  }

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
      loadingbut: true,
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
      query += " and TipoTabla eq 'RemisionTabla2'"
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
          .filter(query )
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
          loadingbut: false,
        });
        Swal.close();
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
    const regex = /^(\d{2}\/\d{2}\/\d{4})$/; // "25/09/2025"

    // Check if the input is already in the "25/09/2025" format
    if (regex.test(input)) {
      const parts = input.split('/');
          if (parts.length === 3) {
            const [day, month, year] = parts;
            const formattedDate = `${year}-${month}-${day}`;
            return formattedDate;
          } 
      
    }
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
      return "1900-01-01";
    }
  }
  
  sendDataToAPI = async (iZELShippingNotification: any) => {
    try {
      
    
    const apiUrl =
      "https://system-customer-api-test.us-w2.cloudhub.io/api/customer/shipping/notification";
      const requestBody :any ={
        "iZELShippingNotification": iZELShippingNotification
      };
    const requestOptions = {
      method: "POST",
      headers: {

        "Content-Type": "application/json",
        "client_id": 'alksd82938-asdf23-ase23ew',
        "client_secret": 'alksd82938-asdf23-ase23ew',
       
      },
    
      body: JSON.stringify(requestBody),
      
    };
   
   let response:any;
    try {
  
    response= await  fetch(apiUrl, requestOptions);
    } catch (error) {
      console.log("Error",error);
      console.error(`HTTP Error: ${response.status}`);
      return response;
      
    }
    console.log("repsonse",response);
    return response;
   } catch (error) {
    console.log("Error",error);
    this.setState({
      loadingbut: false,
    });
    }
     
}

  render() {
    return (
      <><div>
        
        <DefaultButton
          text={this.state.loadingbut? "Enviando": "Enviar a WMS" }
          allowDisabledFocus
          disabled={this.state.loadingbut }
          onClick={this.executeAction}
           />
  
      </div></>
    );
  }
}
