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
export default class Sendata extends React.Component<SendataProps & IHelloWorldProps, any> {
    
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
    // Simulate some action with selectedRows
    
    const { selectedRows } = this.props;
    // Call the callback function with selectedRows
  console.log("selected",selectedRows);
  await this.getRemisionDataTable();
  };

  private formatDate(date: string): string {
    const datef = new Date(date);
    const formattedDate = moment(datef).format('YYYY-MM-DD');
    return formattedDate;
  }

  private async getRemisionDataTable(): Promise<void> {
    const { selectedRows } = this.props;
    const selectedArray = (selectedRows[0] as any)?.selectedRows;
    const validDates = selectedArray
    .map((remision:any) => {
      const date = new Date(remision.Created);
      return date.toString() !== "Invalid Date" ? date.getTime() : null;
    })
    .filter((timestamp:any) => timestamp !== null);
    const maxDate = new Date(Math.max(...validDates));
    const minDate = new Date(Math.min(...validDates));
   
      this.setState({
        loading: true
      });
      let query = "";
  
  
     // let items: any = [];
      let response: any = [];
  
      if (minDate) {
        if (query.length === 0) {
          query = ("Created ge datetime'" + this.formatDate(minDate.toString()) + "T00:00:00'");
          if (maxDate) {
            query += (" and Created le datetime'" + this.formatDate(maxDate.toString()) + "T23:59:00'");
          }
        }
        else {
          query = ("Created ge datetime'" + this.formatDate(minDate.toString()) + "T00:00:00'");
          if (maxDate) {
            query += (" and Created le datetime'" + this.formatDate(maxDate.toString()) + "T23:59:00'");
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
              'Id'
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
console.log("response",response);
        return response;
      } catch (err) {
        this.setState({
          loading: false
        });
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  sendDataToAPI = () => {
    const apiUrl = "https://your-api-url.com"; // Replace with your API URL
    const requestOptions = {
      method: "POST", // Adjust HTTP method as needed
      headers: {
        "Content-Type": "application/json",
        // Add any other headers required by your API
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
    context: this.context
    return (
      <div>
        <button onClick={this.executeAction}>Execute Action</button>
      </div>
    );
  }
}