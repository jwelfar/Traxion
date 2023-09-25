
export interface IZELShippingNotificationItem {
    itemNumber: string;
    sscc: string;
    material: string;
    batch: string;
    uom: string;
    stockCategory: string;
    serialNumber: string;
    quantity: number;
    productionDate: string;
    expireDate: string;
    stockType: string;
    cost: number;
    auxiliary01: string;
    auxiliary02: string;
    auxiliary03: string;
    auxiliary04: string;
    auxiliary05: string;
  }
  
  export interface IZELShippingNotification {
    plantCode: string;
    documentNumber: string;
    documentDate: string;
    documentType: string;
    billOfLading: string;
    containerSeal: string;
    storageLocation: string;
    provider: string;
    providerName: string;
    auxiliarydate01: string;
    auxiliarydate02: string;
    auxiliarydate03: string;
    auxiliarydate04: string;
    auxiliarydate05: string;
    auxiliary01: string;
    auxiliary02: string;
    auxiliary03: string;
    auxiliary04: string;
    auxiliary05: string;
    auxiliary06: string;
    auxiliary07: string;
    auxiliary08: string;
    auxiliary09: string;
    auxiliary10: string;
    auxiliary11: string;
    auxiliary12: string;
    auxiliary13: string;
    auxiliary14: string;
    auxiliary15: string;
    items: IZELShippingNotificationItem[];
  }
  