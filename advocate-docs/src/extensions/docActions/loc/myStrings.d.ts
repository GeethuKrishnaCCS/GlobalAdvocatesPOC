declare interface IDocActionsCommandSetStrings {
  Preview: string;
  Merge: string;
  EmailPdf: string;
  EmailVersion: string;
  NotAvailableYet: string;
}

declare module 'DocActionsCommandSetStrings' {
  const strings: IDocActionsCommandSetStrings;
  export = strings;
}
