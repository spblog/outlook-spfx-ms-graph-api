declare interface IOutlookHostedAddinWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'OutlookHostedAddinWebPartStrings' {
  const strings: IOutlookHostedAddinWebPartStrings;
  export = strings;
}
