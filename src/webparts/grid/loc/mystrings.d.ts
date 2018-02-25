declare interface IGridWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TitleFieldLabel: string;
  NumRowsFieldLabel: string;
  NumColsFieldLabel: string;
}

declare module 'GridWebPartStrings' {
  const strings: IGridWebPartStrings;
  export = strings;
}
