declare interface IPeopleSearchStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'peopleSearchStrings' {
  const strings: IPeopleSearchStrings;
  export = strings;
}
