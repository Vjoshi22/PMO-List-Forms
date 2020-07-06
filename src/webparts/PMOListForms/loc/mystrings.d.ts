declare interface IPmoListFormsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'PmoListFormsWebPartStrings' {
  const strings: IPmoListFormsWebPartStrings;
  export = strings;
}
declare module '*.scss' {
  const content: {[className: string]: string};
  export default content;
}