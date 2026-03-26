declare interface IMyActiveTasksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'MyActiveTasksWebPartStrings' {
  const strings: IMyActiveTasksWebPartStrings;
  export = strings;
}
