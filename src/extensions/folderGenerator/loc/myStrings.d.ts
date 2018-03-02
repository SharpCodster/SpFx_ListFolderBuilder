declare interface IFolderGeneratorCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'FolderGeneratorCommandSetStrings' {
  const strings: IFolderGeneratorCommandSetStrings;
  export = strings;
}
