import pnp, { List, ListEnsureResult, ItemAddResult, FieldAddResult, FolderAddResult, ListAddResult } from "sp-pnp-js";

import { 
  override 
} from '@microsoft/decorators';

import { 
  Log,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import { 
  Dialog 
} from '@microsoft/sp-dialog';

import { 
  SPHttpClient, 
  SPHttpClientResponse, 
  SPHttpClientConfiguration,
  ISPHttpClientOptions 
} from '@microsoft/sp-http';

import * as strings from 'FolderGeneratorCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFolderGeneratorCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'FolderGeneratorCommandSet';

export default class FolderGeneratorCommandSet extends BaseListViewCommandSet<IFolderGeneratorCommandSetProperties> {

  private listId:string;
  private selectedItemId:string; 
  private libraryName:string;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FolderGeneratorCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;

      if (compareOneCommand.visible) {
        this.listId =  this.context.pageContext.list.id.toString();
        this.libraryName = event.selectedRows[0].getValueByName("Title");
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this.CheckListExistance();
        break;
      case 'COMMAND_2':
        this.CreateFolders();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private CheckListExistance() : void {
    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.libraryName}')`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((listObject:any) => {    
          if (listObject.hasOwnProperty("error") && listObject.error.code == "-1, System.ArgumentException") {
            this.CreateLibrary();
          } else {
            Dialog.alert(`The List "${this.libraryName}" alredy exists.`);
          }
        });
      }); 
  }

  private CreateLibrary() : void {
    pnp.sp.web.lists.ensure(this.libraryName, "A document library", 101).then(
      (value: ListEnsureResult) => { 
        console.log(value.created); 
        console.log(value.list.toUrl());
        console.log(value.data);
        if (value.created) {
          this.CreateFolders();
        } else {
          Dialog.alert(`The List "${this.libraryName}" has a problem.`);
        }
      }, 
      (error: any) => { 
        console.log(error);
      });
  }

  private CreateFolders() : void {

    let batch = pnp.sp.web.createBatch();

    pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder`);
    pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/Second Folder`);
    pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder/ComplexFolder`);
  
    batch.execute().then(d => console.log("Done"));
  }

}
