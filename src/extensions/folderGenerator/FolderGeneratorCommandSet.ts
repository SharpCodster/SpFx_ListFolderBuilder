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
  private folderName:string;

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
        this.folderName = event.selectedRows[0].getValueByName("Title");
      }
    }
  }

  private GetList(title:string, description:string) {
      return `{ Title: '${title}', Description: '${description}', BaseTemplate: 100 }`;
  }

  private GetDocLibrary(title:string, description:string) {
    return `{ Title: '${title}', Description: '${description}', BaseTemplate: 101 }`;
  }

  private DoSomething() {
    // let html: string = '';
    // this.context.spHttpClient.get(
    //   this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
    //   .then((response: SPHttpClientResponse) => {
    //     response.json().then((listsObjects: any) => {
    //       listsObjects.value.forEach(listObject => {
    //         html += `
    //                 <ul>
    //                     <li>
    //                         <span class="ms-font-l">${listObject.Title}</span>
    //                     </li>
    //                 </ul>`;
    //       });
    //       this.domElement.querySelector('#lists').innerHTML = html;
    //     });
    //   }); 

    let spOpts: ISPHttpClientOptions = {

      body: this.GetDocLibrary(this.folderName, "A test folder")

    };

    this.context.spHttpClient.post(
      this.context.pageContext.web.absoluteUrl + `/_api/web/lists`,
      SPHttpClient.configurations.v1,
      spOpts
    ).then((response:SPHttpClientResponse) => {
        response.json().then((responseJSON:JSON) => {
          console.log(responseJSON);
        })
    });
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne} + Id of the list: ${this.listId} + Folder Name: ${this.folderName}`);
        this.DoSomething();
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
