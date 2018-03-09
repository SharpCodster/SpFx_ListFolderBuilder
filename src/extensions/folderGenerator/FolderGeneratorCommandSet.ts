import pnp, { 
  List, 
  ListEnsureResult, 
  FolderAddResult
} from "sp-pnp-js";

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
}

const LOG_SOURCE: string = 'FolderGeneratorCommandSet';

export default class FolderGeneratorCommandSet extends BaseListViewCommandSet<IFolderGeneratorCommandSetProperties> {

  private listId:string;
  private selectedItemId:string; 
  private libraryName:string;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FolderGeneratorCommandSet');
    pnp.setup({
      defaultCachingStore: "session", // or "local"
      defaultCachingTimeoutSeconds: 600,
      globalCacheDisable: false // or true to disable caching in case of debugging/testing
    });

    pnp.sp.web.roleDefinitions.filter('').usingCaching().get().then(r => { console.log("Roles Loaded") });
    pnp.sp.web.siteGroups.filter('').usingCaching().get().then(r => { console.log("Groups Loaded") });

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('CREATE_FOLDERS_CMD');
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
      case 'CREATE_FOLDERS_CMD':
        this.CheckListExistance();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private CheckListExistance() : void {
    this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.libraryName}')`, 
    SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
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
    pnp.sp.web.lists.ensure(this.libraryName, "A document library", 101).then((value: ListEnsureResult) => { 
      console.log(`List ${this.libraryName} created`); 
      if (value.created) {
        let batch = pnp.sp.web.createBatch();

        pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder`).then((h:FolderAddResult) => { console.log(`Added: ${h.folder.toUrl()}`); });
        pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/Second Folder`).then((h:FolderAddResult) => { console.log(`Added: ${h.folder.toUrl()}`); });
        pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder/ComplexFolder`).then((h:FolderAddResult) => { console.log(`Added: ${h.folder.toUrl()}`); });
        pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder/ComplexFolder Revenge`).then((h:FolderAddResult) => { console.log(`Added: ${h.folder.toUrl()}`); });
        pnp.sp.web.folders.inBatch(batch).add(`${this.libraryName}/First Folder/ComplexFolder/Sticazzi`).then((h:FolderAddResult) => { console.log(`Added: ${h.folder.toUrl()}`); });

        batch.execute().then(d => { 
          console.log("Created All Folders");
          this.AssignSecurityToFolders(value.list);
        });
      } else {
        Dialog.alert(`The List "${this.libraryName}" has a problem.`);
      }
    }, 
    (error: any) => { console.log(error); });
  }

  private AssignSecurityToFolders(lista:List) : void {
    let batchBreakRole = pnp.sp.web.createBatch();

    lista.inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the list`); });
    lista.items.getById(1).inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the folder with id [1]`); });
    lista.items.getById(2).inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the folder with id [2]`); });
    lista.items.getById(3).inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the folder with id [3]`); });
    lista.items.getById(4).inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the folder with id [4]`); });
    lista.items.getById(5).inBatch(batchBreakRole).breakRoleInheritance(false, false).then(brk => { console.log(`Broke Role Inherithance of the folder with id [5]`); });
    
    batchBreakRole.execute().then(d => { 
      console.log(`All Role Inherithance Broke`);

      pnp.sp.web.roleDefinitions.getByName('Read').usingCaching().get().then(roleRead => {
        pnp.sp.web.roleDefinitions.getByName('Edit').usingCaching().get().then(roleEdit=> {
          pnp.sp.web.siteGroups.getByName('All AM').usingCaching().get().then(group_All_AM => {
            pnp.sp.web.siteGroups.getByName('SG Area North West').usingCaching().get().then(group_SG_Area_NW => {
              pnp.sp.web.siteGroups.getByName('CG HR').usingCaching().get().then(group_CG_HR => {
                
                let batchSecurities = pnp.sp.web.createBatch();
                
                lista.roleAssignments.inBatch(batchSecurities).add(group_All_AM.Id, roleEdit.Id).then(h => { console.log(`Added group [${group_All_AM.LoginName}] with role [${roleEdit.Name}]`); });
                lista.roleAssignments.inBatch(batchSecurities).add(group_SG_Area_NW.Id, roleRead.Id).then(h => { console.log(`Added group [${group_SG_Area_NW.LoginName}] with role [${roleEdit.Name}]`); });
                lista.roleAssignments.inBatch(batchSecurities).add(group_CG_HR .Id, roleRead.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleEdit.Name}]`); });
                
                lista.items.getById(1).roleAssignments.inBatch(batchSecurities).add(group_All_AM.Id, roleEdit.Id).then(h => { console.log(`Added group [${group_All_AM.LoginName}] with role [${roleEdit.Name}]`); });
                lista.items.getById(1).roleAssignments.inBatch(batchSecurities).add(group_CG_HR .Id, roleEdit.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleRead.Name}]`); });
                lista.items.getById(1).roleAssignments.inBatch(batchSecurities).add(group_SG_Area_NW.Id, roleRead.Id).then(h => { console.log(`Added group [${group_SG_Area_NW.LoginName}] with role [${roleEdit.Name}]`); })
              
                lista.items.getById(2).roleAssignments.inBatch(batchSecurities).add(group_All_AM .Id, roleEdit.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleEdit.Name}]`); });
                
                lista.items.getById(3).roleAssignments.inBatch(batchSecurities).add(group_CG_HR .Id, roleEdit.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleRead.Name}]`); });
                lista.items.getById(3).roleAssignments.inBatch(batchSecurities).add(group_SG_Area_NW.Id, roleRead.Id).then(h => { console.log(`Added group [${group_SG_Area_NW.LoginName}] with role [${roleEdit.Name}]`); })
              
                lista.items.getById(4).roleAssignments.inBatch(batchSecurities).add(group_All_AM .Id, roleEdit.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleEdit.Name}]`); });
                
                lista.items.getById(5).roleAssignments.inBatch(batchSecurities).add(group_CG_HR .Id, roleEdit.Id).then(h => { console.log(`Added group [${group_CG_HR .LoginName}] with role [${roleEdit.Name}]`); });

                batchSecurities.execute().then(d => { 
                  console.log("End sercurity assignemnts"); 
                  Dialog.alert("Done"); 
                });
              });
            });
          });
        });
      });
    });
  }
}
