import * as Constants from './Constants';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import ProvisioningAssets from "../provisioning/ProvisioningAssets.json";

export interface IMasterData {
  area: string;
  subject: string;
}

let rootSiteURL: string;
export default class CommonServices {
  private spcontext: WebPartContext;

  public constructor(spcontext: WebPartContext) {
    this.spcontext = spcontext;

    let absoluteUrl = this.spcontext.pageContext.web.absoluteUrl;
    let serverRelativeUrl = this.spcontext.pageContext.web.serverRelativeUrl;
    //  Set context for PNP   
    //  When App is added to a Teams
    if (serverRelativeUrl == "/")
      rootSiteURL = absoluteUrl;
    //  when app is added as personal app
    else
      rootSiteURL = absoluteUrl.replace(serverRelativeUrl, "");

    //  Set up URL for pvss site
    rootSiteURL = rootSiteURL + "/" + ProvisioningAssets.inclusionPath + "/" + ProvisioningAssets.sitename;
    sp.setup({
      sp: {
        baseUrl: rootSiteURL
      },
    });
  }

  //  Get list items based on only a filter
  public getItemsWithOnlyFilter = async (listname: string, filterparametres: any): Promise<any> => {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.filter(filterparametres).getAll();
    return items;
  }

  //  Get list items based on a filter and Top
  public getItemsWithOnlyFilterWithTop = async (listname: string, filterparametres: any, topVal: any): Promise<any> => {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.filter(filterparametres).top(topVal).get();
    return items;
  }

  //  This method gets the data from MasterList and creates assessment items in the target Assessment List
  public createFromMasterList = async (organizationName: string, siteName: string): Promise<any> => {
    return new Promise<any>(async (resolve, reject) => {
      let failedCount = 0;
      try {
        let assessments: any[] = await this.getListItemsWithSpecificColumns(Constants.MasterList, "Title");
        let uniqueAssessments = [...new Set(assessments.map(item => item.Title))];

        for (let assessmentName of uniqueAssessments) {

          let assessmentOverview: any = {
            "Title": organizationName,
            "Status": Constants.NotStartedText,
            "AssessmentName": assessmentName,
            "Site": siteName,
            "CompletionProgress": 0
          };

          await sp.web.lists.getByTitle(Constants.AssessmentOverview).items.add(assessmentOverview);

          const itemsForAssessment: any[] = await this.getItemsWithOnlyFilter(Constants.MasterList, "Title eq '" + assessmentName + "'");

          let assessmentList = sp.web.lists.getByTitle(assessmentName);
          let hasSubject = false;
          let hasArea = false;
          let assessmentOverviewMultiple: any = {};

          const assessmentListItems = await assessmentList.fields.filter("Hidden eq false and ReadOnlyField eq false").get();
          assessmentListItems.forEach(column => {
            if (column.Title == "Subject") { hasSubject = true; }
            if (column.Title == "Area") { hasArea = true; }
          });

          itemsForAssessment.forEach(async function a(v) {
            if (hasArea && hasSubject)
              assessmentOverviewMultiple = {
                Title: organizationName,
                Site: siteName,
                Area: v.Area,
                Subject: v.Subject
              };
            else if (hasSubject && !hasArea)
              assessmentOverviewMultiple = {
                Title: organizationName,
                Site: siteName,
                Subject: v.Subject
              };
            else if (hasArea && !hasSubject)
              assessmentOverviewMultiple = {
                Title: organizationName,
                Site: siteName,
                Area: v.Area,

              };
            else
              assessmentOverviewMultiple = {
                Title: organizationName,
                Site: siteName

              };
            await sp.web.lists.getByTitle(assessmentName).items.add(assessmentOverviewMultiple);
          });
        }
      } catch (error) {

        console.log("PVSS_CommonServices_createFromMasterList \n", error);
        failedCount = failedCount + 1;
      }

      if (failedCount == 0) {
        resolve(true);
      }
      else {
        reject(false);
      }

    });
  }

  //  Get all items from a list
  public getAllListItems = async (listname: string): Promise<any> => {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.getAll();
    return items;
  }
  //  Get updated items from the list filtered by Modified date
  public getUpdatedListItemsOnly = async (listname: string, filter: string): Promise<any> => {
    var items: any[] = [];

    items = await sp.web.lists.getByTitle(listname).items.select("Title", "Site").filter(filter).getAll();

    const uniqueArray = items.filter((value, index) => {
      const _value = JSON.stringify(value);
      return index === items.findIndex(item => {
        return JSON.stringify(item) === _value;
      });
    });
    return uniqueArray;
  }

  //  Get all items from a list with specified columns
  public getListItemsWithSpecificColumns = async (listname: string, columns: string): Promise<any> => {
    return await sp.web.lists
      .getByTitle(listname).items.select(columns).getAll();
  }

  //  Get all items from a list with specified columns and sorting
  public getListItemsWithSpecificColumnsSorted = async (listname: string, columns: string): Promise<any> => {
    return await sp.web.lists.getByTitle(listname).items.select(columns).top(1000).orderBy(columns).get();
  }
  //  Create list item
  public createListItem = async (listname: string, data: any, filter): Promise<any> => {
    const itemsForList: any[] = await this.getItemsWithOnlyFilter(listname, filter);
    if (itemsForList.length == 0) {
      return sp.web.lists.getByTitle(listname).items.add(data);
    }
    else {
      return { errorType: "conflict", data: Constants.AlreadyExists };
    }
  }

  // Get Folder URL from reports libray
  public getFolderURL = async (folderName: string): Promise<any> => {
    return rootSiteURL + "/" + Constants.ReportsLibrary + "/" + folderName;
  }
  // Create folder in the library if not exists
  public createFolder = async (folderName: string): Promise<any> => {
    const folder = await sp.web.getFolderByServerRelativeUrl(Constants.ReportsLibrary + "/" + folderName).select('Exists').get();
   
    if (!folder.Exists) {
      return sp.web.getFolderByServerRelativeUrl(Constants.ReportsLibrary).folders.add(folderName);
    }
    else {
      return { errorType: "conflict", data: Constants.AlreadyExists };
    }
  }

  // Create file in the library folder if not exists
  public createFile = async (folderName: string, fileName: string, data: any): Promise<any> => {
      return sp.web.getFolderByServerRelativeUrl(Constants.ReportsLibrary + "/" + folderName).files.add(fileName + ".xlsx", data, true);
  }

  public async UpdateLastRefreshData(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      let list = sp.web.lists.getByTitle(Constants.RefreshDataList);
      const items: any[] = await list.items.top(1).filter("Title eq 'Last Refreshed Time'").get();

      if (items.length > 0) {
        const item = list.items.getById(items[0].Id).update({
          LastUpdatedTime: new Date()
        });
        resolve(true);
      }
    });
  }
  //  Update list item
  public async bulkUpdateListItem(listNameTobeUpdated: string, updatedData: any[]): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {

      try {
        let listName = sp.web.lists.getByTitle(listNameTobeUpdated);

        for (const eachData of updatedData) {
          try {
            let items = await listName.items.top(1).filter(eachData.filterParams).get();
            if (items.length > 0) {
              listName.items.getById(items[0].Id)
                .update({
                  Title: eachData.OrganizationName,
                  Status: eachData.Status,
                  CompletionProgress: eachData.CompletionStatus
                });
            }
          }
          catch (exception) {
            console.log(exception);
          }
        }
        resolve(true);
      }
      catch (error) {
        reject(false);
      }
    });
  }
}
