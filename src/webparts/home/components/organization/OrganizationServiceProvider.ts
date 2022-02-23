import * as Constant from './../../common/Constants';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from '../../common/CommonServices';
import * as Helper from '../../common/FunctionHelper';
import SiteServiceProvider from '../site-overview/SiteServiceProvider';
import AssessmentServiceProvider from '../assessment-overview/AssessmentServiceProvider';

class OrganizationServiceProvider {
    private spcontext: WebPartContext;
    private commonServiceManager: commonServices;
    private siteServiceProvider: SiteServiceProvider;
    private assessmentServiceProvider: AssessmentServiceProvider;

    public constructor(spcontext: WebPartContext) {
        this.spcontext = spcontext;
        this.commonServiceManager = new commonServices(this.spcontext);
        this.siteServiceProvider = new SiteServiceProvider(this.spcontext);
        this.assessmentServiceProvider = new AssessmentServiceProvider(this.spcontext);
    }

    //  Get the list items for Organization Overview.
    public getOrganizationDetails = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.commonServiceManager
                .getAllListItems(Constant.OrganizationOverview)
                .then(async (response: any[]) => {
                    if (response.length > 0) {
                        response.sort((a, b) => (a.created > b.created) ? 1 : -1);
                        resolve(response);
                    }
                })
                .catch((exception) => {
                    console.error("PVSS_OrganizationServiceProvider_getOrganizationDetails", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while getting the organization details." });
                });
        });
    }

    //  Get the drop down values for Deployed regions.
    public getDeployedRegions = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.commonServiceManager
                .getListItemsWithSpecificColumnsSorted(Constant.DeployedRegion, "Title")
                .then(async (response: any[]) => {
                    if (response.length > 0) {
                        resolve(response);
                    }
                })
                .catch((exception) => {
                    console.error("PVSS_OrganizationServiceProvider_getDeployedRegions", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while getting the deployed regions." });
                });
        });
    }
  //  Create a folder for excel reports in the site
  public createFolderforReport= async (organization: any): Promise<any> => {
    return new Promise(async (resolve, reject) => {

        this.commonServiceManager.createFolder(organization.Title)
        .then(async (response: any) => {
            if (response.data) {
                resolve(response);
            }
        });
    });
}

    //  Create an item for the Organization Overview list.
    public createOrganizationAssessment = async (organization: any): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let filter = "Title eq '" + organization.Title + "'";

            this.commonServiceManager
                .createListItem(Constant.OrganizationOverview, organization, filter)
                .then(async (response: any) => {
                    if (response.data) {
                        resolve(response);
                    }
                })
                .catch((exception) => {
                    console.error("PVSS_OrganizationServiceProvider_createOrganizationAssessment", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while creating the new Organization." });
                });
        });
    }

    //  Update the organization lists.
    public updateOrganizationOverviewLists = async (): Promise<any> => {
        return new Promise(async (resolve) => {

            let updatedSitesDetails = await this.getUpdatedOrganizationOverviewLists();

            this.commonServiceManager
                .bulkUpdateListItem(Constant.OrganizationOverview, updatedSitesDetails)
                .then((result) => { resolve(result); });
        });
    }

    //  Update the organization lists based on the filter condition.
    public updateOrganizationOverviewListsOnFilters = async (filter: string): Promise<any> => {
        return new Promise(async (resolve) => {

            let updatedSitesDetails = await this.getUpdatedOrganizationOverviewListsOnFilters(filter);

            this.commonServiceManager
                .bulkUpdateListItem(Constant.OrganizationOverview, updatedSitesDetails)
                .then((result) => { resolve(result); });
        });
    }

    //  Refresh completion progress and status for all list data.
    public refreshAllListData = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.assessmentServiceProvider.updateAssesmentLists()
                .then(async () => {
                    await this.siteServiceProvider.updateSiteOverviewLists()
                        .then(async () => {
                            await this.updateOrganizationOverviewLists()
                                .then(async (response) => {
                                    resolve(response);
                                })
                                .catch((exception) => {
                                    console.error("PVSS_OrganizationServiceProvider_refreshAllListData", exception);
                                    reject({ reponseCode: 500, Error: "Something went wrong while updating the organization list." });
                                });
                        })
                        .catch((exception) => {
                            console.error("PVSS_OrganizationServiceProvider_refreshAllListData", exception);
                            reject({ reponseCode: 500, Error: "Something went wrong while updating the site list." });
                        });
                })
                .catch((exception) => {
                    console.error("PVSS_OrganizationServiceProvider_refreshAllListData", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while updating the Assesment lists." });
                });
        });
    }

    //  Logic to calculate the site's status and completion progress based on filled columns.
    private getUpdatedOrganizationOverviewLists = async (): Promise<any> => {
        let organizations = [];

        return new Promise(async (resolve) => {
            await this.commonServiceManager
                .getAllListItems(Constant.SiteOverview)
                .then((siteDetails) => {

                    let organaizedItems = Helper.groupBy(siteDetails, item => item.Title);

                    organaizedItems.forEach((organaizedItem) => {

                        let notStartedOrg = organaizedItem.filter(site => Helper.removeSpaceToLowercase(site.Status) == Constant.NotStartedKey).length;
                        let completedOrg = organaizedItem.filter(site => Helper.removeSpaceToLowercase(site.Status) == Constant.CompletedKey).length;

                        let completionProgress = organaizedItem.reduce((a, b) => { return a + b["CompletionProgress"]; }, 0);
                        completionProgress = completionProgress / organaizedItem.length;

                        if (notStartedOrg == organaizedItem.length) {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.NotStartedText, 0));
                        }
                        else if (completedOrg == organaizedItem.length) {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.CompletedText, 100));
                        }
                        else {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.InProgressText, Math.round(completionProgress)));
                        }
                    });
                });
            resolve(organizations);
        });
    }

    //  Logic to calculate the site's status and completion progress based on filter condition.
    private getUpdatedOrganizationOverviewListsOnFilters = async (filter: string): Promise<any> => {
        let organizations = [];

        return new Promise(async (resolve) => {
            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.SiteOverview, filter)
                .then((siteDetails) => {

                    let organaizedItems = Helper.groupBy(siteDetails, item => item.Title);

                    organaizedItems.forEach((organaizedItem) => {

                        let notStartedOrg = organaizedItem.filter(site => Helper.removeSpaceToLowercase(site.Status) == Constant.NotStartedKey).length;
                        let completedOrg = organaizedItem.filter(site => Helper.removeSpaceToLowercase(site.Status) == Constant.CompletedKey).length;

                        let completionProgress = organaizedItem.reduce((a, b) => { return a + b["CompletionProgress"]; }, 0);
                        completionProgress = completionProgress / organaizedItem.length;

                        if (notStartedOrg == organaizedItem.length) {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.NotStartedText, 0));
                        }
                        else if (completedOrg == organaizedItem.length) {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.CompletedText, 100));
                        }
                        else {

                            organizations.push(this.createUpdatedOrganizationOverviewObject(
                                organaizedItem, Constant.InProgressText, Math.round(completionProgress)));
                        }
                    });
                });
            resolve(organizations);
        });
    }

    private createUpdatedOrganizationOverviewObject = (
        organaizedItem: any,
        status: string,
        completionProgress: number) => {

        return {
            OrganizationName: organaizedItem[0].Title,
            Status: status,
            CompletionStatus: completionProgress,
            filterParams: "Title eq '" + organaizedItem[0].Title + "'"
        };
    }
}

export default OrganizationServiceProvider;