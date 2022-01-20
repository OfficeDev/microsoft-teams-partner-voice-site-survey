import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from '../../common/CommonServices';
import * as Helper from "../../common/FunctionHelper";
import AssessmentServiceProvider from "../assessment-overview/AssessmentServiceProvider";
import * as Constant from './../../common/Constants';

class SiteServiceProvider {
    private spcontext: WebPartContext;
    private commonServiceManager: commonServices;
    private assessmentServiceProvider: AssessmentServiceProvider;

    public constructor(spcontext: WebPartContext) {
        this.spcontext = spcontext;
        this.commonServiceManager = new commonServices(this.spcontext);
        this.assessmentServiceProvider = new AssessmentServiceProvider(this.spcontext);
    }

    //  Get the list items for Site Overview.
    public getSiteDetails = async (organizationName: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.SiteOverview, "Title eq '" + organizationName + "'")
                .then(async (response: any[]) => {
                    if (response.length > 0) {
                        response.sort((a, b) => (a.created > b.created) ? 1 : -1);
                        resolve(response);
                    }
                })
                .catch((exception) => {
                    console.error("PVSS_SiteServiceProvider_getSiteDetails", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while getting the site details." });
                });
        });
    }

    public getTenantName = async (organizationName: string): Promise<string> => {
        return new Promise(async (resolve, reject) => {

            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.OrganizationOverview, "Title eq '" + organizationName + "'")
                .then(async (response: any[]) => {
                    if (response.length > 0) {
                        resolve(response[0].Tenant_x0020_Name);
                    }
                });
        });
    }
    //  Create an item for the Site Overview list.
    public createSiteAssessment = async (site: any): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let filter = "Title eq '" + site.Title + "' and Site eq '" + site.Site + "'";

            await this.commonServiceManager
                .createListItem(Constant.SiteOverview, site, filter)
                .then(async (response: any) => {
                    if (response.data) {
                        resolve(response);
                    }
                })
                .catch((exception) => {
                    console.error("PVSS_SiteServiceProvider_createSiteAssessment", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while creating the site." });
                });
        });
    }

    //  This method gets the data from MasterList and creates assessment items in the target Assessment List
    public createAssessmentsForSite = async (site: any): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.commonServiceManager
                .createFromMasterList(site.Title, site.Site)
                .then(async (response: any) => {
                    if (response) {
                        resolve(response);
                    }
                    else {
                        reject(response);
                    }
                })
                .catch((exception) => {
                    console.error("SiteServiceProvider_createAssessmentsForSite", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while creating the assesments for sites." });
                });
        });
    }
    //  Refresh completion progress and status for all list data.
    public refreshSiteListData = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.assessmentServiceProvider.updateAssesmentLists()
                .then(async () => {
                    await this.updateSiteOverviewLists()
                        .then((response) => {
                            resolve(response);
                        })
                        .catch((exception) => {
                            console.error("PVSS_SiteServiceProvider_refreshSiteListData", exception);
                            reject({ reponseCode: 500, Error: "Something went wrong while updating the site list." });
                        });
                })
                .catch((exception) => {
                    console.error("PVSS_SiteServiceProvider_refreshSiteListData", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while updating the Assesment lists." });
                });




        });
    }

    //  Refresh completion progress and status for all list data based on organization name.
    public refreshSiteListDataOnFilters = async (organizationName: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            let filter = "Title eq '" + organizationName + "'";
            await this.assessmentServiceProvider.updateAssesmentLists(organizationName)
                .then(async () => {
                    await this.updateSiteOverviewListsOnFilters(filter)
                        .then((response) => {
                            if (response == true) { resolve(response); }
                        })
                        .catch((exception) => {
                            console.error("PVSS_SiteServiceProvider_refreshSiteListDataOnFilters", exception);
                            reject({ reponseCode: 500, Error: "Something went wrong while updating the site list." });
                        });
                })
                .catch((exception) => {
                    console.error("PVSS_SiteServiceProvider_refreshSiteListDataOnFilters", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while updating the Assesment lists." });
                });
        });
    }

    //  Logic to calculate and update the site's status and completion progress based on filled columns.
    public updateSiteOverviewLists = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let updatedSitesDetails = await this.getUpdatedSiteOverviewLists();

            this.commonServiceManager
                .bulkUpdateListItem(Constant.SiteOverview, updatedSitesDetails)
                .then((result) => { resolve(result); });
        });
    }

    //  Logic to calculate and update the site's status and completion progress based on filter condition.
    public updateSiteOverviewListsOnFilters = async (filter: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let updatedSitesDetails = await this.getUpdatedSiteOverviewListsOnFilters(filter);

            this.commonServiceManager
                .bulkUpdateListItem(Constant.SiteOverview, updatedSitesDetails)
                .then((result) => { resolve(result); });
        });
    }

    private getUpdatedSiteOverviewLists = async (): Promise<any> => {
        return new Promise(async (resolve) => {
            let sites = [];

            await this.commonServiceManager
                .getAllListItems(Constant.AssessmentOverview)
                .then((assesmentDetails) => {
                    let organizedItems = Helper.groupBy(assesmentDetails, item => item.Title);

                    organizedItems.forEach((organizedItem) => {
                        let organizedSites = Helper.groupBy(organizedItem, item => item.Site);
                        organizedSites.forEach((site) => {
                            site = site.filter((filterSite) => !filterSite.AssessmentName.includes(Constant.AnalogDevicesList)
                                && !filterSite.AssessmentName.includes(Constant.UserInformationList));

                            let completionProgress = site.reduce((a, b) => { return a + b["CompletionProgress"]; }, 0);
                            completionProgress = completionProgress / site.length;

                            let notStartedSites = site.filter(siteProperties =>
                                Helper.removeSpaceToLowercase(siteProperties.Status) == Constant.NotStartedKey).length;

                            let completedSites = site.filter(siteProperties =>
                                Helper.removeSpaceToLowercase(siteProperties.Status) == Constant.CompletedKey).length;

                            if (site.length == notStartedSites) {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.NotStartedText, 0));
                            }
                            else if (site.length == completedSites) {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.CompletedText, 100));
                            }
                            else {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.InProgressText, Math.round(completionProgress)));
                            }
                        });
                    });
                });
            resolve(sites);
        });

    }

    private getUpdatedSiteOverviewListsOnFilters = async (filter: string): Promise<any> => {
        return new Promise(async (resolve) => {
            let sites = [];

            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.AssessmentOverview, filter)
                .then((assesmentDetails) => {
                    let organizedItems = Helper.groupBy(assesmentDetails, item => item.Title);

                    organizedItems.forEach((organizedItem) => {
                        let organizedSites = Helper.groupBy(organizedItem, item => item.Site);
                        organizedSites.forEach((site) => {
                            site = site.filter((filterSite) => !filterSite.AssessmentName.includes(Constant.AnalogDevicesList)
                                && !filterSite.AssessmentName.includes(Constant.UserInformationList));

                            let completionProgress = site.reduce((a, b) => { return a + b["CompletionProgress"]; }, 0);
                            completionProgress = completionProgress / site.length;

                            let notStartedSites = site.filter(siteProperties =>
                                Helper.removeSpaceToLowercase(siteProperties.Status) == Constant.NotStartedKey).length;

                            let completedSites = site.filter(siteProperties =>
                                Helper.removeSpaceToLowercase(siteProperties.Status) == Constant.CompletedKey).length;

                            if (site.length == notStartedSites) {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.NotStartedText, 0));
                            }
                            else if (site.length == completedSites) {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.CompletedText, 100));
                            }
                            else {

                                sites.push(this.createUpdatedSiteOverviewObject(
                                    site, Constant.InProgressText, Math.round(completionProgress)));
                            }
                        });
                    });
                });
            resolve(sites);
        });

    }

    private createUpdatedSiteOverviewObject = (
        site: any,
        status: string,
        completionProgress: number) => {

        return {
            OrganizationName: site[0].Title,
            Site: site[0].Site,
            Status: status,
            CompletionStatus: completionProgress,
            filterParams: "Title eq '" + site[0].Title + "' and Site eq '" + site[0].Site + "'"
        };
    }
}

export default SiteServiceProvider;