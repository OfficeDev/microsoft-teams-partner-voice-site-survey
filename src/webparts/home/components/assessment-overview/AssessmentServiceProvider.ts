import { WebPartContext } from "@microsoft/sp-webpart-base";
import CommonServices from "../../common/CommonServices";
import commonServices from '../../common/CommonServices';
import * as Constant from "../../common/Constants";
import * as Helper from "../../common/FunctionHelper";

class AssessmentServiceProvider {
    private spcontext: WebPartContext;
    private commonServiceManager: commonServices;

    public constructor(spcontext: WebPartContext) {
        this.spcontext = spcontext;
        this.commonServiceManager = new commonServices(this.spcontext);
    }

    //  Get the list items for Assessment Overview.
    public getAssessmentDetails = async (organizationName: string, site: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            let filter = "Title eq '" + organizationName + "' and Site eq '" + site + "'";

            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.AssessmentOverview, filter)
                .then(async (response: any[]) => {
                    if (response.length > 0) {
                        resolve(response);
                    }
                })
                .then(await this.refreshAssessmentListDataOnFilters(organizationName, site))
                .catch((exception) => {
                    console.error("PVSS_AssessmentServiceProvider_getAssessmentDetails \n", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while getting the assessment details." });
                });
        });
    }

    //  Logic to calculate and update the assesments status and completion progress based on filled columns.
    public updateAssesmentLists = async (orgName?: string): Promise<any> => {
        return new Promise(async (resolve) => {
            let columnNameForCalculation: string;
            try {
                // Fetching all items from Assessments Master List  
                await this.commonServiceManager
                    .getAllListItems(Constant.AssessmentsMaster)
                    .then(async (assesmentDetails) => {

                        for (let i = 0; i < assesmentDetails.length; i++) {
                            columnNameForCalculation = assesmentDetails[i][Constant.AssessmentColumn];
                            if (columnNameForCalculation == null)
                                columnNameForCalculation = "false";
                            // Getting recent updated assessments since last refresh
                            if (columnNameForCalculation != "false") {
                                let updatedAssesmentDetails = await this.getupdatedAssesmentLists(assesmentDetails[i].Title, columnNameForCalculation, orgName);
                                //Updating the status and % completion for recent updated assessments in Assessment Overview list   
                                await this.commonServiceManager
                                    .bulkUpdateListItem(Constant.AssessmentOverview, updatedAssesmentDetails);
                            }
                        }
                    })
                    .then(async (result) => {
                        // Updating the refresh time in the config list
                        await this.commonServiceManager.UpdateLastRefreshData();
                        resolve(result);
                    });
            }
            catch (exception) {
                console.error("PVSS_AssessmentServiceProvider_updateAssesmentLists \n", exception);

            }
        });

    }

    //  Logic to calculate and update the assesments status and completion progress based on organization and site.
    public updateAssesmentListsOnFilters = async (organizationName: string, site: string): Promise<any> => {
        return new Promise(async (resolve) => {
            let columnNameForCalculation: string;
            let listOfPromise = [];
            try {
                let lastRefreshedTimeStamp = await this.commonServiceManager.getListItemsWithSpecificColumns(Constant.RefreshDataList, "LastUpdatedTime");

                // Fetching all items from Assessments Master List  
                const updateAssesments = async () => await this.commonServiceManager
                    .getAllListItems(Constant.AssessmentsMaster)
                    .then(async (assesmentDetails) => {

                        for (let i = 0; i < assesmentDetails.length; i++) {
                            columnNameForCalculation = assesmentDetails[i][Constant.AssessmentColumn];
                            if (columnNameForCalculation == null)
                                columnNameForCalculation = "false";
                            // Getting recent updated assessments since last refresh
                            if (columnNameForCalculation != "false") {
                                let updatedAssesmentDetails = await this.getupdatedAssesmentListsOnFilters(assesmentDetails[i].Title, columnNameForCalculation, lastRefreshedTimeStamp, organizationName, site);
                                //Updating the status and % completion for recent updated assessments in Assessment Overview list   
                                await this.commonServiceManager
                                    .bulkUpdateListItem(Constant.AssessmentOverview, updatedAssesmentDetails);
                            }
                        }
                    });

                await updateAssesments()
                    .then(async () => {
                        let filter = "Title eq '" + organizationName + "' and Site eq '" + site + "'";
                        let updatedSitesDetails = await this.getUpdatedSiteOverviewListsOnFilters(filter);
                        const siteUpdatePromise = await this.commonServiceManager
                            .bulkUpdateListItem(Constant.SiteOverview, updatedSitesDetails);

                        Promise.resolve(siteUpdatePromise)
                            .then(async (isSiteUpdated) => {
                                if (isSiteUpdated == true) {
                                    let updatedOrgDetails = await this.getUpdatedOrganizationOverviewListsOnFilters("Title eq '" + organizationName + "'", updatedSitesDetails);

                                    const orgUpdatePromise = await this.commonServiceManager
                                        .bulkUpdateListItem(Constant.OrganizationOverview, updatedOrgDetails);

                                    Promise.resolve(orgUpdatePromise)
                                        .then(async (isOrgUpdated) => {
                                            if (isOrgUpdated == true) {
                                                // Updating the refresh time in the config list
                                                await this.commonServiceManager.UpdateLastRefreshData();
                                                resolve(true);
                                            }
                                        });
                                }
                            });
                    });
            }
            catch (exception) {
                console.error("PVSS_AssessmentServiceProvider_updateAssesmentListsOnFilters \n", exception);

            }
        });

    }

    //  Refresh completion progress and status for all Assessments data.
    public refreshAssessmentListData = async (): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.updateAssesmentLists()
                .then((result) => { resolve(result); })
                .catch((exception) => {
                    console.error("PVSS_AssessmentServiceProvider_refreshAssessmentListData \n", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while updating the Assesment lists." });
                });
        });
    }

    //  Refresh completion progress and status for Assessments data based on organization and site.
    public refreshAssessmentListDataOnFilters = async (organizationName: string, site: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {

            await this.updateAssesmentListsOnFilters(organizationName, site)
                .then((result) => { if (result == true) { resolve(result); } })
                .catch((exception) => {
                    console.error("PVSS_AssessmentServiceProvider_refreshAssessmentListDataOnFilters \n", exception);
                    reject({ reponseCode: 500, Error: "Something went wrong while updating the Assesment lists." });
                });
        });
    }

    //Getting Info Icon Text for Assessments from SP list
    public getInfoTextForAssessment = async (assessmentName: string) => {
        const item: any = await this.commonServiceManager.getItemsWithOnlyFilter(Constant.AssessmentsMaster, "Title eq '" + assessmentName + "'");

        if (item.length > 0) {
            return item[0][Constant.InfoIconText];
        }
        else {
            return Constant.InfoContentWithNoCalculation;
        }
    }

    //  Logic to calculate the status and completion progress for the assessments.
    private getupdatedAssesmentLists = async (assesmentName: string, columnNameForCalculation: string, orgName?: string, site?: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let updatedAssesmentOverView = [];

            try {
                let lastRefreshedTimeStamp = await this.commonServiceManager.getListItemsWithSpecificColumns(Constant.RefreshDataList, "LastUpdatedTime");
                let filter = (site && orgName) ? "Title eq '" + orgName + "' and Site eq '" + site + "' and Modified gt '" + lastRefreshedTimeStamp[0].LastUpdatedTime + "'" :
                    orgName ? "Title eq '" + orgName + "' and Modified gt '" + lastRefreshedTimeStamp[0].LastUpdatedTime + "'"
                        : "Modified gt '" + lastRefreshedTimeStamp[0].LastUpdatedTime + "'";

                // Fetching the recently updated items from the assessment list
                await this.commonServiceManager
                    .getUpdatedListItemsOnly(assesmentName, filter)
                    .then(async (updatedAssesments) => {

                        if (updatedAssesments.length > 0) {
                            for (let eachAssessment = 0; eachAssessment < updatedAssesments.length; eachAssessment++) {

                                let filterCondition = "Title eq '" + updatedAssesments[eachAssessment].Title + "' and Site eq '" + updatedAssesments[eachAssessment].Site + "'";
                                //Fetching all items from the assessment list for a specific Org and Site to calculate the % completion
                                let assesmentDetails = await this.commonServiceManager.getItemsWithOnlyFilter(assesmentName, filterCondition);

                                if (assesmentDetails.length > 0 && columnNameForCalculation != "false") {

                                    let emptyAnswers = assesmentDetails.filter((e) => e[columnNameForCalculation] == "" || e[columnNameForCalculation] == null).length;
                                    let completedAnswers = assesmentDetails.filter((e) => e[columnNameForCalculation] != null).length;

                                    if (emptyAnswers == assesmentDetails.length) {

                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.NotStartedText,
                                            updatedAssesments[eachAssessment], 0));
                                    }
                                    else if (completedAnswers == assesmentDetails.length) {


                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.CompletedText,
                                            updatedAssesments[eachAssessment], 100));
                                    }
                                    else {
                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.InProgressText,
                                            updatedAssesments[eachAssessment],
                                            Math.round((completedAnswers * 100) / assesmentDetails.length)));
                                    }
                                }

                            }
                            Promise.all(updatedAssesmentOverView).then(async () => {
                                resolve(updatedAssesmentOverView);
                            });
                        }
                        resolve(updatedAssesmentOverView);
                    });
            }
            catch (exception) {
                reject(false);
                console.error("AssessmentServiceProvider_getUpdateAssesmentLists \n", exception);

            }

        });
    }

    //  Logic to calculate the status and completion progress for the assessments based on organization and site..
    private getupdatedAssesmentListsOnFilters = async (assesmentName: string, columnNameForCalculation: string, lastRefreshedTimeStamp: any, orgName?: string, site?: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let updatedAssesmentOverView = [];
            this.commonServiceManager = new CommonServices(this.spcontext);

            try {

                let filter = "Title eq '" + orgName + "' and Site eq '" + site + "' and Modified gt '" + lastRefreshedTimeStamp[0].LastUpdatedTime + "'";

                // Fetching the recently updated items from the assessment list
                await this.commonServiceManager
                    .getUpdatedListItemsOnly(assesmentName, filter)
                    .then(async (updatedAssesments) => {

                        if (updatedAssesments.length > 0) {
                            for (let eachAssessment = 0; eachAssessment < updatedAssesments.length; eachAssessment++) {

                                let filterCondition = "Title eq '" + updatedAssesments[eachAssessment].Title + "' and Site eq '" + updatedAssesments[eachAssessment].Site + "'";
                                //Fetching all items from the assessment list for a specific Org and Site to calculate the % completion
                                let assesmentDetails = await this.commonServiceManager.getItemsWithOnlyFilter(assesmentName, filterCondition);

                                if (assesmentDetails.length > 0 && columnNameForCalculation != "false") {

                                    let emptyAnswers = assesmentDetails.filter((e) => e[columnNameForCalculation] == "" || e[columnNameForCalculation] == null).length;
                                    let completedAnswers = assesmentDetails.filter((e) => e[columnNameForCalculation] != null).length;

                                    if (emptyAnswers == assesmentDetails.length) {

                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.NotStartedText,
                                            updatedAssesments[eachAssessment], 0));
                                    }
                                    else if (completedAnswers == assesmentDetails.length) {


                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.CompletedText,
                                            updatedAssesments[eachAssessment], 100));
                                    }
                                    else {
                                        updatedAssesmentOverView.push(this.createUpdatedAssesmentObject(
                                            assesmentName,
                                            Constant.InProgressText,
                                            updatedAssesments[eachAssessment],
                                            Math.round((completedAnswers * 100) / assesmentDetails.length)));
                                    }
                                }

                            }
                            Promise.all(updatedAssesmentOverView).then(async () => {
                                resolve(updatedAssesmentOverView);
                            });
                        }
                        resolve(updatedAssesmentOverView);
                    });
            }
            catch (exception) {
                reject(false);
                console.error("PVSS_AssessmentServiceProvider_getupdatedAssesmentListsOnFilters \n", exception);

            }

        });
    }

    private createUpdatedAssesmentObject = (
        assesmentName: string,
        status: string,
        eachOrgDetail: any,
        completionProgress: number) => {

        return {
            AssessmentName: assesmentName,
            Status: status,
            CompletionStatus: completionProgress,
            OrganizationName: eachOrgDetail.Title,
            Site: eachOrgDetail.Site,
            filterParams: "Title eq '" + eachOrgDetail.Title + "' and Site eq '" + eachOrgDetail.Site + "' and AssessmentName eq '" + assesmentName + "'"
        };
    }

    //  Logic to calculate and update the site's status and completion progress based on filter condition.
    public getUpdatedSiteOverviewListsOnFilters = async (filter?: string): Promise<any> => {
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

    // Logic to calculate the site's status and completion progress based on filter condition.
    public getUpdatedOrganizationOverviewListsOnFilters = async (filter?: string, updatedSitesDetails?: any): Promise<any> => {
        let organizations = [];

        return new Promise(async (resolve) => {

            await this.commonServiceManager
                .getItemsWithOnlyFilter(Constant.SiteOverview, filter)
                .then((siteDetails) => {

                    siteDetails.map(site => {
                        const item = updatedSitesDetails.find(({ Site }) => Site === site.Site);
                        if (item) { site.CompletionProgress = item.CompletionStatus; }
                        return site;
                    });

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

export default AssessmentServiceProvider;