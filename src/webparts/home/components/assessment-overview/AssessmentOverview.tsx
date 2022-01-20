import { DirectionalHint, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem } from "office-ui-fabric-react/lib/Pivot";
import React, { useCallback, useEffect, useState } from "react";
import { Redirect } from "react-router-dom";
import * as Constants from "../../common/Constants";
import * as Custom from "../../common/CustomStyles";
import "../organization/OrganizationOverview.scss";
import PageHeader from "../shared/headers/PageHeader";
import PivotTabs from "../shared/pivot-tabs/PivotTabs";
import SearchTextBox from "../shared/search-textbox/SearchTextBox";
import "./AssessmentOverview.scss";
import AssessmentServiceProvider from "./AssessmentServiceProvider";

export interface IAssessmentOverview {
    id?: number;
    processName: string;
    siteName: string;
    orgName: string;
    assessmentName: string;
    status: string;
    completionProgress: number;
    infoIconText: string;
}
export interface IAssessmentOverviewProps {
    orgName?: string;
    siteName?: string;
    context: WebPartContext;
    rootSiteURL: string;
}

const AssessmentOverview = ({
    orgName,
    siteName,
    context,
    rootSiteURL
}: IAssessmentOverviewProps) => {
    const [search, setsearch] = React.useState(String);
    const [filteredTableListByCriteria, setfilteredTableListByCriteria] =
        React.useState(new Array<IAssessmentOverview>());
    const [tableList, setTableList] = React.useState(
        new Array<IAssessmentOverview>()
    );
    const [selectedKey, setSelectedKey] = React.useState("all");
    const [filteredTableListBySearch, setfilteredTableListBySearch] =
        React.useState(new Array<IAssessmentOverview>());
    const [isLoaded, setIsLoaded] = React.useState(false);
    const [objectKeys, setObjectKeys] = React.useState(new Array<string>());
    const [allCount, setAllCount] = React.useState(String);
    const [completedCount, setCompletedCount] = React.useState(String);
    const [inProgressCount, setInprogressCount] = React.useState(String);
    const [notStartedCount, setNotStartedCount] = React.useState(String);
    const [pageintialized, setPageInitialized] = useState(false);
    const [redirectBack, setRedirectBack] = useState(Boolean);

    const tabsItems = [
        {
            header: `${Constants.AllText} | ${allCount}`,
            tableItemkey: `${Constants.AllText.toLowerCase()}`,
        },
        {
            header: `${Constants.CompletedText} | ${completedCount}`,
            tableItemkey: `${Constants.CompletedKey}`,
        },
        {
            header: `${Constants.InProgressText} | ${inProgressCount}`,
            tableItemkey: `${Constants.InProgressKey}`,
        },
        {
            header: `${Constants.NotStartedText} | ${notStartedCount}`,
            tableItemkey: `${Constants.NotStartedKey}`,
        },
    ];

    useEffect(() => {
        if (tableList.length === 0 && !pageintialized) {
            setPageInitialized(true);
            getAssessmentDetails();
        }
    }, [tableList.length, pageintialized]);

    const getAssessmentDetails = async () => {
        try {
            let assessmentLists = new Array<IAssessmentOverview>();
            let objectkeys = new Array<string>();

            let serviceProvider = new AssessmentServiceProvider(context);

            await serviceProvider
                .getAssessmentDetails(orgName, siteName)
                .then(async (response: any) => {
                    let i = 0;

                    while (i < response.length) {
                        if (response[i]) {
                            let infoContent = await serviceProvider.getInfoTextForAssessment(response[i].AssessmentName);

                            assessmentLists.push({
                                processName: Constants.AssessmentOverview,
                                orgName: orgName,
                                siteName: siteName,
                                assessmentName: response[i].AssessmentName,
                                status: response[i].Status.replace(/\s+/g, '').toLowerCase(),
                                completionProgress: response[i].CompletionProgress,
                                infoIconText: infoContent
                            });
                        }
                        i++;
                    }
                })
                .then(() => {
                    if (assessmentLists.length > 0) {
                        objectkeys = Object.keys(assessmentLists[0]);
                        objectkeys = ["ASSESSMENT", ...Constants.GridHeaders];
                    }
                });
            setIsLoaded(true);
            setObjectKeys(objectkeys);
            setTableList(assessmentLists);
            setfilteredTableListByCriteria(assessmentLists);
            setfilteredTableListBySearch(assessmentLists);
        } catch (error) {
            console.error(
                "PVSS_AssessmentOverview_getOrganizationDetails",
                error
            );
        }
    };

    //  Refresh completion progress and status for Assessments data based on organization and site.
    const refreshAssessmentListDataOnFilters = async () => {
        try {
            let serviceProvider = new AssessmentServiceProvider(context);
            setSelectedKey("all");
            setIsLoaded(false);
            await serviceProvider.refreshAssessmentListDataOnFilters(orgName, siteName)
                .then((response) => {
                    //Reload the page.
                    setPageInitialized(false);
                    setTableList(Array<IAssessmentOverview>());
                    setIsLoaded(true);
                });
        } catch (error) {
            console.error(
                "PVSS_AssessmentOverview_refreshAssessmentListDataOnFilters",
                error
            );
        }
    };

    const filterDataByStatus = (
        assessmentData: IAssessmentOverview[],
        status?: string,
        searchedValue?: string
    ) => {
        if (
            (status === `${Constants.AllText.toLowerCase()}` || !status) &&
            !searchedValue
        ) {
            setfilteredTableListByCriteria(assessmentData);
        } else if (
            status === `${Constants.AllText.toLowerCase()}` &&
            searchedValue
        ) {
            setfilteredTableListByCriteria(
                assessmentData.filter((x) => {
                    return (
                        x.assessmentName &&
                        x.assessmentName
                            .toLowerCase()
                            .includes(searchedValue.toLowerCase())
                    );
                })
            );
        } else if (status && searchedValue) {
            setfilteredTableListByCriteria(
                assessmentData.filter((x) => {
                    return (
                        x.status &&
                        x.status.toLowerCase() === status.toLowerCase() &&
                        x.assessmentName &&
                        x.assessmentName
                            .toLowerCase()
                            .includes(searchedValue.toLowerCase())
                    );
                })
            );
        } else if (status && !searchedValue) {
            setfilteredTableListByCriteria(
                assessmentData.filter((x) => {
                    return (
                        x.status &&
                        x.status.toLowerCase() === status.toLowerCase()
                    );
                })
            );
        }
    };

    const searchItems = (value: string) => {
        setsearch(value);
        filterDataByStatus(tableList, selectedKey, value);

        setfilteredTableListBySearch(
            tableList.filter((x) => {
                return (
                    x.assessmentName &&
                    x.assessmentName.toLowerCase().includes(value.toLowerCase())
                );
            })
        );
    };
    const setTabCountByStatus = useCallback(() => {
        let allValue = filteredTableListBySearch.length.toString();
        let completedValue = filteredTableListBySearch
            .filter((x) => {
                return (
                    x.status &&
                    x.status.toLowerCase() === `${Constants.CompletedKey}`
                );
            })
            .length.toString();

        let inprogressValue = filteredTableListBySearch
            .filter((x) => {
                return (
                    x.status &&
                    x.status.toLowerCase() === `${Constants.InProgressKey}`
                );
            })
            .length.toString();

        let notStartedValue = filteredTableListBySearch
            .filter((x) => {
                return (
                    x.status &&
                    x.status.toLowerCase() === `${Constants.NotStartedKey}`
                );
            })
            .length.toString();

        setAllCount(allValue);
        setCompletedCount(completedValue);
        setInprogressCount(inprogressValue);
        setNotStartedCount(notStartedValue);
    }, [filteredTableListBySearch]);
    const handleLinkClick = (item?: PivotItem) => {
        if (item) {
            setSelectedKey(item.props.itemKey!);
            filterDataByStatus(tableList, item.props.itemKey, search);
        }
    };

    return (
        <React.Fragment>
            {redirectBack ? (<Redirect to="/siteOverview" />) :
                (
                    <div className="main-container">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm12">
                                <PageHeader
                                    pageHeaderLabel={Constants.AssessmentOverview}
                                    componentName={Constants.AssessmentOverview}
                                    setRedirectBack={setRedirectBack}
                                />
                            </div>
                        </div>
                        <table className="ms-Table ms-Table--selectable assessment-table">
                            <tbody>
                                <tr>
                                    <td>
                                        <div className="assessment-searchbox">
                                            <SearchTextBox
                                                onSearch={searchItems}
                                                placeholderText={
                                                    Constants.AssessmentSearchBoxHolder
                                                }
                                                searchBoxStyles={Custom.SearchBoxStyles}
                                            />
                                        </div>
                                    </td>
                                    <td>
                                        <span className="assessment-key-label">Organization:</span>
                                        <span className="assessment-value-label">{orgName}</span>
                                    </td>
                                    <td>
                                        <span className="assessment-key-label">Site:</span>
                                        <span className="assessment-value-label">{siteName}</span>
                                    </td>
                                    <td>
                                        <div className={`assessment-cursor`}>
                                            <TooltipHost
                                                content="Click refresh to see latest % completion"
                                                delay={2}
                                                tooltipProps={Custom.tooltipProps}
                                                styles={Custom.tooltipStyles}
                                                directionalHint={DirectionalHint.leftCenter}
                                            >
                                                <Icon aria-label="refresh" iconName="refresh" onClick={refreshAssessmentListDataOnFilters} className={Custom.refreshIconStyles.icon} />
                                            </TooltipHost>
                                        </div>

                                    </td>
                                </tr>
                            </tbody>
                        </table>

                        <div className="ms-Grid-col ms-sm12 grid-panel">
                            {!isLoaded && <Spinner
                                label={Constants.LoadingSpinnerMessage}
                                size={SpinnerSize.large}
                            />}
                            {isLoaded && (
                                <PivotTabs
                                    selectedKey={selectedKey}
                                    objectKeys={objectKeys}
                                    search={search}
                                    filteredTableListByCriteria={
                                        filteredTableListByCriteria
                                    }
                                    tabsItems={tabsItems}
                                    setTabCountByStatus={setTabCountByStatus}
                                    handleLinkClick={handleLinkClick}
                                    pivotPanelStyles={Custom.PivotPanelStyles}
                                    context={context}
                                    orgName=""
                                    rootSiteURL={rootSiteURL}
                                />
                            )}
                        </div>
                    </div>
                )
            }

        </React.Fragment>
    );
};

export default AssessmentOverview;
