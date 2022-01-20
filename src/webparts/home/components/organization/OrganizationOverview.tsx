import { DirectionalHint, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem } from "office-ui-fabric-react/lib/Pivot";
import React, { useCallback, useEffect, useState } from "react";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import { Redirect } from "react-router-dom";
import * as Constants from "../../common/Constants";
import * as Custom from "../../common/CustomStyles";
import PageHeader from "../shared/headers/PageHeader";
import PivotTabs from "../shared/pivot-tabs/PivotTabs";
import SearchTextBox from "../shared/search-textbox/SearchTextBox";
import "./OrganizationOverview.scss";
import OrganizationServiceProvider from "./OrganizationServiceProvider";

export interface IOrganization {
    id?: number;
    processName: string;
    organizationName?: string;
    status: string;
    completionProgress: number;
    tenantName?: string;
    tenantId?: string;
    deployedRegion?: string;
}

interface IOrganizationOverviewProps {
    context: WebPartContext;
    updateOrgState?: (name: string) => void;
}

const OrganizationOverview = ({
    context,
    updateOrgState,
}: IOrganizationOverviewProps) => {
    const [objectKeys, setObjectKeys] = useState(new Array<string>());
    const [tableList, setTableList] = useState(new Array<IOrganization>());
    const [filteredTableListByCriteria, setfilteredTableListByCriteria] =
        useState(new Array<IOrganization>());
    const [filteredTableListBySearch, setfilteredTableListBySearch] = useState(
        new Array<IOrganization>()
    );

    const [pageintialized, setPageInitialized] = useState(false);
    const [search, setsearch] = useState(String);
    const [isLoaded, setIsLoaded] = React.useState(true);
    const [selectedKey, setSelectedKey] = useState("all");
    const [allCount, setAllCount] = useState(String);
    const [completedCount, setCompletedCount] = useState(String);
    const [inProgressCount, setInprogressCount] = useState(String);
    const [notStartedCount, setNotStartedCount] = useState(String);
    const [redirectToOrgAssessment, setRedirectToOrgAssessment] =
        React.useState(String);
    const [redirectToSiteOverview, setRedirectToSiteOverview] =
        React.useState(String);
    const [orgName, setOrgName] = React.useState(String);

    useEffect(() => {
        if (tableList.length === 0 && !pageintialized) {
            setPageInitialized(true);
            getOrganizationDetails();
        }
    }, [tableList.length, pageintialized]);

    useEffect(() => {
        updateOrgState(orgName);
    }, [orgName]);

    const refreshAllListData = async () => {
        try {
            let serviceProvider = new OrganizationServiceProvider(context);
            setSelectedKey("all");
            setIsLoaded(false);
            await serviceProvider.refreshAllListData()
                .then((response) => {
                    setPageInitialized(false);
                    setTableList(Array<IOrganization>());
                    setIsLoaded(true);
                });

        } catch (error) {
            console.error(
                "PVSS_OrganizationOverview_refreshAllListData",
                error
            );
        }
    };

    const getOrganizationDetails = async () => {
        try {
            let organizationLists = new Array<IOrganization>();
            let objectkeys = new Array<string>();
            let serviceProvider = new OrganizationServiceProvider(context);

            await serviceProvider
                .getOrganizationDetails()
                .then((response: any) => {
                    let i = 0;
                    while (i < response.length) {
                        if (response[i]) {
                            organizationLists.push({
                                processName: Constants.OrganizationOverview,
                                organizationName: response[i].Title,
                                status: response[i].Status.replace(/\s+/g, '').toLowerCase(),
                                completionProgress: response[i].CompletionProgress,
                                tenantName: "",
                                tenantId: "",
                                deployedRegion: "",
                            });
                        }
                        i++;
                    }
                })
                .then(() => {
                    if (organizationLists.length > 0) {
                        objectkeys = Object.keys(organizationLists[0]);
                        objectkeys = ["ORGANIZATION", ...Constants.GridHeaders];
                    }
                });

            setObjectKeys(objectkeys);
            setTableList(organizationLists);
            setfilteredTableListByCriteria(organizationLists);
            setfilteredTableListBySearch(organizationLists);
        } catch (error) {
            console.error(
                "PVSS_OrganizationOverview_getOrganizationDetails",
                error
            );
        }
    };

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

    const filterDataByStatus = (
        organizationData: IOrganization[],
        status?: string,
        searchedValue?: string
    ) => {
        try {
            if (
                (status === `${Constants.AllText.toLowerCase()}` || !status) &&
                !searchedValue
            ) {
                setfilteredTableListByCriteria(organizationData);
            } else if (
                status === `${Constants.AllText.toLowerCase()}` &&
                searchedValue
            ) {
                setfilteredTableListByCriteria(
                    organizationData.filter((x) => {
                        return (
                            x.organizationName &&
                            x.organizationName
                                .toLowerCase()
                                .includes(searchedValue.toLowerCase())
                        );
                    })
                );
            } else if (status && searchedValue) {
                setfilteredTableListByCriteria(
                    organizationData.filter((x) => {
                        return (
                            x.status &&
                            x.status.toLowerCase() === status.toLowerCase() &&
                            x.organizationName &&
                            x.organizationName
                                .toLowerCase()
                                .includes(searchedValue.toLowerCase())
                        );
                    })
                );
            } else if (status && !searchedValue) {
                setfilteredTableListByCriteria(
                    organizationData.filter((x) => {
                        return (
                            x.status &&
                            x.status.toLowerCase() === status.toLowerCase()
                        );
                    })
                );
            }
        } catch (error) {
            console.error("PVSS_OrganizationOverview_filterDataByStaus", error);
        }
    };

    const searchItems = (value: string) => {
        try {
            setsearch(value);
            filterDataByStatus(tableList, selectedKey, value);

            setfilteredTableListBySearch(
                tableList.filter((x) => {
                    return (
                        x.organizationName &&
                        x.organizationName
                            .toLowerCase()
                            .includes(value.toLowerCase())
                    );
                })
            );
        } catch (error) {
            console.error("PVSS_OrganizationOverview_SearchItems", error);
        }
    };

    const setTabCountByStatus = useCallback(() => {
        try {
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
        } catch (error) {
            console.error("PVSS_OrganizationOverview_setTabCountByStatus", error);
        }
    }, [filteredTableListBySearch]);

    const handleLinkClick = (item?: PivotItem) => {
        try {
            if (item) {
                setSelectedKey(item.props.itemKey!);
                filterDataByStatus(tableList, item.props.itemKey, search);
            }
        } catch (error) {
            console.error("PVSS_OrganizationOverview_handleLinkClick", error);
        }
    };

    return (
        <React.Fragment>
            {redirectToOrgAssessment ? (
                <Redirect to={`/${redirectToOrgAssessment}`} />
            ) : redirectToSiteOverview ? (
                <Redirect to={`/${redirectToSiteOverview}`} />
            ) : (
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 sub-header">
                        <PageHeader
                            pageHeaderLabel={Constants.OrganizationOverview}
                            buttonLabel={Constants.NewOrganizationAssessment}
                            setRedirectToOrgAssessment={
                                setRedirectToOrgAssessment
                            }
                        />
                    </div>

                    <div className="ms-Grid-col ms-sm12">
                        <div className="table-grid">
                            <Row className="mt-4">
                                <Col sm={6}>
                                    <div className="custom-search-box">
                                        <SearchTextBox
                                            onSearch={searchItems}
                                            placeholderText={
                                                Constants.OrgSearchBoxPlaceHolder
                                            }
                                            searchBoxStyles={Custom.SearchBoxStyles}
                                        />
                                    </div>
                                </Col>
                                <Col sm={6}>
                                    <div className={`refreshIcon`}>
                                        <TooltipHost
                                            content="Click refresh to see latest % completion"
                                            delay={2}
                                            tooltipProps={Custom.tooltipProps}
                                            styles={Custom.tooltipStyles}
                                            directionalHint={DirectionalHint.leftCenter}
                                        >
                                            <Icon aria-label="refresh" iconName="refresh" className={Custom.refreshIconStyles.icon} onClick={refreshAllListData} />
                                        </TooltipHost>
                                    </div>
                                </Col>
                            </Row>

                            <div className="ms-Grid-col ms-sm12 grid-panel">
                                {!isLoaded && <Spinner
                                    label={Constants.RefreshSpinnerMessage}
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
                                        setTabCountByStatus={
                                            setTabCountByStatus
                                        }
                                        handleLinkClick={handleLinkClick}
                                        pivotPanelStyles={Custom.PivotPanelStyles}
                                        context={context}
                                        setOrgName={setOrgName}
                                        setRedirectToSiteOverview={
                                            setRedirectToSiteOverview
                                        }
                                    />
                                )}
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </React.Fragment>
    );
};

export default OrganizationOverview;
