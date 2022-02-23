import { DirectionalHint, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { Stack } from "@fluentui/react/lib/Stack";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem } from "office-ui-fabric-react/lib/Pivot";
import React, { useCallback, useEffect, useState } from "react";
import { Redirect } from "react-router-dom";
import * as Constants from "../../common/Constants";
import * as Custom from "../../common/CustomStyles";
import "../organization/OrganizationOverview.scss";
import PageHeader from "../shared/headers/PageHeader";
import ErrorMessageBar from '../shared/message-dialogs/ErrorMessageBar';
import SuccessMessageBar from "../shared/message-dialogs/SuccessMessageBar";
import PivotTabs from "../shared/pivot-tabs/PivotTabs";
import SearchTextBox from "../shared/search-textbox/SearchTextBox";
import NewSiteModal from "./NewSiteModal";
import "./SiteOverview.scss";
import SiteServiceProvider from "./SiteServiceProvider";


export interface ISiteOverview {
    id?: number;
    processName: string;
    siteName: string;
    status: string;
    completionProgress: number;
}
export interface ISiteOverviewProps {
    orgName?: string;
    updateSiteState?: (site: string) => void;
    context: WebPartContext;
}

const SiteOverview = ({ orgName, updateSiteState, context }: ISiteOverviewProps) => {
    const [search, setsearch] = React.useState(String);
    const [filteredTableListByCriteria, setfilteredTableListByCriteria] =
        React.useState(new Array<ISiteOverview>());
    const [tableList, setTableList] = React.useState(
        new Array<ISiteOverview>()
    );
    const [selectedKey, setSelectedKey] = React.useState("all");
    const [filteredTableListBySearch, setfilteredTableListBySearch] =
        React.useState(new Array<ISiteOverview>());
    const [isLoaded, setIsLoaded] = React.useState(true);
    const [objectKeys, setObjectKeys] = React.useState(new Array<string>());
    const [allCount, setAllCount] = React.useState(String);
    const [completedCount, setCompletedCount] = React.useState(String);
    const [inProgressCount, setInprogressCount] = React.useState(String);
    const [notStartedCount, setNotStartedCount] = React.useState(String);
    const [pageintialized, setPageInitialized] = useState(false);
    const [redirectToAssessmentOverview, setRedirectToAssessmentOverview] =
        useState(String);
    const [redirectBack, setRedirectBack] = useState(Boolean);
    const [site, setSite] = useState(String);
    const [openModal, setOpenModal] = useState(Boolean);
    const [success, setSuccess] = useState(false);
    const [error, setError] = React.useState(false);
    const [reponseMessage, setReponseMessage] = useState(Constants.SiteDataUpdatedSuccessfully);
    const [tenantName, setTenantName] = React.useState(String);



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
            getSiteOverviewDetails();
            setPageInitialized(true);
        }
    }, [tableList.length, pageintialized]);

    useEffect(() => {
        updateSiteState(site);
    }, [site]);

    const refreshSiteListData = async () => {
        try {
            let serviceProvider = new SiteServiceProvider(context);
            setSelectedKey("all");
            setIsLoaded(false);
            await serviceProvider.refreshSiteListDataOnFilters(orgName)
                .then((response) => {
                    setPageInitialized(false);
                    setTableList(new Array<ISiteOverview>());
                    setIsLoaded(true);
                });

        } catch (error) {
            console.error(
                "PVSS_SiteOverview_refreshAllListData",
                error
            );
        }
    };

    const getSiteOverviewDetails = async () => {
        try {
            let siteLists = new Array<ISiteOverview>();
            let objectkeys = new Array<string>();

            let serviceProvider: SiteServiceProvider = new SiteServiceProvider(context);

            await serviceProvider
                .getSiteDetails(orgName)
                .then((response: any) => {
                    let i = 0;
                    while (i < response.length) {
                        if (response[i]) {
                            siteLists.push({
                                processName: Constants.SiteOverview,
                                siteName: response[i].Site,
                                status: response[i].Status.replace(/\s+/g, '').replace('-', '').toLowerCase(),
                                completionProgress: response[i].CompletionProgress
                            });
                        }
                        i++;
                    }
                })
                .then(() => {
                    if (siteLists.length > 0) {
                        objectkeys = ["SITE", ...Constants.GridHeaders];
                    }
                });
            setObjectKeys(objectkeys);
            setTableList(siteLists);
            setfilteredTableListByCriteria(siteLists);
            setfilteredTableListBySearch(siteLists);
        }
        catch (error) {
            console.log("PVSS_SiteOverview_getSiteOverviewDetails \n", error);
        }
    };

    const filterDataByStatus = (
        SiteData: ISiteOverview[],
        status?: string,
        searchedValue?: string
    ) => {
        if (
            (status === `${Constants.AllText.toLowerCase()}` || !status) &&
            !searchedValue
        ) {
            setfilteredTableListByCriteria(SiteData);
        } else if (
            status === `${Constants.AllText.toLowerCase()}` &&
            searchedValue
        ) {
            setfilteredTableListByCriteria(
                SiteData.filter((x) => {
                    return (
                        x.siteName &&
                        x.siteName
                            .toLowerCase()
                            .includes(searchedValue.toLowerCase())
                    );
                })
            );
        } else if (status && searchedValue) {
            setfilteredTableListByCriteria(
                SiteData.filter((x) => {
                    return (
                        x.status &&
                        x.status.toLowerCase() === status.toLowerCase() &&
                        x.siteName &&
                        x.siteName
                            .toLowerCase()
                            .includes(searchedValue.toLowerCase())
                    );
                })
            );
        } else if (status && !searchedValue) {
            setfilteredTableListByCriteria(
                SiteData.filter((x) => {
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
                    x.siteName &&
                    x.siteName.toLowerCase().includes(value.toLowerCase())
                );
            })
        );
    };

    useEffect(() => {
        getTenantName(orgName);
    }, [orgName]);

    const getTenantName = (organizationName: string) => {
        let serviceProvider: SiteServiceProvider = new SiteServiceProvider(context);

        serviceProvider.getTenantName(organizationName)
            .then((result) => { setTenantName(result); });
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
            {redirectBack ? <Redirect to="/" /> :
                redirectToAssessmentOverview ? (
                    <Redirect to={`/${redirectToAssessmentOverview}`} />) : (
                    <div className="main-container">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm12">
                                <PageHeader
                                    pageHeaderLabel={Constants.SiteOverview}
                                    buttonLabel={Constants.NewSiteAssessment}
                                    setOpenModal={setOpenModal}
                                    componentName={Constants.SiteOverview}
                                    setRedirectBack={setRedirectBack}
                                />
                            </div>
                        </div>
                        <Stack styles={Custom.RibbonStyles}>
                            {success && (
                                <SuccessMessageBar
                                    setSuccess={setSuccess}
                                    successMessage={
                                        reponseMessage
                                    }
                                />
                            )}

                            {error && (
                                <ErrorMessageBar
                                    setError={setError}
                                    errorMessage={Constants.SomethingWentWrong}
                                />
                            )}
                        </Stack>

                        <table className="ms-Table ms-Table--selectable site-table">
                            <tbody>
                                <tr>
                                    <td>
                                        <div className="site-searchbox">
                                            <SearchTextBox
                                                onSearch={searchItems}
                                                placeholderText={
                                                    Constants.SiteSearchBoxPlaceHolder
                                                }
                                                searchBoxStyles={Custom.SearchBoxStyles}
                                            />
                                        </div>
                                    </td>
                                    <td>
                                        <span className="siteoverview-key-label">Organization:</span>
                                        <span className="siteoverview-value-label">{orgName}</span>
                                    </td>
                                    <td>
                                        <span className="siteoverview-key-label">Tenant:</span>
                                        <span className="siteoverview-value-label">{tenantName}</span>
                                    </td>
                                    <td>
                                        <div className={`site-cursor`}>
                                            <TooltipHost
                                                content="Click refresh to see latest % completion"
                                                delay={2}
                                                tooltipProps={Custom.tooltipProps}
                                                styles={Custom.tooltipStyles}
                                                directionalHint={DirectionalHint.leftCenter}
                                            >
                                                <Icon aria-label="refresh" iconName="refresh" onClick={refreshSiteListData} className={Custom.refreshIconStyles.icon} />
                                            </TooltipHost>
                                        </div>

                                    </td>
                                </tr>
                            </tbody>
                        </table>

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
                                    setTabCountByStatus={setTabCountByStatus}
                                    handleLinkClick={handleLinkClick}
                                    pivotPanelStyles={Custom.PivotPanelStyles}
                                    setRedirectToAssessmentOverview={
                                        setRedirectToAssessmentOverview
                                    }
                                    setSite={setSite}
                                    context={context}
                                    orgName={orgName}
                                />
                            )}
                        </div>
                        {openModal && (
                            <NewSiteModal
                                openModal={openModal}
                                setOpenModal={setOpenModal}
                                context={context}
                                organizationName={orgName}
                                setPageInitialized={setPageInitialized}
                                setTableList={setTableList}
                            />
                        )}
                    </div>
                )}
        </React.Fragment >
    );
};

export default SiteOverview;
