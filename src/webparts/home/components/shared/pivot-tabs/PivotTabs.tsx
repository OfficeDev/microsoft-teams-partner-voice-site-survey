import { WebPartContext } from "@microsoft/sp-webpart-base";
import { isEmpty } from "lodash";
import {
    IPivotStyleProps,
    IPivotStyles,
    Pivot,
    PivotItem
} from "office-ui-fabric-react/lib/Pivot";
import React, { SetStateAction, useEffect, useState } from "react";
import PaginationComponent from "react-reactstrap-pagination";
import StatusIcon from "../status-icon/StatusIcon";
import TableGrid from "../table-grid/TableGrid";
import "./PivotTabs.scss";

interface IPivotTabs {
    selectedKey: string;
    objectKeys: Array<string>;
    search: string;
    filteredTableListByCriteria: Array<any>;
    tabsItems: Array<any>;
    setTabCountByStatus: () => void;
    handleLinkClick: (item?: PivotItem) => void;
    pivotPanelStyles: (props: IPivotStyleProps) => Partial<IPivotStyles>;
    setOrgName?: React.Dispatch<SetStateAction<string>>;
    setRedirectToSiteOverview?: React.Dispatch<SetStateAction<string>>;
    setRedirectToAssessmentOverview?: React.Dispatch<SetStateAction<string>>;
    setSite?: React.Dispatch<SetStateAction<string>>;
    context?: WebPartContext;
    orgName?: string;
    rootSiteURL?: string;
}

const PivotTabs = ({
    selectedKey,
    objectKeys,
    search,
    filteredTableListByCriteria,
    tabsItems,
    setTabCountByStatus,
    handleLinkClick,
    pivotPanelStyles,
    setOrgName,
    setRedirectToSiteOverview,
    setRedirectToAssessmentOverview,
    setSite,
    context,
    orgName,
    rootSiteURL
}: IPivotTabs) => {
    const [pageSize, setpageSize] = useState(8);
    const [selectedPage, setSelectedPage] = useState(1);
    const [selectedTableMembers, setselectedTableMembers] = useState(
        new Array<any>()
    );
 /* Resetting the selected page to "1" while searching
  to show the refined search results from first page */
    useEffect(() => {
        if (!isEmpty(search)) {
            setSelectedPage(1);
        }
    }, [search]);

    const paginate = (selectedpage: any) => {
        try {
            let orginalevents = Object.assign({}, filteredTableListByCriteria);
            let myevents: any = [];
            let toindex = selectedpage * pageSize;
            let fromindex = toindex - pageSize;
            for (let i = fromindex; i < toindex; i++) {
                if (orginalevents[i]) myevents.push(orginalevents[i]);
            }
            setSelectedPage(selectedpage);
            setselectedTableMembers(myevents);
        } catch (error) {
            console.error("PVSS_PivotTabs_paginate", error);
        }
    };

    const tabLinkClick = (item?: PivotItem) => {
        try {
            setSelectedPage(1);
            setselectedTableMembers([]);
            handleLinkClick(item);
        } catch (error) {
            console.error("PVSS_PivotTabs_tabLinkClick", error);
        }
    };

    const getTabId = (itemKey: string) => {
        return `${itemKey}`;
    };

    const onRenderTabHeader = (tabItem: any) => {
        try {
            let headerText = tabItem.headerText.split("|");
            return (
                <div className="tab-header">
                    <span className="custom-status-icon">
                        <StatusIcon iconName={tabItem.itemKey} />
                    </span>
                    <span className="sm-text">{headerText[0]}</span> |
                    <span className="lg-text">{headerText[1]}</span>
                </div>
            );
        } catch (error) {
            console.error("PivotTabs_onRenderTabHeader", error);
        }
    };

    return (
        <React.Fragment>
            <Pivot
                aria-labelledby={getTabId(selectedKey)}
                selectedKey={selectedKey}
                onLinkClick={tabLinkClick}
                getTabId={getTabId}
                styles={pivotPanelStyles}
            >
                {tabsItems.map((tabItem, idx) => (
                    <PivotItem
                        itemKey={tabItem.tableItemkey}
                        headerText={tabItem.header}
                        onRenderItemLink={onRenderTabHeader}
                    >
                        <TableGrid
                            objectKeys={objectKeys}
                            tableList={
                                filteredTableListByCriteria.length <= pageSize
                                    ? filteredTableListByCriteria.slice(
                                        0,
                                        pageSize
                                    )
                                    : filteredTableListByCriteria.slice(
                                        selectedPage * pageSize - pageSize,
                                        selectedPage * pageSize
                                    )
                            }
                            setTabCountByStatus={() => setTabCountByStatus()}
                            setOrgName={setOrgName}
                            setRedirectToSiteOverview={
                                setRedirectToSiteOverview
                            }
                            setRedirectToAssessmentOverview={
                                setRedirectToAssessmentOverview
                            }
                            setSite={setSite}
                            context={context}
                            orgName={orgName}
                            rootSiteURL={rootSiteURL}
                        />

                        {filteredTableListByCriteria.length > pageSize && (
                            <div className="ms-Grid-col ms-sm6 pagination-grid">
                                <PaginationComponent
                                    totalItems={
                                        filteredTableListByCriteria.length
                                    }
                                    pageSize={pageSize}
                                    onSelect={paginate}
                                    defaultActivePage={selectedPage}
                                />
                            </div>
                        )}
                    </PivotItem>
                ))}
            </Pivot>
        </React.Fragment>
    );
};

export default PivotTabs;
