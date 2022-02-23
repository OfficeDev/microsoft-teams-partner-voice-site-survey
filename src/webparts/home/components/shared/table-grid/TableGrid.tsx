import { DirectionalHint, TooltipHost } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PivotItem } from "office-ui-fabric-react/lib/components/Pivot/PivotItem";
import React, { SetStateAction, useEffect } from "react";
import * as Constants from "../../../common/Constants";
import { tooltipProps, tooltipStyles } from "../../../common/CustomStyles";
import LinearProgressBar from "../linear-progress-bar/LinearProgressBar";
import StatusIcon from "../status-icon/StatusIcon";
import ExportData from "./ExportData";
import "./TableGrid.scss";

interface ITableGridProps {
    objectKeys: Array<string>;
    tableList: Array<any>;
    setTabCountByStatus: (item?: PivotItem) => void;
    setOrgName?: React.Dispatch<SetStateAction<string>>;
    setRedirectToSiteOverview?: React.Dispatch<SetStateAction<string>>;
    setRedirectToAssessmentOverview?: React.Dispatch<SetStateAction<string>>;
    setSite?: React.Dispatch<SetStateAction<string>>;
    context?: WebPartContext;
    orgName?: string;
    rootSiteURL?: string;
}

const TableGrid = ({
    objectKeys,
    tableList,
    setTabCountByStatus,
    setOrgName,
    setRedirectToSiteOverview,
    setRedirectToAssessmentOverview,
    setSite,
    context,
    orgName,
    rootSiteURL,
}: ITableGridProps) => {
    const infoImage: string = require("../../../assets/images/infoicon.png");

    useEffect(() => {
        setTabCountByStatus();
    }, [setTabCountByStatus]);

    const OpenNewTab = (url: string) => {
        const newWindow = window.open(url, "_blank", "noopener,noreferrer");
    };

    return (
        <React.Fragment>
            <table className="ms-Table ms-Table--selectable custom-table">
                <thead>
                    <tr>
                        {objectKeys?.map((key, index) => {
                            return (
                                <th key={index}>
                                    {key} <span className="fa fa-sort"></span>
                                </th>
                            );
                        })}
                    </tr>
                </thead>
                {
                    <tbody>
                        {tableList?.map((item) => {
                            return (
                                <tr key={item.id}>
                                    <td>
                                        <span className="row-label"
                                            onClick={() => {

                                                if (item.processName == Constants.OrganizationOverview) {
                                                    setOrgName(
                                                        item.organizationName
                                                    );
                                                    setRedirectToSiteOverview(
                                                        "siteOverview"
                                                    );
                                                }
                                                else if (item.processName == Constants.SiteOverview) {
                                                    setSite(item.siteName);
                                                    setRedirectToAssessmentOverview(
                                                        "assessmentOverview"
                                                    );
                                                }
                                                else if (item.processName == Constants.AssessmentOverview) {
                                                    OpenNewTab(rootSiteURL + "/lists/" + item.assessmentName + "/AllItems.aspx?FilterField1=LinkTitle&FilterValue1=" + item.orgName + "&FilterField2=Site&FilterValue2=" + item.siteName);
                                                }
                                            }}
                                        >
                                            {item.processName == Constants.OrganizationOverview
                                                ? item.organizationName
                                                : null}
                                            {item.processName == Constants.SiteOverview
                                                ? item.siteName
                                                : null}
                                            {item.processName == Constants.AssessmentOverview
                                                ?
                                                <span>
                                                    {item.assessmentName}{" "}
                                                    <TooltipHost
                                                        content={item.infoIconText}
                                                        delay={2}
                                                        tooltipProps={tooltipProps}
                                                        styles={tooltipStyles}
                                                        directionalHint={DirectionalHint.rightCenter}
                                                    >
                                                        <img src={infoImage} width="auto" height="16" alt="info-icon" />
                                                    </TooltipHost>
                                                </span>

                                                : null
                                            }
                                        </span>
                                    </td>
                                    <td>
                                        <StatusIcon iconName={item.status} />
                                    </td>
                                    <td>
                                        <LinearProgressBar
                                            completionProgressValue={
                                                item.completionProgress
                                            }
                                            maxValue={100}
                                            status={item.status}
                                        />
                                    </td>
                                    <td>
                                        <div className="row-label">
                                            {item.processName == Constants.AssessmentOverview
                                                ? <TooltipHost
                                                    content={Constants.ExportIconTooltip}
                                                    delay={2}
                                                    tooltipProps={tooltipProps}
                                                    styles={tooltipStyles}
                                                    directionalHint={DirectionalHint.rightCenter}
                                                >
                                                    <ExportData context={context}
                                                        processName={item.processName}
                                                        organizationName={item.orgName}
                                                        siteName={item.siteName}
                                                        assessmentName={item.assessmentName}></ExportData></TooltipHost>
                                                : null}
                                            {item.processName == Constants.SiteOverview
                                                ? <TooltipHost
                                                    content={Constants.ExportIconTooltip}
                                                    delay={2}
                                                    tooltipProps={tooltipProps}
                                                    styles={tooltipStyles}
                                                    directionalHint={DirectionalHint.rightCenter}
                                                >
                                                    <ExportData context={context}
                                                        processName={item.processName}
                                                        organizationName={orgName}
                                                        siteName={item.siteName}
                                                        assessmentName=""></ExportData></TooltipHost>
                                                : null}
                                            {item.processName == Constants.OrganizationOverview
                                                ? <TooltipHost
                                                    content={Constants.ExportIconTooltip}
                                                    delay={2}
                                                    tooltipProps={tooltipProps}
                                                    styles={tooltipStyles}
                                                    directionalHint={DirectionalHint.rightCenter}
                                                >
                                                    <ExportData context={context}
                                                        processName={item.processName}
                                                        organizationName={item.organizationName}
                                                        siteName=""
                                                        assessmentName=""></ExportData></TooltipHost>
                                                : null}
                                        </div>
                                    </td>
                                </tr>
                            );
                        })}
                    </tbody>
                }
            </table>
        </React.Fragment >
    );
};

export default TableGrid;
