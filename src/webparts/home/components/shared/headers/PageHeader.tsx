import { DirectionalHint, TooltipHost } from '@fluentui/react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import React, { SetStateAction } from "react";
import * as Constants from "../../../common/Constants";
import * as Custom from '../../../common/CustomStyles';
import { tooltipProps, tooltipStyles } from "../../../common/CustomStyles";
import styles from "./Header.module.scss";

interface PageHeaderProps {
    pageHeaderLabel: string;
    buttonLabel?: string;
    setRedirectToOrgAssessment?: React.Dispatch<SetStateAction<string>>;
    setOpenModal?: React.Dispatch<SetStateAction<boolean>>;
    componentName?: string;
    setRedirectBack?: React.Dispatch<SetStateAction<boolean>>;
}

const PageHeader = ({
    pageHeaderLabel,
    buttonLabel,
    setRedirectToOrgAssessment,
    setOpenModal,
    componentName,
    setRedirectBack
}: PageHeaderProps) => {
    const pageIcon: string = require("../../../assets/images/company.png");
    const insertIcon: string = require("../../../assets/images/insert.png");

    const handleClick = (buttonText: string) => {
        if (buttonText === Constants.NewOrganizationAssessment) {
            setRedirectToOrgAssessment("newOrgAssessment");
        }
        if (buttonText === Constants.NewSiteAssessment) {
            setOpenModal(true);
        }
    };

    const handleRedirect = (component: string) => {
        if (component !== '') {
            setRedirectBack(true);
        }
    };

    return (
        <div className={styles.pageHeader}>
            <div className={styles.black}>
                <div
                    className={
                        buttonLabel
                            ? "ms-Grid-col ms-sm6"
                            : "ms-Grid-col ms-sm12"
                    }
                >
                    <span className={styles.leftBannner}>

                        {componentName &&
                            <div className={`${styles.backBtn}`}>
                                <TooltipHost
                                    content="Back"
                                    delay={2}
                                    tooltipProps={tooltipProps}
                                    styles={tooltipStyles}
                                    directionalHint={DirectionalHint.bottomCenter}
                                >
                                    <Icon aria-label="SkypeArrow" iconName="SkypeArrow" className={`${Custom.backIconClasses.icon}`}
                                        onClick={() => { handleRedirect(componentName); }}
                                    />
                                </TooltipHost>
                            </div>
                        }

                        <span className={styles.bannerImage}>
                            <img
                                className={styles.bannerIcon}
                                src={pageIcon}
                                alt="pageIcon"
                            />
                        </span>
                        <span className={styles.bannerLabel}>
                            {pageHeaderLabel}
                        </span>
                    </span>
                </div>

                {buttonLabel ? (
                    <div className="ms-Grid-col ms-sm6">
                        <span
                            className={styles.pageHeaderButton}
                            onClick={() => {
                                handleClick(buttonLabel);
                            }}
                        >

                            <span className={styles.pageHeaderButtonLabel}>
                                <TooltipHost
                                    content={buttonLabel}
                                    delay={2}
                                    tooltipProps={tooltipProps}
                                    styles={tooltipStyles}
                                    directionalHint={DirectionalHint.bottomCenter}
                                >
                                    {buttonLabel}
                                </TooltipHost>
                            </span>
                            <TooltipHost
                                content={buttonLabel}
                                delay={2}
                                tooltipProps={tooltipProps}
                                styles={tooltipStyles}
                                directionalHint={DirectionalHint.bottomLeftEdge}
                            >
                                <img
                                    className={styles.pageHeaderButtonIcon}
                                    src={insertIcon}
                                    alt="insertIcon"
                                />
                            </TooltipHost>
                        </span>

                    </div>
                ) : null}
            </div>
        </div>
    );
};

export default PageHeader;
