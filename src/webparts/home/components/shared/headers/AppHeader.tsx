import { TooltipHost } from "@fluentui/react";
import "bootstrap/dist/css/bootstrap.min.css";
import { Icon } from "office-ui-fabric-react";
import React from "react";
import Nav from "react-bootstrap/Nav";
import Navbar from "react-bootstrap/Navbar";
import { SiteAppName } from "../../../common/Constants";
import * as Custom from "../../../common/CustomStyles";
import styles from "./Header.module.scss";
import { Callout, Link, Text } from '@fluentui/react';
import { useId } from '@fluentui/react-hooks';
import * as Constants from "../../../common/Constants";


const AppHeader = () => {
    const [isCalloutVisible, setIsCalloutVisible] = React.useState(Boolean);
    const mslogo: string = require("../../../assets/images/mslogo.png");
    const packageSolution: any = require("../../../../../../config/package-solution.json");
    const buttonId = useId('callout-button');
    const labelId = useId('callout-label');
    const descriptionId = useId('callout-description');
    return (
        <Navbar className={styles.navbg}>
            <Navbar.Brand
                href="#/"
                className={styles.bgDecorator}
            >
                <span className={styles.headerPosition}>
                    <img
                        src={mslogo}
                        width="auto"
                        height="40"
                        className="d-inline-block"
                        alt="mslogo"
                        title="Home"
                    />
                    <span className={styles.appHeaderLabel} title="Home">{SiteAppName}</span>
                </span>
            </Navbar.Brand>
            <Nav.Item className={styles.navItem}>
                <Nav.Item className={styles.padding}>
                    <TooltipHost
                        content="More Info"
                        delay={2}
                        styles={Custom.tooltipStyles}
                    >
                        <Icon iconName="Info" id={buttonId} className={Custom.appHeaderIconClasses.Icon} onClick={() => setIsCalloutVisible(!isCalloutVisible)} />
                    </TooltipHost>
                    {isCalloutVisible && (
                        <Callout
                            className={Custom.appHeaderCalloutstyles.callout}
                            ariaLabelledBy={labelId}
                            ariaDescribedBy={descriptionId}
                            gapSpace={20}
                            target={`#${buttonId}`}
                            onDismiss={() => setIsCalloutVisible(!isCalloutVisible)}
                            setInitialFocus
                            directionalHint={3}
                        >
                            <Text block variant="xLarge" className={Custom.appHeaderCalloutstyles.title}>
                                About the Partner Voice Site Survey (PVSS):
                            </Text>
                            <Text block variant="small" className={Custom.appHeaderCalloutstyles.titlebody}>
                                Partner Voice Site Survey is a tool that will enable a partner/customer to complete a site survey across all the sites involved without having to have a SME at each site.
                            </Text>
                            <Text block variant="xLarge" className={Custom.appHeaderCalloutstyles.title}>
                                Additional Resources:
                            </Text>
                            <Text block variant="small" className={Custom.appHeaderCalloutstyles.titlebody}>
                                The Microsoft Teams Customer Advocacy Group is focused on delivering solutions like these to inspire and help you achieve your goals. Follow and join in through these other resources to learn more from us and the community:
                            </Text>
                            <Link href={Constants.M365ChampionCommunity} target="_blank" className={`${Custom.appHeaderCalloutstyles.link} ${Custom.appHeaderCalloutstyles.linkFont}`}>
                                Microsoft 365 Champion Community
                            </Link>
                            <Link href={Constants.DrivingAdoptionM365} target="_blank" className={`${Custom.appHeaderCalloutstyles.link} ${Custom.appHeaderCalloutstyles.linkFont}`}>
                                Driving Adoption on the Microsoft Technical Community
                            </Link>
                            <Text block variant="xLarge" className={Custom.appHeaderCalloutstyles.title}>
                                ----
                            </Text>
                            <Text block variant="small">
                                Current Version: {packageSolution.solution.version}
                            </Text>
                            <Text block variant="small">
                                Latest Version: N/A
                            </Text>
                            <Text block variant="xLarge" className={Custom.appHeaderCalloutstyles.title}>
                                ----
                            </Text>
                            <Text block variant="small">
                                Visit the Partner Voice Site Survey page to learn more:
                            </Text>
                            <Text block variant="small">
                                Overview & Information on our <Link href={Constants.M365AdoptionHub} target="_blank">Microsoft Adoption Hub</Link>
                            </Text>
                            <Text block variant="small">
                                Solution technical documentation and architectural overview on <Link href={Constants.M365GitHub} target="_blank">GitHub</Link>
                            </Text>
                        </Callout>
                    )}
                </Nav.Item>
                <Nav.Item className={styles.padding}>
                    <a href={Constants.SupportUrl} target="_blank">
                        <TooltipHost
                            content="Support"
                            delay={2}
                            styles={Custom.tooltipStyles}
                        >
                            <Icon iconName="Unknown" className={Custom.appHeaderIconClasses.Icon} />
                        </TooltipHost>
                    </a>
                </Nav.Item>
                <Nav.Item className={styles.fbIcon}>
                    <a href={Constants.FeedBackUrl} target="_blank">
                        <TooltipHost
                            content="Feedback"
                            delay={2}
                            styles={Custom.tooltipStyles}
                        >
                            <Icon iconName="Feedback" className={Custom.appHeaderIconClasses.Icon} />
                        </TooltipHost>
                    </a>
                </Nav.Item>
            </Nav.Item>
        </Navbar >
    );
};

export default AppHeader;
