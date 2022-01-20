import { sp } from "@pnp/sp/presets/all";
import "bootstrap/dist/css/bootstrap.min.css";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import { HashRouter as Router, Route } from "react-router-dom";
import AssessmentOverview from "./assessment-overview/AssessmentOverview";
import "./Home.module.scss";
import { IHomeProps } from "./IHomeProps";
import NewOrganizationAssessment from "./organization/NewOrganizationAssessment";
import OrganizationOverview from "./organization/OrganizationOverview";
import AppHeader from "./shared/headers/AppHeader";
import SiteOverview from "./site-overview/SiteOverview";
import ProvisioningHelper from "../provisioning/ProvisioningHelper";
import { Label } from "@fluentui/react/lib/Label";
import Styles from "./Home.module.scss";
import { PrimaryButton } from "@microsoft/office-ui-fabric-react-bundle";
import * as Constants from "../common/Constants";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import ProvisioningAssets from "../provisioning/ProvisioningAssets.json";


interface IHomeStates {
    orgName: string;
    siteName: string;
    showError: boolean;
    showSuccess: boolean;
    setupMessage: string;
    isShowLoader: boolean;
    showProvisioningButton: boolean;
    showPVSSForm: boolean;
}

initializeIcons();

let rootSiteURL: string;
export default class Home extends React.Component<IHomeProps, IHomeStates> {
    private provisioningHelper: ProvisioningHelper;
    private spcontext: WebPartContext;

    constructor(_props: any) {
        super(_props);

        let absoluteUrl = this.props.context.pageContext.web.absoluteUrl;
        let serverRelativeUrl = this.props.context.pageContext.web.serverRelativeUrl;

        if (serverRelativeUrl == "/")
            rootSiteURL = absoluteUrl;
        else
            rootSiteURL = absoluteUrl.replace(serverRelativeUrl, "");

        rootSiteURL = rootSiteURL + "/" + ProvisioningAssets.inclusionPath + "/" + ProvisioningAssets.sitename;

        sp.setup({
            spfxContext: this.props.context,
        });

        this.provisioningHelper = new ProvisioningHelper(this.props.context);
        this.state = {
            orgName: "",
            siteName: "",
            showError: false,
            showSuccess: false,
            isShowLoader: false,
            setupMessage: "",
            showProvisioningButton: false,
            showPVSSForm: false,
        };
        this.updateOrgState = this.updateOrgState.bind(this);
        this.updateSiteState = this.updateSiteState.bind(this);
    }

    public updateOrgState = (organizationName?: string) => {
        this.setState({ ...this.state, orgName: organizationName });
    }

    public updateSiteState = (site?: string) => {
        this.setState({ ...this.state, siteName: site });
    }
    // Check if provisioning is required on App load
    public componentDidMount() {
        // Check if the assets are already provisioned or not.
        this.provisioningHelper.checkProvisioning().then((response) => {
            if (response != undefined) {
                if (!response) {
                    this.setState({ showProvisioningButton: true });
                }
                else {
                    this.setState({ showPVSSForm: true });
                }
            }

        }).catch((err) => {
            console.error("PVSS_Home_componentDidMount. \n ", err);
        });
    }

    //Calling the method in provisioning helper to create the assets
    public enableAppSetup = () => {
        this.setState({ isShowLoader: true, setupMessage: Constants.ProvisioningSetupMessage, showProvisioningButton: false });
        //Creating provisioning assets for the App
        this.provisioningHelper.createSiteAndLists().then((response) => {
            if (response != undefined) {
                if (!response) {
                    this.setState({ showError: true, showSuccess: false, showProvisioningButton: true, isShowLoader: false, setupMessage: Constants.ProvisioningErrorMessage });
                    console.log(Constants.ProvisioningLog, "Error in Provisioning. ");
                }
                else {
                    this.setState({ showError: false, showSuccess: true, showProvisioningButton: false, isShowLoader: false, setupMessage: Constants.ProvisioningSuccessMessage });
                    console.log(Constants.ProvisioningLog, "Provisioning Successful. ");
                }
            }

        }).catch((err) => {
            this.setState({ showError: true, showSuccess: false, showProvisioningButton: true, isShowLoader: false, setupMessage: Constants.ProvisioningErrorMessage });
            console.error("PVSS_Home_enableAppSetup. \n ", err);
        });
    }
    public render(): React.ReactElement<IHomeProps> {
        return (
            <div>
                <div className="container ms-Grid">
                    <AppHeader />
                </div>
                {this.state.showPVSSForm && (
                    <Router>
                        <div className="container ms-Grid">
                            <Route exact path="/">
                                <OrganizationOverview
                                    context={this.props.context}
                                    updateOrgState={this.updateOrgState}
                                />
                            </Route>
                            <Route path="/siteOverview">
                                <SiteOverview
                                    orgName={this.state.orgName}
                                    updateSiteState={this.updateSiteState}
                                    context={this.props.context}
                                />
                            </Route>
                            <Route path="/newOrgAssessment">
                                <NewOrganizationAssessment context={this.props.context} />
                            </Route>
                            <Route path="/assessmentOverview">
                                <AssessmentOverview
                                    orgName={this.state.orgName}
                                    siteName={this.state.siteName}
                                    context={this.props.context}
                                    rootSiteURL={rootSiteURL}
                                />
                            </Route>
                        </div>
                    </Router>
                )}
                <br></br>
                <div className={`container ${Styles.background}`}>
                    {this.state.showProvisioningButton && !this.state.showSuccess && (
                        <PrimaryButton
                            text={Constants.ProvisioningButton}
                            onClick={() => this.enableAppSetup()}
                            allowDisabledFocus
                            className={Styles.button}
                        />
                    )}
                    {this.state.isShowLoader && (
                        <Label className={Styles.setupMessage}>
                            {this.state.setupMessage}
                        </Label>
                    )}
                    {this.state.isShowLoader && (
                        <Spinner
                            label={Constants.SpinnerMessage}
                            size={SpinnerSize.large}
                        />
                    )}
                    {this.state.showError && (
                        <Label className={Styles.errorMessage}>
                            {this.state.setupMessage}
                        </Label>
                    )}
                    {this.state.showSuccess && (
                        <Label className={Styles.successMessage}>
                            {this.state.setupMessage}
                        </Label>
                    )}
                </div>
            </div>
        );
    }
}
