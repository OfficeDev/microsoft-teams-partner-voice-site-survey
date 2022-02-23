import { DirectionalHint, Label, TooltipHost } from '@fluentui/react';
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import {
    Dropdown,
    DropdownMenuItemType,
    IDropdownOption,
    IDropdownProps
} from "@fluentui/react/lib/Dropdown";
import { Stack } from "@fluentui/react/lib/Stack";
import { ITextFieldProps, TextField } from "@fluentui/react/lib/TextField";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import React, { useEffect, useState } from "react";
import { Redirect } from 'react-router-dom';
import * as Constants from "../../common/Constants";
import * as Custom from "../../common/CustomStyles";
import PageHeader from "../shared/headers/PageHeader";
import ErrorMessageBar from "../shared/message-dialogs/ErrorMessageBar";
import SuccessMessageBar from "../shared/message-dialogs/SuccessMessageBar";
import "./OrganizationOverview.scss";
import OrganizationServiceProvider from "./OrganizationServiceProvider";

export const statusOptions: IDropdownOption[] = [
    { key: Constants.NotStartedText, text: Constants.NotStartedText }
];

interface INewOrganizationAssessmentProps {
    context: WebPartContext;
}

const NewOrganizationAssessment = ({ context }: INewOrganizationAssessmentProps) => {
    const [orgName, setOrgName] = useState(String);
    const [tenantName, setTenantName] = useState(String);
    const [tenantId, setTenantId] = useState(String);
    const [deployedRegion, setDeployedRegion] = useState(String);
    const [status, setStatus] = useState(Constants.NotStartedText);
    const [success, setSuccess] = useState(false);
    const [error, setError] = React.useState(false);
    const [errorOrgExists, setErrorOrgExists] = React.useState(false);
    const [redirectBack, setRedirectBack] = React.useState(Boolean);
    const [deployedRegionOptions, setDeployedRegionOptions] = React.useState(new Array<IDropdownOption>());
    const [showErrorForDeployedRegion, setShowErrorForDeployedRegion] = React.useState(false);
    const [isErrorOrgName, setIsErrorOrgName] = React.useState(false);
    const [errorMessageOrgName, setErrorMessageOrgName] = React.useState("");
    const [isErrorTenantName, setIsErrorTenantName] = React.useState(false);
    const [errorMessageTenantName, setErrorMessageTenantName] = React.useState("");
    const [isErrorTenantId, setIsErrorTenantId] = React.useState(false);
    const [errorMessageTenantId, setErrorMessageTenantId] = React.useState("");

    useEffect(() => {
        if (deployedRegionOptions.length == 0)
            getDeployedRegionOptions();
    }, [deployedRegionOptions]);

    const validateOrgName = () => {
        setErrorOrgExists(false);
        if (!orgName) {
            setIsErrorOrgName(true);
            setErrorMessageOrgName(Constants.ErrorRequiredOrganizationError);
        }
        else if (Constants.NotAllowedSpecialChar.test(orgName)) {
            setIsErrorOrgName(true);
            setErrorMessageOrgName(Constants.ErrorMessageSpecialChar);
        }
        else {
            setIsErrorOrgName(false);
            setErrorMessageOrgName("");
        }
    };

    const validateTenentName = () => {
        if (!tenantName) {
            setIsErrorTenantName(true);
            setErrorMessageTenantName(Constants.ErrorRequiredTenantNameError);
        }
        else if (Constants.NotAllowedSpecialChar.test(tenantName)) {
            setIsErrorTenantName(true);
            setErrorMessageTenantName(Constants.ErrorMessageSpecialChar);
        }
        else {
            setIsErrorTenantName(false);
            setErrorMessageTenantName("");
        }
    };

    const validateTenentId = () => {
        if (!tenantId) {
            setIsErrorTenantId(true);
            setErrorMessageTenantId(Constants.ErrorRequiredTenantIdError);
        }
        else if (Constants.NotAllowedSpecialChar.test(tenantId)) {
            setIsErrorTenantId(true);
            setErrorMessageTenantId(Constants.ErrorMessageSpecialChar);
        }
        else {
            setIsErrorTenantId(false);
            setErrorMessageTenantId("");
        }
    };

    const handleSubmit = async (e: any) => {
        try {
            validateOrgName();
            validateTenentName();
            validateTenentId();
            setShowErrorForDeployedRegion(true);

            e.preventDefault();
            let serviceProvider = new OrganizationServiceProvider(context);

            if (validateForm()) {
                //TODO :: Create a class to build the object to maintain consistency.
                let organization: any = {
                    "Title": orgName,
                    "Status": status,
                    "Tenant_x0020_Name": tenantName,
                    "Deployed_x0020_Region": deployedRegion,
                    "Tenant_x0020_ID": tenantId,
                    "CompletionProgress": 0
                };

                await serviceProvider
                    .createOrganizationAssessment(organization)
                    .then((response) => {
                        if (response.data === Constants.AlreadyExists) {
                            setErrorOrgExists(true);
                        }
                        else {
                            setSuccess(true);
                            setError(false);
                            setOrgName(null);
                            setTenantName(null);
                            setTenantId(null);
                            setDeployedRegion(null);
                            setStatus(Constants.NotStartedText);
                            setShowErrorForDeployedRegion(false);
                            setErrorOrgExists(false);
                        }
                    });

                await serviceProvider
                    .createFolderforReport(organization)
                    .then((response) => {
                        if (response.data === Constants.AlreadyExists) {
                            console.log("Folder already exists");
                        }
                    });
            }
        } catch (exception) {
            setError(true);
            setSuccess(false);
            console.log("PVSS_NewOrganizationAssessment_handleSubmit", exception);
        }
    };

    const validateForm = (): boolean => {
        {
            return (
                !isErrorOrgName && !isErrorOrgName && !isErrorTenantId && deployedRegion
            ) ? true : false;
        }
    };

    const handleCancel = () => {
        try {
            setRedirectBack(true);
        } catch (exception) {
            console.log("PVSS_NewOrganizationAssessment_handleCancel", exception);
        }
    };

    const getDeployedRegionOptions = async () => {

        let deployedRegions = new Array<IDropdownOption>();
        let serviceProvider = new OrganizationServiceProvider(context);

        await serviceProvider
            .getDeployedRegions()
            .then((response: any) => {
                let i = 0;
                deployedRegions.push({
                    key: Constants.SelectDeployedRegion,
                    text: Constants.SelectDeployedRegion,
                    itemType: DropdownMenuItemType.Header,
                });
                while (i < response.length) {
                    if (response[i]) {
                        deployedRegions.push({
                            key: response[i].Title,
                            text: response[i].Title
                        });
                    }
                    i++;
                }
                setDeployedRegionOptions(deployedRegions);
            });
    };

    const onRenderLabel = (textProps: ITextFieldProps | IDropdownProps) => {
        return (
            <>
                <Stack horizontal verticalAlign="center">
                    <Label id={textProps.id}>{textProps.label}<span className="asterisk">*</span> &nbsp;</Label>
                    {textProps.label !== Constants.DeployedRegion &&
                        <div className="newOrgInfoTooltip">
                            <TooltipHost
                                content={Constants.AllowedCharacters}
                                delay={2}
                                styles={Custom.tooltipStyles}
                            >
                                <img src={require("../../assets/images/infoicon.png")} alt="info-icon" className="newOrgInfoIcon" />
                            </TooltipHost>
                        </div>
                    }
                </Stack>
            </>
        );
    };

    return (
        <React.Fragment>
            {redirectBack ? <Redirect to="/" /> :
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 sub-header">
                        <PageHeader
                            pageHeaderLabel={Constants.NewOrganizationAssessment}
                            componentName={Constants.NewOrganizationAssessment}
                            setRedirectBack={setRedirectBack}
                        />
                    </div>

                    <div className="ms-Grid-col ms-sm6 form-container">
                        <form onSubmit={handleSubmit}>
                            <Stack
                                horizontal
                                tokens={Custom.StackTokens}
                                styles={Custom.StackStyles}
                            >
                                <Stack {...Custom.FormProps}>
                                    {success && (
                                        <SuccessMessageBar
                                            setSuccess={setSuccess}
                                            successMessage={
                                                Constants.OrganizationCreatedSuccessfully
                                            }
                                        />
                                    )}
                                    {error && (
                                        <ErrorMessageBar
                                            setError={setError}
                                            errorMessage={Constants.SomethingWentWrong}
                                        />
                                    )}

                                    <TextField
                                        label={Constants.OrganizationName}
                                        placeholder={Constants.EnterOrganizationName}
                                        maxLength={50}
                                        value={orgName}
                                        onChange={(e: any) =>
                                            setOrgName(e.target.value)
                                        }
                                        onRenderLabel={onRenderLabel}
                                        onKeyUp={validateOrgName}
                                        errorMessage={errorOrgExists ? Constants.ErrorOrganizationAlreadyExits
                                            : isErrorOrgName ? errorMessageOrgName : ""}
                                        className="ms-Grid-col ms-sm12"
                                    />

                                    <TextField
                                        label={Constants.TenantName}
                                        placeholder={Constants.EnterTenantName}
                                        maxLength={50}
                                        value={tenantName}
                                        onRenderLabel={onRenderLabel}
                                        onChange={(e: any) =>
                                            setTenantName(e.target.value)
                                        }
                                        onKeyUp={validateTenentName}
                                        errorMessage={isErrorTenantName ? errorMessageTenantName : ""}
                                        className="ms-Grid-col ms-sm12"
                                    />

                                    <TextField
                                        label={Constants.TenantId}
                                        placeholder={Constants.EnterTenantId}
                                        maxLength={50}
                                        value={tenantId}
                                        onRenderLabel={onRenderLabel}
                                        onChange={(e: any) =>
                                            setTenantId(e.target.value)
                                        }
                                        onKeyUp={validateTenentId}
                                        errorMessage={isErrorTenantId ? errorMessageTenantId : ""}
                                        className="ms-Grid-col ms-sm12"
                                    />

                                    <Dropdown
                                        placeholder={Constants.SelectDeployedRegion}
                                        label={Constants.DeployedRegion}
                                        options={deployedRegionOptions}
                                        selectedKey={deployedRegion}
                                        onChange={(e: object, selectedOption: any) => {
                                            setDeployedRegion(selectedOption.key);
                                            setShowErrorForDeployedRegion(true);
                                        }}
                                        onRenderLabel={onRenderLabel}
                                        styles={Custom.DropdownStyles}
                                        errorMessage={showErrorForDeployedRegion && !deployedRegion ? Constants.ErrorRequiredDeployedRegionError : undefined}
                                        className="ms-Grid-col ms-sm12"
                                    />

                                    <Stack
                                        horizontal
                                        className="ms-Grid-col ms-sm12"
                                        {...Custom.ButtonProps}
                                    >
                                        <TooltipHost
                                            content={Constants.Save}
                                            delay={2}
                                            tooltipProps={Custom.tooltipProps}
                                            styles={Custom.tooltipStyles}
                                            directionalHint={DirectionalHint.topCenter}
                                        >
                                            <PrimaryButton
                                                iconProps={{ iconName: "Save" }}
                                                text={Constants.Save}
                                                type={Constants.Submit}
                                                allowDisabledFocus
                                                className="ms-Grid-col ms-sm12 saveBtn"
                                            />
                                        </TooltipHost>
                                        <TooltipHost
                                            content={Constants.Back}
                                            delay={2}
                                            tooltipProps={Custom.tooltipProps}
                                            styles={Custom.tooltipStyles}
                                            directionalHint={DirectionalHint.topCenter}
                                        >
                                            <DefaultButton
                                                text={Constants.Back}
                                                iconProps={{ iconName: "NavigateBack" }}
                                                onClick={handleCancel}
                                                allowDisabledFocus
                                                className="ms-Grid-col ms-sm12 cancelBtn"
                                                styles={Custom.cancelBtnStyles}
                                            />
                                        </TooltipHost>
                                    </Stack>
                                </Stack>
                            </Stack>
                        </form>
                    </div>
                </div>
            }
        </React.Fragment>
    );
};
export default NewOrganizationAssessment;
