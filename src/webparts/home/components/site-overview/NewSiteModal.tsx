import { DirectionalHint, Modal, Spinner, SpinnerSize, Stack, TooltipHost } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";
import {
    DefaultButton,
    IconButton,
    PrimaryButton
} from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as Constants from "../../common/Constants";
import * as Custom from "../../common/CustomStyles";
import "../organization/OrganizationOverview.scss";
import ErrorMessageBar from "../shared/message-dialogs/ErrorMessageBar";
import SuccessMessageBar from "../shared/message-dialogs/SuccessMessageBar";
import { ISiteOverview } from "./SiteOverview";
import SiteServiceProvider from "./SiteServiceProvider";

interface NewSiteModalProps {
    openModal: boolean;
    setOpenModal: React.Dispatch<React.SetStateAction<boolean>>;
    setPageInitialized: React.Dispatch<React.SetStateAction<boolean>>;
    setTableList: React.Dispatch<React.SetStateAction<Array<ISiteOverview>>>;
    context: WebPartContext;
    organizationName: string;

}

const NewSiteModal = ({
    openModal,
    setOpenModal,
    setPageInitialized,
    setTableList,
    context,
    organizationName,

}: NewSiteModalProps) => {
    const [siteName, setSiteName] = React.useState(String);
    const [errorSiteName, setErrorSiteName] = React.useState(false);
    const [errorMessageSiteName, setErrorMessageSiteName] = React.useState("");
    const [errorSiteExists, setErrorSiteExists] = React.useState(false);
    const titleId = useId("title");
    const [loader, setLoader] = React.useState(false);
    const [success, setSuccess] = React.useState(false);
    const [error, setError] = React.useState(false);
    const [reponseMessage, setReponseMessage] = React.useState(Constants.SiteDataUpdatedSuccessfully);

    const validateSiteName = () => {
        setErrorSiteExists(false);
        setSuccess(false);
        setError(false);
        if (!siteName) {
            setErrorSiteName(true);
            setErrorMessageSiteName(Constants.ErrorRequiredSiteError);
        }
        else if (Constants.NotAllowedSpecialChar.test(siteName)) {
            setErrorSiteName(true);
            setErrorMessageSiteName(Constants.ErrorMessageSpecialChar);
        }
        else {
            setErrorSiteName(false);
            setErrorMessageSiteName("");
        }
    };

    const handleSubmit = async (e: any) => {
        validateSiteName();

        e.preventDefault();
        let serviceProvider = new SiteServiceProvider(context);
        if (validateForm()) {
            let siteAssessment: any = {
                "Title": organizationName,
                "Status": Constants.NotStartedText,
                "Site": siteName,
                "CompletionProgress": 0
            };
         
            await serviceProvider
                .createSiteAssessment(siteAssessment)
                .then(async (response) => {
                    if (response.data === Constants.AlreadyExists) {
                        setErrorSiteExists(true);
                    }
                    else {
                        setLoader(true);
                        //Reload the site overview page.
                        setPageInitialized(false);
                        setTableList(Array<ISiteOverview>());
                        setErrorSiteExists(false);

                        await serviceProvider
                            .createAssessmentsForSite(siteAssessment)
                            .then((resp) => {
                                if (resp) {
                                    setError(false);
                                    //Clearing sitename
                                    setSiteName(null);
                                    setLoader(false);
                                    setReponseMessage(Constants.MessageNewAssessmentsCreationSuccess);
                                    setSuccess(true);
                                }
                                else {
                                    setError(true);
                                    setSuccess(false);
                                    setLoader(false);
                                    setReponseMessage(Constants.MessageNewAssessmentsCreationFailure);
                                }
                            })
                            .catch((exception) => {
                                setError(true);
                                setSuccess(false);
                                setLoader(false);
                                setReponseMessage(Constants.MessageNewAssessmentsCreationFailure);
                                console.log("NewSiteModal_handleSubmit \n", exception);
                            });
                    }
                })
                .catch((exception) => {
                    setSuccess(false);
                    setError(true);
                    setReponseMessage(Constants.MessageNewSiteCreationFailure);
                    console.log("NewSiteModal_handleSubmit \n", exception);
                });
        }

    };

    const validateForm = (): boolean => {
        {
            return (siteName && !errorSiteName) ? true : false;
        }
    };

    return (
        <div>
            <Modal
                titleAriaId={titleId}
                isOpen={openModal}
                isBlocking={false}
                containerClassName={Custom.contentStyles.container}
            >
                <div className={Custom.contentStyles.header}>
                    <span id={titleId}>New Site Assessment</span>
                    <IconButton
                        styles={Custom.iconButtonStyles}
                        iconProps={Custom.cancelIcon}
                        ariaLabel="Close popup modal"
                        onClick={() => setOpenModal(false)}
                    />
                </div>

                <div className={Custom.contentStyles.body}>
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
                                errorMessage={reponseMessage}
                            />
                        )}
                    </Stack>
                    <form
                        onSubmit={handleSubmit}
                        className="ms-Grid-col ms-sm12 form-container"
                    >
                        {loader && (
                            <Stack className="new-site-spinner"> <Spinner
                                label={Constants.AssessmentLoaderMessage}
                                size={SpinnerSize.large}
                            /> </Stack>
                        )}
                        <TextField
                            placeholder={Constants.EnterSiteName}
                            maxLength={50}
                            label={Constants.SiteName}
                            value={siteName}
                            onChange={(e: any) => setSiteName(e.target.value)}
                            onKeyUp={validateSiteName}
                            errorMessage={errorSiteExists ? Constants.ErrorSiteAlreadyExits :
                                errorSiteName ? errorMessageSiteName : ""}
                            className="ms-Grid-col ms-sm12"
                            styles={Custom.getNewSiteStyles}
                            disabled={loader}
                        />
                        <Stack
                            horizontal
                            {...Custom.PopupButtonProps}
                            className="ms-Grid-col ms-sm12"
                        >

                            <TooltipHost
                                content={Constants.Save}
                                delay={2}
                                tooltipProps={Custom.tooltipProps}
                                styles={Custom.tooltipStyles}
                                directionalHint={DirectionalHint.bottomCenter}
                            >
                                <PrimaryButton
                                    iconProps={{ iconName: "Save" }}
                                    text={Constants.Save}
                                    type={Constants.Submit}
                                    className="saveBtn"
                                    disabled={loader}
                                />
                            </TooltipHost>
                            <TooltipHost
                                content={Constants.Close}
                                delay={2}
                                tooltipProps={Custom.tooltipProps}
                                styles={Custom.tooltipStyles}
                                directionalHint={DirectionalHint.bottomCenter}
                            >
                                <DefaultButton
                                    text={Constants.Close}
                                    onClick={() => {
                                        setOpenModal(false);
                                    }}
                                    allowDisabledFocus={true}
                                    className="cancelBtn"
                                    styles={Custom.cancelBtnStyles}
                                />
                            </TooltipHost>
                        </Stack>
                    </form>
                </div>
            </Modal>
        </div>
    );
};

export default NewSiteModal;
