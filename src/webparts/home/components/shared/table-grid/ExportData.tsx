import { WebPartContext } from "@microsoft/sp-webpart-base";
import React, { useRef, useState } from "react";
import { CSVLink } from "react-csv";
import commonServices from '../../../common/CommonServices';
import * as Constants from "../../../common/Constants";
import * as XLSX from 'xlsx';
import { Dialog, DialogType, IconButton, Label, Modal, Stack } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";
import * as Custom from "../../../common/CustomStyles";
import ErrorMessageBar from "../message-dialogs/ErrorMessageBar";
import { DefaultButton, DirectionalHint, Spinner, SpinnerSize, TextField, TooltipHost } from "@microsoft/office-ui-fabric-react-bundle";
import "./TableGrid.scss";
interface IExportDataProps {
  context: WebPartContext;
  processName: string;
  organizationName: string;
  siteName: string;
  assessmentName: string;
}

const ExportData = ({
  context,
  processName,
  organizationName,
  siteName,
  assessmentName,
}: IExportDataProps) => {

  const [exportData, setExportData] = useState([]);
  const [csvFilename, setCsvFilename] = useState("");
  const [openModal, setOpenModal] = useState(false);
  const [folderPath, setFolderPath] = useState("");
  const titleId = useId("title");
  const [error, setError] = React.useState(false);
  const [success, setSuccess] = React.useState(false);
  const [reponseMessage, setReponseMessage] = React.useState("");
  const [loader, setLoader] = React.useState(false);



  const exportIcon: string = require("../../../assets/images/export.png");
  const clickRef = useRef(null);

  const ExportToFile = async () => {

    try {
      let commonServiceManager: commonServices = new commonServices(context);
      let fileName = "";
      let folderName = "";
      let masterName = "";
      let filterName = "";
      let site = "";


      folderName = organizationName;
      setError(false);
      setSuccess(false);
      setReponseMessage("");


      if (processName == Constants.AssessmentOverview) {
        masterName = assessmentName;
        filterName = "Title eq '" + organizationName + "' and Site eq '" + siteName + "'";
        setCsvFilename(organizationName + "_" + siteName + "_" + assessmentName + ".csv");

        await exportAssessmentDetails(filterName, masterName).then((response: any) => {
          if (response) {
            setExportData(response);
            clickRef.current.link.click();
          }
        });
      }
      else if (processName == Constants.SiteOverview) {
        filterName = "Title eq '" + organizationName + "' and Site eq '" + siteName + "'";
        fileName = organizationName + "_" + siteName;
        let excelData: any;
        const workBook = XLSX.utils.book_new();

        setOpenModal(true);
        setLoader(true);

        // Creating folder if not exists
        await commonServiceManager.createFolder(organizationName)
          .then(async (response) => {
            if (response.data === Constants.AlreadyExists) {
              console.log("Folder already exists");
            }
          });

        // Getting reports folder name
        await commonServiceManager.getFolderURL(folderName)
          .then(async (folderURL: any) => {
            if (folderURL)
              setFolderPath(folderURL + "/" + fileName + ".xlsx?web=1");
          });

        // Fetching all items from Assessments Master List  
        await commonServiceManager
          .getAllListItems(Constants.AssessmentsMaster)
          .then(async (assesmentDetails) => {
            for (let i = 0; i < assesmentDetails.length; i++) {
              masterName = assesmentDetails[i]["Title"];

              await exportAssessmentDetails(filterName, masterName).then(async (response: any) => {
                if (response) {
                  //Adding data to the report
                  excelData = await createExcel(response, workBook, masterName);
                }
              });
            }
            //Creating excel file
            commonServiceManager.createFile(folderName, fileName, excelData)
              .then(async (response: any) => {
                if (response) {
                  setSuccess(true);
                  setLoader(false);
                }
              }).catch((exception) => {
                setError(true);
                setLoader(false);
                setReponseMessage(Constants.ExportErrorMessage);
                console.error("PVSS_ExportData_ExportToFile \n", exception);

              });
          });
      }
      else {
        setOpenModal(true);

        // Creating folder if not exists
        await commonServiceManager.createFolder(organizationName)
          .then(async (response) => {
            if (response.data === Constants.AlreadyExists) {
              console.log("Folder already exists");
            }
          });

        // Getting reports folder name
        await commonServiceManager.getFolderURL(folderName)
          .then(async (folderURL: any) => {
            if (folderURL)
              setFolderPath(folderURL);
          });

        // Fetching all items from Site Overview List 
        await commonServiceManager
          .getItemsWithOnlyFilter(Constants.SiteOverview, "Title eq '" + organizationName + "'")
          .then(async (siteDetails) => {
            for (let i = 0; i < siteDetails.length; i++) {
              site = siteDetails[i]["Site"];
              {
                filterName = "Title eq '" + organizationName + "' and Site eq '" + site + "'";
                fileName = organizationName + "_" + site;
                let excelData: any;
                const workBook = XLSX.utils.book_new();
                // Fetching all items from Assessments Master List  
                await commonServiceManager
                  .getAllListItems(Constants.AssessmentsMaster)
                  .then(async (assesmentDetails) => {
                    for (let j = 0; j < assesmentDetails.length; j++) {
                      masterName = assesmentDetails[j]["Title"];

                      await exportAssessmentDetails(filterName, masterName).then(async (response: any) => {
                        if (response) {
                          //Adding data to the report
                          excelData = await createExcel(response, workBook, masterName);
                        }
                      });
                    }
                    //Creating excel file
                    commonServiceManager.createFile(folderName, fileName, excelData)
                      .then(async (response: any) => {
                        if (response) {
                          setSuccess(true);
                        }
                      }).catch((exception) => {
                        setError(true);
                        setReponseMessage(Constants.ExportErrorMessage);
                        console.error("PVSS_ExportData_ExportToFile \n", exception);

                      });
                  });
              }
            }
          });
      }

    }
    catch (error) {
      console.error("PVSS_ExportData_ExportToFile \n", error);
    }
  };

  //Creating the data for excel file
  const createExcel = async (response: any, workBook: any, masterName: string): Promise<any> => {
    return new Promise(async (resolve, reject) => {
      const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
      let data = [];
      let dataHeader = [];
      let excelData: any;
      dataHeader.push(response[0]);
      for (let i = 1; i < response.length; i++) {
        data.push(response[i]);
      }

      const workSheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet([]);
      XLSX.utils.sheet_add_aoa(workSheet, dataHeader);
      XLSX.utils.sheet_add_json(workSheet, data, { origin: 'A2', skipHeader: true });
      XLSX.utils.book_append_sheet(workBook, workSheet, masterName);
      const excelBuffer = XLSX.write(workBook, { bookType: 'xlsx', type: 'array' });
      excelData = new Blob([excelBuffer], { type: fileType });
      resolve(excelData);
    });

  };
 //Getting data from SharePoint lists to generate report
  const exportAssessmentDetails = async (
    filterName: string,
    masterName: string): Promise<any> => {
    return new Promise(async (resolve, reject) => {

      let commonServiceManager: commonServices = new commonServices(context);
      let data = [];
      let dataHeader = [];
      let valueIndexOf = -1;
      const notUsedFields = ["Attachments", "AuthorId", "ComplianceAssetId", "ContentTypeId", "Created", "EditorId", "FileSystemObjectType", "GUID", "ID", "Id", "Modified", "OData__UIVersionString", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl", "odata.editLink", "odata.etag", "odata.id", "odata.type"];

      await commonServiceManager
        .getItemsWithOnlyFilterWithTop(masterName, filterName, 1000)
        .then((response: any) => {
          if (response.length > 0) {
            let count = 0;

            while (count < response.length) {

              if (response[count]) {

                if (count == 0) {
                  dataHeader.push(Object.keys(response[count]));
                  notUsedFields.forEach((notUsedField) => {
                    valueIndexOf = dataHeader[0].indexOf(notUsedField);
                    if (valueIndexOf > -1)
                      dataHeader[0].splice(valueIndexOf, 1);
                  });
                  data.push(dataHeader[0]);
                }
                let assessmentData = [];
                Object.keys(response[count]).forEach((field) => {
                  if (dataHeader[0].includes(field)) {
                    assessmentData.push(response[count][field]);
                  }
                });
                data.push(assessmentData);
                count++;
              }
            }
          }
        })
        .then(() => {
          if (data.length > 0) {
            data[0].forEach((element) => {

              if (element == "Title") {
                data[0][data[0].indexOf(element)] = Constants.Organization;
              }
              if (element.includes('_x0020_')) {
                data[0][data[0].indexOf(element)] = element.replaceAll("_x0020_", " ");
              }
            });
          }
        });
      resolve(data);
    });
  };

  return (
    <React.Fragment>
      <img className="export-icon"
        alt="export"
        src={exportIcon}
        onClick={ExportToFile} />

      <CSVLink
        data={exportData}
        filename={csvFilename}
        ref={clickRef}
        target='_blank'
      />
      <Modal
        titleAriaId={titleId}
        isOpen={openModal}
        isBlocking={false}
        containerClassName={Custom.contentStyles.container}
      >
        <div className={Custom.contentStyles.header}>
          <span id={titleId}>
            {processName == Constants.SiteOverview ? Constants.SiteReportTitle : Constants.OrgReportTitle}
          </span>
          <IconButton
            styles={Custom.iconButtonStyles}
            iconProps={Custom.cancelIcon}
            ariaLabel="Close popup modal"
            onClick={() => setOpenModal(false)}
          />
        </div>
        <Stack styles={Custom.RibbonStyles}>
          {error && (
            <ErrorMessageBar
              setError={setError}
              errorMessage={reponseMessage}
            />
          )}
        </Stack>
        <div className="exportPopupText">
          {loader && (
            <Stack className="site-spinner">
              <Spinner
                label={Constants.SiteExportSpinnerMessage}
                size={SpinnerSize.large}
              />
            </Stack>
          )}
          {processName == Constants.OrganizationOverview && (
            <p> Reports can be found <a href={folderPath} target='_blank'>here</a>.</p>)}
          {processName == Constants.OrganizationOverview && (
            <p>{Constants.ExportReportMessage}</p>)}
          {processName == Constants.SiteOverview && success && (
            <p> Reports can be found <a href={folderPath} target='_blank'>here</a>.</p>)}
          {processName == Constants.SiteOverview && !error && (
            <p>{Constants.ExportReportMessage}</p>)}

          <div className="exportPopupButton">
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
                hidden={
                  processName == Constants.SiteOverview && loader ? true : false
                }
              />
            </TooltipHost>
          </div>
        </div>
      </Modal>

    </React.Fragment>
  );
};

export default ExportData;