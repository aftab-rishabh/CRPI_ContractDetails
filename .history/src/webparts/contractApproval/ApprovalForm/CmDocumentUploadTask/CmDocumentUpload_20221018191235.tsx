import * as React from "react";
import styles1 from "../../components/ContractApproval.module.scss";
import styles from "./CmDocumentUploadTaskFormWebPart.module.scss";
import { css } from "@uifabric/utilities/lib/css";
import {
  DefaultButton,
  FontWeights,
  IIconProps,
  PrimaryButton,
} from "office-ui-fabric-react";
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from "office-ui-fabric-react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as $ from "jquery";
import * as moment from "moment";
import { Dropdown, TextField } from "@fluentui/react";
// require('./css/jquery-ui.css');
require(".././css/jquery-ui.css");
let cssURL =
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript(
  "https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js"
);
/**
 * Icon styles. Feel free to change them
 */

export interface ICmDocumentUploadProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ICmDocumentUploadState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class CmDocumentUpload extends React.Component<
  ICmDocumentUploadProps,
  ICmDocumentUploadState
> {
  [x: string]: any;
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";
  OpportunityID = "";
  constructor(props: ICmDocumentUploadProps) {
    super(props);

    this.state = {
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
    };
  }
  public async componentDidMount() {
    this.setForm();
    this.viewItem2();
    this.getCurrency();
  }
  public render(): React.ReactElement<ICmDocumentUploadProps> {
    return (
      <div className={styles.cmDocumentUpload}>
        <div className={styles.container}>
          <div className={styles.row} style={{ paddingTop: "0px" }}>
            <div
              className={styles.pagetitle_wrap}
              style={{ marginBottom: "0px" }}
            >
              <div id="status" className="status ${styles.pagetitle}"></div>
              <div
                id="oppID"
                style={{ display: "none" }}
                className="oppID ${styles.pagesubtitle}"
              ></div>
            </div>

            <div id="cm-documents-upload-awaited-buttons">
              <div id="ProjectDetails" style={{ display: "inline-block" }}>
                <div className={styles.sectionblockCMPI}>
                  <Label
                    className={styles.headersCMPI}
                    style={{ marginTop: "0px" }}
                  >
                    <u>PI Note Details Form</u>
                  </Label>
                </div>
                <h3 className={styles.subheading3}>
                  1. Project Details (For Projects Admin Group)
                </h3>
                <div className="row ${styles.sectionEnd}">
                  <div className="col-md-6 ">
                    <div className={styles.viewdetail_block}>
                      <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                          <ul>
                            <li>
                              <div className={styles.viewlable}>
                                Project Name:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="ProjectName"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>
                                Short Name:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="ShortName"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>
                                Client Organization:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="ClientOrganisation"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>
                                Estimated Hours:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="EstimatedHours"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>Industry:</div>
                              <div
                                className={styles.viewlable_info}
                                id="Industry"
                              >
                                {" "}
                              </div>
                            </li>
                          </ul>
                        </div>
                      </div>
                    </div>{" "}
                  </div>
                  <div className="col-md-6">
                    <div className={styles.viewdetail_block}>
                      <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                          <ul>
                            <li>
                              <div className={styles.viewlable}>
                                Start Date:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="StartDate"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>End Date:</div>
                              <div
                                className={styles.viewlable_info}
                                id="EndDate"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>
                                Project Type:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="ProjectType"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>Practice:</div>
                              <div
                                className={styles.viewlable_info}
                                id="Practice"
                              >
                                {" "}
                              </div>
                            </li>
                            <li>
                              <div className={styles.viewlable}>
                                Engagement Type:
                              </div>
                              <div
                                className={styles.viewlable_info}
                                id="EngagementType"
                              >
                                {" "}
                              </div>
                            </li>
                          </ul>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>{" "}
                <br></br>
                <div className="row">
                  <div className="col-md-6">
                    <div className={styles.form_group}>
                      <Label className={styles.lablecontrol}>
                        Approved Hours:<span className={styles.estric}>*</span>
                      </Label>
                      <input
                        className={styles.form_control}
                        type="number"
                        placeholder="Approved Hours"
                        id="ApprovedHours"
                      ></input>
                      <Label
                        className={styles.errorlable}
                        id="ApprovedHoursErr"
                        style={{ display: "none" }}
                      >
                        This field is required.
                      </Label>
                    </div>
                  </div>
                  <div className="col-md-6">
                    <div className={styles.form_group}>
                      <Label className={styles.lablecontrol}>
                        Comments For Hours Difference:
                        <span
                          className={styles.estric}
                          id="Just"
                          style={{ display: "inline" }}
                        >
                          *
                        </span>
                      </Label>
                      <textarea
                        className={styles.form_control}
                        placeholder="Comments"
                        rows={2}
                        cols={50}
                        id="Justification"
                      ></textarea>
                      <Label
                        className={styles.errorlable}
                        id="JustificationErr"
                        style={{ display: "none" }}
                      >
                        This field is required.
                      </Label>
                    </div>
                  </div>
                </div>
                <div className="row ${styles.sectionEnd}">
                  <div className="col-md-6" id="CurrencyDiv">
                    <div className={styles.form_group}></div>
                  </div>
                  <div className="col-md-6">
                    <div className={styles.form_group}>
                      <Label className={styles.lablecontrol}>
                        Contract Date: <span className={styles.estric}>*</span>
                      </Label>
                      <input
                        type="date"
                        className={styles.form_control}
                        id="ContractDate"
                      ></input>
                      <Label
                        className={styles.errorlable}
                        id="ContractDateErr"
                        style={{ display: "none" }}
                      >
                        This field is required.
                      </Label>
                    </div>
                  </div>
                </div>
                <br></br>
                <div className="row">
                  <div className="col-lg-12 ">
                    <div className={styles.viewdetail_block}>
                      <div className={styles.view_listing}>
                        <ul>
                          <li>
                            <div className={styles.viewlable}>
                              Technology-Primary:
                            </div>
                            <div
                              className={styles.viewlable_info}
                              id="Technology-Primary"
                            >
                              {" "}
                            </div>
                          </li>
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
                <div className="row ${styles.sectionEnd}">
                  <div className="col-lg-12 ">
                    <div className={styles.viewdetail_block}>
                      <div className={styles.view_listing}>
                        <ul>
                          <li>
                            <div className={styles.viewlable}>
                              Technology-Secondary:
                            </div>
                            <div
                              className={styles.viewlable_info}
                              id="Technology-Secondary"
                            >
                              {" "}
                            </div>
                          </li>
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
                <br></br>
                <div className="row">
                  <div className="col-md-6">
                    <div className={styles.form_group}>
                      <Label className={styles.lablecontrol}>
                        Order No.: <span className={styles.estric}>*</span>
                      </Label>
                      <input
                        type="text"
                        placeholder="OrderNo"
                        id="OrderNo"
                        className={styles.form_control}
                        disabled
                      ></input>
                      <Label
                        className={styles.errorlable}
                        id="OrderNoErr"
                        style={{ display: "none" }}
                      >
                        This field is required.
                      </Label>
                      <span
                        id="OrderNoErr2"
                        style={{ display: "none" }}
                        className={styles.errorlable}
                      >
                        Contract for this Order Number Exists.
                      </span>
                    </div>
                  </div>
                  <div className="col-md-6">
                    <div className={styles.form_group}>
                      <Label className={styles.lablecontrol}>
                        Contract No.: <span className={styles.estric}>*</span>
                      </Label>
                      <input
                        type="text"
                        placeholder="ContractNo"
                        id="ContractNo"
                        className={styles.form_control}
                        disabled
                      ></input>
                      <Label
                        className={styles.errorlable}
                        id="ContractNoErr"
                        style={{ display: "none" }}
                      >
                        This field is required.
                      </Label>
                      <span
                        id="ContractNoErr2"
                        style={{ display: "none" }}
                        className={styles.errorlable}
                      >
                        Contract for this Contract Number Exists.
                      </span>
                    </div>
                  </div>

                  <div className="row ${styles.sectionEnd}">
                    <div className="col-md-6 ">
                      <div className={styles.sectionblock}>
                        <div className={styles.viewdetail_block}>
                          <div className={styles.view_listing}>
                            <ul>
                              <li>
                                <div className={styles.viewlable}>
                                  Account Manager:
                                </div>
                                <div
                                  className={styles.viewlable_info}
                                  id="AccountManager"
                                ></div>
                              </li>
                            </ul>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="col-md-6">
                      <div className={styles.sectionblock}>
                        <div className={styles.viewdetail_block}>
                          <div className={styles.view_listing}>
                            <ul>
                              <li>
                                <div className={styles.viewlable}>
                                  Business Manager:
                                </div>
                                <div
                                  className={styles.viewlable_info}
                                  id="BusinessManager"
                                ></div>
                              </li>
                            </ul>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div id="ClientAndBilling" style={{ display: "block" }}>
                    <h3 className={styles.subheading3}>
                      2 Client & Billing Details (For Projects Admin Group)
                    </h3>
                    <h4 className={styles.subheading4}>2.1 Client details</h4>

                    <div className="row ${styles.sectionEnd}">
                      <div className="col-md-4">
                        <div className={styles.form_group}>
                          <Label className={styles.lablecontrol}>
                            Client Contact:{" "}
                            <span className={styles.estric}>*</span>
                          </Label>
                          <input
                            className={styles.form_control}
                            type="text"
                            placeholder="Client Contact"
                            id="ClientContract"
                          ></input>
                          <Label
                            className={styles.errorlable}
                            id="ClientContractErr"
                            style={{ display: "none" }}
                          >
                            This field is required.
                          </Label>
                        </div>
                      </div>

                      <div className="col-md-4">
                        <div className={styles.form_group}>
                          <Label className={styles.lablecontrol}>
                            Client E-Mail:{" "}
                            <span className={styles.estric}>*</span>
                          </Label>
                          <input
                            className={styles.form_control}
                            type="text"
                            placeholder="Client E-Mail"
                            id="ClientEMail"
                          ></input>
                          <Label
                            className={styles.errorlable}
                            id="ClientEMailErr"
                            style={{ display: "none" }}
                          >
                            This field is required.
                          </Label>
                          <Label
                            className={styles.errorlable}
                            id="ClientEMailErr2"
                            style={{ display: "none" }}
                          >
                            Invalid email address!
                          </Label>
                        </div>
                      </div>

                      <div className="col-md-4">
                        <div className={styles.form_group}>
                          <Label className={styles.lablecontrol}>
                            Contact Number:{" "}
                            <span className={styles.estric}>*</span>
                          </Label>
                          <input
                            className={styles.form_control}
                            type="number"
                            placeholder="ContactNumber"
                            id="ContactNumbers"
                          ></input>
                          <Label
                            className={styles.errorlable}
                            id="ContactNumberErr"
                            style={{ display: "none" }}
                          >
                            This field is required.
                          </Label>
                          <Label
                            className={styles.errorlable}
                            id="ContactNumberErr2"
                            style={{ display: "none" }}
                          >
                            Minimum 10 digits or Maximum 15 digits allowed.
                          </Label>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div id="DeliveryCommitments" style={{ display: "block" }}>
                    <h3 className={styles.subheading3}>
                      3. Delivery Commitments
                    </h3>

                    <div className="row ${styles.Dept}">
                      <div className="col-md-6 ${styles.Dept}">
                        <div className={styles.form_group}>
                          <h4 className={styles.subheading4}>
                            3.1 Commitments made to Customer (By Marketing –
                            only those not mentioned in the contract. by PMO by
                            referring Contract)
                          </h4>
                        </div>
                      </div>
                      <div className="col-md-6">
                        <div className={styles.form_group}>
                          <h4 className={styles.subheading4}>
                            3.2 Risks and Issues
                          </h4>
                        </div>
                      </div>
                    </div>

                    <div className="row ${styles.Dept}">
                      <div className="col-md-6 ${styles.Dept}">
                        <div className={styles.form_group}>
                          <Label className={styles.lablecontrol}>
                            By marketing:{" "}
                            <span className={styles.estric}>*</span>
                          </Label>
                          <textarea
                            className={styles.form_control}
                            placeholder="By marketing"
                            rows={4}
                            cols={50}
                            id="Bymarketing"
                          ></textarea>
                          <Label
                            className={styles.errorlable}
                            id="BymarketingErr"
                            style={{ display: "none" }}
                          >
                            This field is required.
                          </Label>
                        </div>
                      </div>

                      <div className="col-md-6">
                        <div className={styles.form_group}>
                          <Label className={styles.lablecontrol}>
                            Risks and Issues:{" "}
                            <span className={styles.estric}>*</span>
                          </Label>
                          <textarea
                            className={styles.form_control}
                            placeholder="By marketing"
                            rows={4}
                            cols={50}
                            id="RisksandIssues"
                          ></textarea>
                          <Label
                            className={styles.errorlable}
                            id="RisksandIssuesErr"
                            style={{ display: "none" }}
                          >
                            This field is required.
                          </Label>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="row">
                    <div className="col col-md-6">
                      <div className={styles.form_group}>
                        <div className={styles.mb_1}>
                          <input
                            type="file"
                            className={styles.form_control}
                            id="uploadFile"
                          ></input>
                          <span
                            id="fileErr1"
                            style={{ display: "block" }}
                            className={styles.errorlable}
                          >
                            Please select file to upload.
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className="col col-md-6">
                      <div className={styles.form_group}>
                        <div className={styles.mb_1}>
                          <DefaultButton className="fileUpload-Button ${styles.btn} btn-primary">
                            <span
                              className=""
                              onClick={() => this.uploadFileFromControl1()}
                            >
                              Upload
                            </span>
                          </DefaultButton>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col-lg-12">
                      <div className={styles.file_list}>
                        <ul id="filesList"></ul>
                      </div>
                    </div>
                  </div>
                  <br></br>

                  <div id="loader" className={styles.modal}>
                    <div className="">
                      <div
                        className={styles.loader}
                        style={{ margin: "200px auto" }}
                      ></div>
                    </div>
                  </div>

                  <div id="SubmittedModal" className={styles.modal}>
                    <div className={styles.modalcontent}>
                      <span
                        className={styles.close}
                        onClick={() => this.Close1()}
                      >
                        &times;
                      </span>
                      <Label className={styles.header2}>
                        Feedback Submitted!
                      </Label>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col col-lg-12">
                      <div className={styles.form_footer}>
                        <PrimaryButton
                          text="Submit"
                          className={styles.btn}
                          style={{ borderRadius: "10px", marginRight: "10px" }}
                          id="Submit"
                          onClick={() => this.CMSubmit()}
                        ></PrimaryButton>
                        <PrimaryButton
                          text="Cancel"
                          className={styles.btn}
                          style={{ borderRadius: "10px" }}
                          id="Cancel"
                          onClick={() => this.CMCancel()}
                        ></PrimaryButton>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  //================== Currency Dropdown Starts ======================

  private _getCurrency(): any {
    let web = Web(this.props.webURL);
    return web.lists
      .getByTitle("Currency")
      .items.getAll()
      .then((response) => {
        return response;
      });
  }

  private getCurrency(): any {
    this._getCurrency().then((response) => {
      this._renderCurrency(response);
    });
  }

  private _renderCurrency(items: any): void {
    let html: string = "";
    let html1: string = "";
    var lablecontrol = styles.lablecontrol;
    var Select = styles.form_control;
    var errorlable = styles.errorlable;
    var Errorastrict = styles.estric;
    html += `<option value="Select Currency">Select Currency</option>`;
    items.forEach((item: any) => {
      html += `
         <option value="${item.Currency}">${item.Currency}</option>`;
    });
    html1 +=
      '<lable class="' +
      lablecontrol +
      '">Currency: <span class="' +
      Errorastrict +
      '">*</span></lable>' +
      '<select name="Currency" class="' +
      Select +
      '" id="Currency">' +
      html +
      "</select>" +
      '<lable class="' +
      errorlable +
      '" id="CurrencyErr" style="display:none; ">"This field is required."</lable>';
    $("#CurrencyDiv").append(html1);
  }
  //================== Currency Dropdown Ends ======================

  private CMSubmit() {
    var isvalid = true;
    var EstimatedHours = Number($("#EstimatedHours").text());
    var AppHours = Number($("#ApprovedHours").val());

    if (EstimatedHours - AppHours != 0) {
      if ($("#Justification").val() === "") {
        $("#JustificationErr").show();
        isvalid = false;
      } else $("#JustificationErr").hide();
    }

    if ($("#ApprovedHours").val() === "") {
      $("#ApprovedHoursErr").show();
      isvalid = false;
    } else $("#ApprovedHoursErr").hide();

    if ($("#Currency").val() === "Select Currency") {
      $("#CurrencyErr").show();
      isvalid = false;
    } else $("#CurrencyErr").hide();

    if ($(".OrderNo").val() === "") {
      $("#OrderNoErr").show();
      isvalid = false;
    } else $("#OrderNoErr").hide();

    if ($(".ContractNo").val() === "") {
      $("#ContractNoErr").show();
      isvalid = false;
    } else $("#ContractNoErr").hide();

    if ($(".ContractDate").val() === "") {
      $("#ContractDateErr").show();
      isvalid = false;
    } else $("#ContractDateErr").hide();

    if ($("#ClientContract").val() === "") {
      $("#ClientContractErr").show();
      isvalid = false;
    } else $("#ClientContractErr").hide();

    if ($("#ClientEMail").val() === "") {
      $("#ClientEMailErr2").hide();
      $("#ClientEMailErr").show();
      isvalid = false;
    } else if (this.validateEmail($("#ClientEMail").val()) == false) {
      $("#ClientEMailErr").hide();
      $("#ClientEMailErr2").show();
      isvalid = false;
    } else if (this.validateEmail($("#ClientEMail").val()) == true) {
      $("#ClientEMailErr").hide();
      $("#ClientEMailErr2").hide();
    }

    if ($("#ContactNumbers").val() === "") {
      $("#ContactNumberErr").show();
      $("#ContactNumberErr2").hide();
      isvalid = false;
    } else $("#ContactNumberErr").hide();

    if ($("#ContactNumbers").val().toLocaleString().length < 10) {
      $("#ContactNumberErr").hide();
      $("#ContactNumberErr2").show();
      isvalid = false;
    }

    if ($("#ContactNumbers").val().toLocaleString().length >= 10) {
      $("#ContactNumberErr").hide();
      $("#ContactNumberErr2").hide();
    }

    if ($("#Bymarketing").val() === "") {
      $("#BymarketingErr").show();
      isvalid = false;
    } else $("#BymarketingErr").hide();

    if ($("#RisksandIssues").val() === "") {
      $("#RisksandIssuesErr").show();
      isvalid = false;
    } else $("#RisksandIssuesErr").hide();

    if (isvalid) {
      this.CMSubmit2();
    }
  }

  private CMSubmit2(): void {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    var ProjectName = $("#ProjectName").text();
    var ShortName = $("#ShortName").text();
    var ClientOrganisation = $("#ClientOrganisation").text();
    var Industry = $("#Industry").text();
    var Practice = $("#Practice").text();
    var EngagementType = $("#EngagementType").text();
    var StartDate = $("#StartDate").text();
    var EndDate = $("#EndDate").text();
    var ProjectType = $("#ProjectType").text();
    var EstimatedHours = $("#EstimatedHours").text();
    var Justification = $("#Justification").val();
    var ApprovedHours = $("#ApprovedHours").val();
    var PrimaryTechnology = $("#Technology-Primary").text();
    var SecondaryTechnology = $("#Technology-Secondary").text();
    var AccountManager = $("#AccountManager").text();
    var BusinessManager = $("#BusinessManager").text();
    var Currency = $("#Currency").val();
    var OrderNo = $("#OrderNo").val();
    var ContractNo = $("#ContractNo").val();
    var ContractDate = $("#ContractDate").val();
    var ClientContract = $("#ClientContract").val();
    var ClientEMail = $("#ClientEMail").val();
    var ContactNumber = $("#ContactNumbers").val();
    var Bymarketing = $("#Bymarketing").val();
    var RisksandIssues = $("#RisksandIssues").val();
    let web = Web(this.props.webURL);
    web.lists
      .getByTitle("PINote")
      .items.add({
        ItemId: itemID,
        ProjectName: ProjectName,
        ShortName: ShortName,
        ClientOrganisation: ClientOrganisation,
        Industry: Industry,
        Practice: Practice,
        EngagementType: EngagementType,
        StartDate: StartDate,
        EndDate: EndDate,
        ProjectType: ProjectType,
        EstimatedHours: EstimatedHours,
        Justification: Justification,
        ApprovedHours: ApprovedHours,
        PrimaryTechnology: PrimaryTechnology,
        SecondaryTechnology: SecondaryTechnology,
        AccountManager: AccountManager,
        BusinessManager: BusinessManager,
        Currency: Currency,
        OrderNo: OrderNo,
        ContractNo: ContractNo,
        ContractDate: ContractDate,
        ClientContract: ClientContract,
        ClientEMail: ClientEMail,
        ContactNumber: ContactNumber,
        Bymarketing: Bymarketing,
        RisksandIssues: RisksandIssues,
      })
      .then((i) => {
        let web = Web(this.props.webURL);
        web.lists
          .getByTitle("CompleteTask")
          .items.add({
            ProjectID: parseInt(itemID),
            TaskType: "CM",
          })
          .then((j) => {
            let web = Web(this.props.webURL);
            web.lists
              .getByTitle("Projects")
              .items.getById(parseInt(itemID))
              .update({
                Status: `PM Assignment Awaited`,
              })
              .then((newListItem) => {
                $("#loader").show();
                setTimeout(() => {
                  $("#loader").hide();
                }, 3000);
                setTimeout(() => {
                  $("#SubmittedModal").show();
                }, 3500);
              });
          });
      });
  }

  private validateEmail(email) {
    var mailformt =
      /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1, 3}\.[0-9]{1, 3}\.[0-9]{1, 3}\.[0-9]{1, 3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    if (mailformt.test(email)) {
      return true;
    } else {
      return false;
    }
  }

  private Close1() {
    var modal = document.getElementById("SubmittedModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $("#cm-documents-upload-awaited-buttons").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2000);
  }

  private EnableComment() {
    var EstimatedHours = Number($(".EstimatedHours")[0].innerText);
    var AppHours = Number($("#ApprovedHours").val());
    if (EstimatedHours - AppHours == 0) {
      $("#Justification").attr("disabled", "");
      $("#Justification").attr("placeholder", "Not Required");
      $("#Justification").attr("style", "cursor: no-drop;");
      $("#Just").hide();
    } else {
      $("#Justification").removeAttr("disabled");
      $("#Justification").removeAttr("style");
      $("#Justification").removeAttr("placeholder");
      $("#Just").show();
      $("#Justification").attr("placeholder", "Comments");
    }
  }

  private uploadFileFromControl1() {
    $("#loader").show();

    var files = (document.getElementById("uploadFile") as HTMLInputElement)
      .files;
    var fileArr = [];
    if (files.length !== 0) {
      $("#fileErr1").hide();
      for (var g = 0; g < files.length; g++) {
        var file = files[g];
        fileArr.push(file);
        var hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
        let web = Web(this.props.webURL);
        web
          .getFolderByServerRelativeUrl(hostUrl)
          .files.add(file.name, file, true)
          .then((f) => {
            f.file.getItem().then((item) => {
              item
                .update({
                  Title: "Project Related Document from CM",
                })
                .then((h) => {
                  $("#uploadFile").val("");
                  this.getFiles1();
                });
            });
          });
      }
    } else {
      $("#fileErr1").show();
      setTimeout(() => {
        $("#loader").hide();
      }, 1000);
    }
  }

  private getFiles1(): void {
    let attachmentfiles: string = "";
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl + '/ProjectDocuments/' + $('#oppID').text();
    let web = Web(this.props.webURL);
    var hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
    web
      .getFolderByServerRelativeUrl(hostUrl)
      .files.filter("Title eq 'Project Related Document from CM'")
      .get()
      .then((files) => {
        if (files.length == 0) {
          $("#filesList").empty();
        } else {
          for (var i = 0; i < files.length; i++) {
            var title = files[i].Title;
            var valu = files[i].Name;
            attachmentfiles +=
              `<li id="FileListItem` +
              i +
              `" value=` +
              valu +
              `><button class="CloseBtn"` +
              `id="CloseBtn" style="padding: 0;">&times;</button><lable style="display:none;">` +
              files[i].UniqueId +
              `</lable>` +
              `<lable id="filesnames` +
              i +
              `">` +
              `&nbsp&nbsp&nbsp&nbsp${title}&nbsp-&nbsp</lable>` +
              `<a href="${files[i].ServerRelativeUrl}" target="_blank" id="FileName` +
              i +
              `">` +
              `${files[i].Name}</a></li>`;
            (
              document.getElementById("filesList") as HTMLInputElement
            ).innerHTML = attachmentfiles;
            // $("#filesList").append(attachmentfiles);
          }
        }
        this.DeleteFileEvent();
      })
      .then(() => {
        setTimeout(() => {
          $("#loader").hide();
        }, 1000);
      });
  }

  private DeleteFileEvent() {
    $("#loader").show();
    let web = Web(this.props.webURL);
    var hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
    web
      .getFolderByServerRelativeUrl(hostUrl)
      .files.filter("Title eq 'Project Related Document from CM'")
      .get()
      .then((files) => {
        for (var i = 0; i < files.length; i++) {
          var btnClass = $("#CloseBtn")[i] as HTMLElement;
          btnClass.addEventListener("click", (f) => {
            var xyz = f.currentTarget as HTMLElement;
            var fName =
              xyz.nextElementSibling.nextElementSibling.nextElementSibling
                .textContent;
            let web = Web(this.props.webURL);
            web
              .getFolderByServerRelativeUrl(hostUrl)
              .files.getByName(fName)
              .delete()
              .then((g) => {})
              .then((check) => {
                this.getFiles1();
              })
              .then(() => {
                setTimeout(() => {
                  $("#loader").hide();
                }, 1000);
              });
          });
        }
      });
  }

  private CMCancel() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }

  private async viewItem2(): Promise<void> {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select("*", "AccountManager/Title", "ProjectManager/Title")
      .expand("ProjectManager/Id", "AccountManager/Id")
      .get()
      .then((item) => {
        var SD = item.StartDate;
        var ED = item.EndDate;

        // primary technology array
        var PT = item.PrimaryTechnology;
        var pt: any = [];
        for (var x = 0; x < PT.length; x++) {
          pt.push(PT[x].split("-")[1]);
        }

        // secondary technology array
        var ST = item.SecondaryTechnology;
        var st: any = [];
        if (item.SecondaryTechnology != null) {
          for (var y = 0; y < ST.length; y++) {
            st.push(ST[y].split("-")[1]);
          }
        } else {
          st.push("");
        }

        SD = new Date(SD).toLocaleDateString();
        ED = new Date(ED).toLocaleDateString();
        $("#ProjectName").text(item.ProjectName);
        $("#ShortName").text(item.ShortName);
        $("#ClientOrganisation").text(item.ClientOrganization);
        $("#Industry").text(item.Industry);
        $("#Practice").text(item.Practice);
        $("#EngagementType").text(item.EngagementType);
        $("#StartDate").text(SD);
        $("#EndDate").text(ED);
        $("#ProjectType").text(item.ProjectType);
        $("#EstimatedHours").text(item.EstimatedHours);
        $("#Technology-Primary").text(pt);
        $("#Technology-Secondary").text(st);
        $("#AccountManager").text(item.AccountManager.Title);
        $("#BusinessManager").text(item.BusinessManager0);
      });
  }

  private async setForm() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    var _this = this;
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select("Id", "Status", "OpportunityID")
      .get()
      .then((data) => {
        $("#oppID").text(data.OpportunityID);

        // if (data.Status == "CM Documents Upload Awaited") {
        //   this.document.getElementById("cm-documents-upload-awaited-buttons").style.display = "block";
        // }
        this.getFiles1();
        this.importFromCrm();
      });
  }

  private async importFromCrm() {
    var filterStr = "OpportunityID eq '" + $("#oppID").text() + "'";
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("CRMProjects")
      .items.filter(filterStr)
      .get()
      .then((data) => {
        if (data.length) {
          var item = data[0];

          if (item.OrderNo != null) {
            $("#OrderNo").val(item.OrderNo);
          } else {
            $("#OrderNo").removeAttr("disabled");
            $("#OrderNo").attr("style", "color:black;");
          }

          if (item.ContractNo != null) {
            $("#ContractNo").val(item.ContractNo);
          } else {
            $("#ContractNo").attr("style", "color:black;");
            $("#ContractNo").removeAttr("disabled");
          }
        }
      });
  }
}
