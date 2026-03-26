import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './ManagementTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, FontWeights, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import * as moment from 'moment';
import { TextField } from '@fluentui/react';
// require('./css/jquery-ui.css');
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


export interface IManagementProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface IManagementState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Management extends React.Component<IManagementProps, IManagementState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";
  OpportunityID = "";
  constructor(props: IManagementProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {
    this.setForm();


  }
  public render(): React.ReactElement<IManagementProps> {
    return (
      <div className={styles.ManagementTaskForm}>
        <div className="panel-body" style={{ padding: '0' }}>
          <div className={styles.pagetitle_wrap} style={{ marginBottom: '0' }}>
            <div id="status" className={styles.pagetitle}></div>
            <div id="oppID" className={styles.pagesubtitle}></div>
          </div>


          <div id="ManagementApprovalTaskForm" style={{ display: 'none' }}>
              <div className="row">
                <div className="col-lg-12">
                  <div id="ManagementApprovalForm" className={styles.form_group} style={{ display: 'block' }}>
                    <div className={styles.sectionblock}>
                      <Label className={styles.headers}><u>Project Approval Form</u></Label>
                    </div>
                    <div style={{ display: 'none' }} id="currnetDateTime"></div>

                    <div className="NoComments" id="NoComments">

                      <Label className={styles.header2}>!! Below Comments Need Attention !!</Label><br></br>

                      <Label className={styles.Action}>Sales Comments:</Label><br></br>
                      <div className="SalesHtml" id="SalesHtml"></div>

                      <Label className={styles.Action}>Technical Comments:</Label>
                      <div className="TechnicalHtml" id="TechnicalHtml"></div>

                      <Label className={styles.Action}>Delivery Comments:</Label><br></br>
                      <div className="DeliveryHtml" id="DeliveryHtml"></div>



                      <div className="RejectionSection" >
                        <Label className={styles.lablecontrol}><b>Contract Rejection Comment:</b></Label>
                        <p style={{ color: 'red' }} id="MGTComment"></p><br></br>
                      </div>


                      <div id="Commercialdocs" style={{ display: 'none' }}>

                        <Label className={styles.Action}>Contract Files:</Label><br></br>

                        <div id="fileList" className={styles.file_list}></div>

                      </div>


                        <div className="row">
                          <div className="col-md-8">
                            <div className="form-group">
                              <Label className={styles.lablecontrol}>Attachment:</Label>
                              <input type="file" className={styles.form_control} id="file" name="file"></input>
                              <span id="fileErr0" style={{ display: 'none' }} className={styles.errorlable} >Please Upload Mandatory
                                Files.</span>
                              <span id="fileErr1" style={{ display: 'none' }} className={styles.errorlable} >Please select file to
                                upload.</span>
                              <span id="fileErr2" style={{ display: 'none' }} className={styles.errorlable} >Opportunity ID is
                                required to upload file.</span>
                            </div>
                          </div>

                          <div className="col-md-4">
                            <Label className={styles.lablecontrol}>&nbsp;</Label>
                              <PrimaryButton id="UploadDocs" text="Upload" className={styles.btn} onClick={() => this.UploadFiles()} >
                  </PrimaryButton>
                          </div>
                        </div>
                        <div className="row">
                          <div className="col-lg-12">
                            <div className={styles.file_list}>
                              <div>
                                <ul className="FileList2" id="FileList2" style={{ display: 'none' }}></ul>
                              </div>
                            </div>
                          </div>
                        </div><br /> 
                         <Label className={styles.lablecontrol}>Comments: <span className={styles.estric}>*</span></Label>
                        <textarea rows={4} cols={50} id="MgtComments" className={styles.form_control}
                          required ></textarea>
                          {/* <textarea placeholder="LegalComments" id="LegalComments" className={styles.form_control} required></textarea>  */}
                        {/* <Label id="MgtErr" className={styles.errorlable}>Please add comments.</Label><br> </br>  */}
 
                      <div id="ConflictModal" className={styles.modal}>
                        <div className={styles.modalcontent}>
                          <span className={styles.close} id="Close2">&times;</span>
                          <Label className={styles.headerZ}>Delivery Team has submitted their feedback after you opened this page. Webpage needs to be refreshed to view latest feedback. Thank you.</Label>
                        </div>
                      </div>

                      <div id="TechConflictModal" className={styles.modal}>
                        <div className={styles.modalcontent}>
                          <span className={styles.close} id="Close3">&times;</span>
                          <Label className={styles.headerZ}>Technical Team has submitted their feedback after you opened this page. Webpage needs to be refreshed to view latest feedback. Thank you
                          </Label>
                        </div>
                      </div>

                     
                      <div id="loader" className={styles.modal} >
                        <div className="">
                          <div className={styles.loader} style={{ margin: '200px auto' }}></div>
                        </div>
                      </div>

                      <div id="Approveloader" className={styles.modal} >
                        <div className="">
                          <div className={styles.loader} style={{ margin: '200px auto' }}></div>
                        </div>
                      </div>

                      <div id="ApprovalModal" className={styles.modal}>
                        <div className={styles.modalcontent}>
                          <span className={styles.close} onClick={() => this.Close0()}>
                            &times;
                          </span>
                          <Label className={styles.headerZ}>Project Approved!</Label>
                        </div>
                      </div>


                      <div id="mgtinputs2">
                    <div id="Rejectloader" className={styles.modal} >
                      <div
                        className={styles.loader}
                        style={{
                          marginTop: "200px",
                          marginBottom: "200px",
                          marginLeft: "auto",
                          marginRight: "auto",
                        }}
                      ></div>
                    </div>
                  </div>


                  <div id="RejectionModal" className={styles.modal}>
                    <div className={styles.modalcontent}>
                      <span className={styles.close} onClick={() => this.Close1()}>
                        &times;
                      </span>
                      <Label className={styles.headerZ}>Project Rejected!</Label>
                    </div>
                  </div>


                 
                <div className="row">
                  <div className="col col-lg-12">
                  <div className={styles.form_footer}>
                <PrimaryButton text="Approve" className={styles.btn}  style={{ borderRadius: "10px",marginRight:"10px" }}  id="MgtApprove"   onClick={() => this.MgtApprove()} ></PrimaryButton>
              
                <PrimaryButton text="Reject" className={styles.btn} id="MgtReject" style={{ borderRadius: "10px"}}  onClick={() => this.MgtReject()}></PrimaryButton>
                </div></div></div>
              </div>



                  </div>


                </div>

              </div>
            </div>
          </div>
      </div>
    );
  }


  private async setForm() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var _this = this;
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).select('Id', 'MgtComments', 'Status', 'LegalDone', 'TechnicalDone', 'DeliveryDone', 'OpportunityID')
      .get().then((data) => {


        if (data.Status === "Management Action Awaited" || data.Status == "Escalated to management" || data.Status == "Delivery Action Awaited") {
          document.getElementById("ManagementApprovalTaskForm").style.display = "block";
          _this.GetNoCommentsCR();
          _this.GetNoCommentsTech();
          _this.GetNoCommentsDelivery();
          _this.OpportunityID = data.OpportunityID;
          $("#MGTComment").text(data.MgtComments);
          if (data.MgtComments != null && data.MgtComments != "" && data.MgtComments != undefined) {
            $(".RejectionSection").show();
          }

        }
        var current = moment(new Date()).format('YYYY-MM-DDTHH:mm:ss');
        var Now = "";
        Now += '<lable>' + current + '</lable>';
        document.getElementById("currnetDateTime").innerHTML = Now;
        this.getFiles1();
      });
  }

  private getFiles1() {
    let attachmentfiles: string = "";
    var hostUrl =  "/sites/demo/ProjectDocuments/" + this.OpportunityID
    let web = Web(this.props.webURL);
    web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
      for (var i = 0; i < files.length; i++) {
        var title = files[i].Title;
        if (title == "Contract - Commercial") {
          $("#Commercialdocs").show();
          attachmentfiles += `<li>${title}&nbsp-&nbsp<a href="${files[i].ServerRelativeUrl}" target="_blank">${files[i].Name}</a></li>`;
        }
      }
      attachmentfiles = `<ul>${attachmentfiles}</ul>`;
      $("#fileList").append(attachmentfiles);
      // this.renderData(attachmentfiles);
    });
  }

  private renderData(strResponse: string): void {
    const htmlElement = document.getElementById("#fileList");
    htmlElement.innerHTML = strResponse;
  }


  private async GetNoCommentsCR() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');

    var Detailblock = styles.viewdetail_block;
    var sectionblock = styles.sectionblock;
    var viewanswer_listing = styles.viewanswer_listing;
    var viewanswer = styles.viewanswer;
    var viewcomment = styles.viewcomment;
    var viewquestion = styles.viewquestion;
    var fltrStrSales = "ProjectID eq '" + itemID + "'";
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("SalesFeedback").items.filter(fltrStrSales).get().then((item) => {
      if (item.length !== 0) {
        for (var j = 0; j < item.length; j++) {
          if (item[j].Requirement == "IfNoComment" && item[j].Comment !== null) {
            var html2 = "";
            html2 += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#SalesHtml").append(html2);
          }

          if (item[j].Requirement == "IfYesComment" && item[j].Comment !== null) {
            var htmla = "";
            htmla += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#SalesHtml").append(htmla);
          }

          if (item[j].Requirement == "textbox" && item[j].Comment !== null) {
            var htmlb = "";
            htmlb += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#SalesHtml").append(htmlb);
          }

          if (item[j].Requirement == "Dropdown") {
            var htmlc = "";
            htmlc += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#SalesHtml").append(htmlc);
          }



        }
      }
      else {
        $(".SalesH").empty();
      }

    });
  }

  private async GetNoCommentsTech() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');

    var Detailblock = styles.viewdetail_block;
    var sectionblock = styles.sectionblock;
    var viewanswer_listing = styles.viewanswer_listing;
    var viewanswer = styles.viewanswer;
    var viewcomment = styles.viewcomment;
    var viewquestion = styles.viewquestion;
    var filterStr = "ProjectID eq '" + parseInt(itemID) + "' and " + "FeedbackStatus eq 'Submit'";

    let web = Web(this.props.webURL);
    await web.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then((item) => {
      if (item.length !== 0) {

        for (var j = 0; j < item.length; j++) {

          if (item[j].Requirement == "IfNoComment" && item[j].Comment !== null) {
            var html2 = "";
            html2 += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">' + item[j].Department.split("-")[1] + '&nbspComment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#TechnicalHtml").append(html2);
          }

          if (item[j].Requirement == "IfYesComment" && item[j].Comment !== null) {
            var htmld = "";
            htmld += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">' + item[j].Department.split("-")[1] + '&nbspAnswer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">' + item[j].Department.split("-")[1] + '&nbspComment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#TechnicalHtml").append(htmld);
          }
          if (item[j].Requirement == "textbox") {
            var htmle = "";
            htmle += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">' + item[j].Department.split("-")[1] + '&nbspComment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#TechnicalHtml").append(htmle);
          }
          if (item[j].Requirement == "Dropdown") {
            var htmlf = "";
            htmlf += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">' + item[j].Department.split("-")[1] + '&nbspAnswer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#TechnicalHtml").append(htmlf);
          }
        }
      }
      else {
        $(".TechH").empty();
      }
    });
  }

  private async GetNoCommentsDelivery() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');

    var Detailblock = styles.viewdetail_block;
    var sectionblock = styles.sectionblock;
    var viewanswer_listing = styles.viewanswer_listing;
    var viewanswer = styles.viewanswer;
    var viewcomment = styles.viewcomment;
    var viewquestion = styles.viewquestion;


    var fltrStrSales = "ProjectID eq '" + itemID + "'";
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(fltrStrSales).get().then((item) => {
      if (item.length !== 0) {
        for (var j = 0; j < item.length; j++) {

          if (item[j].Requirement == "IfNoComment" && item[j].Comment !== null) {
            var html2 = "";
            html2 += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#DeliveryHtml").append(html2);
          }

          if (item[j].Requirement == "IfYesComment" && item[j].Comment !== null) {
            var htmlg = "";
            htmlg += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#DeliveryHtml").append(htmlg);
          }
          if (item[j].Requirement == "textbox") {
            var htmlh = "";
            htmlh += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<div class="Techdiv' + item[j].ID + '" id="Techdiv' + item[j].ID + '" style="display:block">'
              + '<lable class="' + viewcomment + '">Comment:</lable>'
              + '<lable class="' + viewanswer + '" style="color:red" id="Techcomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
              + '</div>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#DeliveryHtml").append(htmlh);
          }
          if (item[j].Requirement == "Dropdown") {
            var htmli = "";
            htmli += '<div class="' + sectionblock + '">'
              + '<div class="' + Detailblock + ' ">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<div class="' + viewquestion + '">' + item[j].Title + '</div>'
              + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" style="color:red" id = "TechAnswer' + item[j].ID + '" ><b>' + item[j].Answer + '</b></lable></lable>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>'
              + '</div>';
            $("#DeliveryHtml").append(htmli);
          }
        }
      }
      else {
        $(".DelH").empty();
      }
    });
  }




  private Close0() {
    var modal = document.getElementById("ApprovalModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#ManagementApprovalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }
  private Close1() {
    var modal = document.getElementById("RejectionModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#ManagementApprovalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }

  private Close2() {
    var modal = document.getElementById("ConflictModal");
    modal.style.display = "none";
    setTimeout(() => { window.location.reload(); }, 1);
  }

  private Close3() {
    var modal = document.getElementById("TechConflictModal");
    modal.style.display = "none";
    setTimeout(() => { window.location.reload(); }, 1);
  }


  private ApproveContract() {

    document.getElementById("mgtinputs2").style.display = "none";
    if (document.getElementById("NoComments").style.display == "none") {
      document.getElementById("NoComments").style.display = "block";
    }
    else {
      this.hiboth();
    }
  }

  private hiboth() {
    document.getElementById("NoComments").style.display = "none";
    document.getElementById("mgtinputs2").style.display = "none";

  }

  private RejectContract() {
    document.getElementById("NoComments").style.display = "none";

    if (document.getElementById("mgtinputs2").style.display == "none") {
      document.getElementById("mgtinputs2").style.display = "block";
    }
    else {
      this.hiMgt();
    }
  }

  private hiMgt() {
    document.getElementById("mgtinputs2").style.display = "none";
  }

  //Management Approval Block

  private async MgtReject() {

    document.getElementById("mgtinputs2").style.display = "block";

    var isvalid = true;
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    if ($("#MgtComments").val() === "") {
      $("#MgtErr").show();
      isvalid = false;
    }
    else
      $("#MgtErr").hide();
    if (isvalid) {
      let web = Web(this.props.webURL);
      await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
        MgtComments: $("#MgtComments").val().toString(),
        Status: "Rejected to Sales",
        TechnicalDone: 2,
        DeliveryDone: 2
      }).then(i => {
        let web = Web(this.props.webURL);
        web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "Management/Marketing"
        }).then(j => {
          let web = Web(this.props.webURL);
          web.lists.getByTitle('CompleteTask').items.add({
            ProjectID: parseInt(itemID),
            TaskType: "Marketing"
          }).then(k => {
            let web = Web(this.props.webURL);
            web.lists.getByTitle('CompleteTask').items.add({
              ProjectID: parseInt(itemID),
              TaskType: "Delivery"
            });
          }).then(l => {
            let web = Web(this.props.webURL);
            web.lists.getByTitle('CompleteTask').items.add({
              ProjectID: parseInt(itemID),
              TaskType: "Technical"
            }).then(n => {
              let web = Web(this.props.webURL);
              web.lists.getByTitle('CompleteTask').items.add({
                ProjectID: parseInt(itemID),
                TaskType: "Legal"
              });

            }).then(m => {
              $("#Rejectloader").show();
              setTimeout(() => { $('#Rejectloader').hide(); }, 3000);
              setTimeout(() => { $('#RejectionModal').show(); }, 3500);
            });
          });
        });
      });
    }
  }

  private MgtApprove() {

    document.getElementById("NoComments").style.display = "block";
    document.getElementById("mgtinputs2").style.display = "none";
    var isvalid = true;
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var fltrStrSales = "ProjectID eq '" + itemID + "' and " + "QuestionNumber eq 'DEL1'";
    if ($("#MgtComments").val() === "") {
      $("#MgtErr").show();
      isvalid = false;
    }
    else
      $("#MgtErr").hide();
    if (isvalid) {
      this.checkIfAnyFeedbackisUpdated();
    }
  }

  private UploadFiles() {
    $("#loader").show();

    var isvalid = true;
    var OID = this.OpportunityID;

    if (this.OpportunityID === "") {
      $("#OpportunityIDErr").show();
      $("#FileErr2").show();
      isvalid = false;
    }
    else
      $("#OpportunityIDErr").hide();


    if (isvalid == true) {
      var files = (document.getElementById('file') as HTMLInputElement).files;
      if (files.length !== 0) {
        $("#fileErr1").hide();
        let web = Web(this.props.webURL);
        var hostUrl = '/sites/demo/ProjectDocuments/' + this.OpportunityID;
        // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
        web.folders.getByName('ProjectDocuments').folders.add(OID).then((data) => {
        }).then(b => {
          var fileArr = [];
          for (var g = 0; g < files.length; g++) {
            var file = files[g];
            fileArr.push(file);

            let web = Web(this.props.webURL);
            var hostUrl = '/sites/demo/ProjectDocuments/' + this.OpportunityID;
            // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
            web.getFolderByServerRelativeUrl(hostUrl).files.add(encodeURIComponent(file.name), file, true).then(f => {
              f.file.getItem().then(item => {
                item.update({
                  Title: "Management"
                }).then(r => {
                  $("#fileErr0").hide();
                  $("#fileErr").hide();
                  setTimeout(() => { $('#loader').hide(); }, 1000);
                  this.getFiles();
                  (document.getElementById('file') as HTMLInputElement).value = "";
                });
              });
            }).catch((e) => {
              e.console.error();
            });
          }
          }).then(y => {
            this.getFiles();
        });
      }
      else {
        $("#fileErr1").show();
        setTimeout(() => { $('#loader').hide(); }, 1000);
      }

    }

  }


  private getFiles() {
    let attachmentfiles: string = "";
    var OID = this.OpportunityID;
    var noFile = [];
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl + '/ProjectDocuments/' + OID;
    let web = Web(this.props.webURL);
    var hostUrl = '/sites/demo/ProjectDocuments/' + OID;
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
      if (files.length == 0) {
        $("#FileList2").empty();
      }
      else {
        for (var i = 0; i < files.length; i++) {
          var title = files[i].Title;
          if (files[i].Title == "Management") {
            $("#FileList2").show();
            var Uid = "DeleteFile" + files[i].UniqueId;
            var valu = files[i].Name;
            attachmentfiles += `<li id="FileListItem` + i + `" value=` + valu + `><button class="CloseBtn"`
              + `id="CloseBtn">&times;</button><lable style="display:none;">` + files[i].UniqueId + `</lable>`
              + `<lable id="filesnames` + i + `">`
              + `&nbsp&nbsp&nbsp&nbsp${title}&nbsp-&nbsp</lable>`
              + `<a href="${files[i].ServerRelativeUrl}" target="_blank" id="FileName` + i + `">`
              + `${files[i].Name}</a></li>`;
              $("#FileList2").append(attachmentfiles);
            // document.getElementById("#FileList2").innerHTML = attachmentfiles;
          }
        }
      }

    }).then(noF => {
      web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
        for (var i = 0; i < files.length; i++) {
          var title = files[i].Title;
          if (files[i].Title == "Management") {
            noFile.push(1);
          }
        }
      }).then(hell => {
        if (noFile.length == 0) {
          $("#FileList2").empty();
        }
      });
    }).then(help => {
      this.DeleteFileEvent();
    }).then(() => {
      setTimeout(() => { $('#loader').hide(); }, 1000);
    });
  }

  private DeleteFileEvent() {
    $("#loader").show();
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl + '/ProjectDocuments/' + this.OpportunityID;
    let web = Web(this.props.webURL);
    var hostUrl = '/sites/demo/ProjectDocuments/' + this.OpportunityID;
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
      var d = 0;
      for (var i = 0; i < files.length; i++) {

        if (files[i].Title == "Management") {

          var btnClass = document.getElementsByClassName("CloseBtn")[d] as HTMLElement;
          d++;
          btnClass.addEventListener('click', (f) => {
            var xyz = f.currentTarget as HTMLElement;
            var fName = xyz.nextElementSibling.nextElementSibling.nextElementSibling.textContent;
            web.getFolderByServerRelativeUrl(hostUrl).files.getByName(fName).delete().then((g) => {
            }).then(check => {
              this.getFiles();
            }).then(() => {
              setTimeout(() => { $('#loader').hide(); }, 1000);
            });
          });
        }
      }
    });
  }

  private async checkIfAnyFeedbackisUpdated() {
    var isvalid = true;
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var fltrStrSales = "ProjectID eq '" + itemID + "' and " + "QuestionNumber eq 'DEL1'";
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(fltrStrSales).select('Modified').get().then((items1) => {
      var OpenTime = document.getElementById("currnetDateTime").innerText;
      if (items1.length !== 0) {
        if (items1[0].Modified !== null) {
          var DateModified = moment(items1[0].Modified).format("YYYY-MM-DDTHH:mm:ss");
          var ManagementOpenTime = moment().diff(OpenTime, 'seconds');
          var DeliveryUpdateTime1 = moment().diff(DateModified, 'seconds');
          var DeliveryUpdateTime2 = DeliveryUpdateTime1 + 60;
          console.log("Opened at:" + OpenTime);
          console.log("Delivery resubmitted:" + DateModified);
          console.log("Management difference:" + ManagementOpenTime);
          console.log("del diff:" + DeliveryUpdateTime2);

          if (ManagementOpenTime < DeliveryUpdateTime2) {
            isvalid = true;
          }
          else {
            var modal = document.getElementById("ConflictModal");
            modal.style.display = "block";
            window.stop();
            isvalid = false;
          }
        }
      }
    }).then(async () => {
      var fltrStrTech = "ProjectID eq '" + itemID + "'";
      let web = Web(this.props.webURL);
      await web.lists.getByTitle("TechnicalComments").items.filter(fltrStrTech).select('Modified').get().then((items1) => {
        var OpenTime = document.getElementById("currnetDateTime").innerText;
        if (items1.length !== 0) {
          for (var t = 0; t < items1.length; t++) {
            var DateModified = moment(items1[t].Modified).format("YYYY-MM-DDTHH:mm:ss");
            var ManagementOpenTime = moment().diff(OpenTime, 'seconds');
            var TechnicalUpdateTime1 = moment().diff(DateModified, 'seconds');
            var TechnicalUpdateTime2 = TechnicalUpdateTime1 + 60;
            if (ManagementOpenTime < TechnicalUpdateTime2 && isvalid !== false) {
              isvalid = true;
            }
            else if (isvalid !== false) {
              var modal = document.getElementById("TechConflictModal");
              modal.style.display = "block";
              window.stop();
              isvalid = false;
              break;
            }
          }
        }
      }).then(() => {
        if (isvalid) {
          this.MgtApprove1();
        }
      });
    });

  }

  private async MgtApprove1() {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + (today.getDate() + 1);
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var DueDateTime = date + ' | ' + time;
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
      MgtComments: $("#MgtComments").val().toString(),
      DueDateTime: DueDateTime,
      Status: "Approved PI Workflow Awaited",
      TechnicalDone: 2,
      DeliveryDone: 2
    }).then(i => {
      let web = Web(this.props.webURL);
      web.lists.getByTitle('CompleteTask').items.add({
        ProjectID: parseInt(itemID),
        TaskType: "Management/Marketing"
      }).then(j => {
        let web = Web(this.props.webURL);
        web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "Marketing"
          // }).then(k => {
          //   pnp.sp.web.lists.getByTitle('CompleteTask').items.add({
          //     ProjectID: parseInt(itemID),
          //     TaskType: "Delivery"
          //   });
          // }).then(l => {
          //   pnp.sp.web.lists.getByTitle('CompleteTask').items.add({
          //     ProjectID: parseInt(itemID),
          //     TaskType: "Technical"
          //   });
          // }).then(n => {
          //   pnp.sp.web.lists.getByTitle('CompleteTask').items.add({
          //     ProjectID: parseInt(itemID),
          //     TaskType: "Legal"
          //   });
        }).then(m => {
          $("#Approveloader").show();
          setTimeout(() => { $('#Approveloader').hide(); }, 3000);
          setTimeout(() => { $('#ApprovalModal').show(); }, 3500);
        });
      });
    });
  }

}


