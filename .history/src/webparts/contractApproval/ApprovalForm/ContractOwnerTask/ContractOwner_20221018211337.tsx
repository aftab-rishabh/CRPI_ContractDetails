import * as React from 'react';

import styles from './ContractOwnerTaskFormWebPart.module.scss';

import { PrimaryButton } from 'office-ui-fabric-react';
import {  Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import * as moment from 'moment';

// require('./css/jquery-ui.css');
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


export interface IContractOwnerProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface IContractOwnerState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class ContractOwner extends React.Component<IContractOwnerProps, IContractOwnerState> {

  ContractStatus = "";
  OpportunityID = "";
  webURL:any="";
  constructor(props: IContractOwnerProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {
    this.webURL = Web(this.props.webURL);
    this.setForm();


  }
  public render(): React.ReactElement<IContractOwnerProps> {
    return (
      <div className={styles.contractOwnerForm}>
        <div className={styles.container}>
          <div className={styles.row} style={{ paddingTop: '0px' }}>


            <div id="signedContractDecission" style={{ display: 'none' }}>
              <div className="panel-body" style={{ paddingTop: '0px' }}>

                <div className="row">
                  <div className={styles.sectionblockContractOwner}>
                    <Label className={styles.headersContractOwner}><u>Contract Owner Form</u></Label>
                  </div>

                  <div className="col-lg-12">
                    <div className={styles.form_group}>
                      <Label id="Question" className={styles.lablecontrol}>Signed Contract
                        Received?</Label>
                      <div className="radio">
                        <Label className={styles.lablecontrol}>
                          <input type="radio" id="answer1" name="ContractReceived" value="Yes" checked={true}
                            className="answer" onClick={() => this.radioClicked()}></input>
                          Yes (Proceed with PI Note and Documents)
                        </Label>
                      </div>

                      <div className="radio">
                        <Label className={styles.lablecontrol}>
                          <input type="radio" id="answer2" name="ContractReceived" value="No"
                            onClick={() => this.radioClicked()} className="answer"></input>
                          No (Send to management starting without contract)
                        </Label>
                      </div>
                    </div>
                  </div>
                </div>

                <div id="withContractBlock">
                  <div className="row">
                    <div className="col-lg-12">
                      <div className={styles.form_group}>
                        <Label className={styles.lablecontrol}>Details of Amendment:<span
                          className={styles.estric}>*</span></Label>
                        <textarea name="DetailsOfAmendment" className={styles.form_control}
                          id="DetailsOfAmendment" placeholder="Details of Amendment"></textarea>
                        <Label id="DetailsOfAmendmentErr" style={{ display: 'none', fontSize: 'x-small', color: 'red' }}
                          className="error">This field is required.</Label>
                      </div>
                    </div>
                  </div>
                  <div className="row">
                    <div className="col-lg-12">
                      <div className={styles.form_group}>
                        <Label className={styles.lablecontrol}>Is the Contract without any amendments? </Label>
                        <select id="WithoutAnyAmendments" onChange={() => this.amendmentsChange()}
                          className={styles.form_control}>
                          <option>Yes</option>
                          <option>No</option>
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col-lg-12" style={{ display: 'none' }} id="AmendmentCommentdiv">
                      <div className={styles.form_group}>
                        <Label className={styles.lablecontrol}>Comment:<span className={styles.estric}>*</span>
                        </Label>
                        <textarea rows={4} cols={50} name="AmendmentComment" className={styles.form_control}
                          id="AmendmentComment" placeholder="Comment"></textarea>
                        <Label id="AmendmentCommentErr" style={{ display: 'none', fontSize: 'x-small', color: 'red' }}
                          className="error">This field is required.</Label>
                      </div>
                    </div>
                  </div>


                  <div id="loader1" className={styles.modal} >
                    <div className="">
                      <div className={styles.loader}  style={{ margin: '200px auto' }}></div>
                    </div>
                  </div>

                  <div id="SignedSubmitModal" className={styles.modal}>
                    <div className={styles.modalcontent}>
                      <span className={styles.close} onClick={() => this.Close0()}>
                        &times;
                      </span>
                      <Label className={styles.headerZ}>Response Submitted!</Label>
                    </div>
                  </div>


                  <div className="row">
                    <div className="col col-lg-12">
                    <div className={styles.form_footer}>
                      <PrimaryButton text="Submit" className={styles.btn} style={{ borderRadius: "10px",marginRight:"10px" }} id="SignedSubmit" onClick={() => this.Submit1()}></PrimaryButton>
                      <PrimaryButton text="Cancel" className={styles.btn} style={{ borderRadius: "10px"}} id="SignedCancel" onClick={() => this.Cancel1()}></PrimaryButton>
                      </div></div>
                  </div>
                </div>
                <div id="withoutContractBlock" style={{ display: 'none' }}>
              <div className="row">
                <div className="col-lg-12">
                <div className={styles.form_group}>
                  <Label className={styles.lablecontrol}>Key pointers to start the project without
                    contract: <span className={styles.estric}>*</span></Label>
                  <textarea rows={4} cols={50} name="KeyPointersToStWithoutContract"
                    className={styles.form_control} id="KeyPointersToStWithoutContract"
                    placeholder="Key Pointers"></textarea>
                  <Label id="KeyPointersToStWithoutContractErr"
                    style={{ display: 'none', fontSize: 'x-small', color: 'red' }} className="error">This
                    field is required.</Label>
                </div>
                </div>
              </div>

              <div className="row">
                <div className="col-lg-12">
                <div className={styles.form_group}>
                  <Label className={styles.lablecontrol}>Due date until contract gets signed: <span
                    className={styles.estric}>*</span></Label>
                  <div className="input-group date" data-provide="datepicker">
                    <input type="date" className={styles.form_control} id="ExpectedContractSignDueDt"
                      name="txtDate" placeholder="DD/MM/YYYY"></input>
                  </div>
                  <Label id="ExpectedContractSignDueDtErr"
                    style={{ display: 'none', fontSize: 'x-small', color: 'red' }} className="error" >This
                    field is required.</Label>
                  <Label id="ExpectedContractSignDueDtErr2"
                    style={{ display: 'none', fontSize: 'x-small', color: 'red' }} className="error">Due
                    date cannot be less than today's date.</Label>
                </div>
              </div>
              </div>


              <div id="loader2" className={styles.modal} >
                <div className="">
                  <div className={styles.loader}  style={{ margin: '200px auto' }}></div>
                </div>
              </div>

              <div id="UnSignedSubmitModal" className={styles.modal}>
                <div className={styles.modalcontent}>
                  <span className={styles.close} onClick={() => this.Close1()}>
                    &times;
                  </span>
                  <Label className={styles.headerZ}>Key-pointers are sent to management for
                    approval!</Label>
                </div>
              </div>


              <div className="row">
                <div className="col col-lg-12">
                <div className={styles.form_footer}>
                  <PrimaryButton text="Submit" className={styles.btn} style={{ borderRadius: "10px",marginRight:"10px" }} id="UnSignedSubmit" onClick={() => this.Submit2()}></PrimaryButton>
                  <PrimaryButton text="Cancel" className={styles.btn} style={{ borderRadius: "10px" }} id="UnSignedCancel" onClick={() => this.Cancel2()}></PrimaryButton>
                </div> </div>
              </div>

            </div>
              </div>
            </div>

           

          </div>
        </div>
      </div>
    );
  }



  private Close0() {
    var modal = document.getElementById("SignedSubmitModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#signedContractDecission').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }
  private Close1() {
    var modal = document.getElementById("UnSignedSubmitModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#signedContractDecission').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }


  private Submit1() {

    var isvalid = true;

    if ($("#answer1:checked").val() !== undefined && $("#DetailsOfAmendment").val() === "") {
      $("#DetailsOfAmendmentErr").show();
      isvalid = false;
    }
    else
      $("#DetailsOfAmendmentErr").hide();

    if ($("#answer1:checked").val() !== undefined && $("#AmendmentComment").val() === "" && $("#WithoutAnyAmendments").val() == "No") {
      $("#AmendmentCommentErr").show();
      isvalid = false;
    }
    else
      $("#AmendmentCommentErr").hide();

    if (isvalid) {
      this.SignedSubmit();
    }
  }

  private SignedSubmit() {

    var ContractRec = $("input[name='ContractReceived']:checked").val();
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    this.webURL.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).update({
        WithoutAnyAmendments: $('#WithoutAnyAmendments').val(),
        DetailsOfAmendment: $('#DetailsOfAmendment').val(),
        AmendmentComment: $('#AmendmentComment').val(),
        MgtComment4Unsigned: 'NA',
        Status: "CM Documents Upload Awaited"
      }).then(i => {
        let web = Web(this.props.webURL);
        web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "CM"
        }).then(newListItem => {
          $("#loader1").show();
          setTimeout(() => { $('#loader1').hide(); }, 3000);
          setTimeout(() => { $('#SignedSubmitModal').show(); }, 3500);
        });
      });
  }

  private Cancel1() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }

  private Submit2() {

    var isvalid = true;

    if ($("#answer2:checked").val() !== undefined && $("#KeyPointersToStWithoutContract").val() === "") {
      $("#KeyPointersToStWithoutContractErr").show();
      isvalid = false;
    }
    else
      $("#KeyPointersToStWithoutContractErr").hide();

    if ($("#answer2:checked").val() !== undefined && $("#ExpectedContractSignDueDt").val() === "") {
      $("#ExpectedContractSignDueDtErr").show();
      isvalid = false;
    }
    else
      $("#ExpectedContractSignDueDtErr").hide();



    var datest = moment($("#ExpectedContractSignDueDt").val()).format('YYYY-MM-DD');
    var datetd = moment(new Date()).format('YYYY-MM-DD');

    if ($("#answer2:checked").val() !== undefined && datest < datetd) {
      $("#ExpectedContractSignDueDtErr2").show();
      isvalid = false;
    }
    else
      $("#ExpectedContractSignDueDtErr2").hide();



    if (isvalid) {
      this.UnSignedSubmit();
    }
  }

  private UnSignedSubmit() {
    // var ContractRec = $("input[name='ContractReceived']:checked").val();
    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + (today.getDate() + 1);
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var DueDateTime = date + ' | ' + time;

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    this.webURL.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).update({
        DueDateTime: DueDateTime,
        KeyPointersToStWithoutContract: $('#KeyPointersToStWithoutContract').val(),
        EcpectedContractSignDueDt: $('#ExpectedContractSignDueDt').val(),
        Status: "Management Approval Awaited to Start Without Signed Contract"
      }).then(i => {
        let web = Web(this.props.webURL);
        web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "CM"
        }).then(newListItem => {
          $("#loader2").show();
          setTimeout(() => { $('#loader2').hide(); }, 3000);
          setTimeout(() => { $('#UnSignedSubmitModal').show(); }, 3500);
        });
      });
  }

  private Cancel2() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }



  private amendmentsChange() {

    if ($("#WithoutAnyAmendments").val() == "No") {
      $("#AmendmentCommentdiv").show();
    }
    else {
      $("#AmendmentCommentdiv").hide();
      $("#AmendmentCommentErr").hide();
    }

  }


  //Signed Contract Block

  private radioClicked() {

    if ($('input:radio[name=ContractReceived]:checked').val() == "Yes") {
      $('#withoutContractBlock').hide();
      $('#withContractBlock').show();
    }
    else {
      $('#withoutContractBlock').show();
      $('#withContractBlock').hide();
    }
  }


  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');

    this.webURL.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).select('Id', 'Status', 'LegalDone', 'TechnicalDone', 'DeliveryDone', 'OpportunityID')
      .get().then((data) => {

        if (data.Status === "Approved PI Workflow Awaited") {
          document.getElementById("signedContractDecission").style.display = "block";
        }
        else if (data.Status == "Rejected To Start Without Contract") {
          document.getElementById("signedContractDecission").style.display = "block";
        }
      });
  }




}


