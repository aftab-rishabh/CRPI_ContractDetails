import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './managementConsentTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, FontSizes, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import * as moment from 'moment';
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


export interface ImanagementConsentTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ImanagementConsentTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class managementConsentTask extends React.Component<ImanagementConsentTaskProps, ImanagementConsentTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";

  constructor(props: ImanagementConsentTaskProps) {
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
  public render(): React.ReactElement<ImanagementConsentTaskProps> {
    return (
      <div className={styles.managementConsentTaskForm}>
        <div className={styles.container}>
          <div className={styles.row}>

            <div id="managementConsent" style={{ display: 'none' }}>

              <div className={styles.sectionblockConsent}>
                <Label className={styles.headersConsent}><u>Consent For Unsigned Contract</u></Label>
              </div>
              <div >
                <Label className={styles.viewlable}>Comments from CM:</Label>
                <Label className="CommentCM"></Label>
              </div><br></br>

              <div className="">
                <Label className={styles.viewlable}>Signed Contract will be received by:</Label>
                <Label className="SignedDate"></Label>
              </div><br></br>


              <Label className={styles.viewlable}>Rejection/Approval Comments:<span className={styles.estric}>*</span>
              </Label><br />
              <textarea className={styles.form_control} style={{ width: '50%' }} placeholder="Comment" id="MgtCommentsUnsigned" rows={4}></textarea><br></br>
              <label id="MgtErr" style={{ display: 'none', color: 'red' }} className="error" >Please add
                comments.</label><br /><br />
            </div>


            <div id="loader" className={styles.modal} >
              <div className="">
                <div className={styles.loader} style={{ margin: '200px auto' }}></div>
              </div>
            </div>


            <div id="ApproveModal" className={styles.modal}>
              <div className={styles.modalcontent}>
                <span className={styles.close} onClick={() => this.Close0()}>
                  &times;
                </span>
                <Label className={styles.header2}>Approved to start without signed contract from customer!</Label>
              </div>
            </div>



            <div id="RejectModal" className={styles.modal}>
              <div className={styles.modalcontent}>
                <span className={styles.close} onClick={() => this.Close1()}>
                  &times;
                </span>
                <Label className={styles.header2}>Rejected to start without signed contract from customer!</Label>
              </div>
            </div>



            <div className="row">
              <div className="col col-lg-12">
                <div className={styles.form_footer}>
                  <PrimaryButton text="Approve" className={styles.btn} style={{ borderRadius: "10px", marginRight: "10px" }} id="Approve" onClick={() => this.Approve()} ></PrimaryButton>

                  <PrimaryButton text="Reject" className={styles.btn} id="Reject" style={{ borderRadius: "10px" }} onClick={() => this.Reject()}></PrimaryButton>
                </div></div></div>




          </div>
        </div>
      </div>

    );
  }



  private viewItem(): void {

    const itemID = new URLSearchParams(window.location.search).get('itemid');

    sp.web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('*', 'AccountManager/Title', 'ProjectManager/Title').expand('ProjectManager/Id', 'AccountManager/Id').get().then((item) => {
      $('.CommentCM').text(item.KeyPointersToStWithoutContract);
      $('.SignedDate').text(moment(item.ExpectedContractSignDueDt).format('LL'));

    });
  }


  private Approve() {

    var isvalid = true;

    if ($("#MgtComments").val() === "") {
      $("#MgtErr").show();
      isvalid = false;
    }
    else
      $("#MgtErr").hide();

    if (isvalid) {
      this.Approve1();
    }

  }

  private Approve1(): void {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + (today.getDate() + 1);
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var DueDateTime = date + ' | ' + time;

    var commentcheck = $("#MgtCommentsUnsigned").val().toString();
    let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).update({
        'Status': `CM Documents Upload Awaited`,
        DueDateTime: DueDateTime,
        'MgtComment4Unsigned': commentcheck,
        'WithoutAnyAmendments': "NA",
        'DetailsOfAmendment': "NA",
        'AmendmentComment': "NA"
      }).then(i => {
        let web = Web(this.props.webURL);
        web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "Management"
        }).then(j => {
          let web = Web(this.props.webURL);
          web.lists.getByTitle('CompleteTask').items.add({
            ProjectID: parseInt(itemID),
            TaskType: "Marketing"
          }).then(s => {
            $("#loader").show();
            setTimeout(() => { $('#loader').hide(); }, 3000);
            setTimeout(() => { $('#ApproveModal').show(); }, 3500);
          });
        });
      });
  }

  private Reject() {
    var isvalid = true;

    if ($("#MgtCommentsUnsigned").val() === "") {
      $("#MgtErr").show();
      isvalid = false;
    }
    else
      $("#MgtErr").hide();
    if (isvalid) {
      this.Reject2();
    }
  }

  private Reject2(): void {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).update({
        'Status': `Rejected To Start Without Contract`,
        MgtComment4Unsigned: $("#MgtCommentsUnsigned").val().toString(),
      }).then(i => {
       let web = Web(this.props.webURL);
         web.lists.getByTitle('CompleteTask').items.add({
          ProjectID: parseInt(itemID),
          TaskType: "Management"
        }).then(j => {
          let web = Web(this.props.webURL);
         web.lists.getByTitle('CompleteTask').items.add({
            ProjectID: parseInt(itemID),
            TaskType: "Marketing"
          }).then(newListItem => {
            $("#loader").show();
            setTimeout(() => { $('#loader').hide(); }, 3000);
            setTimeout(() => { $('#RejectModal').show(); }, 3500);
          });
        });
      });
  }



  private setForm() {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    sp.web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'Status', 'OpportunityID').get().then((data) => {

      if (data.Status == "Management Approval Awaited to Start Without Signed Contract") {
        document.getElementById("managementConsent").style.display = "block";
      }
    });
  }


  private Close0() {
    var modal = document.getElementById("ApproveModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#managementConsent').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 1000);
  }
  private Close1() {
    var modal = document.getElementById("RejectModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#managementConsent').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 1000);

  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#LegalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 1000);
  }






}


