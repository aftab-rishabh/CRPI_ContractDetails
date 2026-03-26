import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './DeliveryTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


export interface IDeliveryTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface IDeliveryTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class DeliveryTask extends React.Component<IDeliveryTaskProps, IDeliveryTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";

  constructor(props: IDeliveryTaskProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {
    this.setForm();
    this.loadQuestions();

  }
  public render(): React.ReactElement<IDeliveryTaskProps> {
    return (
      <div className={styles.deliveryTaskForm}>
        <div className={styles.container}>
        <div className={styles.row}>
                  <div className={styles.viewdetail_block}>
          {/* <input type="hidden" id="userId" name="userId" value=""> */}
          <div id="DeliveryTaskForm">

            <div className={styles.sectionblock}>
              <Label className={styles.headers}><u><Label id="TechTeamName"></Label>Delivery Task Form</u></Label>
            </div>

            <div className="Questions" id="Questions"></div>

            <div id="loader" className={styles.modal}>
            <div className="">
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

            <div id="ApprovedModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close0()}>
                &times;
              </span>
                <label className={styles.header2}>Former feedback is already approved!</label>
              </div>
            </div>

            <div id="SubmittedModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close1()}>
                &times;
              </span>
                <label className={styles.header2}>Feedback Submitted!</label>
              </div>
            </div>

            <div id="DraftModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close2()}>
                &times;
              </span>
                <label className={styles.header2}>Draft Saved!</label>
              </div>
            </div>


            <div className="row">
              <div className="col col-lg-12 {styles.form_footer}">
                <PrimaryButton text="Draft" className={styles.btn} style={{ borderRadius: "10px",marginRight:"10px" }} id="DelDraft"   onClick={() => this.DelDraft()} >Draft</PrimaryButton>
                <PrimaryButton text="Submit" className={styles.btn}  style={{ borderRadius: "10px",marginRight:"10px" }}id="DelSubmit"   onClick={() => this.DelSubmitData()}></PrimaryButton>
                <PrimaryButton text="Cancel" className={styles.btn} style={{ borderRadius: "10px" }} id="DelCancel"   onClick={() => this.DelCancelForm()}></PrimaryButton>
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
    var filterStr = "ProjectID eq '" + itemID + "'";
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr).get().then(data => {
      for (var b = 0; b < data.length; b++) {
        var DELNo = data[b].QuestionNumber;
        $("#Answer" + DELNo).val(data[b].Answer);
        if (data[b].Comment !== null) {
          $("#DelComment" + DELNo).show();
          $("#DelComment" + DELNo).val(data[b].Comment);
        }
        else {
          $("#DelComment" + DELNo).hide();
        }
      } 

    }).then(async ()=>{
      var filterStr4 = "ProjectID eq '" + itemID + "' and " + "FeedbackStatus eq 'Submit'";
       let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr4).getAll().then((items) => {
        if(items.length !== 0){
          $("#DelDraft").hide();
        }
        else{
          $("#DelDraft").show();
        }
      });  
    });
  }

  private async loadQuestions() {
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.orderBy('QuestionNumber').filter("Team eq 'Delivery'").orderBy('Sequence').get().then((item) => {
      var QueLength = item.length;
      for (var s = 0; s < QueLength; s++) {
        var ID = item[s].ID;
        var QuestionNumber = item[s].QuestionNumber;
        var lablecontrol = styles.lablecontrol;
        var Styles = styles.form_control;
        var errorlable = styles.errorlable;
        var Questions = "";

        if (item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          Questions += '<lable class="' + lablecontrol + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable><span id="Requirement'+QuestionNumber+'"style="display:none">'+item[s].Requirement+'</span>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px;" id="Answer' + QuestionNumber + '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="DelComment' + QuestionNumber + '" style="display: none;"></textarea>'
            + '<span id="DelCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>';
          $("#Questions").append(Questions);
        }

        if (item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          Questions += '<lable class="' + lablecontrol + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable><span id="Requirement'+QuestionNumber+'"style="display:none">'+item[s].Requirement+'</span>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px;" id="Answer' + QuestionNumber + '"><option>Select</option><option>No</option><option>Yes</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="DelComment' + QuestionNumber + '" style="display: none;"></textarea>'
            + '<span id="DelCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>';
          $("#Questions").append(Questions);
        }

        if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          Questions += '<lable class="' + lablecontrol + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable><span id="Requirement'+QuestionNumber+'"style="display:none">'+item[s].Requirement+'</span>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="DelComment' + QuestionNumber + '" style="display: none;"></textarea>'
            + '<span id="DelCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>';
          $("#Questions").append(Questions);
        }

        if (item[s].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          var Options = "";
          Questions += '<lable class="' + lablecontrol + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable><span id="Requirement'+QuestionNumber+'"style="display:none">'+item[s].Requirement+'</span>'
          + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px;" id="Answer' + QuestionNumber + '"></select>'
          + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
          + '<span id="DelCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>';
          $("#Questions").append(Questions);
          for (var w = 0; w < item[s].OptionsForDropDown.split(";").length; w++) {
            Options += '<option>' + item[s].OptionsForDropDown.split(";")[w] + '</option>';
          }
          $("#Answer" + QuestionNumber).append(Options);
        }


        if(item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()){
          this.IfNoCommentChange(QuestionNumber);
        }

        if(item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()){
          this.IfYesCommentChange(QuestionNumber);
        }

        if(item[s].Requirement.toLowerCase() == "textbox".toLowerCase()){
        $("#Question"+QuestionNumber).append('<lable class="" style="color:red;">*</lable>');
        }

      }
    }).then(() =>{
      this.setForm();
    });
  }

  private IfNoCommentChange(QuestionNumber) {
    var ChangeFunction = document.getElementById("Answer" + QuestionNumber);
    ChangeFunction.addEventListener('change', () => this.ShowCommentBoxOnNo(QuestionNumber));
  }

  private ShowCommentBoxOnNo(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#DelComment" + QuestionNumber).show();
    }
    else {
      $("#DelComment" + QuestionNumber).val("");
      $("#DelComment" + QuestionNumber).hide();
      $("#DelCommentErr" + QuestionNumber).hide();
    }
  }

  private IfYesCommentChange(QuestionNumber) {
    var ChangeFunction = document.getElementById("Answer" + QuestionNumber);
    ChangeFunction.addEventListener('change', () => this.ShowCommentBoxOnYes(QuestionNumber));
  }

  private ShowCommentBoxOnYes(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#DelComment" + QuestionNumber).show();
    }
    else {
      $("#DelComment" + QuestionNumber).val("");
      $("#DelComment" + QuestionNumber).hide();
      $("#DelCommentErr" + QuestionNumber).hide();
    }
  }

  private Close0() {
    var modal = document.getElementById("ApprovedModal");
    modal.style.display = "none";
  }
  private Close1() {
    var modal = document.getElementById("SubmittedModal");
    var hostUrl = this.props.webURL;

    modal.style.display = "none"; 
    $('#DeliveryTaskForm').fadeOut(2500); 
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);

  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#DeliveryTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }

  private DelCancelForm() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }


  private async DelSubmitData() {

    var isvalid = true;
    // =========== Questions Validations ====================

     let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then( (DelData) => {

      var DelDatalength = DelData.length;

      for (var i = 0; i < DelDatalength; i++) {
        var ID = DelData[i].QuestionNumber;
        var QuestionNumber = DelData[i].QuestionNumber;
        var Answer = $('#Answer' + [ID]).val();
        var DelComment = $('#DelComment' + [ID]).val();
        var DelCommentErr = $("#DelCommentErr" + [ID]);
        var CRQuestionErr = $("#CRQuestionErr" + [ID]);
        if (Answer !== "Select") {
          CRQuestionErr.hide();
        if (DelData[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          if (Answer === "No" && DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          if (Answer === "Yes" && DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          if (DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          if (Answer === "Select") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }
      }else{
        CRQuestionErr.show();
        $('#Answer' + QuestionNumber).focus();
        isvalid = false;
        window.stop();
      }


      }
    }).then(checkit => {
      if (isvalid == true) {
        this.DelSubmitData1();
      }
    });
  }

  private async DelSubmitData1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {

      if (Pdata.DeliveryDone != 2) {
        this.DelSubmitData2();
      }
      else {
        var modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }

  private async DelSubmitData2() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var filterStr = "ProjectID eq '" + itemID + "'";
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr).get().then(async (feedbacks) => {
      if (feedbacks.length == 0) {
         let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then( async (DelData) => {
          var DelDatalength = DelData.length;
          for (var i = 0; i < DelDatalength; i++) {
            var ID = DelData[i].QuestionNumber;
            var Title = $('#Question' + [ID])[0].innerHTML;
            var Answer = $('#Answer' + [ID]).val();
            var DelComment = $('#DelComment' + [ID]).val();
            var Qnumber = DelData[i].QuestionNumber;
            if (Answer !== undefined && DelComment !== undefined) {
               let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.add({
                Title: Title,
                ProjectID: itemID,
                QuestionNumber: Qnumber,
                Answer: Answer,
                Requirement: $("#Requirement"+ID)[0].innerHTML,
                Comment: DelComment,
                FeedbackStatus: "Submit"
              });
            }
            else {
              console.log("xyz");
            }
          }
        }).then(async i => {
           let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'Status', 'DeliveryDone','LegalDone').get().then(async (data) => {
            status = data.Status;
             let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
              DeliveryDone: 1
            }).then(async changestatus =>{
              if (data.LegalDone !== -1 && data.Status == "Delivery Action Awaited" && data.Status !== "Management Action Awaited" && data.Status !== "Escalated to management") {
                 let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
                    Status: "Management Action Awaited", 
                });
              }  
            });
          });
        }).then(async f => {
          var filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Delivery'";
           let web = Web(this.props.webURL);
    await web.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then( (items) => {
            if (items.length > 0) {
              items.forEach( async (item2, index2) => {
                 let web = Web(this.props.webURL);
    await web.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                  // TaskType: "Management/Marketing",
                  // Action: "Approve/ Reject"
                  Status: "Completed"
                });
              });
            }
          });
        }).then(async ()=> {
          var filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Technical'";
           let web = Web(this.props.webURL);
    await web.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then( (items) => {
            if (items.length > 0) {
              items.forEach( async (item2, index2) => {
                 let web = Web(this.props.webURL);
    await web.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                  // TaskType: "Management/Marketing",
                  // Action: "Approve/ Reject"
                  Status: "Completed"
                });
              });
            }
          });
        }).then(r => {
          $("#loader").show();
          setTimeout( ()=> {$('#loader').hide();}, 3000);  
          setTimeout( ()=> {$('#SubmittedModal').show();}, 3500);
        });
      }
      else {
         let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr).get().then(async (DeleteOld) => {
          for (var remove = 0; remove < DeleteOld.length; remove++) {
             let web = Web(this.props.webURL);
    await web.lists.getByTitle('DeliveryFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then(() => {
             let web = Web(this.props.webURL);
     web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then( (DelData2) => {
              var Deliverydatalength = DelData2.length;
              for (var i = 0; i < Deliverydatalength; i++) {
                var ID = DelData2[i].QuestionNumber;
                var Title = $('#Question' + [ID])[0].innerHTML;
                var Answer = $('#Answer' + [ID]).val();
                var DelComment = $('#DelComment' + [ID]).val();
                var Qnumber = DelData2[i].QuestionNumber;
                 let web = Web(this.props.webURL);
     web.lists.getByTitle("DeliveryFeedback").items.add({
                  Title: Title,
                  ProjectID: itemID,
                  QuestionNumber: Qnumber,
                  Answer: Answer,
                  Requirement: $("#Requirement"+ID)[0].innerHTML,
                  Comment: DelComment,
                  FeedbackStatus: "Submit"
                });
              }
            });
        }).then(f => {
          var filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Delivery'";
           let web = Web(this.props.webURL);
     web.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then((items) => {
            if (items.length > 0) {
              items.forEach((item2, index2) => {
                 let web = Web(this.props.webURL);
     web.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                  // TaskType: "Management/Marketing",
                  // Action: "Approve/ Reject",
                  Status: "Completed"
                });
              });
            }
          });
        }).then(()=> {
          var filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Technical'";
           let web = Web(this.props.webURL);
     web.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then( (items) => {
            if (items.length > 0) {
              items.forEach( (item2, index2) => {
                 let web = Web(this.props.webURL);
     web.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                  // TaskType: "Management/Marketing",
                  // Action: "Approve/ Reject",
                  Status: "Completed"
                });
              });
            }
          });
        }).then(n => {
           let web = Web(this.props.webURL);
     web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'Status', 'DeliveryDone','LegalDone').get().then((data) => {
            status = data.Status;
             let web = Web(this.props.webURL);
     web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
              DeliveryDone: 1
            }).then(changestatus =>{
              if (data.LegalDone !== -1 && data.Status == "Delivery Action Awaited" && data.Status !== "Management Action Awaited" && data.Status !== "Escalated to management") {
                 let web = Web(this.props.webURL);
     web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
                    Status: "Management Action Awaited", 
                });
              }  
            });
          });
        }).then(r => {
          $("#loader").show();
          setTimeout(()=> {$('#loader').hide();}, 3000);  
          setTimeout( ()=> {$('#SubmittedModal').show();}, 3500);
        });
      }
    });
  }

  private async DelDraft() {
    var isvalid = true;
    // =========== Questions Validations ====================

     let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then( (DelData) => {

      var DelDatalength = DelData.length;

      for (var i = 0; i < DelDatalength; i++) {
        var ID = DelData[i].QuestionNumber;
        var Answer = $('#Answer' + [ID]).val();
        var DelComment = $('#DelComment' + [ID]).val();
        var DelCommentErr = $("#DelCommentErr" + [ID]);

        if (DelData[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          if (Answer === "No" && DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          if (Answer === "Yes" && DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          if (DelComment === "") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }

        if (DelData[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          if (Answer === "Select") {
            DelCommentErr.show();
            isvalid = false;
            window.stop();
          }
          else {
            DelCommentErr.hide();
          }
        }
      }

    }).then(() =>{
      if (isvalid) {
        this.DelDraft1();
      }  
    });
  }

  private async DelDraft1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {

      if (Pdata.DeliveryDone != 2) {
        this.DelDraft2();
      }
      else {
        var modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }



  private async DelDraft2() {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var filterStr = "ProjectID eq '" + itemID + "'";
     let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr).get().then(async (feedbacks) => {
      if (feedbacks.length == 0) {
         let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then( async (DelData) => {
          var DelDatalength = DelData.length;
          for (var i = 0; i < DelDatalength; i++) {
            var ID = DelData[i].QuestionNumber;
            var Title = $('#Question' + [ID])[0].innerHTML;
            var Answer = $('#Answer' + [ID]).val();
            var DelComment = $('#DelComment' + [ID]).val();
            var Qnumber = DelData[i].QuestionNumber;
               let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.add({
                Title: Title,
                ProjectID: itemID,
                QuestionNumber: Qnumber,
                Answer: Answer,
                Requirement: $("#Requirement"+ID)[0].innerHTML,
                Comment: DelComment,
                FeedbackStatus: "Draft"
              });
            }
        }).then(r => {
          $("#loader").show();
          setTimeout( ()=> {$('#loader').hide();}, 3000);
          setTimeout( ()=> {$('#DraftModal').show();}, 3500);
        });
      }
      else {
         let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.filter(filterStr).get().then(async (DeleteOld) => {
          for (var remove = 0; remove < DeleteOld.length; remove++) {
             let web = Web(this.props.webURL);
    await web.lists.getByTitle('DeliveryFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then(async () => {
             let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").get().then(async (DelData2) => {
              var Deliverydatalength = DelData2.length;
              for (var i = 0; i < Deliverydatalength; i++) {
                var ID = DelData2[i].QuestionNumber;
                var Title = $('#Question' + [ID])[0].innerHTML;
                var Answer = $('#Answer' + [ID]).val();
                var DelComment = $('#DelComment' + [ID]).val();
                var Qnumber = DelData2[i].QuestionNumber;
                 let web = Web(this.props.webURL);
    await web.lists.getByTitle("DeliveryFeedback").items.add({
                  Title: Title,
                  ProjectID: itemID,
                  QuestionNumber: Qnumber,
                  Answer: Answer,
                  Requirement: $("#Requirement"+ID)[0].innerHTML,
                  Comment: DelComment,
                  FeedbackStatus: "Draft"
                });
              }
            });
        }).then(r => {
          $("#loader").show();
          setTimeout( ()=> {$('#loader').hide();}, 3000);  
          setTimeout( ()=> {$('#DraftModal').show();}, 3500);
        });
      }
    });


  }


}


