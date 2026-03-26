import * as React from 'react';
import styles from './TechnicalTaskFormWebPart.module.scss';
import {  PrimaryButton } from 'office-ui-fabric-react';
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import { IStackTokens, mergeStyleSets} from '@fluentui/react';
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */

export interface ITechnicalTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ITechnicalTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

export class TechnicalTask extends React.Component<ITechnicalTaskProps, ITechnicalTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  webURL:any="";
  ContractStatus = "";

  constructor(props: ITechnicalTaskProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
    
  }
  public  componentDidMount() {
    this.webURL = Web(this.props.webURL);
    this.SetFormName();
    this.loadQuestions();

  }
  public render(): React.ReactElement<ITechnicalTaskProps> {
    return (
      <div className={styles.technicalTaskForm}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.viewdetail_block}>
              {/* <input type="hidden" id="userId" name="userId" value=""> */}
              <div id="TechnicalTaskForm">

                <div className={styles.sectionblock}>
                  <Label className={styles.headers}><u><Label id="TechTeamName"></Label> Technical Task Form</u></Label>
                </div>

                <div className="Questions" id="Questions"></div>

                <div id="loader" className={styles.modal}>
                  <div className="">
                    <div
                      className={styles.loader}
                      style={{ margin: '200px auto' }}
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
                  <div className="col col-lg-12">
                      <div className={styles.form_footer}>
                        <PrimaryButton text="Draft" className={styles.btn} style={{ borderRadius: "10px",marginRight:"10px" }} id="btnDraft" onClick={() => this.TechDraftData()} >Draft</PrimaryButton>
                        <PrimaryButton text="Submit" className={styles.btn} id="btnSubmit" style={{ borderRadius: "10px",marginRight:"10px" }} onClick={() => this.TechSubmitData()}></PrimaryButton>
                        <PrimaryButton text="Cancel" className={styles.btn} id="btnCancel" style={{ borderRadius: "10px"}}  onClick={() => this.CancelForm()}></PrimaryButton>
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

  private loadQuestions() {

    this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").orderBy('Sequence').get().then((item) => {
      let QueLength = item.length;
      for (let s = 0; s < QueLength; s++) {
        let QuestionNumber = item[s].QuestionNumber;
        let lablecontrol = styles.lablecontrol;
        let Styles = styles.form_control;
        let errorlable = styles.errorlable;
        let Questions = "";
        if (item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          Questions += '<div class="Div' + QuestionNumber + '" id="Div' + QuestionNumber + '" style="display: block;"><span id="Requirement' + QuestionNumber + '"style="display:none">' + item[s].Requirement + '</span>'
            + '<lable class="' + lablecontrol + '" value="' + item[s].QuestionNumber + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px; display: block;" id="Answer' + QuestionNumber + '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="TechComment' + QuestionNumber + '" style="display: none;"></textarea>'
            + '<span id="TechCommentErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">This field is required.</span><br>'
            + '</div>';
          $("#Questions").append(Questions);
        }
        if (item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          Questions += '<div class="Div' + QuestionNumber + '" id="Div' + QuestionNumber + '" style="display: block;"><span id="Requirement' + QuestionNumber + '"style="display:none">' + item[s].Requirement + '</span>'
            + '<lable class="' + lablecontrol + '" value="' + item[s].QuestionNumber + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px; display: block;" id="Answer' + QuestionNumber + '"><option>Select</option><option>No</option><option>Yes</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="TechComment' + QuestionNumber + '" style="display: none;"></textarea>'
            + '<span id="TechCommentErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">This field is required.</span><br>'
            + '</div>';
          $("#Questions").append(Questions);
        }
        if (item[s].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          let Options = "";
          Questions += '<div class="Div' + QuestionNumber + '" id="Div' + QuestionNumber + '" style="display: block;"><span id="Requirement' + QuestionNumber + '"style="display:none">' + item[s].Requirement + '</span>'
            + '<lable class="' + lablecontrol + '" value="' + item[s].QuestionNumber + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px; display: block;" id="Answer' + QuestionNumber + '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<span id="TechCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>'
            + '</div>';
          $("#Questions").append(Questions);
          for (let w = 0; w < item[s].OptionsForDropDown.split(";").length; w++) {
            Options += '<option>' + item[s].OptionsForDropDown.split(";")[w] + '</option>';
          }
          $("#Answer" + QuestionNumber).append(Options);
        }
        if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          Questions += '<div class="Div' + QuestionNumber + '" id="Div' + QuestionNumber + '" style="display: block;"><span id="Requirement' + QuestionNumber + '"style="display:none">' + item[s].Requirement + '</span>'
            + '<lable class="' + lablecontrol + '" value="' + item[s].QuestionNumber + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable>'
            + '<textarea rows="4" cols="50" class="' + Styles + '" placeholder="Please enter comments" id="TechComment' + QuestionNumber + '"></textarea>'
            + '<span id="TechCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>'
            + '</div>';
          $("#Questions").append(Questions);
        }

        if (item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          this.IfNoCommentChange(QuestionNumber);
        }

        if (item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          this.IfYesCommentChange(QuestionNumber);
        }

        if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          $("#Question" + QuestionNumber).append('<lable class="" style="color:red;">*</lable>');
        }

      }
      this.setForm();
    });
  }



  private IfNoCommentChange(QuestionNumber) {
    document.getElementById("Answer" + QuestionNumber).addEventListener('change', () => this.ShowCommentBoxOnNo(QuestionNumber));
  }

  private ShowCommentBoxOnNo(QuestionNumber) {

    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#TechComment" + QuestionNumber).show();
    }
    else {
      $("#TechComment" + QuestionNumber).val("");
      $("#TechComment" + QuestionNumber).hide();
      $("#TechCommentErr" + QuestionNumber).hide();
    }
  }


  private IfYesCommentChange(QuestionNumber) {
    document.getElementById("Answer" + QuestionNumber).addEventListener('change', () => this.ShowCommentBoxOnYes(QuestionNumber));
  }

  private ShowCommentBoxOnYes(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#TechComment" + QuestionNumber).show();
    }
    else {
      $("#TechComment" + QuestionNumber).val("");
      $("#TechComment" + QuestionNumber).hide();
      $("#TechCommentErr" + QuestionNumber).hide();
    }
  }


  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    let filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(data => {
      for (let b = 0; b < data.length; b++) {
        let TECNo = data[b].QuestionNumber;
        $("#Answer" + TECNo).val(data[b].Answer);
        if (data[b].Comment !== null) {
          $("#TechComment" + TECNo).show();
          $("#TechComment" + TECNo).val(data[b].Comment);
        }
        else {
          $("#TechComment" + TECNo).hide();
        }
      }
    }).then( () => {
      let filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'  and " + "FeedbackStatus eq 'Submit'";
      this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr4).getAll().then((items) => {
        if (items.length !== 0) {
          $("#btnDraft").hide();
        }
        else {
          $("#btnDraft").show();
        }
      });
    });
  }

  private SetFormName() {
    let TechName = new URLSearchParams(window.location.search).get('dpt');
    let TechNameWOP = TechName.split("-");
    document.getElementById("TechTeamName").innerHTML = TechNameWOP[1];

  }

  private CancelForm() {
    let hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }



  private TechSubmitData() {
    let isvalid = true;
    // =========== Questions Validations ====================
    this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then((Technicaldata) => {
      let Technicaldatalength = Technicaldata.length;
      for (let i = 0; i < Technicaldatalength; i++) {
        let ID = Technicaldata[i].ID;
        let QuestionNumber = Technicaldata[i].QuestionNumber;
        let Answer = $('#Answer' + QuestionNumber).val();
        let Comment = $('#TechComment' + QuestionNumber).val();
        let CommentError = $("#TechCommentErr" + QuestionNumber);
        let CRQuestionErr = $("#CRQuestionErr" + QuestionNumber);
        if (Answer !== "Select") {
          CRQuestionErr.hide();
          if (Technicaldata[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
            if (Answer === "No" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            }
            else {
              CommentError.hide();
            }
          }

          if (Technicaldata[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
            if (Answer === "Yes" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            }
            else {
              CommentError.hide();
            }
          }

          if (Technicaldata[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            if (Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            }
            else {
              CommentError.hide();
            }
          }

          if (Technicaldata[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
            if (Answer === "Select") {
              CommentError.show();
              isvalid = false;
              window.stop();
            }
            else {
              CommentError.hide();
            }
          }
        } else {
          CRQuestionErr.show();
          $('#Answer' + QuestionNumber).focus();
          isvalid = false;
          window.stop();
        }
      }

    }).then(check => {
      if (isvalid == true) {
        this.TechSubmitData1();
      }
    });
  }

  private  TechSubmitData1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    this.webURL.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {
      this.ContractStatus = Pdata.Status;
      if (Pdata.TechnicalDone != 2) {
        this.TechSubmitData2();
      }
      else {
        let modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }

  private TechSubmitData2() {
    $("#loader").show();
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    let filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then( (feedbacks) => {
      if (feedbacks.length == 0) {
        this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then( (Technicaldata) => {
          let Tdatalength = Technicaldata.length;
          for (let i = 0; i < Tdatalength; i++) {
            let TitleHTML = $('#Question' + Technicaldata[i].QuestionNumber)[0].innerText;
            let QuestionNo = $('#Question' + Technicaldata[i].QuestionNumber)[0].attributes[1].value;
            let TAnswer = $('#Answer' + Technicaldata[i].QuestionNumber).val();
            let TechtxtComments = $('#TechComment' + Technicaldata[i].QuestionNumber).val();
            this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
              Title: TitleHTML,
              ProjectID: itemID,
              Department: Dpt,
              QuestionNumber: QuestionNo,
              Requirement: $("#Requirement" + Technicaldata[i].QuestionNumber)[0].innerHTML,
              Answer: TAnswer,
              Comment: TechtxtComments,
              FeedbackStatus: "Submit"
            });
          }
        });
      }
      else {
        this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then( (DeleteOld) => {
          for (let remove = 0; remove < DeleteOld.length; remove++) {
            this.webURL.lists.getByTitle('TechnicalFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then( () => {
          this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then( (Technicaldata2) => {
            let Technicaldatalength = Technicaldata2.length;
            for (let i = 0; i < Technicaldatalength; i++) {
              let TitleHTML = $('#Question' + Technicaldata2[i].QuestionNumber)[0].innerText;
              let QuestionNo = $('#Question' + Technicaldata2[i].QuestionNumber)[0].attributes[1].value;
              let TAnswer = $('#Answer' + Technicaldata2[i].QuestionNumber).val();
              let TechtxtComments = $('#TechComment' + Technicaldata2[i].QuestionNumber).val();
              this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
                Title: TitleHTML,
                ProjectID: itemID,
                Department: Dpt,
                QuestionNumber: QuestionNo,
                Requirement: $("#Requirement" + Technicaldata2[i].QuestionNumber)[0].innerHTML,
                Answer: TAnswer,
                Comment: TechtxtComments,
                FeedbackStatus: "Submit"
              });
            }
          });
        });
      }
    }).then( r => {
      let filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
      this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr4).getAll().then((items) => {
        if (items.length > 0) {
          items.forEach( (item4, index4) => {
            this.webURL.lists.getByTitle("TechnicalComments").items.getById(item4.Id).update({
              IsFinished: true,
            }).then( s => {
              let TechName = new URLSearchParams(window.location.search).get('dpt');
              let filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Technical'";
              this.webURL.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then(itemz => {
                if (itemz.length > 0) {
                  itemz.forEach( (item2, index2) => {
                    if (TechName == item2.VIew.split("&")[1].split("=")[1]) {
                      this.webURL.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                        // TaskType: "Technical",
                        // Action: "Technical Feedback",
                        Status: "Completed"
                      });
                    }
                  });
                }
              }).then( t => {
                let filterStr3 = "ProjectID eq '" + itemID + "' and " + "IsFinished eq '" + false + "' and " + "Title eq 'PT'";
                this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr3).get().then( data3 => {
                  if (data3.length == 0) {
                    let filterStr6 = "ProjectID eq '" + itemID + "' and " + "Title eq 'Delivery Action Awaited'";
                    this.webURL.lists.getByTitle("ProjectTasks").items.filter(filterStr6).getAll().then( DelTask => {
                      if (this.ContractStatus == "Technical Action Awaited") {
                        this.webURL.lists.getByTitle("Projects").items.getById(parseInt(itemID)).update({
                          Status: `Delivery Action Awaited`,
                          TechnicalDone: 1
                        }).then(newListItem => {
                          setTimeout(() => { $('#loader').hide(); }, 5000);
                          setTimeout(() => { $('#SubmittedModal').show(); }, 3000);
                        });
                      }
                      else {
                        setTimeout(() => { $('#loader').hide(); }, 5000);
                        setTimeout(() => { $('#SubmittedModal').show(); }, 3000);
                      }
                    });
                  }
                  else {
                    setTimeout(() => { $('#loader').hide(); }, 5000);
                    setTimeout(() => { $('#SubmittedModal').show(); }, 3000);
                  }
                });

              });
            });
          });
        }
      });
    });
  }


  private  TechDraftData() {
    let isvalid = true;
    // =========== Questions Validations ====================
    this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then((Technicaldata) => {
      let Technicaldatalength = Technicaldata.length;
      for (let i = 0; i < Technicaldatalength; i++) {
        let ID = Technicaldata[i].ID;
        let QuestionNumber = Technicaldata[i].QuestionNumber;
        let Answer = $('#Answer' + QuestionNumber).val();
        let Comment = $('#TechComment' + QuestionNumber).val();
        let CommentError = $("#TechCommentErr" + QuestionNumber);

        if (Technicaldata[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase()) {
          if (Answer === "No" && Comment === "") {
            CommentError.show();
            isvalid = false;
            window.stop();
          }
          else {
            CommentError.hide();
          }
        }

        if (Technicaldata[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase()) {
          if (Answer === "Yes" && Comment === "") {
            CommentError.show();
            isvalid = false;
            window.stop();
          }
          else {
            CommentError.hide();
          }
        }

        if (Technicaldata[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          if (Comment === "") {
            CommentError.show();
            isvalid = false;
            window.stop();
          }
          else {
            CommentError.hide();
          }
        }

        if (Technicaldata[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          if (Answer === "Select") {
            CommentError.show();
            isvalid = false;
            window.stop();
          }
          else {
            CommentError.hide();
          }
        }
      }
    }).then(check => {
      if (isvalid) {
        this.TechDraftData1();
      }
    });
  }

  private  TechDraftData1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    this.webURL.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {

      if (Pdata.TechnicalDone != 2) {
        this.TechDraftData2();
      }
      else {
        let modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }

  private Close0() {
    let modal = document.getElementById("ApprovedModal");
    modal.style.display = "none";
  }


  private Close1() {
    let modal = document.getElementById("SubmittedModal");
    let hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#TechnicalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);

  }

  private Close2() {
    let modal = document.getElementById("DraftModal");
    let hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#TechnicalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }


  private TechDraftData2() {
    $("#loader").show();
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    let filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then( (feedbacks) => {
      if (feedbacks.length == 0) {

        this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then( (Technicaldata) => {
          let Tdatalength = Technicaldata.length;
          for (let i = 0; i < Tdatalength; i++) {
            let TitleHTML = $('#Question' + Technicaldata[i].QuestionNumber)[0].innerText;
            let QuestionNo = $('#Question' + Technicaldata[i].QuestionNumber)[0].attributes[1].value;
            let TAnswer = $('#Answer' + Technicaldata[i].QuestionNumber).val();
            let TechtxtComments = $('#TechComment' + Technicaldata[i].QuestionNumber).val();
            this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
              Title: TitleHTML,
              ProjectID: itemID,
              Department: Dpt,
              QuestionNumber: QuestionNo,
              Answer: TAnswer,
              Requirement: $("#Requirement" + Technicaldata[i].QuestionNumber)[0].innerHTML,
              Comment: TechtxtComments,
              FeedbackStatus: "Draft"
            });
          }
        });

      }

      else {
        this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then( (DeleteOld) => {
          for (let remove = 0; remove < DeleteOld.length; remove++) {
            this.webURL.lists.getByTitle('TechnicalFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then( () => {
          this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then( (Technicaldata2) => {
            let Technicaldatalength = Technicaldata2.length;
            for (let i = 0; i < Technicaldatalength; i++) {
              let TitleHTML = $('#Question' + Technicaldata2[i].QuestionNumber)[0].innerText;
              let QuestionNo = $('#Question' + Technicaldata2[i].QuestionNumber)[0].attributes[1].value;
              let TAnswer = $('#Answer' + Technicaldata2[i].QuestionNumber).val();
              let TechtxtComments = $('#TechComment' + Technicaldata2[i].QuestionNumber).val();
              this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
                Title: TitleHTML,
                ProjectID: itemID,
                Department: Dpt,
                QuestionNumber: QuestionNo,
                Requirement: $("#Requirement" + Technicaldata2[i].QuestionNumber)[0].innerHTML,
                Answer: TAnswer,
                Comment: TechtxtComments,
                FeedbackStatus: "Draft"
              });
            }
          });
        });
      }
    }).then( () => {
      let TechName = new URLSearchParams(window.location.search).get('dpt');
      let filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
      this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr4).getAll().then((items) => {
        if (items.length > 0) {
          items.forEach( (item4, index4) => {
            this.webURL.lists.getByTitle("TechnicalComments").items.getById(item4.Id).update({
              IsFinished: false,
            });
          });
        }
      });
    }).then(r => {
      setTimeout(() => { $('#loader').hide(); }, 3000);
      setTimeout(() => { $('#DraftModal').show(); }, 3500);
    });
  }


}


