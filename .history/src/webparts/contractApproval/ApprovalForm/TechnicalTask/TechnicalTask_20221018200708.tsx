import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './TechnicalTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import {  PrimaryButton } from 'office-ui-fabric-react';
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import { IStackTokens, mergeStyleSets} from '@fluentui/react';
// require('./css/jquery-ui.css');
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


 const classes = mergeStyleSets({
  cell: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    margin: '80px',
    float: 'left',
    height: '50px',
    width: '50px',
  },
  icon: {
    fontSize: '50px',
  },
  code: {
    background: '#f2f2f2',
    borderRadius: '4px',
    padding: '4px',
  },
  navigationText: {
    width: 100,
    margin: '0 5px',
  },
});
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

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };
const stackTokens: IStackTokens = { childrenGap: 40 };
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
  public async componentDidMount() {
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
      var QueLength = item.length;
      for (var s = 0; s < QueLength; s++) {
        var QuestionNumber = item[s].QuestionNumber;
        var lablecontrol = styles.lablecontrol;
        var Styles = styles.form_control;
        var errorlable = styles.errorlable;
        var Questions = "";
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
          var Options = "";
          Questions += '<div class="Div' + QuestionNumber + '" id="Div' + QuestionNumber + '" style="display: block;"><span id="Requirement' + QuestionNumber + '"style="display:none">' + item[s].Requirement + '</span>'
            + '<lable class="' + lablecontrol + '" value="' + item[s].QuestionNumber + '" id="Question' + QuestionNumber + '">' + item[s].Question + '</lable>'
            + '<select name="Answer' + QuestionNumber + '" class="' + Styles + '" style = "width: 200px; display: block;" id="Answer' + QuestionNumber + '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>'
            + '<span id="CRQuestionErr' + QuestionNumber + '" style="display:none; " class="' + errorlable + '" for="Author">Please select the value from the drop-down.</span><br>'
            + '<span id="TechCommentErr' + QuestionNumber + '" style="display:none;" class="' + errorlable + '" for="Author">This field is required.</span><br>'
            + '</div>';
          $("#Questions").append(Questions);
          for (var w = 0; w < item[s].OptionsForDropDown.split(";").length; w++) {
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
    var filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(data => {
      for (var b = 0; b < data.length; b++) {
        var TECNo = data[b].QuestionNumber;
        $("#Answer" + TECNo).val(data[b].Answer);
        if (data[b].Comment !== null) {
          $("#TechComment" + TECNo).show();
          $("#TechComment" + TECNo).val(data[b].Comment);
        }
        else {
          $("#TechComment" + TECNo).hide();
        }
      }
    }).then(async () => {
      var filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'  and " + "FeedbackStatus eq 'Submit'";
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
    var TechName = new URLSearchParams(window.location.search).get('dpt');
    var TechNameWOP = TechName.split("-");
    document.getElementById("TechTeamName").innerHTML = TechNameWOP[1];

  }

  private CancelForm() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }



  private TechSubmitData() {
    var isvalid = true;
    // =========== Questions Validations ====================
    this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then((Technicaldata) => {
      var Technicaldatalength = Technicaldata.length;
      for (var i = 0; i < Technicaldatalength; i++) {
        var ID = Technicaldata[i].ID;
        var QuestionNumber = Technicaldata[i].QuestionNumber;
        var Answer = $('#Answer' + QuestionNumber).val();
        var Comment = $('#TechComment' + QuestionNumber).val();
        var CommentError = $("#TechCommentErr" + QuestionNumber);
        var CRQuestionErr = $("#CRQuestionErr" + QuestionNumber);
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
        var modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }

  private TechSubmitData2() {
    $("#loader").show();
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    var filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(async (feedbacks) => {
      if (feedbacks.length == 0) {
        this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then(async (Technicaldata) => {
          var Tdatalength = Technicaldata.length;
          for (var i = 0; i < Tdatalength; i++) {
            var TitleHTML = $('#Question' + Technicaldata[i].QuestionNumber)[0].innerText;
            var QuestionNo = $('#Question' + Technicaldata[i].QuestionNumber)[0].attributes[1].value;
            var TAnswer = $('#Answer' + Technicaldata[i].QuestionNumber).val();
            var TechtxtComments = $('#TechComment' + Technicaldata[i].QuestionNumber).val();
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
        this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(async (DeleteOld) => {
          for (var remove = 0; remove < DeleteOld.length; remove++) {
            this.webURL.lists.getByTitle('TechnicalFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then(async () => {
          this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then(async (Technicaldata2) => {
            var Technicaldatalength = Technicaldata2.length;
            for (var i = 0; i < Technicaldatalength; i++) {
              var TitleHTML = $('#Question' + Technicaldata2[i].QuestionNumber)[0].innerText;
              var QuestionNo = $('#Question' + Technicaldata2[i].QuestionNumber)[0].attributes[1].value;
              var TAnswer = $('#Answer' + Technicaldata2[i].QuestionNumber).val();
              var TechtxtComments = $('#TechComment' + Technicaldata2[i].QuestionNumber).val();
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
    }).then(async r => {
      var filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
      this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr4).getAll().then((items) => {
        if (items.length > 0) {
          items.forEach(async (item4, index4) => {
            this.webURL.lists.getByTitle("TechnicalComments").items.getById(item4.Id).update({
              IsFinished: true,
            }).then(async s => {
              var TechName = new URLSearchParams(window.location.search).get('dpt');
              var filterStr5 = "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Technical'";
              this.webURL.lists.getByTitle("ProjectTasks").items.filter(filterStr5).getAll().then(itemz => {
                if (itemz.length > 0) {
                  itemz.forEach(async (item2, index2) => {
                    if (TechName == item2.VIew.split("&")[1].split("=")[1]) {
                      this.webURL.lists.getByTitle("ProjectTasks").items.getById(item2.Id).update({
                        // TaskType: "Technical",
                        // Action: "Technical Feedback",
                        Status: "Completed"
                      });
                    }
                  });
                }
              }).then(async t => {
                var filterStr3 = "ProjectID eq '" + itemID + "' and " + "IsFinished eq '" + false + "' and " + "Title eq 'PT'";
                this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr3).get().then(async data3 => {
                  if (data3.length == 0) {
                    var filterStr6 = "ProjectID eq '" + itemID + "' and " + "Title eq 'Delivery Action Awaited'";
                    this.webURL.lists.getByTitle("ProjectTasks").items.filter(filterStr6).getAll().then(async DelTask => {
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


  private async TechDraftData() {
    var isvalid = true;
    // =========== Questions Validations ====================
    this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then((Technicaldata) => {
      var Technicaldatalength = Technicaldata.length;
      for (var i = 0; i < Technicaldatalength; i++) {
        var ID = Technicaldata[i].ID;
        var QuestionNumber = Technicaldata[i].QuestionNumber;
        var Answer = $('#Answer' + QuestionNumber).val();
        var Comment = $('#TechComment' + QuestionNumber).val();
        var CommentError = $("#TechCommentErr" + QuestionNumber);

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

  private async TechDraftData1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    this.webURL.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {

      if (Pdata.TechnicalDone != 2) {
        this.TechDraftData2();
      }
      else {
        var modal = document.getElementById("ApprovedModal");
        modal.style.display = "block";
      }
    });
  }

  private Close0() {
    var modal = document.getElementById("ApprovedModal");
    modal.style.display = "none";
  }


  private Close1() {
    var modal = document.getElementById("SubmittedModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#TechnicalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);

  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#TechnicalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }


  private TechDraftData2() {
    $("#loader").show();
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    var filterStr = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(async (feedbacks) => {
      if (feedbacks.length == 0) {

        this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then(async (Technicaldata) => {
          var Tdatalength = Technicaldata.length;
          for (var i = 0; i < Tdatalength; i++) {
            var TitleHTML = $('#Question' + Technicaldata[i].QuestionNumber)[0].innerText;
            var QuestionNo = $('#Question' + Technicaldata[i].QuestionNumber)[0].attributes[1].value;
            var TAnswer = $('#Answer' + Technicaldata[i].QuestionNumber).val();
            var TechtxtComments = $('#TechComment' + Technicaldata[i].QuestionNumber).val();
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
        this.webURL.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then(async (DeleteOld) => {
          for (var remove = 0; remove < DeleteOld.length; remove++) {
            this.webURL.lists.getByTitle('TechnicalFeedback').items.getById(DeleteOld[remove].ID).delete();
          }
        }).then(async () => {
          this.webURL.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").get().then(async (Technicaldata2) => {
            var Technicaldatalength = Technicaldata2.length;
            for (var i = 0; i < Technicaldatalength; i++) {
              var TitleHTML = $('#Question' + Technicaldata2[i].QuestionNumber)[0].innerText;
              var QuestionNo = $('#Question' + Technicaldata2[i].QuestionNumber)[0].attributes[1].value;
              var TAnswer = $('#Answer' + Technicaldata2[i].QuestionNumber).val();
              var TechtxtComments = $('#TechComment' + Technicaldata2[i].QuestionNumber).val();
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
    }).then(async () => {
      var TechName = new URLSearchParams(window.location.search).get('dpt');
      var filterStr4 = "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
      this.webURL.lists.getByTitle("TechnicalComments").items.filter(filterStr4).getAll().then((items) => {
        if (items.length > 0) {
          items.forEach(async (item4, index4) => {
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


