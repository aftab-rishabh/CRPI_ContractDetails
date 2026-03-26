import * as React from "react";
import styles from "./TechnicalTaskFormWebPart.module.scss";
import { PrimaryButton } from "office-ui-fabric-react";
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from "office-ui-fabric-react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import * as $ from "jquery";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as ReactDOM from "react-dom";
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

export interface ITechnicalTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
  context: WebPartContext;
}


interface IQuestion {
  Sequence : string,
  Question : string,
  ProjectID : string,
  QuestionNumber : string,
  Answer : string,
  Comment : string
}


export interface ITechnicalTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
  RichtextData: IQuestion[];
}

export class TechnicalTask extends React.Component<
  ITechnicalTaskProps,
  ITechnicalTaskState
> {
  webURL: any = "";
  currentUserDetails: any = null;
  ContractStatus = "";




  constructor(props: ITechnicalTaskProps) {
    super(props);

    this.state = {
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
      RichtextData: []
    };

    this.onTextChange = this.onTextChange.bind(this); 
    this.setForm = this.setForm.bind(this); 
    this.renderRichTextControl = this.renderRichTextControl.bind(this); 
    this.TechSubmitData = this.TechSubmitData.bind(this);
    this.TechSubmitData1 = this.TechSubmitData1.bind(this); 
    this.TechSubmitData2 = this.TechSubmitData2.bind(this); 
    this.loadQuestions = this.loadQuestions.bind(this); 

  }


  public async componentDidMount() {
    this.currentUserDetails = await this.spLoggedInUserDetails();
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
                  <Label className={styles.headers}>
                    <u>
                      <Label id="TechTeamName"></Label> Technical Task Form
                    </u>
                  </Label>
                </div>

                <div className="Questions" id="Questions"></div>

                <div id="loader" className={styles.modal}>
                  <div className="">
                    <div
                      className={styles.loader}
                      style={{ margin: "200px auto" }}
                    ></div>
                  </div>
                </div>

                <div id="ApprovedModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span
                      className={styles.close}
                      onClick={() => this.Close0()}
                    >
                      &times;
                    </span>
                    <label className={styles.header2}>
                      Former feedback is already approved!
                    </label>
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
                    <label className={styles.header2}>
                      Feedback Submitted!
                    </label>
                  </div>
                </div>

                <div id="DraftModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span
                      className={styles.close}
                      onClick={() => this.Close2()}
                    >
                      &times;
                    </span>
                    <label className={styles.header2}>Draft Saved!</label>
                  </div>
                </div>

                <div className="row">
                  <div className="col col-lg-12">
                    <div className={styles.form_footer}>
                      <PrimaryButton
                        text="Draft"
                        className={styles.btn}
                        style={{ borderRadius: "10px", marginRight: "10px" }}
                        id="btnDraft"
                        onClick={() => this.TechDraftData()}
                      >
                        Draft
                      </PrimaryButton>
                      <PrimaryButton
                        text="Submit"
                        className={styles.btn}
                        id="btnSubmit"
                        style={{ borderRadius: "10px", marginRight: "10px" }}
                        onClick={() => this.TechSubmitData()}
                      ></PrimaryButton>
                      <PrimaryButton
                        text="Cancel"
                        className={styles.btn}
                        id="btnCancel"
                        style={{ borderRadius: "10px" }}
                        onClick={() => this.CancelForm()}
                      ></PrimaryButton>
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





  public onTextChange(text: string, parentDivId: string){
    
   

    this.setState(prevState => {
    
      const updatedRichtextData =  prevState.RichtextData.map(item => 
        item.QuestionNumber === parentDivId
          ? { ...item, Comment: text.replace(/<\/?[^>]+(>|$)/g, "").replace(/&nbsp;/g, "").trim().length === 0 ? "" : text}  // Update the Comment field
          : item
      );
     
      return { RichtextData: updatedRichtextData };
    
    });
  
  
    return text;
  }



  private renderRichTextControl(elementId: string) {
    const element = document.getElementById(elementId);

    const item = this.state.RichtextData.find(item => item.QuestionNumber === elementId);
  
    const comment = item ? item.Comment : ''; // Returns the Comment if the item is found, otherwise an empty string


    if (element) {
      ReactDOM.render(
        <RichText className={`${elementId}`} value={comment} onChange={(text)=> {
          
          let parentDivId = document.getElementById(elementId).getAttribute("id");
          
          return this.onTextChange(text,parentDivId);
        
        } }
        />,
        element
      );
    }
  }



  private async loadQuestions() {

    

    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let temp = false;



    let filterStr = "ProjectID eq '" + itemID + "'";

    await this.webURL.lists
      .getByTitle("TechnicalFeedback")
      .items.filter(filterStr)
      .get()
      .then((data) => {

        if (data.length > 0) {
          temp = true;
        }
        for (let b = 0; b < data.length; b++) {

          let CommentVal = "";

          if (data[b].Comment === null || data[b].Comment?.replace(/<\/?[^>]+(>|$)/g, "").replace(/&nbsp;/g, "").trim() === "") {
            CommentVal = "";
          } else {
            CommentVal = `${data[b].Comment}`;
          }



          this.setState(prevState => ({
            RichtextData: [...prevState.RichtextData,
            {
              Sequence: ``,
              Question: `${data[b].Title}`,
              ProjectID: `${itemID}`,
              QuestionNumber: `TechComment${data[b].QuestionNumber}`,
              Answer: `${data[b].Answer}`,
              Comment: CommentVal

            }
            ]
          }));


        }
      });







   await this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.filter("Team eq 'Technical'")
      .orderBy("Sequence")
      .get()
      .then((item) => {
        let QueLength = item.length;
        for (let s = 0; s < QueLength; s++) {




          if(temp){

          }else {
          
            this.setState(prevState => ({
              RichtextData:  [...prevState.RichtextData ,
                {
                  Sequence : `${item[s].Sequence}`,
                  Question : `${item[s].Question}`,
                  ProjectID : `${itemID}`,
                  QuestionNumber : `TechComment${item[s].QuestionNumber}`,
                  Answer : ``,
                  Comment : ``
                }
              ]
            }));
        
          }


          let QuestionNumber = item[s].QuestionNumber;
          let lablecontrol = styles.lablecontrol;
          let Styles = styles.form_control;
          let errorlable = styles.errorlable;
          let Questions = "";


// Logic for IfYesNoComment starts

if (
  item[s].Requirement.toLowerCase() == "IfYesNoCommentYes".toLowerCase() ||  item[s].Requirement.toLowerCase() == "IfYesNoCommentNo".toLowerCase()
) {
  Questions +=
              '<div class="Div' +
              QuestionNumber +
              '" id="Div' +
              QuestionNumber +
              '" style="display: block;"><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<lable class="' +
              lablecontrol +
              '" value="' +
              item[s].QuestionNumber +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              "</lable>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px; display: block;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
                 '<div id="TechComment' + QuestionNumber + '" class="' + styles.richTextContainer +'" ></div>'
              +
              '<span id="TechCommentErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>' +
              "</div>";
            $("#Questions").append(Questions);

            this.renderRichTextControl(`TechComment${QuestionNumber}`);
}

// Logic for IfYesNoComment Ends


          if (
            item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()
          ) {
            Questions +=
              '<div class="Div' +
              QuestionNumber +
              '" id="Div' +
              QuestionNumber +
              '" style="display: block;"><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<lable class="' +
              lablecontrol +
              '" value="' +
              item[s].QuestionNumber +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              "</lable>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px; display: block;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<div id="TechComment' + QuestionNumber + '" class="' + styles.richTextContainer +'" style="display: none;"></div>' +
              '<span id="TechCommentErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>' +
              "</div>";
            $("#Questions").append(Questions);
            this.renderRichTextControl(`TechComment${QuestionNumber}`);
          }
          if (
            item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()
          ) {
            Questions +=
              '<div class="Div' +
              QuestionNumber +
              '" id="Div' +
              QuestionNumber +
              '" style="display: block;"><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<lable class="' +
              lablecontrol +
              '" value="' +
              item[s].QuestionNumber +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              "</lable>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px; display: block;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<div id="TechComment' + QuestionNumber + '" class="' + styles.richTextContainer +'" style="display: none;"></div>' +
              '<span id="TechCommentErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>' +
              "</div>";
            $("#Questions").append(Questions);
            this.renderRichTextControl(`TechComment${QuestionNumber}`);

          }
          if (item[s].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
            let Options = "";
            Questions +=
              '<div class="Div' +
              QuestionNumber +
              '" id="Div' +
              QuestionNumber +
              '" style="display: block;"><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<lable class="' +
              lablecontrol +
              '" value="' +
              item[s].QuestionNumber +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              "</lable>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px; display: block;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none; " class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<span id="TechCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>' +
              "</div>";
            $("#Questions").append(Questions);
            for (
              let w = 0;
              w < item[s].OptionsForDropDown.split(";").length;
              w++
            ) {
              Options +=
                "<option>" +
                item[s].OptionsForDropDown.split(";")[w] +
                "</option>";
            }
            $("#Answer" + QuestionNumber).append(Options);
          }
          if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            Questions +=
              '<div class="Div' +
              QuestionNumber +
              '" id="Div' +
              QuestionNumber +
              '" style="display: block;"><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<lable class="' +
              lablecontrol +
              '" value="' +
              item[s].QuestionNumber +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              "</lable>" +
              '<div id="TechComment' + QuestionNumber + '" class="' + styles.richTextContainer +'" style="display: block;"></div>' +
              '<span id="TechCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>' +
              "</div>";
            $("#Questions").append(Questions);
            this.renderRichTextControl(`TechComment${QuestionNumber}`);
          }

          if (
            item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()
          ) {
            this.IfNoCommentChange(QuestionNumber);
          }

          if (
            item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()
          ) {
            this.IfYesCommentChange(QuestionNumber);
          }

          if (
            item[s].Requirement.toLowerCase() == "IfYesNoCommentYes".toLowerCase()
          ) {
            this.IfYesNoCommentChange(QuestionNumber);
          }

          if (
            item[s].Requirement.toLowerCase() == "IfYesNoCommentNo".toLowerCase()
          ) {
            this.IfYesNoCommentChange2(QuestionNumber);
          }

          if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            $("#Question" + QuestionNumber).append(
              '<lable class="" style="color:red;">*</lable>'
            );
          }
        }
        this.setForm();
      });

     
  }

  private IfNoCommentChange(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnNo(QuestionNumber)
      );
  }


//************New changes*******************************************************/

  private IfYesNoCommentChange(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnYesNo(QuestionNumber)
      );
  }


  private ShowCommentBoxOnYesNo(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("");
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("Not Applicable");
    } else {
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("");
    }
  }

  private IfYesNoCommentChange2(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnYesNo2(QuestionNumber)
      );
  }


  private ShowCommentBoxOnYesNo2(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("");
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("Not Applicable");
    } else {
      $("#TechComment" + QuestionNumber).find(".ql-editor").html("");
    }
  }


//*******************************************************************/


  private ShowCommentBoxOnNo(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#TechComment" + QuestionNumber).show();
    } else {
      $("#TechComment" + QuestionNumber).val("");
      $("#TechComment" + QuestionNumber).hide();
      $("#TechCommentErr" + QuestionNumber).hide();
    }
  }

  private IfYesCommentChange(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnYes(QuestionNumber)
      );
  }

  private ShowCommentBoxOnYes(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#TechComment" + QuestionNumber).show();
    } else {
      $("#TechComment" + QuestionNumber).val("");
      $("#TechComment" + QuestionNumber).hide();
      $("#TechCommentErr" + QuestionNumber).hide();
    }
  }

  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    const Dpt = new URLSearchParams(window.location.search).get("dpt");
    let filterStr =
      "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
    this.webURL.lists
      .getByTitle("TechnicalFeedback")
      .items.filter(filterStr)
      .get()
      .then((data) => {
        for (let b = 0; b < data.length; b++) {
          let TECNo = data[b].QuestionNumber;
          $("#Answer" + TECNo).val(data[b].Answer);

  if(data[b].Requirement.toLowerCase() == "IfYesNoCommentYes".toLowerCase() || data[b].Requirement.toLowerCase() == "IfYesNoCommentNo".toLowerCase() )
{
}
else{
  if (data[b].Comment !== null) {
    $("#TechComment" + TECNo).show();
  //  $("#TechComment" + TECNo).val(data[b].Comment);
  $("#TechComment" + TECNo).val();
  } else{
    $("#TechComment" + TECNo).hide();
  }
}

     
        }
      })
      .then(() => {
        let filterStr4 =
          "ProjectID eq '" +
          itemID +
          "' and " +
          "Department eq '" +
          Dpt +
          "'  and " +
          "FeedbackStatus eq 'Submit'";
        this.webURL.lists
          .getByTitle("TechnicalFeedback")
          .items.filter(filterStr4)
          .getAll()
          .then((items) => {
            if (items.length !== 0) {
              $("#btnDraft").hide();
            } else {
              $("#btnDraft").show();
            }
          });
      });
  }

  private SetFormName() {
    let TechName = new URLSearchParams(window.location.search).get("dpt");
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
    this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.filter("Team eq 'Technical'")
      .get()
      .then(async (Technicaldata) => {
        let Technicaldatalength = Technicaldata.length;

        for (let i = 0; i < Technicaldatalength; i++) {


          let CommentVal = "";
          let item = await this.state.RichtextData.find(Data => Data.QuestionNumber === "TechComment" + Technicaldata[i].QuestionNumber);

          if(item){
            CommentVal = item.Comment == null ? "" : item.Comment;
          }
          else{
            CommentVal = "";
          }

          let ID = Technicaldata[i].ID;
          let QuestionNumber = Technicaldata[i].QuestionNumber;
          let Answer = $("#Answer" + QuestionNumber).val();
          let Comment = CommentVal
          let CommentError = $("#TechCommentErr" + QuestionNumber);
          let CRQuestionErr = $("#CRQuestionErr" + QuestionNumber);
          if (Answer !== "Select") {
            CRQuestionErr.hide();
            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "IfNoComment".toLowerCase()
            ) {
              if (Answer === "No" && Comment === "") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }

            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "IfYesNoCommentYes".toLowerCase()
            ) {
              if (Answer === "No" && Comment === "") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }

            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "IfYesNoCommentNo".toLowerCase()
            ) {
              if (Answer === "Yes" && Comment === "") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }

            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "IfYesComment".toLowerCase()
            ) {
              if (Answer === "Yes" && Comment === "") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }

            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "textbox".toLowerCase()
            ) {
              if (Comment === "") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }

            if (
              Technicaldata[i].Requirement.toLowerCase() ==
              "Dropdown".toLowerCase()
            ) {
              if (Answer === "Select") {
                CommentError.show();
                isvalid = false;
                window.stop();
              } else {
                CommentError.hide();
              }
            }
          } else {
            CRQuestionErr.show();
            $("#Answer" + QuestionNumber).focus();
            isvalid = false;
            window.stop();
          }
        }
      })
      .then(() => {
        if (isvalid == true) {
          this.TechSubmitData1();
        }
      });
  }

  private TechSubmitData1() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    this.webURL.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .get()
      .then((Pdata) => {
        this.ContractStatus = Pdata.Status;
        if (Pdata.TechnicalDone != 2) {
          this.TechSubmitData2();
        } else {
          let modal = document.getElementById("ApprovedModal");
          modal.style.display = "block";
        }
      });
  }

  private TechSubmitData2() {
    $("#loader").show();

    let commentdata = [];

    commentdata = this.state.RichtextData;

    const itemID = new URLSearchParams(window.location.search).get("itemid");
    const Dpt = new URLSearchParams(window.location.search).get("dpt");
    let filterStr =
      "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists
      .getByTitle("TechnicalFeedback")
      .items.filter(filterStr)
      .get()
      .then((feedbacks) => {
        if (feedbacks.length == 0) {
          this.webURL.lists
            .getByTitle("FeedbackQuestions")
            .items.filter("Team eq 'Technical'")
            .get()
            .then((Technicaldata) => {
              let Tdatalength = Technicaldata.length;
              for (let i = 0; i < Tdatalength; i++) {
                let TitleHTML = $(
                  "#Question" + Technicaldata[i].QuestionNumber
                )[0].innerText;
                let QuestionNo = $(
                  "#Question" + Technicaldata[i].QuestionNumber
                )[0].attributes[1].value;
                let TAnswer = $(
                  "#Answer" + Technicaldata[i].QuestionNumber
                ).val();
               


                let Techtextcommemntelement = '';
                let matchvalue = "TechComment" + Technicaldata[i].QuestionNumber;
                for (let i = 0; i < commentdata.length; i++) {
                  if (commentdata[i].QuestionNumber === matchvalue) {
                    Techtextcommemntelement = commentdata[i].Comment;
                    break;  // Exit the loop once the matching item is found
                  }
                }
                
                
              
                let TechtxtComments = Techtextcommemntelement ? Techtextcommemntelement : '';
             


                this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
                  Title: TitleHTML,
                  ProjectID: itemID,
                  Department: Dpt,
                  QuestionNumber: QuestionNo,
                  Requirement: $(
                    "#Requirement" + Technicaldata[i].QuestionNumber
                  )[0].innerHTML,
                  Answer: TAnswer,
                  Comment: TechtxtComments,
                  FeedbackStatus: "Submit",
                });
              }
            });
        } else {
          this.webURL.lists
            .getByTitle("TechnicalFeedback")
            .items.filter(filterStr)
            .get()
            .then((DeleteOld) => {
              for (let remove = 0; remove < DeleteOld.length; remove++) {
                this.webURL.lists
                  .getByTitle("TechnicalFeedback")
                  .items.getById(DeleteOld[remove].ID)
                  .delete();
              }
            })
            .then(() => {
              this.webURL.lists
                .getByTitle("FeedbackQuestions")
                .items.filter("Team eq 'Technical'")
                .get()
                .then((Technicaldata2) => {
                  let Technicaldatalength = Technicaldata2.length;
                  for (let i = 0; i < Technicaldatalength; i++) {
                    let TitleHTML = $(
                      "#Question" + Technicaldata2[i].QuestionNumber
                    )[0].innerText;
                    let QuestionNo = $(
                      "#Question" + Technicaldata2[i].QuestionNumber
                    )[0].attributes[1].value;
                    let TAnswer = $(
                      "#Answer" + Technicaldata2[i].QuestionNumber
                    ).val();

                    let Techtextcommemntelement2 = '';
                    let matchvalue2 = "TechComment" + Technicaldata2[i].QuestionNumber;

                    for (let i = 0; i < commentdata.length; i++) {
                   
                      if (commentdata[i].QuestionNumber === matchvalue2) {
                        Techtextcommemntelement2 = commentdata[i].Comment;
                        break;  // Exit the loop once the matching item is found
                      }
                    }          
               
                   
                    let TechtxtComments2 = Techtextcommemntelement2 ? Techtextcommemntelement2 : '';
                   

                    this.webURL.lists
                      .getByTitle("TechnicalFeedback")
                      .items.add({
                        Title: TitleHTML,
                        ProjectID: itemID,
                        Department: Dpt,
                        QuestionNumber: QuestionNo,
                        Requirement: $(
                          "#Requirement" + Technicaldata2[i].QuestionNumber
                        )[0].innerHTML,
                        Answer: TAnswer,
                        Comment: TechtxtComments2,
                        FeedbackStatus: "Submit",
                      });
                  }
                });
            });
        }
      })
      .then(() => {
        let filterStr4 =
          "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
        this.webURL.lists
          .getByTitle("TechnicalComments")
          .items.filter(filterStr4)
          .getAll()
          .then((items) => {
            if (items.length > 0) {
              items.forEach((item4, index4) => {
                this.webURL.lists
                  .getByTitle("TechnicalComments")
                  .items.getById(item4.Id)
                  .update({
                    IsFinished: true,
                  })
                  .then((s) => {
                    let TechName = new URLSearchParams(
                      window.location.search
                    ).get("dpt");
                    let filterStr5 =
                      "ProjectID eq '" +
                      itemID +
                      "' and " +
                      "TaskType eq 'Technical'";
                    this.webURL.lists
                      .getByTitle("ProjectTasks")
                      .items.filter(filterStr5)
                      .getAll()
                      .then((itemz) => {
                        if (itemz.length > 0) {
                          itemz.forEach((item2, index2) => {
                            if (
                              TechName == item2.VIew.split("&")[1].split("=")[1]
                            ) {
                              this.webURL.lists
                                .getByTitle("ProjectTasks")
                                .items.getById(item2.Id)
                                .update({
                                  // TaskType: "Technical",
                                  // Action: "Technical Feedback",
                                  Status: "Completed",
                                  LastActionTakeBy:
                                    this.currentUserDetails.Title,
                                  LastActionTakenOn: new Date(),
                                });
                            }
                          });
                        }
                      })
                      .then((t) => {
                        let filterStr3 =
                          "ProjectID eq '" +
                          itemID +
                          "' and " +
                          "IsFinished eq '" +
                          false +
                          "' and " +
                          "Title eq 'PT'";
                        this.webURL.lists
                          .getByTitle("TechnicalComments")
                          .items.filter(filterStr3)
                          .get()
                          .then((data3) => {
                            if (data3.length == 0) {
                              let filterStr6 =
                                "ProjectID eq '" +
                                itemID +
                                "' and " +
                                "Title eq 'Delivery Action Awaited'";
                              this.webURL.lists
                                .getByTitle("ProjectTasks")
                                .items.filter(filterStr6)
                                .getAll()
                                .then((DelTask) => {
                                  if (
                                    this.ContractStatus ==
                                    "Technical Action Awaited"
                                  ) {
                                    this.webURL.lists
                                      .getByTitle("Projects")
                                      .items.getById(parseInt(itemID))
                                      .update({
                                        Status: `Delivery Action Awaited`,
                                        TechnicalDone: 1,
                                      })
                                      .then((newListItem) => {
                                        setTimeout(() => {
                                          $("#loader").hide();
                                        }, 5000);
                                        setTimeout(() => {
                                          $("#SubmittedModal").show();
                                        }, 3000);
                                      });
                                  } else {
                                    setTimeout(() => {
                                      $("#loader").hide();
                                    }, 5000);
                                    setTimeout(() => {
                                      $("#SubmittedModal").show();
                                    }, 3000);
                                  }
                                });
                            } else {
                              setTimeout(() => {
                                $("#loader").hide();
                              }, 5000);
                              setTimeout(() => {
                                $("#SubmittedModal").show();
                              }, 3000);
                            }
                          });
                      });
                  });
              });
            }
          });
      });
  }

  private TechDraftData() {
    let isvalid = true;
    // =========== Questions Validations ====================
    this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.filter("Team eq 'Technical'")
      .get()
      .then(async (Technicaldata) => {
        let Technicaldatalength = Technicaldata.length;
        for (let i = 0; i < Technicaldatalength; i++) {


          let CommentVal = "";
          let item = await this.state.RichtextData.find(Data => Data.QuestionNumber === "TechComment" + Technicaldata[i].QuestionNumber);

          if(item){
            CommentVal = item.Comment == null ? "" : item.Comment;
          }
          else{
            CommentVal = "";
          }

          let ID = Technicaldata[i].ID;
          let QuestionNumber = Technicaldata[i].QuestionNumber;
          let Answer = $("#Answer" + QuestionNumber).val();
          let Comment = CommentVal;
          let CommentError = $("#TechCommentErr" + QuestionNumber);

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "IfNoComment".toLowerCase()
          ) {
            if (Answer === "No" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "IfYesNoCommentYes".toLowerCase()
          ) {
            if (Answer === "No" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "IfYesNoCommentNo".toLowerCase()
          ) {
            if (Answer === "Yes" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "IfYesComment".toLowerCase()
          ) {
            if (Answer === "Yes" && Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "textbox".toLowerCase()
          ) {
            if (Comment === "") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }

          if (
            Technicaldata[i].Requirement.toLowerCase() ==
            "Dropdown".toLowerCase()
          ) {
            if (Answer === "Select") {
              CommentError.show();
              isvalid = false;
              window.stop();
            } else {
              CommentError.hide();
            }
          }
        }
      })
      .then((check) => {
        if (isvalid) {
          this.TechDraftData1();
        }
      });
  }

  private TechDraftData1() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    this.webURL.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .get()
      .then((Pdata) => {
        if (Pdata.TechnicalDone != 2) {
          this.TechDraftData2();
        } else {
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
    $("#TechnicalTaskForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2500);
  }

  private Close2() {
    let modal = document.getElementById("DraftModal");
    let hostUrl = this.props.webURL;
    modal.style.display = "none";
    $("#TechnicalTaskForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2500);
  }

  private TechDraftData2() {
    $("#loader").show();

    let commentdata = [];

    commentdata = this.state.RichtextData;

    const itemID = new URLSearchParams(window.location.search).get("itemid");
    const Dpt = new URLSearchParams(window.location.search).get("dpt");
    let filterStr =
      "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";

    this.webURL.lists
      .getByTitle("TechnicalFeedback")
      .items.filter(filterStr)
      .get()
      .then((feedbacks) => {
        if (feedbacks.length == 0) {
          this.webURL.lists
            .getByTitle("FeedbackQuestions")
            .items.filter("Team eq 'Technical'")
            .get()
            .then((Technicaldata) => {
              let Tdatalength = Technicaldata.length;
              for (let i = 0; i < Tdatalength; i++) {



                let TitleHTML = $(
                  "#Question" + Technicaldata[i].QuestionNumber
                )[0].innerText;
                let QuestionNo = $(
                  "#Question" + Technicaldata[i].QuestionNumber
                )[0].attributes[1].value;
                let TAnswer = $(
                  "#Answer" + Technicaldata[i].QuestionNumber
                ).val();
               
                let Techtextcommemntelement = '';
                let matchvalue = "TechComment" + Technicaldata[i].QuestionNumber;
                for (let i = 0; i < commentdata.length; i++) {
                  if (commentdata[i].QuestionNumber === matchvalue) {
                    Techtextcommemntelement = commentdata[i].Comment;
                    break;  // Exit the loop once the matching item is found
                  }
                }             
              
                let TechtxtComments = Techtextcommemntelement ? Techtextcommemntelement : '';
             

                this.webURL.lists.getByTitle("TechnicalFeedback").items.add({
                  Title: TitleHTML,
                  ProjectID: itemID,
                  Department: Dpt,
                  QuestionNumber: QuestionNo,
                  Answer: TAnswer,
                  Requirement: $(
                    "#Requirement" + Technicaldata[i].QuestionNumber
                  )[0].innerHTML,
                  Comment: TechtxtComments,
                  FeedbackStatus: "Draft",
                });
              }
            });
        } else {
          this.webURL.lists
            .getByTitle("TechnicalFeedback")
            .items.filter(filterStr)
            .get()
            .then((DeleteOld) => {
              for (let remove = 0; remove < DeleteOld.length; remove++) {
                this.webURL.lists
                  .getByTitle("TechnicalFeedback")
                  .items.getById(DeleteOld[remove].ID)
                  .delete();
              }
            })
            .then(() => {
              this.webURL.lists
                .getByTitle("FeedbackQuestions")
                .items.filter("Team eq 'Technical'")
                .get()
                .then((Technicaldata2) => {
                  let Technicaldatalength = Technicaldata2.length;
                  for (let i = 0; i < Technicaldatalength; i++) {
                    let TitleHTML = $(
                      "#Question" + Technicaldata2[i].QuestionNumber
                    )[0].innerText;
                    let QuestionNo = $(
                      "#Question" + Technicaldata2[i].QuestionNumber
                    )[0].attributes[1].value;
                    let TAnswer = $(
                      "#Answer" + Technicaldata2[i].QuestionNumber
                    ).val();
                
                    
                    let Techtextcommemntelement2 = '';
                    let matchvalue2 = "TechComment" + Technicaldata2[i].QuestionNumber;

                    for (let i = 0; i < commentdata.length; i++) {
                   
                      if (commentdata[i].QuestionNumber === matchvalue2) {
                        Techtextcommemntelement2 = commentdata[i].Comment;
                        break;  // Exit the loop once the matching item is found
                      }
                    }
                                  
                   
                    let TechtxtComments2 = Techtextcommemntelement2 ? Techtextcommemntelement2 : '';
                   


                    this.webURL.lists
                      .getByTitle("TechnicalFeedback")
                      .items.add({
                        Title: TitleHTML,
                        ProjectID: itemID,
                        Department: Dpt,
                        QuestionNumber: QuestionNo,
                        Requirement: $(
                          "#Requirement" + Technicaldata2[i].QuestionNumber
                        )[0].innerHTML,
                        Answer: TAnswer,
                        Comment: TechtxtComments2,
                        FeedbackStatus: "Draft",
                      });
                  }
                });
            });
        }
      })
      .then(() => {
        let TechName = new URLSearchParams(window.location.search).get("dpt");
        let filterStr4 =
          "ProjectID eq '" + itemID + "' and " + "Department eq '" + Dpt + "'";
        this.webURL.lists
          .getByTitle("TechnicalComments")
          .items.filter(filterStr4)
          .getAll()
          .then((items) => {
            if (items.length > 0) {
              items.forEach((item4, index4) => {
                this.webURL.lists
                  .getByTitle("TechnicalComments")
                  .items.getById(item4.Id)
                  .update({
                    IsFinished: false,
                  });
              });
            }
          });
      })
      .then((r) => {
        setTimeout(() => {
          $("#loader").hide();
        }, 3000);
        setTimeout(() => {
          $("#DraftModal").show();
        }, 3500);
      });
  }

  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.context.pageContext.web.absoluteUrl);
    return await web.currentUser.get();
  }
}
