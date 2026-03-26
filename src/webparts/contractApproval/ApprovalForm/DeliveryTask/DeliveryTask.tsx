import * as React from "react";
import styles from "./DeliveryTaskFormWebPart.module.scss";
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
  "https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js",
);
/**
 * Icon styles. Feel free to change them
 */

export interface IDeliveryTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
  context: WebPartContext;
}

interface IQuestion {
  Sequence: string;
  Question: string;
  ProjectID: string;
  QuestionNumber: string;
  Answer: string;
  Comment: string;
}

export interface IDeliveryTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
  RichtextData: IQuestion[];
}

// declare global {
//   interface Window {
//     componentInstance: DeliveryTask; // Replace `MyComponent` with your actual component type
//   }
// }

export class DeliveryTask extends React.Component<
  IDeliveryTaskProps,
  IDeliveryTaskState
> {
  ContractStatus = "";
  webURL: any = "";
  currentUserDetails: any = null;

  constructor(props: IDeliveryTaskProps) {
    super(props);

    this.state = {
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
      RichtextData: [],
    };

    this.onTextChange = this.onTextChange.bind(this);
    this.setForm = this.setForm.bind(this);
    this.renderRichTextControl = this.renderRichTextControl.bind(this);
    this.DelSubmitData = this.DelSubmitData.bind(this);
    this.DelSubmitData1 = this.DelSubmitData1.bind(this);
    this.DelSubmitData2 = this.DelSubmitData2.bind(this);
    this.loadQuestions = this.loadQuestions.bind(this);
  }

  public async componentDidMount() {
    this.currentUserDetails = await this.spLoggedInUserDetails();
    this.webURL = Web(this.props.webURL);
    this.setForm();
    this.loadQuestions();

    // window.componentInstance = this;
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
                  <Label className={styles.headers}>
                    <u>
                      <Label id="TechTeamName"></Label>Delivery Task Form
                    </u>
                  </Label>
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

                <div id="LegalPendingModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span
                      className={styles.close}
                      onClick={() => this.Close4()}
                    >
                      &times;
                    </span>
                    <label className={styles.header2}>
                      Legal feedback is pending!
                    </label>
                  </div>
                </div>

                <div className="row">
                  <div className="col col-lg-12 {styles.form_footer}">
                    <PrimaryButton
                      text="Draft"
                      className={styles.btn}
                      style={{ borderRadius: "10px", marginRight: "10px" }}
                      id="DelDraft"
                      onClick={() => this.DelDraft()}
                    >
                      Draft
                    </PrimaryButton>
                    <PrimaryButton
                      text="Submit"
                      className={styles.btn}
                      style={{ borderRadius: "10px", marginRight: "10px" }}
                      id="DelSubmit"
                      onClick={() => this.DelSubmitData()}
                    ></PrimaryButton>
                    <PrimaryButton
                      text="Cancel"
                      className={styles.btn}
                      style={{ borderRadius: "10px" }}
                      id="DelCancel"
                      onClick={() => this.DelCancelForm()}
                    ></PrimaryButton>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public onTextChange(text: string, parentDivId: string) {
    this.setState((prevState) => {
      const updatedRichtextData = prevState.RichtextData.map((item) =>
        item.QuestionNumber === parentDivId
          ? {
              ...item,
              Comment:
                text
                  .replace(/<\/?[^>]+(>|$)/g, "")
                  .replace(/&nbsp;/g, "")
                  .trim().length === 0
                  ? ""
                  : text,
            } // Update the Comment field
          : item,
      );

      return { RichtextData: updatedRichtextData };
    });

    return text;
  }

  private renderRichTextControl(elementId: string) {
    const element = document.getElementById(elementId);

    const item = this.state.RichtextData.find(
      (item) => item.QuestionNumber === elementId,
    );

    const comment = item ? item.Comment : ""; // Returns the Comment if the item is found, otherwise an empty string

    if (element) {
      ReactDOM.render(
        <RichText
          className={`${elementId}`}
          value={comment}
          onChange={(text) => {
            let parentDivId = document
              .getElementById(elementId)
              .getAttribute("id");

            return this.onTextChange(text, parentDivId);
          }}
        />,
        element,
      );
    }
  }

  private async setForm() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    var filterStr = "ProjectID eq '" + itemID + "'";
    this.webURL.lists
      .getByTitle("DeliveryFeedback")
      .items.filter(filterStr)
      .get()
      .then((data) => {
        for (var b = 0; b < data.length; b++) {
          var DELNo = data[b].QuestionNumber;
          $("#Answer" + DELNo).val(data[b].Answer);

          if (
            data[b].Requirement.toLowerCase() ==
              "IfYesNoCommentYes".toLowerCase() ||
            data[b].Requirement.toLowerCase() ==
              "IfYesNoCommentNo".toLowerCase()
          ) {
          } else {
            if (data[b].Comment !== null) {
              $("#DelComment" + DELNo).show();
              $("#DelComment" + DELNo).val(data[b].Comment);
            } else {
              $("#DelComment" + DELNo).hide();
            }
          }
        }
      })
      .then(async () => {
        var filterStr4 =
          "ProjectID eq '" + itemID + "' and " + "FeedbackStatus eq 'Submit'";
        this.webURL.lists
          .getByTitle("DeliveryFeedback")
          .items.filter(filterStr4)
          .getAll()
          .then((items) => {
            if (items.length !== 0) {
              $("#DelDraft").hide();
            } else {
              $("#DelDraft").show();
            }
          });
      });
  }

  private async loadQuestions() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let temp = false;

    let filterStr = "ProjectID eq '" + itemID + "'";

    await this.webURL.lists
      .getByTitle("DeliveryFeedback")
      .items.filter(filterStr)
      .get()
      .then((data) => {
        if (data.length > 0) {
          temp = true;
        }
        for (let b = 0; b < data.length; b++) {
          let CommentVal = "";

          if (
            data[b].Comment === null ||
            data[b].Comment?.replace(/<\/?[^>]+(>|$)/g, "")
              .replace(/&nbsp;/g, "")
              .trim() === ""
          ) {
            CommentVal = "";
          } else {
            CommentVal = `${data[b].Comment}`;
          }

          this.setState((prevState) => ({
            RichtextData: [
              ...prevState.RichtextData,
              {
                Sequence: ``,
                Question: `${data[b].Title}`,
                ProjectID: `${itemID}`,
                QuestionNumber: `DelComment${data[b].QuestionNumber}`,
                Answer: `${data[b].Answer}`,
                Comment: CommentVal,
              },
            ],
          }));
        }
      });

    await this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.orderBy("QuestionNumber")
      .filter("Team eq 'Delivery'")
      .orderBy("Sequence")
      .get()
      .then((item) => {
        var QueLength = item.length;
        for (var s = 0; s < QueLength; s++) {
          if (temp) {
          } else {
            this.setState((prevState) => ({
              RichtextData: [
                ...prevState.RichtextData,
                {
                  Sequence: `${item[s].Sequence}`,
                  Question: `${item[s].Question}`,
                  ProjectID: `${itemID}`,
                  QuestionNumber: `DelComment${item[s].QuestionNumber}`,
                  Answer: ``,
                  Comment: ``,
                },
              ],
            }));
          }

          var QuestionNumber = item[s].QuestionNumber;
          var lablecontrol = styles.lablecontrol;
          var Styles = styles.form_control;
          var errorlable = styles.errorlable;
          var Questions = "";

          // Logic for IfYesNoComment starts

          if (
            item[s].Requirement.toLowerCase() ==
              "IfYesNoCommentYes".toLowerCase() ||
            item[s].Requirement.toLowerCase() ==
              "IfYesNoCommentNo".toLowerCase()
          ) {
            Questions +=
              '<label class="' +
              lablecontrol +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              '</label><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<div id="DelComment' +
              QuestionNumber +
              '" class="' +
              styles.richTextContainer +
              '" ></div>' +
              '<span id="DelCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>';
            $("#Questions").append(Questions);
            this.renderRichTextControl(`DelComment${QuestionNumber}`);
          }

          // Logic for IfYesNoComment Ends

          if (
            item[s].Requirement.toLowerCase() == "IfNoComment".toLowerCase()
          ) {
            Questions +=
              '<label class="' +
              lablecontrol +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              '</label><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<div id="DelComment' +
              QuestionNumber +
              '" class="' +
              styles.richTextContainer +
              '" style="display: none;"></div>' +
              '<span id="DelCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>';
            $("#Questions").append(Questions);
            this.renderRichTextControl(`DelComment${QuestionNumber}`);
          }

          if (
            item[s].Requirement.toLowerCase() == "IfYesComment".toLowerCase()
          ) {
            Questions +=
              '<label class="' +
              lablecontrol +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              '</label><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px;" id="Answer' +
              QuestionNumber +
              '"><option>Select</option><option>Yes</option><option>No</option><option>NA</option></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<div id="DelComment' +
              QuestionNumber +
              '" class="' +
              styles.richTextContainer +
              '" style="display: none;"></div>' +
              '<span id="DelCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>';
            $("#Questions").append(Questions);
            this.renderRichTextControl(`DelComment${QuestionNumber}`);
          }

          if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            Questions +=
              '<label class="' +
              lablecontrol +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              '</label><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<div id="DelComment' +
              QuestionNumber +
              '" class="' +
              styles.richTextContainer +
              '" style="display: block;"></div>' +
              '<span id="DelCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>';
            $("#Questions").append(Questions);
            this.renderRichTextControl(`DelComment${QuestionNumber}`);
          }

          if (item[s].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
            var Options = "";
            Questions +=
              '<label class="' +
              lablecontrol +
              '" id="Question' +
              QuestionNumber +
              '">' +
              item[s].Question +
              '</label><span id="Requirement' +
              QuestionNumber +
              '"style="display:none">' +
              item[s].Requirement +
              "</span>" +
              '<select name="Answer' +
              QuestionNumber +
              '" class="' +
              Styles +
              '" style = "width: 200px;" id="Answer' +
              QuestionNumber +
              '"></select>' +
              '<span id="CRQuestionErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">Please select the value from the drop-down.</span><br>' +
              '<span id="DelCommentErr' +
              QuestionNumber +
              '" style="display:none;" class="' +
              errorlable +
              '" for="Author">This field is required.</span><br>';
            $("#Questions").append(Questions);
            for (
              var w = 0;
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
            item[s].Requirement.toLowerCase() ==
            "IfYesNoCommentYes".toLowerCase()
          ) {
            this.IfYesNoCommentChange(QuestionNumber);
          }

          if (
            item[s].Requirement.toLowerCase() ==
            "IfYesNoCommentNo".toLowerCase()
          ) {
            this.IfYesNoCommentChange2(QuestionNumber);
          }

          if (item[s].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            $("#Question" + QuestionNumber).append(
              '<lable class="" style="color:red;">*</lable>',
            );
          }
        }
      })
      .then(() => {
        this.setForm();
      });

    console.log("Data after load question: ", this.state.RichtextData);
  }

  private IfNoCommentChange(QuestionNumber) {
    var ChangeFunction = document.getElementById("Answer" + QuestionNumber);
    ChangeFunction.addEventListener("change", () =>
      this.ShowCommentBoxOnNo(QuestionNumber),
    );
  }

  //************New changes*******************************************************/

  private IfYesNoCommentChange(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnYesNo(QuestionNumber),
      );
  }

  private ShowCommentBoxOnYesNo(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("");
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("Not Applicable");
    } else {
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("");
    }
  }

  private IfYesNoCommentChange2(QuestionNumber) {
    document
      .getElementById("Answer" + QuestionNumber)
      .addEventListener("change", () =>
        this.ShowCommentBoxOnYesNo2(QuestionNumber),
      );
  }

  private ShowCommentBoxOnYesNo2(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("");
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("Not Applicable");
    } else {
      $("#DelComment" + QuestionNumber)
        .find(".ql-editor")
        .html("");
    }
  }

  //*******************************************************************/

  private ShowCommentBoxOnNo(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "No") {
      $("#DelComment" + QuestionNumber).show();
    } else {
      $("#DelComment" + QuestionNumber).val("");
      $("#DelComment" + QuestionNumber).hide();
      $("#DelCommentErr" + QuestionNumber).hide();
    }
  }

  private IfYesCommentChange(QuestionNumber) {
    var ChangeFunction = document.getElementById("Answer" + QuestionNumber);
    ChangeFunction.addEventListener("change", () =>
      this.ShowCommentBoxOnYes(QuestionNumber),
    );
  }

  private ShowCommentBoxOnYes(QuestionNumber) {
    if ($("#Answer" + QuestionNumber).val() == "Yes") {
      $("#DelComment" + QuestionNumber).show();
    } else {
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
    $("#DeliveryTaskForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2500);
  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $("#DeliveryTaskForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2500);
  }
  private Close4() {
    var modal = document.getElementById("LegalPendingModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    // $('#DeliveryTaskForm').fadeOut(2500);
    // setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
  }

  private DelCancelForm() {
    var hostUrl = this.props.webURL;
    window.parent.location.href = hostUrl;
  }

  private async DelSubmitData() {
    var isvalid = true;
    // =========== Questions Validations ====================

    this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.filter("Team eq 'Delivery'")
      .get()
      .then(async (DelData) => {
        var DelDatalength = DelData.length;

        for (var i = 0; i < DelDatalength; i++) {
          let CommentVal = "";
          let item = await this.state.RichtextData.find(
            (Data) =>
              Data.QuestionNumber === "DelComment" + DelData[i].QuestionNumber,
          );

          if (item) {
            CommentVal = item == undefined ? "" : item.Comment;
          } else {
            CommentVal = "";
          }

          var ID = DelData[i].QuestionNumber;
          var QuestionNumber = DelData[i].QuestionNumber;
          var Answer = $("#Answer" + [ID]).val();
          var DelComment = CommentVal;
          var DelCommentErr = $("#DelCommentErr" + [ID]);
          var CRQuestionErr = $("#CRQuestionErr" + [ID]);

          if (Answer !== "Select") {
            CRQuestionErr.hide();
            if (
              DelData[i].Requirement.toLowerCase() ==
              "IfNoComment".toLowerCase()
            ) {
              if (Answer === "No" && DelComment === "") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                //    item.Comment = '';
                DelCommentErr.hide();
              }
            }

            if (
              DelData[i].Requirement.toLowerCase() ==
              "IfYesNoCommentYes".toLowerCase()
            ) {
              if (Answer === "No" && DelComment === "") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                DelCommentErr.hide();
              }
            }

            if (
              DelData[i].Requirement.toLowerCase() ==
              "IfYesNoCommentNo".toLowerCase()
            ) {
              if (Answer === "Yes" && DelComment === "") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                DelCommentErr.hide();
              }
            }

            if (
              DelData[i].Requirement.toLowerCase() ==
              "IfYesComment".toLowerCase()
            ) {
              if (Answer === "Yes" && DelComment === "") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                //   item.Comment = '';
                DelCommentErr.hide();
              }
            }

            if (
              DelData[i].Requirement.toLowerCase() == "textbox".toLowerCase()
            ) {
              if (DelComment === "") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                DelCommentErr.hide();
              }
            }

            if (
              DelData[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()
            ) {
              if (Answer === "Select") {
                DelCommentErr.show();
                isvalid = false;
                window.stop();
              } else {
                DelCommentErr.hide();
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
          this.DelSubmitData1();
        }
      });
  }

  private async DelSubmitData1() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    this.webURL.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select("Id", "Status", "DeliveryDone", "LegalDone")
      .get()
      .then((Pdata) => {
        if (Pdata.DeliveryDone != 2) {
          // if (
          //   Pdata.LegalDone !== -1 &&
          //   Pdata.Status == "Delivery Action Awaited" &&
          //   Pdata.Status !== "Management Action Awaited" &&
          //   Pdata.Status !== "Escalated to management"
          // )
          if (
            Pdata.LegalDone !== -1 &&
            (Pdata.Status == "Delivery Action Awaited" ||
              Pdata.Status == "Management Action Awaited" ||
              Pdata.Status == "Escalated to management")
          ) {
            this.DelSubmitData2();
          } else {
            var modal = document.getElementById("LegalPendingModal");
            modal.style.display = "block";
          }
        } else {
          var modal = document.getElementById("ApprovedModal");
          modal.style.display = "block";
        }
      });
  }

  private async DelSubmitData2() {
    console.log("DelSubmitData2 start: ", this.state.RichtextData);
    let commentdata = [];

    commentdata = this.state.RichtextData;

    const itemID = new URLSearchParams(window.location.search).get("itemid");
    var filterStr = "ProjectID eq '" + itemID + "'";
    this.webURL.lists
      .getByTitle("DeliveryFeedback")
      .items.filter(filterStr)
      .get()
      .then(async (feedbacks) => {
        if (feedbacks.length == 0) {
          this.webURL.lists
            .getByTitle("FeedbackQuestions")
            .items.filter("Team eq 'Delivery'")
            .get()
            .then(async (DelData) => {
              var DelDatalength = DelData.length;
              for (var i = 0; i < DelDatalength; i++) {
                var ID = DelData[i].QuestionNumber;
                var Title = $("#Question" + [ID])[0].innerHTML;
                var Answer = $("#Answer" + [ID]).val();
                var Qnumber = DelData[i].QuestionNumber;

                let Deltextcommemntelement = "";
                let matchvalue = "DelComment" + DelData[i].QuestionNumber;
                for (let i = 0; i < commentdata.length; i++) {
                  if (commentdata[i].QuestionNumber === matchvalue) {
                    Deltextcommemntelement = commentdata[i].Comment;
                    break; // Exit the loop once the matching item is found
                  }
                }

                let DelComment = Deltextcommemntelement
                  ? Deltextcommemntelement
                  : "";

                if (Answer !== undefined && DelComment !== undefined) {
                  this.webURL.lists.getByTitle("DeliveryFeedback").items.add({
                    Title: Title,
                    ProjectID: itemID,
                    QuestionNumber: Qnumber,
                    Answer: Answer,
                    Requirement: $("#Requirement" + ID)[0].innerHTML,
                    Comment: DelComment,
                    FeedbackStatus: "Submit",
                  });
                } else {
                }
              }
            })
            .then(async () => {
              this.webURL.lists
                .getByTitle("Projects")
                .items.getById(parseInt(itemID))
                .select("Id", "Status", "DeliveryDone", "LegalDone")
                .get()
                .then(async (data) => {
                  status = data.Status;
                  this.webURL.lists
                    .getByTitle("Projects")
                    .items.getById(parseInt(itemID))
                    .update({
                      DeliveryDone: 1,
                    })
                    .then(async () => {
                      // if (data.LegalDone !== -1 && data.Status == "Delivery Action Awaited" && data.Status !== "Management Action Awaited" && data.Status !== "Escalated to management") {
                      //PCR:18.11.2022: Added below if condition, if project status is "Escalated to management" so the task for management
                      //                are already been created, so new task should not be created.
                      if (data.Status == "Delivery Action Awaited") {
                        this.webURL.lists
                          .getByTitle("Projects")
                          .items.getById(parseInt(itemID))
                          .update({
                            Status: "Management Action Awaited",
                          });
                      }
                    });
                });
            })
            .then(async (f) => {
              var filterStr5 =
                "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Delivery'";
              this.webURL.lists
                .getByTitle("ProjectTasks")
                .items.filter(filterStr5)
                .getAll()
                .then((items) => {
                  if (items.length > 0) {
                    items.forEach(async (item2) => {
                      this.webURL.lists
                        .getByTitle("ProjectTasks")
                        .items.getById(item2.Id)
                        .update({
                          // TaskType: "Management/Marketing",
                          // Action: "Approve/ Reject"
                          Status: "Completed",
                          LastActionTakeBy: this.currentUserDetails.Title,
                          LastActionTakenOn: new Date(),
                        });
                    });
                  }
                });
            })
            .then(async () => {
              var filterStr5 =
                "ProjectID eq '" +
                itemID +
                "' and " +
                "TaskType eq 'Technical'";
              this.webURL.lists
                .getByTitle("ProjectTasks")
                .items.filter(filterStr5)
                .getAll()
                .then((items) => {
                  if (items.length > 0) {
                    items.forEach(async (item2) => {
                      this.webURL.lists
                        .getByTitle("ProjectTasks")
                        .items.getById(item2.Id)
                        .update({
                          // TaskType: "Management/Marketing",
                          // Action: "Approve/ Reject"
                          Status: "Completed",
                        });
                    });
                  }
                });
            })
            .then((r) => {
              $("#loader").show();
              setTimeout(() => {
                $("#loader").hide();
              }, 3000);
              setTimeout(() => {
                $("#SubmittedModal").show();
              }, 3500);
            });
        } else {
          this.webURL.lists
            .getByTitle("DeliveryFeedback")
            .items.filter(filterStr)
            .get()
            .then(async (DeleteOld) => {
              for (var remove = 0; remove < DeleteOld.length; remove++) {
                this.webURL.lists
                  .getByTitle("DeliveryFeedback")
                  .items.getById(DeleteOld[remove].ID)
                  .delete();
              }
            })
            .then(() => {
              this.webURL
                .getByTitle("FeedbackQuestions")
                .items.filter("Team eq 'Delivery'")
                .get()
                .then((DelData2) => {
                  var Deliverydatalength = DelData2.length;
                  for (var i = 0; i < Deliverydatalength; i++) {
                    var ID = DelData2[i].QuestionNumber;
                    var Title = $("#Question" + [ID])[0].innerHTML;
                    var Answer = $("#Answer" + [ID]).val();
                    var Qnumber = DelData2[i].QuestionNumber;
                    let web = Web(this.props.webURL);

                    let Deltextcommemntelement = "";
                    let matchvalue = "DelComment" + DelData2[i].QuestionNumber;
                    for (let i = 0; i < commentdata.length; i++) {
                      if (commentdata[i].QuestionNumber === matchvalue) {
                        Deltextcommemntelement = commentdata[i].Comment;
                        break; // Exit the loop once the matching item is found
                      }
                    }

                    let DelComment = Deltextcommemntelement
                      ? Deltextcommemntelement
                      : "";

                    web.lists.getByTitle("DeliveryFeedback").items.add({
                      Title: Title,
                      ProjectID: itemID,
                      QuestionNumber: Qnumber,
                      Answer: Answer,
                      Requirement: $("#Requirement" + ID)[0].innerHTML,
                      Comment: DelComment,
                      FeedbackStatus: "Submit",
                    });
                  }
                });
            })
            .then((f) => {
              var filterStr5 =
                "ProjectID eq '" + itemID + "' and " + "TaskType eq 'Delivery'";
              this.webURL.lists
                .getByTitle("ProjectTasks")
                .items.filter(filterStr5)
                .getAll()
                .then((items) => {
                  if (items.length > 0) {
                    items.forEach((item2, index2) => {
                      let web = Web(this.props.webURL);
                      web.lists
                        .getByTitle("ProjectTasks")
                        .items.getById(item2.Id)
                        .update({
                          // TaskType: "Management/Marketing",
                          // Action: "Approve/ Reject",
                          Status: "Completed",
                          LastActionTakeBy: this.currentUserDetails.Title,
                          LastActionTakenOn: new Date(),
                        });
                    });
                  }
                });
            })
            .then(() => {
              var filterStr5 =
                "ProjectID eq '" +
                itemID +
                "' and " +
                "TaskType eq 'Technical'";
              this.webURL.lists
                .getByTitle("ProjectTasks")
                .items.filter(filterStr5)
                .getAll()
                .then((items) => {
                  if (items.length > 0) {
                    items.forEach((item2, index2) => {
                      let web = Web(this.props.webURL);
                      web.lists
                        .getByTitle("ProjectTasks")
                        .items.getById(item2.Id)
                        .update({
                          // TaskType: "Management/Marketing",
                          // Action: "Approve/ Reject",
                          Status: "Completed",
                          LastActionTakeBy: "System",
                          LastActionTakenOn: new Date(),
                        });
                    });
                  }
                });
            })
            .then((n) => {
              this.webURL.lists
                .getByTitle("Projects")
                .items.getById(parseInt(itemID))
                .select("Id", "Status", "DeliveryDone", "LegalDone")
                .get()
                .then((data) => {
                  status = data.Status;
                  this.webURL.lists
                    .getByTitle("Projects")
                    .items.getById(parseInt(itemID))
                    .update({
                      DeliveryDone: 1,
                    })
                    .then((changestatus) => {
                      // if (data.LegalDone !== -1 && data.Status == "Delivery Action Awaited" && data.Status !== "Management Action Awaited" && data.Status !== "Escalated to management") {
                      this.webURL.lists
                        .getByTitle("Projects")
                        .items.getById(parseInt(itemID))
                        .update({
                          Status: "Management Action Awaited",
                        });
                      // }
                    });
                });
            })
            .then((r) => {
              $("#loader").show();
              setTimeout(() => {
                $("#loader").hide();
              }, 3000);
              setTimeout(() => {
                $("#SubmittedModal").show();
              }, 3500);
            });
        }
      });
  }

  private async DelDraft() {
    var isvalid = true;
    // =========== Questions Validations ====================

    this.webURL.lists
      .getByTitle("FeedbackQuestions")
      .items.filter("Team eq 'Delivery'")
      .get()
      .then(async (DelData) => {
        var DelDatalength = DelData.length;

        for (var i = 0; i < DelDatalength; i++) {
          let CommentVal = "";
          let item = await this.state.RichtextData.find(
            (Data) =>
              Data.QuestionNumber === "DelComment" + DelData[i].QuestionNumber,
          );

          if (item) {
            CommentVal = item == undefined ? "" : item.Comment;
          } else {
            CommentVal = "";
          }

          var ID = DelData[i].QuestionNumber;
          var Answer = $("#Answer" + [ID]).val();
          var DelComment = CommentVal;
          var DelCommentErr = $("#DelCommentErr" + [ID]);

          if (
            DelData[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase()
          ) {
            if (Answer === "No" && DelComment === "") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              //  item.Comment = '';
              DelCommentErr.hide();
            }
          }

          if (
            DelData[i].Requirement.toLowerCase() ==
            "IfYesNoCommentYes".toLowerCase()
          ) {
            if (Answer === "No" && DelComment === "") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              DelCommentErr.hide();
            }
          }

          if (
            DelData[i].Requirement.toLowerCase() ==
            "IfYesNoCommentNo".toLowerCase()
          ) {
            if (Answer === "Yes" && DelComment === "") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              DelCommentErr.hide();
            }
          }

          if (
            DelData[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase()
          ) {
            if (Answer === "Yes" && DelComment === "") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              // item.Comment = '';
              DelCommentErr.hide();
            }
          }

          if (DelData[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            if (DelComment === "") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              DelCommentErr.hide();
            }
          }

          if (
            DelData[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()
          ) {
            if (Answer === "Select") {
              DelCommentErr.show();
              isvalid = false;
              window.stop();
            } else {
              DelCommentErr.hide();
            }
          }
        }
      })
      .then(() => {
        if (isvalid) {
          this.DelDraft1();
        }
      });
  }

  private async DelDraft1() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    this.webURL.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .get()
      .then((Pdata) => {
        if (Pdata.DeliveryDone != 2) {
          this.DelDraft2();
        } else {
          var modal = document.getElementById("ApprovedModal");
          modal.style.display = "block";
        }
      });
  }

  private async DelDraft2() {
    let commentdata = [];

    commentdata = this.state.RichtextData;

    const itemID = new URLSearchParams(window.location.search).get("itemid");
    var filterStr = "ProjectID eq '" + itemID + "'";
    this.webURL.lists
      .getByTitle("DeliveryFeedback")
      .items.filter(filterStr)
      .get()
      .then(async (feedbacks) => {
        if (feedbacks.length == 0) {
          this.webURL.lists
            .getByTitle("FeedbackQuestions")
            .items.filter("Team eq 'Delivery'")
            .get()
            .then(async (DelData) => {
              var DelDatalength = DelData.length;
              for (var i = 0; i < DelDatalength; i++) {
                var ID = DelData[i].QuestionNumber;
                var Title = $("#Question" + [ID])[0].innerHTML;
                var Answer = $("#Answer" + [ID]).val();
                var Qnumber = DelData[i].QuestionNumber;

                let Deltextcommemntelement = "";
                let matchvalue = "DelComment" + DelData[i].QuestionNumber;
                for (let i = 0; i < commentdata.length; i++) {
                  if (commentdata[i].QuestionNumber === matchvalue) {
                    Deltextcommemntelement = commentdata[i].Comment;
                    break; // Exit the loop once the matching item is found
                  }
                }

                let DelComment = Deltextcommemntelement
                  ? Deltextcommemntelement
                  : "";

                this.webURL.lists.getByTitle("DeliveryFeedback").items.add({
                  Title: Title,
                  ProjectID: itemID,
                  QuestionNumber: Qnumber,
                  Answer: Answer,
                  Requirement: $("#Requirement" + ID)[0].innerHTML,
                  Comment: DelComment,
                  FeedbackStatus: "Draft",
                });
              }
            })
            .then((r) => {
              $("#loader").show();
              setTimeout(() => {
                $("#loader").hide();
              }, 3000);
              setTimeout(() => {
                $("#DraftModal").show();
              }, 3500);
            });
        } else {
          this.webURL.lists
            .getByTitle("DeliveryFeedback")
            .items.filter(filterStr)
            .get()
            .then(async (DeleteOld) => {
              for (var remove = 0; remove < DeleteOld.length; remove++) {
                this.webURL.lists
                  .getByTitle("DeliveryFeedback")
                  .items.getById(DeleteOld[remove].ID)
                  .delete();
              }
            })
            .then(async () => {
              this.webURL.lists
                .getByTitle("FeedbackQuestions")
                .items.filter("Team eq 'Delivery'")
                .get()
                .then(async (DelData2) => {
                  var Deliverydatalength = DelData2.length;
                  for (var i = 0; i < Deliverydatalength; i++) {
                    var ID = DelData2[i].QuestionNumber;
                    var Title = $("#Question" + [ID])[0].innerHTML;
                    var Answer = $("#Answer" + [ID]).val();
                    var Qnumber = DelData2[i].QuestionNumber;

                    let Deltextcommemntelement = "";
                    let matchvalue = "DelComment" + DelData2[i].QuestionNumber;
                    for (let i = 0; i < commentdata.length; i++) {
                      if (commentdata[i].QuestionNumber === matchvalue) {
                        Deltextcommemntelement = commentdata[i].Comment;
                        break; // Exit the loop once the matching item is found
                      }
                    }

                    let DelComment = Deltextcommemntelement
                      ? Deltextcommemntelement
                      : "";

                    this.webURL.lists.getByTitle("DeliveryFeedback").items.add({
                      Title: Title,
                      ProjectID: itemID,
                      QuestionNumber: Qnumber,
                      Answer: Answer,
                      Requirement: $("#Requirement" + ID)[0].innerHTML,
                      Comment: DelComment,
                      FeedbackStatus: "Draft",
                    });
                  }
                });
            })
            .then((r) => {
              $("#loader").show();
              setTimeout(() => {
                $("#loader").hide();
              }, 3000);
              setTimeout(() => {
                $("#DraftModal").show();
              }, 3500);
            });
        }
      });
  }

  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.context.pageContext.web.absoluteUrl);
    return await web.currentUser.get();
  }
}
