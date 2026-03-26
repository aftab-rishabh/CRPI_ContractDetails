import * as React from "react";
import styles1 from "../../components/ContractApproval.module.scss";
import styles from "./KickOffDocTaskFormWebPart.module.scss";
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

export interface IKickOffDocUploadProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface IKickOffDocUploadState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class KickOffDocUpload extends React.Component<
  IKickOffDocUploadProps,
  IKickOffDocUploadState
> {
  [x: string]: any;
  
  ContractStatus = "";
  OpportunityID = "";
  webURL:any="";
  constructor(props: IKickOffDocUploadProps) {
    super(props);

    this.state = {
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
    };
  }
  public async componentDidMount() {
    this.webURL = Web(this.props.webURL);
    this.setForm();
  }
  public render(): React.ReactElement<IKickOffDocUploadProps> {
    return (
      <div className={styles.KickOffDocUpload}>
        <div className={styles.container}>
          <div className={styles.row} style={{ paddingTop: "0px" }}>
            <div
              id="oppID"
              style={{ display: "none" }}
              className={styles.pagesubtitle}
            ></div>

            <div id="kickoffdocs" style={{ display: "none" }}>
              <div className={styles.sectionblockPMAction}>
                <Label
                  className={styles.headersPMAction}
                  style={{ marginTop: "0px" }}
                >
                  <u>Kick-off Document Upload Form</u>
                </Label>
              </div>

              <div id="" className="row">
                <div className="col col-md-6 ${styles.form_group}">
                  <div className={styles.mb_1}>
                    <input
                      type="file"
                      className={styles.form_control}
                      id="uploadFile"
                    ></input>
                  </div>
                  <span
                    id="fileErr1"
                    style={{ display: "none" }}
                    className={styles.errorlable}
                  >
                    Please select file to upload.
                  </span>
                  <div className={styles.mb_1}>
                    <PrimaryButton
                      className="fileUpload-Button ${styles.btn} btn-primary"
                      onClick={() => this.uploadFileFromControl()}
                    >
                      <span>Upload</span>
                    </PrimaryButton>
                  </div>
                </div>
              </div>
            </div>

            <div className="row">
              <div className="col-lg-12">
                <div className={styles.file_list}>
                  <div>
                    <ul className="KickOfffileList" id="KickOfffileList"></ul>
                  </div>
                </div>
              </div>
            </div>
            <br />
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
              </div>{" "}
            </div>
            <div id="SubmittedModal" className={styles.modal}>
              <div className={styles.modalcontent}>
                <span className={styles.close} onClick={() => this.Close1()}>
                  &times;
                </span>
                <label className={styles.header2}>Project Initiated!</label>
              </div>
            </div>

            <div className="row">
              <div className="col-lg-12">
                <div className={styles.form_footer}>
                  <PrimaryButton
                    text="Initiate Project"
                    style={{ borderRadius: "10px", marginRight: "10px" }}
                    className={styles.btn}
                    id="Submit"
                    onClick={() => this.KickoffDocsUploded()}
                  ></PrimaryButton>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private KickoffDocsUploded(): void {
    let attachmentfiles: string = "";
  
    let hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
    // let hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    this.webURL
      .getFolderByServerRelativeUrl(hostUrl)
      .files.filter("Title eq 'Kick-off document'")
      .get()
      .then((files) => {
        if (files.length !== 0) {
          $("#fileErr1").hide();
          const itemID = new URLSearchParams(window.location.search).get(
            "itemid"
          );
          this.webURL.lists
            .getByTitle("Projects")
            .items.getById(parseInt(itemID))
            .update({
              Status: `Project Initiated`,
            })
            .then((i) => {
              this.webURL.lists
                .getByTitle("CompleteTask")
                .items.add({
                  ProjectID: parseInt(itemID),
                  TaskType: "PM",
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
        } else {
          $("#fileErr1").show();
          setTimeout(() => {
            $("#loader").hide();
          }, 1000);
        }
      });
  }

  private Close1() {
    let modal = document.getElementById("SubmittedModal");
    let hostUrl = this.props.webURL;
    modal.style.display = "none";
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 2500);
  }

  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    this.webURL.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select("Id", "Status", "OpportunityID")
      .get()
      .then((data) => {
        $("#oppID").text(data.OpportunityID);

        if (data.Status == "Kickoff Docs Awaited") {
          document.getElementById("kickoffdocs").style.display = "block";
        }
        this.getFiles1();
      });
  }

  private uploadFileFromControl() {
    $("#loader").show();

    let files = (document.getElementById("uploadFile") as HTMLInputElement)
      .files;
    let fileArr = [];
    if (files.length !== 0) {
      $("#fileErr1").hide();

      for (let g = 0; g < files.length; g++) {
        let file = files[g];
        fileArr.push(file);
       
        let hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
        // let hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
        this.webURL
          .getFolderByServerRelativeUrl(hostUrl)
          .files.add(file.name, file, true)
          .then((f) => {
            f.file.getItem().then((item) => {
              item
                .update({
                  Title: "Kick-off document",
                })
                .then((h) => {
                  $("#uploadFile").val("");
                  this.getFiles1();
                });
            });
          })
          .then((y) => {
            this.getFiles1();
          });
      }
    } else {
      $("#fileErr1").show();
      setTimeout(() => {
        $("#loader").hide();
      }, 1000);
    }
  }

  private getFiles1() {
    let attachmentfiles: string = "";
    // let hostUrl = this.context.pageContext.site.serverRelativeUrl + '/ProjectDocuments/' + $('#oppID').text();
    
    let hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    this.webURL
      .getFolderByServerRelativeUrl(hostUrl)
      .files.filter("Title eq 'Kick-off document'")
      .get()
      .then((files) => {
        if (files.length == 0) {
          $("#KickOfffileList").empty();
        } else {
          for (let i = 0; i < files.length; i++) {
            let title = files[i].Title;
            let valu = files[i].Name;
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
              document.getElementById("KickOfffileList") as HTMLInputElement
            ).innerHTML = attachmentfiles;
            // $("#KickOfffileList").append(attachmentfiles);
          }
        }
        this.DeleteFileEvent();
      });
    setTimeout(() => {
      $("#loader").hide();
    }, 1000);
  }

  private DeleteFileEvent() {
    $("#loader").show();

   
    let hostUrl = "/sites/demo/ProjectDocuments/" + $("#oppID").text();
    // let hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    this.webURL
      .getFolderByServerRelativeUrl(hostUrl)
      .files.filter("Title eq 'Kick-off document'")
      .get()
      .then((files) => {
        for (let i = 0; i < files.length; i++) {
          let btnClass = $("#CloseBtn")[i] as HTMLElement;
          btnClass.addEventListener("click", (f) => {
            let xyz = f.currentTarget as HTMLElement;
            let fName =
              xyz.nextElementSibling.nextElementSibling.nextElementSibling
                .textContent;
                this.webURL
              .getFolderByServerRelativeUrl(hostUrl)
              .files.getByName(fName)
              .delete()
              .then((g) => {})
              .then((check) => {
                this.getFiles1();
              });
          });
        }
      });
  }
}
