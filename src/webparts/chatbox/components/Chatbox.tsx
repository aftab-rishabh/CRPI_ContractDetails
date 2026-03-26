import * as React from "react";
import styles from "./Chatbox.module.scss";
import { IChatboxProps } from "./IChatboxProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IChatboxState } from "./IChatboxState";
import { SPOperations } from "../Services/SPOps";

import {
  TextField,
  PrimaryButton,
  IconButton,
  Dialog,
  Label,
  DialogFooter,
  DefaultButton,
  DialogType,
} from "office-ui-fabric-react";

import * as $ from "jquery";

const modelProps = {
  isBlocking: false,
  styles: { main: { maxWidth: 450 } },
};
const dialogContentProps = {
  type: DialogType.largeHeader,
  title: "Confirm!",
  subText: "Are you sure you want to delete note?",
};

export default class Chatbox extends React.Component<
  IChatboxProps,
  IChatboxState,
  {}
> {
  private _spServices: SPOperations;

  constructor(props: IChatboxProps) {
    super(props);
    this._spServices = new SPOperations(this.props.context);

    this.state = {
      listResult: [],
      Status: "",
      EditedNote: "",
      EditedId: 0,
      ProjectId: parseInt(
        new URLSearchParams(window.location.search).get("itemid")
      ),
      isSalesPerson: false,
      isManagementPerson: false,
      hideDialog: true,
      DeletedId: 0,
    };
    this.onChange = this.onChange.bind(this);
    this.onUpdateConversation = this.onUpdateConversation.bind(this);
    this.onDeleteConversation = this.onDeleteConversation.bind(this);
  }

  private onChange(
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void {
    if ($.trim(newValue).length > 0) {
      this.setState({ EditedNote: newValue });
    } else if ($.trim(newValue).length == 0) {
      this.setState({ EditedNote: "" });
    }
  }

  private onDeleteConversation(): void {
    this.setState({
      hideDialog: true,
    });
    $("#loader").show();
    this._spServices.DeleteConversation(this.state.DeletedId).then((result) => {
      this.setState({ Status: result });
      this._spServices.getConversations(this.state.ProjectId).then((result) => {
        setTimeout(() => {
          this.setState({
            listResult: result,
            DeletedId: 0,
          });
          $("#loader").hide();
        }, 1000);
      });
    });
  }

  private onUpdateConversation(Id: number, note: string): void {
    this.setState({ EditedNote: note, EditedId: Id });
  }

  private Close0() {
    var modal = document.getElementById("ContractDeletedModal");
    var hostUrl = this.props.context.pageContext.web.absoluteUrl;
    modal.style.display = "none";
    $("#TestForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 1000);
  }

  public async componentDidMount() {
    let salesPerson: boolean = false;

    if (this.state.ProjectId > 0) {
      await this._spServices
        .isUserSales(
          this.state.ProjectId,
          this.props.context.pageContext.user.email
        )
        .then((isSales) => {
          this.setState({ isSalesPerson: isSales });
          this._spServices
            .isUserManagementTeam(this.props.context.pageContext.user.email)
            .then((isManagement) => {
              this.setState({ isManagementPerson: isManagement });
              this._spServices
                .getConversations(this.state.ProjectId)
                .then((result) => {
                  setTimeout(() => {
                    this.setState({ listResult: result });
                  }, 1000);
                });
            });
        });
    }
  }

  public render(): React.ReactElement<IChatboxProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    const projectId = this.state.ProjectId;
    const showNote = this.state.isSalesPerson || this.state.isManagementPerson;

    return (
      <section className={styles.chatbox}>
        {projectId > 0 && showNote ? (
          <>
            <div className={styles.chatbox}>
              <div className={styles.container}>
                <div className={styles.row}>
                  <div className={styles.sectionblock}>
                    <Label className={styles.headers}>
                      <u>Conversations</u>
                    </Label>
                  </div>
                  <div>
                    <TextField
                      value={this.state.EditedNote}
                      id="txtNote"
                      multiline={true}
                      onChange={this.onChange}
                    ></TextField>
                  </div>
                  <div>
                    <PrimaryButton
                      id="btnSave"
                      iconProps={{ iconName: "Save" }}
                      disabled={this.state.EditedNote.length == 0}
                      className={styles.buttonStyle}
                      text="Save"
                      onClick={() =>
                        this._spServices
                          .AddEditConversation(
                            this.state.EditedId,
                            this.state.EditedNote,
                            this.state.ProjectId,
                            this.state.isSalesPerson
                              ? "Sales"
                              : this.state.isManagementPerson
                              ? "Management"
                              : ""
                          )
                          .then((result) => {
                            //this.setState({ Status: result });
                            this._spServices
                              .getConversations(this.state.ProjectId)
                              .then((result) => {
                                setTimeout(() => {
                                  this.setState({
                                    listResult: result,
                                    EditedNote: "",
                                    EditedId: 0,
                                  });
                                }, 1000);
                              });
                          })
                      }
                    ></PrimaryButton>
                  </div>
                  <div className={styles.dataList}>
                    {this.state.listResult.map((item) => {
                      return (
                        <div className={styles.mainBox}>
                          <div className={styles.noteBox}>
                            <div className={styles.leftbox}>{item.Notes}</div>
                            <div className={styles.middlebox}>
                              {this.props.context.pageContext.user.email ==
                              item.AddedByEmail ? (
                                <IconButton
                                  iconProps={{ iconName: "Edit" }}
                                  title="Edit"
                                  onClick={(event) => {
                                    this.onUpdateConversation(
                                      item.ID,
                                      item.Notes
                                    );
                                  }}
                                />
                              ) : (
                                ""
                              )}
                            </div>
                            <div className={styles.rightbox}>
                              {this.props.context.pageContext.user.email ==
                              item.AddedByEmail
                                ? //HIDE Delete Button in future if requried then uncomment and below button
                                  // <IconButton
                                  //   iconProps={{ iconName: "Delete" }}
                                  //   title="Delete"
                                  //   onClick={() =>
                                  //     this.setState({
                                  //       hideDialog: false,
                                  //       DeletedId: item.ID,
                                  //     })
                                  //   }
                                  // />
                                  ""
                                : ""}
                            </div>
                          </div>
                          <div className={styles.addedByValue}>
                            Added by: <b>{item.AddedBy}</b> on {item.CreatedDt}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
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
              <div id="ContractDeletedModal" className={styles.modal}>
                <div className={styles.modalcontent}>
                  <span className={styles.close} onClick={() => this.Close0()}>
                    &times;
                  </span>
                  <label className={styles.header2}>Contract Deleted!</label>
                </div>
              </div>
              <Dialog
                hidden={this.state.hideDialog}
                // onDismiss={toggleHideDialog}
                dialogContentProps={dialogContentProps}
                modalProps={modelProps}
              >
                <DialogFooter>
                  <PrimaryButton
                    onClick={(event) => {
                      this.onDeleteConversation();
                    }}
                    text="OK"
                    style={{ borderRadius: "20px" }}
                  />
                  <DefaultButton
                    onClick={() => this.setState({ hideDialog: true })}
                    text="Cancel"
                    style={{ borderRadius: "20px" }}
                  />
                </DialogFooter>
              </Dialog>
            </div>
          </>
        ) : (
          <h1></h1>
        )}
      </section>
    );
  }
}
