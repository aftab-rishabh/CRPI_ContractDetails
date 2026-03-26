import * as React from "react";
import styles from "./ContractDetailsNew.module.scss";
import styles1 from "../../accordion/Accordion.module.scss";
import { css } from "@uifabric/utilities/lib/css";
import { IContractDetailsNewProps } from "./IContractDetailsNewProps";
import { escape } from "@microsoft/sp-lodash-subset";
// import { Accordion } from '../../accordion';
import {
  Nav,
  INavLink,
  INavStyles,
  INavLinkGroup,
} from "@fluentui/react/lib/Nav";
// import SplitPane, { Pane } from 'react-split-pane';
import { Technical } from "../Feedback/Technical";
import { Delivery } from "../Feedback/Delivery";
import { Legal } from "../Feedback/Legal";
import { Sales } from "../Feedback/Sales";
import {
  DefaultButton,
  IIconProps,
  Link,
  PrimaryButton,
} from "office-ui-fabric-react";
import { Management } from "../Feedback/Management";
import { AllTask } from "../Feedback/AllTask";
import { ProjectDetails } from "../Feedback/ProjectDetails";
import { Web } from "@pnp/sp/webs";
import Chatbox from "../../chatbox/components/Chatbox";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import * as $ from "jquery";

const collapsedIcon: IIconProps = {
  iconName: "ChevronRight",
  className: styles.accordionChevron,
};
const expandedIcon: IIconProps = {
  iconName: "ChevronDown",
  className: styles.accordionChevron,
};
export interface IContractDetailsNewState {
  Sales: boolean;
  Technical: boolean;
  Delivery: boolean;
  Legal: boolean;
  Management: boolean;
  expanded: boolean;
  AllTask: boolean;
  ChatBox: boolean;
  IsSales: boolean;
  IsManagement: boolean;
  hideDialog: boolean;
  ShowEditContractBtn: boolean;
  ShowPDFBtn: boolean;
}

const modelProps = {
  isBlocking: false,
  styles: { main: { maxWidth: 450 } },
};
const dialogContentProps = {
  type: DialogType.largeHeader,
  title: "Confirm!",
  subText:
    "Once you delete this contract all the related details permanant deleted from database",
};

const navLinkGroups: INavLinkGroup[] = [
  {
    links: [
      {
        name: "Sales Feedback",
        url: "",
        key: "key1",
        isExpanded: true,
        target: "_blank",
      },
      {
        name: "Legal Feedback",
        url: "",
        key: "key1",
        disabled: true,
      },
      {
        name: "Technical Feedback",
        url: "",
        key: "key3",
        target: "_blank",
      },

      {
        name: "Delivery Feedback",
        url: "",
        key: "key4",
        target: "_blank",
      },
    ],
  },
];

export default class ContractDetailsNew extends React.Component<
  IContractDetailsNewProps,
  IContractDetailsNewState
> {
  private _drawerDiv: HTMLDivElement = undefined;
  currentuseremail: any;
  constructor(props: IContractDetailsNewProps) {
    super(props);

    this.state = {
      Sales: false,
      Technical: false,
      Delivery: false,
      Legal: false,
      Management: false,
      AllTask: true,
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      ChatBox: false,
      IsSales: false,
      IsManagement: false,
      hideDialog: true,
      ShowEditContractBtn: false,
      ShowPDFBtn: false,
    };
  }
  public async componentDidMount() {
    this.setData();
  }

  public render(): React.ReactElement<IContractDetailsNewProps> {
    return (
      <div className={styles.contractDetailsNew}>
        <div className={styles.viewdetail_block}>
          <div className={styles.sectionblock}>
            <div className={styles.divTextRight}>
              {this.state.ShowPDFBtn && (
                <PrimaryButton
                  iconProps={{ iconName: "PDF" }}
                  text="Generate PDF"
                  onClick={() => this.ExportToPDF()}
                  hidden={false}
                  className={styles.buttonPDF}
                >
                  {" "}
                </PrimaryButton>
              )}
              {this.state.IsSales && (
                <PrimaryButton
                  iconProps={{ iconName: "Delete" }}
                  text="Delete Contract"
                  onClick={() => this.setState({ hideDialog: false })}
                  className={styles.buttonDelete}
                >
                  {" "}
                </PrimaryButton>
              )}
            </div>
          </div>
        </div>
        <div className={styles.viewdetail_block}>
          <div className={styles.sectionblock}>
            <div className={styles.section_left}>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.expanded}
                    text={"Project Details"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    iconProps={
                      this.state.expanded ? expandedIcon : collapsedIcon
                    }
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        AllTask: true,
                        Sales: false,
                        Technical: false,
                        Delivery: false,
                        Legal: false,
                        Management: false,
                        ChatBox: false,
                      });
                    }}
                    aria-expanded={this.state.expanded}
                    aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                  {this.state.IsSales && this.state.ShowEditContractBtn && (
                    <Link
                      href={
                        this.props.webURL +
                        "/SitePages/UpdateContract.aspx?itemid=" +
                        new URLSearchParams(window.location.search).get(
                          "itemid"
                        )
                      }
                      target="_blank"
                      data-interception="off"
                      style={{
                        position: "absolute",
                        marginTop: "0.6%",
                        marginLeft: "-12%",
                      }}
                    >
                      Edit contract
                    </Link>
                  )}
                  {this.state.expanded && (
                    <div
                      className={styles1.drawer}
                      ref={(el) => {
                        this._drawerDiv = el;
                      }}
                    >
                      {/* {this.state.HTML} */}
                      {/* <AllTask webURL={this.props.webURL}></AllTask> */}
                      <ProjectDetails
                        webURL={this.props.webURL}
                        context={this.props.context}
                      ></ProjectDetails>
                    </div>
                  )}
                </div>
              </div>
              {/* <Accordion title={'Project Details'} webURL={this.props.webURL}>

              </Accordion> */}
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.Sales}
                    text={"Sales Feedback"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        AllTask: false,
                        Sales: true,
                        Technical: false,
                        Delivery: false,
                        Legal: false,
                        Management: false,
                        ChatBox: false,
                      });
                    }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div>
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.Legal}
                    text={"Legal Feedback"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        AllTask: false,
                        Sales: false,
                        Technical: false,
                        Delivery: false,
                        Legal: true,
                        Management: false,
                        ChatBox: false,
                      });
                    }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div>
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.Technical}
                    text={"Technical Feedback"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        AllTask: false,
                        Sales: false,
                        Technical: true,
                        Delivery: false,
                        Legal: false,
                        Management: false,
                        ChatBox: false,
                      });
                    }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div>
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.Delivery}
                    text={"Delivery Feedback"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        AllTask: false,
                        Sales: false,
                        Technical: false,
                        Delivery: true,
                        Legal: false,
                        Management: false,
                        ChatBox: false,
                      });
                    }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div>
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.Management}
                    text={"Management Feedback"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        AllTask: false,
                        Sales: false,
                        Technical: false,
                        Delivery: false,
                        Legal: false,
                        Management: true,
                        ChatBox: false,
                      });
                    }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div>
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  {(this.state.IsSales || this.state.IsManagement) && (
                    <DefaultButton
                      toggle
                      checked={this.state.Management}
                      text={"Conversation"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      onClick={(e) => {
                        this.setState({
                          AllTask: false,
                          Sales: false,
                          Technical: false,
                          Delivery: false,
                          Legal: false,
                          Management: false,
                          ChatBox: true,
                        });
                      }}
                      // aria-expanded={this.state.expanded}
                      // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  )}
                </div>
              </div>
            </div>
            <div className={styles.section_right}>
              {this.state.Sales && <Sales webURL={this.props.webURL}></Sales>}
              {this.state.Technical && (
                <Technical webURL={this.props.webURL}></Technical>
              )}
              {this.state.Delivery && (
                <Delivery webURL={this.props.webURL}></Delivery>
              )}
              {this.state.Legal && <Legal webURL={this.props.webURL}></Legal>}
              {this.state.Management && (
                <Management
                  webURL={this.props.webURL}
                  context={this.props.context}
                ></Management>
              )}
              {this.state.AllTask && (
                <AllTask webURL={this.props.webURL}></AllTask>
              )}
              {this.state.ChatBox &&
                (this.state.IsSales || this.state.IsManagement) && (
                  <Chatbox
                    description={""}
                    context={this.props.context}
                  ></Chatbox>
                )}
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
                onClick={() => this.DeleteContract()}
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
      </div>
    );
  }

  private async DeleteContract() {
    this.setState({
      hideDialog: true,
    });
    $("#loader").show();
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .update({
        Status: "Deleted",
      })
      .then(async (i) => {
        let web = Web(this.props.webURL);
        await web.lists
          .getByTitle("CompleteTask")
          .items.add({
            ProjectID: parseInt(itemID),
            TaskType: "Deleted",
          })
          .then((newListItem) => {
            $("#loader").show();
            setTimeout(() => {
              $("#loader").hide();
            }, 3000);
            setTimeout(() => {
              $("#ContractDeletedModal").show();
            }, 3500);
          });
      });
  }

  private Close0() {
    var modal = document.getElementById("ContractDeletedModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $("#TestForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 1000);
  }

  private async setData() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select(
        "*",
        "Author/Id",
        "Author/EMail",
        "DeliveryManager/Title",
        "AccountManager/Title"
      )
      .expand("DeliveryManager/Id", "AccountManager/Id", "Author/EMail")
      .get()
      .then(async (item) => {
        var CreatedEmail = item.Author.EMail;
        let userDetails = await this.spLoggedInUserDetails();
        this.currentuseremail = userDetails.Email;

        if (item.Status == "Project Initiated") {
          this.setState({ ShowPDFBtn: true });
        }

        if (this.currentuseremail.toLowerCase() == CreatedEmail.toLowerCase()) {
          if (
            [
              "Feedback Action Awaited",
              "Management Action Awaited",
            ].indexOf(item.Status) != -1
          ) {
            this.setState({ IsSales: true, ShowEditContractBtn: true });
          } else {
            this.setState({
              IsSales: true,
            });
          }
        }
      })
      .then((x) => {
        if (!this.state.IsSales) {
          web.lists
            .getByTitle("Management Team")
            .items.select("Manager/EMail")
            .expand("Manager")
            .getAll()
            .then((items) => {
              //console.log(items);
              items.forEach((item) => {
                if (
                  item.Manager.EMail.toLowerCase() ==
                  this.currentuseremail.toLowerCase()
                ) {
                  this.setState({ IsManagement: true });
                }
              });
            });
        }
      });
  }
  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.webURL);
    return await web.currentUser.get();
  }
  ChatBox(): void {
    throw new Error("Method not implemented.");
  }

  ExportToPDF(): void {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .update({ TriggerPDFCreation: "trigger" })
      .then((result: any) => {
        alert(
          "PDF generated successfully.\n You will receive an email and generated PDF can be downloaded from the link provided inside the email."
        );
      });
    // .then(({ item }) =>
    //   item.update({ TriggerPDFCreation: "" }).then((result: any) => {
    //     alert("PDF Generated Successfully");
    //   })
    // );
  }
}
