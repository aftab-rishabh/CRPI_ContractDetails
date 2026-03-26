import * as React from "react";
import styles from "./Feedback.module.scss";

import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from "office-ui-fabric-react";
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from "office-ui-fabric-react";
// import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import * as $ from "jquery";
import { WebPartContext } from "@microsoft/sp-webpart-base";
/**
 * Icon styles. Feel free to change them
 */

export interface IManagementProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
  context: WebPartContext;
}

export interface IManagementState {
  expanded: boolean;
  Items: any;
  HTML: any;
  files: any;
  ManagementFlag: boolean;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Management extends React.Component<
  IManagementProps,
  IManagementState
> {
  private _drawerDiv: HTMLDivElement = undefined;

  constructor(props: IManagementProps) {
    super(props);

    this.state = {
      expanded:
        props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
      files: [],
      ManagementFlag: false,
    };
  }
  public async componentDidMount() {
    this.fetchData();
  }
  public render(): React.ReactElement<IManagementProps> {
    return (
      <div className={styles.Feedback}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.viewdetail_block}>
              <div className="ManagementData">
                <div className={styles.sectionblock}>
                  <Label className={styles.headers}>
                    <u>Management Feedback</u>
                  </Label>
                </div>

                <div id="managementDetailsSection" style={{ display: "none" }}>
                  <div className={styles.sectionblock}>
                    <div className={styles.viewanswer_listing}>
                      <ul>
                        <li>
                          <div className={styles.viewcomment}>
                            Management comment:
                          </div>
                          <div
                            className={styles.viewanswer}
                            id="ManagementComment"
                            dangerouslySetInnerHTML={{
                              __html: this.state.Items,
                            }}
                          ></div>
                        </li>
                      </ul>
                    </div>
                  </div>
                  {this.state.ManagementFlag && (
                    <div className={styles.sectionblock}>
                      <div className={styles.view_listing}>
                        <ul>
                          <li>
                            <div className={styles.viewlable}>All Files:</div>
                          </li>
                        </ul>
                        <div id="fileList3" className={styles.file_list}>
                          <ul>
                            {this.state.files.map(function (item, key) {
                              if (
                                item.Title == "Management" ||
                                item.Title == "Contract - Commercial"
                              ) {
                                return (
                                  <li>
                                    {item.Title} -{" "}
                                    <a
                                      href={item.ServerRelativeUrl}
                                      target="_blank"
                                    >
                                      {item.Name}
                                    </a>
                                  </li>
                                );
                              }
                            })}
                          </ul>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public async fetchData() {
    let userDetails = await this.spLoggedInUserDetails();

    this.GetManagementTeam(userDetails.Email);
  }

  private async setData() {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(itemID))
      .select("*", "DeliveryManager/Title", "AccountManager/Title")
      .expand("DeliveryManager/Id", "AccountManager/Id")
      .get()
      .then((item) => {
        var MgtComments = item.MgtComments;
        if (MgtComments === null) {
          MgtComments = "";
        } else if (item.FeedbackStatus === "Draft") {
          MgtComments = "";
        }
        this.getFiles(item, this.props.context);
        // this.setState({ Items: MgtComments });

        const cleanHtml = MgtComments?.replace(
          /<script.*?>.*?<\/script>/gi,
          "",
        );
        this.setState({ Items: cleanHtml });
        $("#managementDetailsSection").show();
      })
      .then((x) => {
        // this.HideNAData();
      });
  }

  private async GetManagementTeam(CurrentUserEmail) {
    const itemID = new URLSearchParams(window.location.search).get("itemid");
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Management Team")
      .items.select("*", "Manager/Title", "Manager/EMail")
      .expand("Manager")
      .get()
      .then((data) => {
        for (var i = 0; i < data.length; i++) {
          if (
            CurrentUserEmail.toLowerCase() ==
            data[i].Manager.EMail.toLowerCase()
          ) {
            this.setState({
              ManagementFlag: true,
            });
          }
        }

        this.setData();
      });
  }

  private getFiles(items, context) {
    // this.context.pageContext.site.serverRelativeUrl
    let web = Web(this.props.webURL);
    var hostUrl =
      context.pageContext.site.serverRelativeUrl +
      "/ProjectDocuments/" +
      items.OpportunityID;
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web
      .getFolderByServerRelativeUrl(hostUrl)
      .files.get()
      .then(async (files) => {
        this.setState({
          files: files,
        });
      });
  }

  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.webURL);
    return await web.currentUser.get();
  }
}
