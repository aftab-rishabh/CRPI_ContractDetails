import * as React from "react";
import styles from "./MyActiveTasks.module.scss";
import { IMyActiveTasksProps } from "./IMyActiveTasksProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IMyActiveTasksState } from "./IMyActiveTasksState";
import { DetailsDialog } from "./Dialog/DetailsDialog";
import {
  ConstrainMode,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  DialogType,
  IColumn,
  IconButton,
  PrimaryButton,
  SelectionMode,
} from "@fluentui/react";

import { SPOperations } from "../Services/SPOps";

import * as $ from "jquery";
import { IMyActiveTasksItems } from "./IMyActiveTasksItems";
import { List } from "@fluentui/react";

const modelProps = {
  isBlocking: false,
  styles: { main: { maxWidth: 450 } },
};
const dialogContentProps = {
  type: DialogType.normal,
  title: "Audit",
};

export default class MyActiveTasks extends React.Component<
  IMyActiveTasksProps,
  IMyActiveTasksState,
  {}
> {
  public onRenderDetailsHeader = (headerProps, defaultRender) => {
    if (!headerProps || !defaultRender) {
      //technically these may be undefined...
      return null;
    }
    return defaultRender({
      ...headerProps,
      styles: {
        root: {
          selectors: {
            ".ms-DetailsHeader-cell": {
              whiteSpace: "normal",
              textOverflow: "clip",
              lineHeight: "normal",
              background: "#4672c4",
              borderRight: "1px solid #00008B",
            },
            ":hover .ms-DetailsHeader-cell": {
              whiteSpace: "normal",
              textOverflow: "clip",
              lineHeight: "normal",
              background: "#4672c4",
              borderRight: "1px solid #00008B",
            },
            ".ms-DetailsHeader-cellTitle": {
              height: "100%",
              color: "white",
              justifyContent: "center",
              alignItems: "center",
              "font-family":
                '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
              "-webkit-font-smoothing": "antialiased",
              fontSize: "12px",
              fontWeight: "400",
            },
          },
        },
      },
    });
  };

  public onRenderRow = (RowProps, defaultRender) => {
    if (!RowProps || !defaultRender) {
      //technically these may be undefined...
      return null;
    }
    return defaultRender({
      ...RowProps,
      styles: {
        root: {
          selectors: {
            ".ms-DetailsRow-cell": {
              whiteSpace: "normal",
              textOverflow: "clip",
              lineHeight: "normal",
              color: "#4672c4",
              borderRight: "1px solid #4672c4",
              "font-family":
                '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
              "-webkit-font-smoothing": "antialiased",
              fontSize: "12px",
              fontWeight: "400",
            },
          },
        },
      },
    });
  };
  private _spServices: SPOperations;
  private _columns: IColumn[];
  private _columns2: IColumn[];

  constructor(props: IMyActiveTasksProps) {
    super(props);
    this._spServices = new SPOperations(this.props.context);

    this.state = {
      ProjectListResult: [],
      CurrentTab: "Pending",
      openDialog: false,
      auditItems: [],
    };

    this._columns = [
      {
        key: "OpportunityNumber",
        name: "Opportunity No",
        fieldName: "OpportunityNumber",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ProjectName",
        name: "Project Name",
        fieldName: "ProjectName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ClientOrganization",
        name: "Client Organization",
        fieldName: "ClientOrganization",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ProjectType",
        name: "Project Type",
        fieldName: "ProjectType",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },

      {
        key: "AssignedOn",
        name: "Assigned On",
        fieldName: "AssignedOn",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
        onRender: (item) => item.AssignedOn,
      },

      {
        key: "Action",
        name: "Action",
        fieldName: "Action",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
        onRender: (item) =>
          item.Status == "In-Progress" ? (
            <>
              <a
                href={
                  this.props.context.pageContext.web.absoluteUrl +
                  "/sitePages/contract-details-New.aspx?itemid=" +
                  item.ProjectId
                }
                target="__blank"
              >
                <IconButton
                  iconProps={{ iconName: "BulletedList" }}
                  title="View Contract Details"
                  // onClick={() => {this.onPendingEditButtonClick(item.ProjectId)}}
                  styles={{
                    icon: { color: "rgba(40, 40, 225, 0.888)", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: "rgba(40, 40, 225, 0.888)",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
              <a href={item.TaskLink} target="__blank">
                <IconButton
                  iconProps={{ iconName: "TextDocumentEdit" }}
                  title="View Task"
                  styles={{
                    icon: { color: " #4672c4", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: " #4672c4",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
            </>
          ) : (
            <div>
              <a href={item.ContractDetail} target="__blank">
                <IconButton
                  iconProps={{ iconName: "BulletedList" }}
                  title="Contract Details"
                  styles={{
                    icon: { color: " #4672c4", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: " #4672c4",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
              <a>
                <IconButton
                  iconProps={{ iconName: "ComplianceAudit" }}
                  title="Audit Details"
                  onClick={() => this.openAuditDialog(item.ProjectId)}
                  styles={{
                    icon: { color: "#4672c4", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: "#4672c4",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
              {item.TaskLink.lenght > 0 ? (
                <a href={item.TaskLink} target="__blank">
                  <IconButton
                    iconProps={{ iconName: "TextDocumentEdit" }}
                    title="View Task"
                    styles={{
                      icon: { color: "#4672c4", fontSize: 18 },
                      root: {
                        selectors: {
                          ":hover .ms-Button-icon": {
                            color: "#4672c4",
                            fontSize: 20,
                          },
                        },
                      },
                    }}
                  />
                </a>
              ) : (
                ""
              )}
            </div>
          ),
      },
    ];

    this._columns2 = [
      {
        key: "OpportunityNumber",
        name: "Opportunity No",
        fieldName: "OpportunityNumber",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ProjectName",
        name: "Project Name",
        fieldName: "ProjectName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ClientOrganization",
        name: "Client Organization",
        fieldName: "ClientOrganization",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "ProjectType",
        name: "Project Type",
        fieldName: "ProjectType",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      // {
      //   key: "ContractLink",
      //   name: "Contract Link",
      //   fieldName: "ContractLink",
      //   minWidth: 100,
      //   maxWidth: 200,
      //   isResizable: true,
      //   data: "string",
      //   isPadded: true,
      //   onRender: (item) =>
      //   (
      //   <a href={this.props.context.pageContext.web.absoluteUrl+"/sitePages/contract-details-New.aspx?itemid="+item.ProjectId} target="__blank">
      //   <IconButton
      //     iconProps={{ iconName: "BulletedList" }}
      //     title="View Contract Details"
      //    // onClick={() => {this.onPendingEditButtonClick(item.ProjectId)+''}}
      //     styles={{
      //      icon: { color: "#4672c4", fontSize: 18 },
      //      root: {
      //        selectors: {
      //          ":hover .ms-Button-icon": {
      //            color: "#4672c4",
      //            fontSize: 20,
      //          },
      //        },
      //      },
      //    }}
      //   />
      //   </a>)
      // }
      {
        key: "Contractstatus",
        name: "Contract status",
        fieldName: "Contractstatus",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
        onRender: (item) => item.Status,
      },
      ,
      {
        key: "Action",
        name: "Action",
        fieldName: "Action",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        data: "string",
        isPadded: true,
        onRender: (item) =>
          item.Status == "In-Progress" ? (
            <a href={item.TaskLink} target="__blank">
              <IconButton
                iconProps={{ iconName: "TextDocumentEdit" }}
                title="View Task"
                styles={{
                  icon: { color: "#4672c4", fontSize: 18 },
                  root: {
                    selectors: {
                      ":hover .ms-Button-icon": {
                        color: "#4672c4",
                        fontSize: 20,
                      },
                    },
                  },
                }}
              />
            </a>
          ) : (
            <div>
              <a href={item.ContractDetail} target="__blank">
                <IconButton
                  iconProps={{ iconName: "BulletedList" }}
                  title="Contract Details"
                  styles={{
                    icon: { color: "#4672c4", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: "#4672c4",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
              <a>
                <IconButton
                  iconProps={{ iconName: "ComplianceAudit" }}
                  title="Audit Details"
                  onClick={() => this.openAuditDialog(item.ProjectId)}
                  styles={{
                    icon: { color: "#4672c4", fontSize: 18 },
                    root: {
                      selectors: {
                        ":hover .ms-Button-icon": {
                          color: "#4672c4",
                          fontSize: 20,
                        },
                      },
                    },
                  }}
                />
              </a>
              {item.TaskLink != "" ? (
                <a href={item.TaskLink} target="__blank">
                  <IconButton
                    iconProps={{ iconName: "TextDocumentEdit" }}
                    title="Edit"
                    styles={{
                      icon: { color: "#4672c4", fontSize: 18 },
                      root: {
                        selectors: {
                          ":hover .ms-Button-icon": {
                            color: "#4672c4",
                            fontSize: 20,
                          },
                        },
                      },
                    }}
                  />
                </a>
              ) : (
                ""
              )}
            </div>
          ),
      },
    ];

    this.onPendingEditButtonClick = this.onPendingEditButtonClick.bind(this);
    this.openAuditDialog = this.openAuditDialog.bind(this);
  }

  public componentDidMount(): void {
    this.ShowPendingTasks();
  }

  public render(): React.ReactElement<IMyActiveTasksProps> {
    return (
      <div className={styles.myActiveTasks}>
        <div className={styles.tabHeader}>
          <PrimaryButton
            text="Pending"
            onClick={() => this.ShowPendingTasks()}
            hidden={false}
            className={
              this.state.CurrentTab == "Pending"
                ? styles.tabButtonSelected
                : styles.tabButton
            }
          >
            {" "}
          </PrimaryButton>
          <PrimaryButton
            text="Acomplished"
            onClick={() => this.ShowCompletedTasks()}
            hidden={false}
            className={
              this.state.CurrentTab == "Completed"
                ? styles.tabButtonSelected
                : styles.tabButton
            }
          >
            {" "}
          </PrimaryButton>
        </div>
        <div className={styles.tabContainer}>
          {
            <DetailsList
              items={this.state.ProjectListResult}
              columns={
                this.state.CurrentTab == "Pending"
                  ? this._columns
                  : this._columns2
              }
              compact={true}
              layoutMode={DetailsListLayoutMode.justified}
              onRenderDetailsHeader={this.onRenderDetailsHeader}
              onRenderRow={this.onRenderRow}
              isHeaderVisible={true}
              selectionMode={SelectionMode.none}
              //onRenderItemColumn={this._renderItemColumn}
            ></DetailsList>
          }
        </div>
        {this.state.openDialog ? (
          <DetailsDialog
            open={this.state.openDialog}
            audit={this.state.auditItems}
            onClose={this.closeDialog.bind(this)}
          ></DetailsDialog>
        ) : (
          <></>
        )}
      </div>
    );
  }

  public closeDialog() {
    this.setState({ openDialog: false, auditItems: [] });
  }

  // public _renderItemColumn(
  //   item: IMyActiveTasksItems,
  //   index: number,
  //   column: IColumn
  // ) {
  //   const fieldContent = item[
  //     column.fieldName as keyof IMyActiveTasksItems
  //   ] as string;

  //   switch (column.key) {
  //     case "Action":
  //       return item.Status == "In-Progress" ? (
  //         <a href={item.TaskLink} target="__blank">
  //           <IconButton
  //             iconProps={{ iconName: "TextDocumentEdit" }}
  //             title="View Task"
  //           />
  //         </a>
  //       ) : (
  //         <div>
  //           <a href={item.TaskLink} target="__blank">
  //             <IconButton
  //               iconProps={{ iconName: "TextDocumentEdit" }}
  //               title="View Task"
  //             />
  //           </a>
  //           <a>
  //             <IconButton
  //               iconProps={{ iconName: "ComplianceAudit" }}
  //               title="View Task"
  //               onClick={() => this.openAuditDialog(item.ProjectId)}
  //             />
  //           </a>
  //           <a href={item.ContractDetail} target="__blank">
  //             <IconButton
  //               iconProps={{ iconName: "BulletedList" }}
  //               title="Contract Details"
  //             />
  //           </a>
  //         </div>
  //       );
  //     case "AssignedOn":
  //       return <span>{fieldContent}</span>;
  //     default:
  //       return <span>{fieldContent}</span>;
  //   }
  // }

  public openAuditDialog(projectId: number) {
    this._spServices.getDialogListItems(projectId).then((items) => {
      this.setState({ openDialog: true, auditItems: items });
    });
  }

  public onPendingEditButtonClick(itemId: number): void {
    window.location.href =
      this.props.context.pageContext.web.absoluteUrl +
      "/sitePages/contract-details-New.aspx?itemid=" +
      itemId;
  }

  ShowPendingTasks(): void {
    this._spServices
      .getMyTaskItems(
        this.props.context.pageContext.web.absoluteUrl,
        "In-Progress"
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ProjectListResult: result, CurrentTab: "Pending" });
        }, 1000);
      });
  }

  ShowCompletedTasks(): void {
    this._spServices
      .getMyTaskItems(
        this.props.context.pageContext.web.absoluteUrl,
        "Completed"
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ProjectListResult: result, CurrentTab: "Completed" });
        }, 1000);
      });
  }
}
