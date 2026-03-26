import * as React from "react";
import styles from "./MyActiveTasks.module.scss";
import { IMyActiveTasksProps } from "./IMyActiveTasksProps";
import { IMyActiveTasksState } from "./IMyActiveTasksState";
import { DetailsDialog } from "./Dialog/DetailsDialog";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IconButton,
  PrimaryButton,
  SelectionMode,
} from "@fluentui/react";
import { Web } from "@pnp/sp/presets/all";

import { SPOperations } from "../Services/SPOps";
import { SPComponentLoader } from "@microsoft/sp-loader";
import "DataTables.net";

require("../assets/CustomDatatableStyle.css");

let that: any = null;
let table: any = null;

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
          paddingTop: "0px!important",
          selectors: {
            "ms-DetailsHeader": {
              paddingTop: "0px!important",
            },
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
  currentUserId: any;

  constructor(props: IMyActiveTasksProps) {
    super(props);
    that = this;
    this._spServices = new SPOperations(this.props.context);

    this.alertOnClick = this.alertOnClick.bind(this);

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
                target="_blank"
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
              <a href={item.TaskLink} target="_blank">
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
              <a href={item.ContractDetail} target="_blank">
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
                <a href={item.TaskLink} target="_blank">
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
      //   <a href={this.props.context.pageContext.web.absoluteUrl+"/sitePages/contract-details-New.aspx?itemid="+item.ProjectId} target="_blank">
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
            <a href={item.TaskLink} target="_blank">
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
              <a href={item.ContractDetail} target="_blank">
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
                <a href={item.TaskLink} target="_blank">
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

  public async componentDidMount() {
    this.ShowPendingTasks();
  }

  public render(): React.ReactElement<IMyActiveTasksProps> {
    SPComponentLoader.loadCss(
      "https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css"
    );
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
            text="Accomplished"
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
        <table id="exampleActive" className="cell-border" width="100%"></table>

        {/* <div className={styles.tabContainer}>
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
        </div> */}
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
  //         <a href={item.TaskLink} target="_blank">
  //           <IconButton
  //             iconProps={{ iconName: "TextDocumentEdit" }}
  //             title="View Task"
  //           />
  //         </a>
  //       ) : (
  //         <div>
  //           <a href={item.TaskLink} target="_blank">
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
  //           <a href={item.ContractDetail} target="_blank">
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

  private alertOnClick(): void {
    alert();
  }

  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.context.pageContext.web.absoluteUrl);
    return await web.currentUser.get();
  }

  private async ShowPendingTasks() {
    let userDetails = await this.spLoggedInUserDetails();
    this.currentUserId = userDetails.Id;

    await this._spServices
      .getMyTaskItems(
        this.props.context.pageContext.web.absoluteUrl,
        "In-Progress",
        this.currentUserId
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ProjectListResult: result, CurrentTab: "Pending" });

          if ($.fn.dataTable.isDataTable("#exampleActive")) {
            table.destroy();
            $("#exampleActive").empty();
          }

          table = $("#exampleActive").DataTable({
            data: result,
            order: [[0, "desc"]],
            columns: [
              {
                title: "Created",
                data: "Created",
                visible: false,
                searchable: false,
              },
              {
                title: "Opportunity Number",
                data: "OpportunityNumber",
                orderable: false,
                visible: false,
                searchable: false,
              },
              { title: "Project Name", data: "ProjectName", orderable: false },
              // {
              //   title: "Initiated Date",
              //   data: "CreatedDisplay",
              //   orderable: false,
              //   searchable: false,
              // },
              {
                title: "Client Organization",
                data: "ClientOrganization",
                orderable: false,
              },
              { title: "Project Type", data: "ProjectType", orderable: false },
              { title: "Assigned on", data: "AssignedOn", orderable: false },
              {
                title: "Action",
                data: null,
                orderable: false,
                render: function (data, type, row, meta) {
                  return (
                    // '<input type="button" class="salary" id=s-"' +
                    // meta.row +
                    // '" value="Details"/>&nbsp<input type="button" class="name" id=n-"' +
                    // meta.row +
                    // '" value="Audit"/>'
                    '<a href="' +
                    row.ContractDetail +
                    '" target="_blank" tabindex="-1"><button type="button" style="width:20px height:20px; background:transparent; border:none" title="View Contract Details"><span style="margin-left:-6px"><i  data-icon-name="BulletedList" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>' +
                    '<a href="' +
                    row.TaskLink +
                    '" target="_blank" tabindex="-1"><button type="button" style="width:20px height:20px; background:transparent; border:none" title="View Task"><span style="margin-left:-6px"><i data-icon-name="TextDocumentEdit" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>'
                  );
                },
              },
            ],
            // columnDefs: [
            //   {
            //     targets: 5,
            //     render: function (data, type, row) {
            //       //Hyper link to Column First
            //       if (type === "display") {
            //         return (
            //           "<a href='" +
            //           row.TaskLink +
            //           "' target='_blank'>Edit</a>&nbsp<a href='#' onClick={this.alertOnClick} >Audit</a>"
            //         );
            //       }
            //       return data;
            //     },
            //   },
            // ],
          });
        }, 1000);
      });
  }

  private async ShowCompletedTasks() {
    let userDetails = await this.spLoggedInUserDetails();
    this.currentUserId = userDetails.Id;

    await this._spServices
      .getMyTaskItems(
        this.props.context.pageContext.web.absoluteUrl,
        "Completed",
        this.currentUserId
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ProjectListResult: result, CurrentTab: "Completed" });

          if ($.fn.dataTable.isDataTable("#exampleActive")) {
            table.destroy();
            $("#exampleActive").empty();
          }

          table = $("#exampleActive").DataTable({
            data: result,
            order: [[0, "desc"]],
            columns: [
              {
                title: "Created",
                data: "Created",
                visible: false,
                searchable: false,
              },
              {
                title: "Opportunity Number",
                data: "OpportunityNumber",
                visible: false,
                searchable: false,
                orderable: false,
              },
              { title: "Project Name", data: "ProjectName", orderable: false },
              // {
              //   title: "Initiated Date",
              //   data: "CreatedDisplay",
              //   orderable: false,
              //   searchable: false,
              // },
              {
                title: "Client Organization",
                data: "ClientOrganization",
                orderable: false,
              },
              { title: "Project Type", data: "ProjectType", orderable: false },
              { title: "Contract status", data: "Status", orderable: false },
              {
                title: "Action",
                data: null,
                orderable: false,
                render: function (data, type, row, meta) {
                  let buttons: string = "";

                  buttons =
                    buttons +
                    '<a href="' +
                    row.ContractDetail +
                    '" target="_blank"><button type="button" style="width:20px height:20px; background:transparent; border:none" title="View Contract Details"><span style="margin-left:-6px"><i data-icon-name="BulletedList" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  if (row.TaskLink != "") {
                    buttons =
                      buttons +
                      '<a href="' +
                      row.TaskLink +
                      '" target="_blank"><button type="button" style="width:20px height:20px; background:transparent; border:none" title="View Task" ><span style="margin-left:-6px"><i data-icon-name="TextDocumentEdit" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                  }

                  buttons =
                    buttons +
                    '<a class="name" id=n-"' +
                    meta.row +
                    '"><button type="button" style="width:20px height:20px; background:transparent; border:none" title="Audit Details"><span style="margin-left:-6px"><i data-icon-name="ComplianceAudit" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  return buttons;
                },
              },
            ],
            // columnDefs: [
            //   {
            //     targets: 5,
            //     render: function (data, type, row) {
            //       //Hyper link to Column First
            //       if (type === "display") {
            //         return (
            //           "<a href='" +
            //           row.TaskLink +
            //           "' target='_blank'>Edit</a>&nbsp<a href='#' onClick={this.alertOnClick} >Audit</a>"
            //         );
            //       }
            //       return data;
            //     },
            //   },
            // ],
          });

          $("#exampleActive tbody").on("click", ".name", function () {
            var data = table.row($(this).closest("tr")).data();

            that._spServices
              .getDialogListItems(data.ProjectId)
              .then((items) => {
                that.setState({ openDialog: true, auditItems: items });
              });
          });
        }, 1000);
      });
  }
}
