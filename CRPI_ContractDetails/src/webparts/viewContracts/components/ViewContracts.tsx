import * as React from "react";
import styles from "./ViewContracts.module.scss";
import { IViewContractsProps } from "./IViewContractsProps";
import { IViewContractsState } from "../components/IViewContractsState";
import { SPOperations } from "./Services/SPOps";
import { AuditDetailsDialog } from "./Dialog/AuditDetailsDialog";
import { Dialog, DialogFooter } from "@fluentui/react/lib/Dialog";
import { SPComponentLoader } from "@microsoft/sp-loader";
import "DataTables.net";
require("../assets/CustomDatatableStyle.css");

import { Web } from "@pnp/sp/webs";
import {
  DefaultButton,
  Link,
  Sticky,
  StickyPositionType,
} from "office-ui-fabric-react";
import {
  DetailsList,
  DialogType,
  IColumn,
  DetailsListLayoutMode,
  IconButton,
  SelectionMode,
  PrimaryButton,
} from "@fluentui/react";

import * as $ from "jquery";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

let that: any = null;
let table: any = null;
let _stausobject: object = {
  "Technical Action Awaited": "Pending for Technical Feedback",
  "Legal Action Awaited": "Pending for legal feedback",
  "Delivery Action Awaited": "Pending for delivery feedback",
  "Management Action Awaited": "Pending for management feedback",
  "CM Documents Upload Awaited": "Pending for CM to upload doc",
  "PM Assignment Awaited": "Pending for PM assignment",
  "PM Action Awaited": "Pending for PM fill the PI note",
  "Kickoff Docs Awaited": "Pending for Kick off doc to upload",
  "Contract Rejection": "Rejected by Management",
  "Rejected to Sales": "Rejected by Management",
  "Project Initiated": "Completed",
  "Escalated to management": "Pending for management feedback",
  "Approved PI Workflow Awaited": "Pending for Approved PI Workflow",
  "Management Approval Awaited to Start Without Signed Contract" : "Pending for management feedback",
};

export default class ViewContracts extends React.Component<
  IViewContractsProps,
  IViewContractsState,
  {}
> {
  public onRenderDetailsHeader = (headerProps, defaultRender) => {
    if (!headerProps || !defaultRender) {
      //technically these may be undefined...
      return null;
    }
    return (
      <Sticky
        stickyPosition={StickyPositionType.Header}
        isScrollSynced={true}
        stickyBackgroundColor="transparent"
      >
        {defaultRender({
          ...headerProps,
          styles: {
            root: {
              paddingTop: "0px!important",
              selectors: {
                ".ms-DetailsHeader-cell": {
                  whiteSpace: "normal",
                  textAlign: "center",
                  textOverflow: "clip",
                  lineHeight: "normal",
                  background: "#4672c4",
                  borderRight: "1px solid #00008B",
                },
                ":hover .ms-DetailsHeader-cell": {
                  whiteSpace: "normal",
                  textAlign: "center",
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
                  textAlign: "center",
                  "font-family":
                    '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                  "-webkit-font-smoothing": "antialiased",
                  fontSize: "12px",
                  fontWeight: "400",
                },
              },
            },
          },
        })}
      </Sticky>
    );
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
              //0.65625rem
            },
          },
        },
      },
    });
  };

  public modelProps = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
  };
  public dialogContentProps = {
    type: DialogType.largeHeader,
    title: "Confirm!",
    subText:
      "Once you delete this contract all the related details permanant deleted from database",
  };
  private _columns: IColumn[];

  private _ContextUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _spServices: SPOperations;

  public constructor(props: IViewContractsProps) {
    super(props);
    that = this;
    this._spServices = new SPOperations(this.props.context);
    this.state = {
      ListData: [],
      ListAllData: [],
      openDialog: false,
      CurrentTab: "In Progress",
      auditItems: [],
      hideDialog: true,
      itemID: "",
    };

    this._columns = [
      {
        key: "OpportunityNo",
        name: "Opportunity No",
        fieldName: "OpportunityNo",
        minWidth: 100,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
      },
      {
        key: "ProjectName",
        name: "Project Name",
        fieldName: "ProjectName",
        minWidth: 100,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
      },
      {
        key: "ClientOrganization",
        name: "Client Org.",
        fieldName: "ClientOrganization",
        minWidth: 100,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
      },
      {
        key: "ProjectType",
        name: "ProjectType",
        fieldName: "ProjectType",
        minWidth: 100,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
      },
      {
        key: "Stage",
        name: "Stage",
        fieldName: "Stage",
        minWidth: 185,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
        onRender: (data) =>
          _stausobject[data.Stage] == null
            ? data.Stage
            : _stausobject[data.Stage],
      },
      {
        key: "EstimatedHours",
        name: "Est. Hours",
        fieldName: "EstimatedHours",
        minWidth: 100,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
      },
      {
        key: "Action",
        name: "Action",
        styles: {},
        fieldName: "Action",
        minWidth: 160,
        maxWidth: 200,
        isPadded: true,
        isResizable: true,
        onRender: (data) => (
          <>
            <Link
              href={
                this._ContextUrl +
                "/SitePages/contract-details-New.aspx?itemid=" +
                data.ID
              }
              target="_blank"
              data-interception="off"
              style={{
                padding: "5px",
                fontFamily:
                  '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                WebkitFontSmoothing: "antialiased",
                fontSize: "13px",
                fontWeight: "400",
                color: "#4672c4",
              }}
            >
              View-link
            </Link>

            {data.CreatedBY.toLowerCase() ==
            this.props.userEmailId.toLowerCase() ? (
              //   ||
              // this.props.userEmailId.toLowerCase() ==
              //   "Sapan.Sehgal@rishabhsoft.com".toLowerCase()
              data.Stage == "Technical Action Awaited" ||
              data.Stage == "Legal Action Awaited" ||
              data.Stage == "Delivery Action Awaited" ||
              data.Stage == "Management Action Awaited" ? (
                <>
                  <a
                    href={
                      this._ContextUrl +
                      "/SitePages/UpdateContract.aspx?itemid=" +
                      data.ID
                    }
                    target="_blank"
                  >
                    <IconButton
                      iconProps={{ iconName: "TextDocumentEdit" }}
                      title="Contract Edit"
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

                  <IconButton
                    title="Contract Delete"
                    iconProps={{ iconName: "Delete" }}
                    onClick={() =>
                      this.setState({ hideDialog: false, itemID: data.ID })
                    }
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
                </>
              ) : (
                <IconButton
                  title="Contract Delete"
                  iconProps={{ iconName: "Delete" }}
                  onClick={() =>
                    this.setState({ hideDialog: false, itemID: data.ID })
                  }
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
              )
            ) : (
              <></>
            )}

            <a>
              <IconButton
                iconProps={{ iconName: "ComplianceAudit" }}
                title="Audit Details"
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
                onClick={() => this.openAuditDialog(data.ID)}
              />
            </a>
          </>
        ),
      },
    ];

    this.openAuditDialog = this.openAuditDialog.bind(this);
  }

  public componentDidMount() {
    this.inprogress(this._ContextUrl);
    // this._spServices.ProjectsListData(this.props.context.pageContext.web.absoluteUrl,this.props.userEmailId).then((result) => {
    //   setTimeout(() => {
    //     this.setState({ ListAllData: result });
    //   }, 1000);
    //   console.log("Final =" + this.state.ListAllData);
    // });

    //   "columnDefs": [{
    //     "targets": 6,
    //     "render": function (data, type, row) {//Hyper link to Column First
    //         // if (type === "display") {
    //         //     return "<a style=\"text-decoration: none; border-bottom: 1px solid #337ab7;\" href=\"https://pointerone.sharepoint.com/sites/SPFXDemo/Lists/EmployeeInformation/DispForm.aspx?ID=" + encodeURIComponent(row[3]) + "\">" + data + "</a>";
    //         // }
    //         return    `
    //         <a
    //           href=${
    //             row.WebUrl +
    //             "/SitePages/contract-details-New.aspx?itemid=" +
    //             row.ID
    //           }
    //           target="_blank"

    //         >
    //           View-link
    //         </a>
    //         ${row.CreatedBY.toLowerCase() ==
    //           row.useremaildata.toLowerCase() ||
    //         row.useremaildata.toLowerCase() ==
    //           "Sapan.Sehgal@rishabhsoft.com".toLowerCase() ? (
    //           row.Stage == "Technical Action Awaited" ||
    //           row.Stage == "Legal Action Awaited" ||
    //           row.Stage == "Delivery Action Awaited" ||
    //           row.Stage == "Management Action Awaited" ? (
    //             `
    //               <a
    //                 href=${
    //                   row.WebUrl +
    //                   "/SitePages/UpdateContract.aspx?itemid=" +
    //                   row.ID
    //                 }
    //                 target="_blank"
    //               >

    //                   link update
    //               </a>
    //               <button
    //               onClick=${() => this.setState({ hideDialog: false, itemID: row.ID})}

    //                 >Contract Delete</button>
    //             `
    //           ) : (
    //           ` <button

    //               onClick=${() => this.setState({ hideDialog: false, itemID: row.ID})}

    //             >Contract Delete</button>`
    //           )
    //         ) : (
    //           ` `
    //         )}

    //        <a>
    //           <button
    //             onClick=${() => { this.openAuditDialog(row.ID)}}
    //           >Audit Details</button>

    //       `;

    //     }
    // }
    // ]
  }

  public openAuditDialog(projectId: number) {
    this._spServices.getDialogListItems(projectId).then((items) => {
      this.setState({ openDialog: true, auditItems: items });
    });
  }

  inprogress(contexturl: string): void {
    this._spServices
      .ProjectsListDataInProgress(
        this.props.context.pageContext.web.absoluteUrl,
        this.props.userEmailId
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ListData: result, CurrentTab: "In Progress" });

          if ($.fn.dataTable.isDataTable("#example")) {
            table.destroy();
            $("#example").empty();
          }
          console.log("data is coming ...." + this.state.ListData);
          table = $("#example").DataTable({
            data: this.state.ListData,
            order: [[0, "desc"]],
            columns: [
              {
                // This column is used for sorting purpose only and not visible
                title: "Created",
                data: "Created",
                visible: false,
                searchable: false,
              },
              {
                title: "Opportunity No",
                data: "OpportunityNo",
                orderable: false,
                visible: false,
                searchable: false,
              },
              { title: "Project Name", data: "ProjectName", orderable: false },
              {
                title: "Client Org.",
                data: "ClientOrganization",
                orderable: false,
              },
              { title: "Project Type", data: "ProjectType", orderable: false },
              {
                title: "Stage",
                data: "Stage",
                orderable: false,
                render: function (data, type, row) {
                  var stage: string = "";
                  if (_stausobject[row.Stage] == "") {
                    stage = row.Stage;
                  } else {
                    stage = _stausobject[row.Stage];
                  }
                  return stage;
                },
              },
              { title: "Est. Hours", data: "EstimatedHours", orderable: false },
              {
                title: "Contract Initiated",
                data: "CreatedDisplay",
                orderable: false,
                searchable: false,
              },
              {
                title: "Action",
                data: null,
                "width":"80px",
                orderable: false,
                render: function (data, type, row, meta) {
                  let buttons: string = "";
                  buttons =
                    buttons +
                    '<a href="' +
                    row.WebUrl +
                    "/sitePages/contract-details-New.aspx?itemid=" +
                    row.ID +
                    '" target="_blank"><button type="button" title="View Contract Details" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="BulletedList"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  if (
                    row.CreatedBY.toLowerCase() ==
                      row.useremaildata.toLowerCase() ||
                    row.useremaildata.toLowerCase() ==
                      "Sapan.Sehgal@rishabhsoft.com".toLowerCase()
                  ) {
                    if (
                      row.Stage == "Technical Action Awaited" ||
                      row.Stage == "Legal Action Awaited" ||
                      row.Stage == "Delivery Action Awaited" ||
                      row.Stage == "Management Action Awaited"
                    ) {
                      if (
                        row.CreatedBY.toLowerCase() ==
                        row.useremaildata.toLowerCase()
                      ) {
                        buttons =
                          buttons +
                          '<a href="' +
                          contexturl +
                          "/SitePages/UpdateContract.aspx?itemid=" +
                          row.ID +
                          '" target="_blank"><button type="button" title="Contract Edit" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="TextDocumentEdit" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                      }

                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete" style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    } else {
                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    }
                  }

                  buttons =
                    buttons +
                    '<a class="name" id=n-"' +
                    meta.row +
                    '"><button type="button" title="Audit Details"style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="ComplianceAudit"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  return buttons;
                },
              },
            ],
          });

          $("#example tbody").on("click", ".name", function () {
            var data = table.row($(this).closest("tr")).data();

            that._spServices.getDialogListItems(data.ID).then((items) => {
              that.setState({ openDialog: true, auditItems: items });
            });
          });

          $("#example tbody").on("click", ".Deleteclass", function () {
            var data = table.row($(this).closest("tr")).data();

            that.setState({ hideDialog: false, itemID: data.ID });
          });
        }, 1000);
      });
  }

  Compleated(contexturl: string): void {
    this._spServices
      .ProjectsListDataCompleated(
        this.props.context.pageContext.web.absoluteUrl,
        this.props.userEmailId
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ListData: result, CurrentTab: "Completed" });

          if ($.fn.dataTable.isDataTable("#example")) {
            table.destroy();
            $("#example").empty();
          }
          console.log("data is coming ...." + this.state.ListData);
          table = $("#example").DataTable({
            data: this.state.ListData,
            order: [[0, "desc"]],
            columns: [
              {
                title: "Created",
                data: "Created",
                visible: false,
                searchable: false,
              },
              {
                title: "Opportunity No",
                data: "OpportunityNo",
                orderable: false,
                visible: false,
                searchable: false,
              },
              { title: "Project Name", data: "ProjectName", orderable: false },
              {
                title: "Client Org.",
                data: "ClientOrganization",
                orderable: false,
              },
              { title: "Project Type", data: "ProjectType", orderable: false },
              {
                title: "Stage",
                data: "Stage",
                orderable: false,
                render: function (data, type, row) {
                  var stage: string = "";
                  if (_stausobject[row.Stage] == "") {
                    stage = row.Stage;
                  } else {
                    stage = _stausobject[row.Stage];
                  }
                  return stage;
                },
              },
              { title: "Est. Hours", data: "EstimatedHours", orderable: false },
              {
                title: "Contract Initiated",
                data: "CreatedDisplay",
                orderable: false,
                searchable: false,
              },
              {
                title: "Action",
                data: null,
                "width":"80px",
                orderable: false,
                render: function (data, type, row, meta) {
                  let buttons: string = "";
                  buttons =
                    buttons +
                    '<a href="' +
                    row.WebUrl +
                    "/sitePages/contract-details-New.aspx?itemid=" +
                    row.ID +
                    '" target="_blank"><button type="button" title="View Contract Details" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="BulletedList"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  if (
                    row.CreatedBY.toLowerCase() ==
                      row.useremaildata.toLowerCase() ||
                    row.useremaildata.toLowerCase() ==
                      "Sapan.Sehgal@rishabhsoft.com".toLowerCase()
                  ) {
                    if (
                      row.Stage == "Technical Action Awaited" ||
                      row.Stage == "Legal Action Awaited" ||
                      row.Stage == "Delivery Action Awaited" ||
                      row.Stage == "Management Action Awaited"
                    ) {
                      if (
                        row.CreatedBY.toLowerCase() ==
                        row.useremaildata.toLowerCase()
                      ) {
                        buttons =
                          buttons +
                          '<a href="' +
                          contexturl +
                          "/SitePages/UpdateContract.aspx?itemid=" +
                          row.ID +
                          '" target="_blank"><button type="button" title="Contract Edit"style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="TextDocumentEdit"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                      }

                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    } else {
                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    }
                  }

                  buttons =
                    buttons +
                    '<a class="name" id=n-"' +
                    meta.row +
                    '"><button type="button" title="Audit Details"style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="ComplianceAudit"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  return buttons;
                },
              },
            ],
          });

          $("#example tbody").on("click", ".name", function () {
            var data = table.row($(this).closest("tr")).data();

            that._spServices.getDialogListItems(data.ID).then((items) => {
              that.setState({ openDialog: true, auditItems: items });
            });
          });

          $("#example tbody").on("click", ".Deleteclass", function () {
            var data = table.row($(this).closest("tr")).data();

            that.setState({ hideDialog: false, itemID: data.ID });
          });
        }, 1000);
      });
  }

  AllData(contexturl: string): void {
    this._spServices
      .ProjectsListData(
        this.props.context.pageContext.web.absoluteUrl,
        this.props.userEmailId
      )
      .then((result) => {
        setTimeout(() => {
          this.setState({ ListAllData: result, CurrentTab: "All Data" });

          if ($.fn.dataTable.isDataTable("#example")) {
            table.destroy();
            $("#example").empty();
          }
          console.log("data is coming ...." + this.state.ListData);
          table = $("#example").DataTable({
            data: this.state.ListAllData,
            order: [[0, "desc"]],
            columns: [
              {
                title: "Created",
                data: "Created",
                visible: false,
                searchable: false,
              },
              {
                title: "Opportunity No",
                data: "OpportunityNo",
                orderable: false,
                visible: false,
                searchable: false,
              },
              { title: "Project Name", data: "ProjectName", orderable: false },
              {
                title: "Client Org.",
                data: "ClientOrganization",
                orderable: false,
              },
              { title: "Project Type", data: "ProjectType", orderable: false },
              {
                title: "Stage",
                data: "Stage",
                orderable: false,
                render: function (data, type, row) {
                  var stage: string = "";
                  if (_stausobject[row.Stage] == "") {
                    stage = row.Stage;
                  } else {
                    stage = _stausobject[row.Stage];
                  }
                  return stage;
                },
              },
              { title: "Est. Hours", data: "EstimatedHours", orderable: false },
              {
                title: "Contract Initiated",
                data: "CreatedDisplay",
                orderable: false,
                searchable: false,
              },
              {
                title: "Action",
                data: null,
                "width":"80px",
                orderable: false,
                render: function (data, type, row, meta) {
                  let buttons: string = "";
                  buttons =
                    buttons +
                    '<a href="' +
                    row.WebUrl +
                    "/sitePages/contract-details-New.aspx?itemid=" +
                    row.ID +
                    '" target="_blank"><button type="button" title="View Contract Details" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="BulletedList"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  if (
                    row.CreatedBY.toLowerCase() ==
                      row.useremaildata.toLowerCase() ||
                    row.useremaildata.toLowerCase() ==
                      "Sapan.Sehgal@rishabhsoft.com".toLowerCase()
                  ) {
                    if (
                      row.Stage == "Technical Action Awaited" ||
                      row.Stage == "Legal Action Awaited" ||
                      row.Stage == "Delivery Action Awaited" ||
                      row.Stage == "Management Action Awaited"
                    ) {
                      if (
                        row.CreatedBY.toLowerCase() ==
                        row.useremaildata.toLowerCase()
                      ) {
                        buttons =
                          buttons +
                          '<a href="' +
                          contexturl +
                          "/SitePages/UpdateContract.aspx?itemid=" +
                          row.ID +
                          '" target="_blank"><button type="button" title="Contract Edit"style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="TextDocumentEdit"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                      }

                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    } else {
                      buttons =
                        buttons +
                        '<a class="Deleteclass" id=n-"' +
                        meta.row +
                        '"><button type="button" title="Contract Delete" style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="Delete"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';
                    }
                  }

                  buttons =
                    buttons +
                    '<a class="name" id=n-"' +
                    meta.row +
                    '"><button type="button" title="Audit Details"style="width:20px height:20px; background:transparent; border:none"><span style="margin-left:-6px"><i data-icon-name="ComplianceAudit"style="font-size: 16px; font-style:normal; font-family: &quot;Fluent MDL2 Hybrid Icons&quot;;"></i></span></button></a>';

                  return buttons;
                },
              },
            ],
          });

          $("#example tbody").on("click", ".name", function () {
            var data = table.row($(this).closest("tr")).data();

            that._spServices.getDialogListItems(data.ID).then((items) => {
              that.setState({ openDialog: true, auditItems: items });
            });
          });

          $("#example tbody").on("click", ".Deleteclass", function () {
            var data = table.row($(this).closest("tr")).data();

            that.setState({ hideDialog: false, itemID: data.ID });
          });
        }, 1000);
      });
  }

  public render(): React.ReactElement<IViewContractsProps> {
    SPComponentLoader.loadCss(
      "https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css"
    );
    const sbWidth = 6;
    const sbHeight = 6;
    const sbBg = "pink";
    const sbThumbBg = "red";

    return (
      <>
        <div className={styles.viewContracts}>
          <div className={styles.tabHeader}>
            <PrimaryButton
              className={
                this.state.CurrentTab == "In Progress"
                  ? styles.tabButtonSelected
                  : styles.tabButton
              }
              onClick={() => {
                this.inprogress(this._ContextUrl);
              }}
            >
              In-Progress
            </PrimaryButton>

            {"   "}
            <PrimaryButton
              className={
                this.state.CurrentTab == "Completed"
                  ? styles.tabButtonSelected
                  : styles.tabButton
              }
              onClick={() => {
                this.Compleated(this._ContextUrl);
              }}
            >
              Completed
            </PrimaryButton>
            {"   "}
            <PrimaryButton
              className={
                this.state.CurrentTab == "All Data"
                  ? styles.tabButtonSelected
                  : styles.tabButton
              }
              onClick={() => {
                this.AllData(this._ContextUrl);
              }}
            >
              All
            </PrimaryButton>
          </div>

          <table id="example" className="cell-border" width="100%"></table>

          {/* <div>
            <DetailsList
              className={styles.tabContainer}
              columns={this._columns}
              items={
                this.state.CurrentTab == "All Data"
                  ? this.state.ListAllData
                  : this.state.ListData
              }
              compact={true}
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              onRenderDetailsHeader={this.onRenderDetailsHeader}
              onRenderRow={this.onRenderRow}
              selectionMode={SelectionMode.none}
              //onRenderDetailsFooter={this._onRenderDetailsFooter}
            />
          </div> */}
          {this.state.openDialog ? (
            <AuditDetailsDialog
              open={this.state.openDialog}
              audit={this.state.auditItems}
              onClose={this.closeDialog.bind(this)}
            ></AuditDetailsDialog>
          ) : (
            <></>
          )}

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
            dialogContentProps={this.dialogContentProps}
            modalProps={this.modelProps}
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
      </>
    );
  }

  public closeDialog() {
    this.setState({ openDialog: false, auditItems: [] });
  }

  private async DeleteContract() {
    this.setState({
      hideDialog: true,
    });
    $("#loader").show();
    var _itemID = this.state.itemID;
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(_itemID))
      .update({
        Status: "Deleted",
      })
      .then(async (i) => {
        console.log("loading...");
        let web = Web(this.props.webURL);
        await web.lists
          .getByTitle("CompleteTask")
          .items.add({
            ProjectID: parseInt(_itemID),
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

  public Close0() {
    var modal = document.getElementById("ContractDeletedModal");
    var hostUrl = this._ContextUrl;
    modal.style.display = "none";
    $("#TestForm").fadeOut(2500);
    setTimeout(() => {
      window.parent.location.href = hostUrl;
    }, 1000);
  }
}
