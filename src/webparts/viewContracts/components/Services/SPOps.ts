import { ISPList } from "../IViewContractsProps";
import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as $ from "jquery";
import * as moment from "moment";

export class SPOperations {
  constructor(context: WebPartContext) {
    sp.setup({
      sp: {
        baseUrl: context.pageContext.web.absoluteUrl,
      },
    });
  }

  public ProjectsListData = (
    Weburl: string,
    useremail: string
  ): Promise<ISPList[]> => {
    let ContractDetails: ISPList[] = [];
    return new Promise<ISPList[]>(async (resolve, reject) => {
      await sp.web.lists
        .getByTitle("Projects")
        .items.select("*", "Author/EMail,FieldValuesAsText/Created")
        .expand("Author/Id,FieldValuesAsText")
        .filter(
          "Status eq 'Project Initiated' or Status eq 'Technical Action Awaited' or Status eq 'Delivery Action Awaited' or Status eq 'Management Action Awaited' or Status eq 'Legal Action Awaited' or Status eq 'CM Documents Upload Awaited' or Status eq 'PM Assignment Awaited' or Status eq 'PM Action Awaited' or Status eq 'Kickoff Docs Awaited' or Status eq 'Contract Rejection' or Status eq 'Rejected to Sales' or Status eq 'Escalated to management' or Status eq 'Approved PI Workflow Awaited'"
        )
        .orderBy("Created desc")
        .getAll()
        .then(
          (Results) => {
            let i = 0;
            Results.map((data) => {
              ContractDetails.push({
                OpportunityNo: data.OpportunityID,
                ProjectName: data.ProjectName,
                ClientOrganization: data.ClientOrganization,
                ProjectType: data.ProjectType,
                Stage: data.Status,
                EstimatedHours: data.EstimatedHours,
                ID: data.ID,
                useremaildata: useremail,
                WebUrl: Weburl,
                CreatedBY: data.Author.EMail,
                IndexNo: i++,
                Created: data.Created,
                CreatedDisplay: moment(data.FieldValuesAsText.Created).format(
                  "DD/MM/YYYY"
                ),
                // CreatedDisplay:
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[1] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[0] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[2],
              });
            });
            resolve(ContractDetails);
            console.log(ContractDetails);
          },
          (error: any) => {
            reject("Error Occured" + error);
          }
        );
    });
  };
  public ProjectsListDataInProgress = (
    Weburl: string,
    useremail: string
  ): Promise<ISPList[]> => {
    let ContractDetails: ISPList[] = [];
    return new Promise<ISPList[]>(async (resolve, reject) => {
      sp.web.lists
        .getByTitle("Projects")
        .items.select("*", "Author/EMail,FieldValuesAsText/Created")
        .expand("Author/Id,FieldValuesAsText")
        .filter(
          "Status eq 'Technical Action Awaited' or Status eq 'Delivery Action Awaited' or Status eq 'Management Action Awaited' or Status eq 'Legal Action Awaited' or Status eq 'CM Documents Upload Awaited' or Status eq 'PM Assignment Awaited' or Status eq 'PM Action Awaited' or Status eq 'Kickoff Docs Awaited' or Status eq 'Contract Rejection' or Status eq 'Rejected to Sales' or Status eq 'Escalated to management' or Status eq 'Approved PI Workflow Awaited'"
        )
        .orderBy("Created desc")
        .getAll()
        .then(
          (Results) => {
            let i = 0;
            Results.map((data) => {
              ContractDetails.push({
                OpportunityNo: data.OpportunityID,
                ProjectName: data.ProjectName,
                ClientOrganization: data.ClientOrganization,
                ProjectType: data.ProjectType,
                Stage: data.Status,
                EstimatedHours: data.EstimatedHours,
                ID: data.ID,
                useremaildata: useremail,
                WebUrl: Weburl,
                CreatedBY: data.Author.EMail,
                IndexNo: i++,
                Created: data.Created,
                CreatedDisplay: moment(data.FieldValuesAsText.Created).format(
                  "DD/MM/YYYY"
                ),
                // CreatedDisplay:
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[1] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[0] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[2],
              });
            });
            resolve(ContractDetails);
            console.log(ContractDetails);
          },
          (error: any) => {
            reject("Error Occured" + error);
          }
        );
    });
  };

  public ProjectsListDataCompleated = (
    Weburl: string,
    useremail: string
  ): Promise<ISPList[]> => {
    let ContractDetails: ISPList[] = [];
    return new Promise<ISPList[]>(async (resolve, reject) => {
      sp.web.lists
        .getByTitle("Projects")
        .items.select("*", "Author/EMail,FieldValuesAsText/Created")
        .expand("Author/Id,FieldValuesAsText")
        .filter("Status eq 'Project Initiated'")
        .orderBy("Created desc")
        .getAll()
        .then(
          (Results) => {
            let i = 0;
            Results.map((data) => {
              ContractDetails.push({
                OpportunityNo: data.OpportunityID,
                ProjectName: data.ProjectName,
                ClientOrganization: data.ClientOrganization,
                ProjectType: data.ProjectType,
                Stage: data.Status,
                EstimatedHours: data.EstimatedHours,
                ID: data.ID,
                useremaildata: useremail,
                WebUrl: Weburl,
                CreatedBY: data.Author.EMail,
                IndexNo: i++,
                Created: data.Created,
                CreatedDisplay: moment(data.FieldValuesAsText.Created).format(
                  "DD/MM/YYYY"
                ),

                // CreatedDisplay:
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[1] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[0] +
                //   "/" +
                //   data.FieldValuesAsText.Created.split(" ")[0].split("/")[2],
              });
            });
            resolve(ContractDetails);
            console.log(ContractDetails);
          },
          (error: any) => {
            reject("Error Occured" + error);
          }
        );
    });
  };

  public async getDialogListItems(projectId: number) {
    const auditItems = [];
    await sp.web.lists
      .getByTitle("TechnicalComments")
      .items.filter("ProjectID eq " + projectId)
      .select("*", "HOD/Title")
      .expand("HOD/Id")
      .get()
      .then(async (item) => {
        var TechTask = item;
        await sp.web.lists
          .getByTitle("ProjectTasks")
          .items.filter("ProjectID eq " + projectId)
          .top(5000)
          .select(
            "*",
            "AssignedTo/Title,Editor/Title,FieldValuesAsText/Modified,FieldValuesAsText/LastActionTakenOn"
          )
          .expand("AssignedTo/Id,Editor/Id,FieldValuesAsText")
          .get()
          .then((item) => {
            var StatusArray = [];
            var TechnologyArray = [];

            for (var i = 0; i < item.length; i++) {
              var AName = " " + item[i].Title + " ";
              var Status1 = " " + item[i].Status + " ";

              var Technology = "";
              var TechnologyCompare = "";

              if (item[i].Title == "Contract Rejection") {
                StatusArray = [];
                TechnologyArray = [];
              }

              if (item[i].Title == "Technical Action Awaited") {
                Technology = new URLSearchParams(item[i].VIew).get("dpt");
                TechnologyCompare = new URLSearchParams(item[i].VIew).get(
                  "dpt"
                );
                Technology = Technology.replace("ST", "Secondary").replace(
                  "PT",
                  "Primary"
                );
                var Technologynew = Technology.split("-");
                Technology = Technologynew[1] + " (" + Technologynew[0] + ")";
              }
              if (
                $.inArray(AName, StatusArray) >= 0 &&
                $.inArray(TechnologyCompare, TechnologyArray) >= 0
              ) {
              } else {
                TechnologyArray.push(TechnologyCompare);
                StatusArray.push(AName);
                if (item[i].Status == "Completed") {
                  AName = AName.replace("Action Awaited", "Feedback").replace(
                    "Awaited",
                    ""
                  );
                }
                // let options ={
                //   hour: 'numeric 2-digit', minute: 'numeric 2-digit', second: 'numeric 2-digit',
                //   day:'2-digit',month:'2-digit',
                //   timeZoneName: 'short'
                // };
                // let ISTtime:any = new Date(item[i].Modified.split("Z")[0]+"+05:30");
                // ISTtime = ISTtime.toLocaleString('hi-IN', { hour24: false, options });
                // let testvar:any = new Date(item[i].Modified);

                var lastActionDtString = "";

                if (item[i].FieldValuesAsText.LastActionTakenOn) {
                  var lstDt = new Date(
                    item[i].FieldValuesAsText.LastActionTakenOn
                  );
                  var date =
                    lstDt.getDate() <= 9
                      ? "0" + lstDt.getDate()
                      : lstDt.getDate().toString();
                  var month =
                    lstDt.getMonth() + 1 <= 9
                      ? "0" + (lstDt.getMonth() + 1)
                      : (lstDt.getMonth() + 1).toString();
                  var year = lstDt.getFullYear().toString();

                  lastActionDtString = date + "/" + month + "/" + year;
                }

                auditItems.push({
                  Action: AName,
                  Team: Technology.trim() == "" ? "-" : Technology,
                  ActionTakenBy:
                    item[i].Status == "In-Progress"
                      ? "-"
                      : item[i].LastActionTakeBy,
                  ActionTakenOn:
                    item[i].Status == "In-Progress"
                      ? "-"
                      : item[i].FieldValuesAsText.LastActionTakenOn != ""
                      ? moment(
                          item[i].FieldValuesAsText.LastActionTakenOn
                        ).format("DD/MM/YYYY")
                      : "",
                });
              }
            }
          });
      });
    console.log("ActionTakenOn value: " + auditItems[0].ActionTakenOn);
    return auditItems;
  }
}
