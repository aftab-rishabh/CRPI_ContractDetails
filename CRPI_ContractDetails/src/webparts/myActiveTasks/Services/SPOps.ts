import { ICamlQuery, sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IMyActiveTasksItems } from "../components/IMyActiveTasksItems";
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

  public getMyTaskItems = (
    contextUrl: string,
    status: string,
    userID: any
  ): Promise<IMyActiveTasksItems[]> => {
    const activeItems: IMyActiveTasksItems[] = [];

    return new Promise<IMyActiveTasksItems[]>(async (resolve, reject) => {
      await sp.web.lists
        .getByTitle("ProjectTasks")
        .items.select("*,FieldValuesAsText/Created")
        .expand("FieldValuesAsText")
        .filter("AssignedTo eq " + userID + " and Status eq '" + status + "'")
        .orderBy("Created", false)
        .top(5000)
        .get()
        .then(async (taskItems) => {
          console.log(taskItems);

          var projectId = taskItems
            .map((taskItem) => {
              return taskItem.ProjectID;
            })
            .filter((value, index, self) => {
              return self.indexOf(value) === index;
            });

          var camlInClause = "";

          for (
            let index = 0;
            projectId.slice().splice(index * 5, 5).length > 0;
            index++
          ) {
            var inQuery = projectId
              .slice()
              .splice(index * 5, 5)
              .map((id) => {
                return "<Value Type='Number'>" + id + "</Value>";
              });

            camlInClause +=
              "<In><FieldRef Name='ID' /><Values>" +
              inQuery.join("") +
              "</Values></In>";

            if (index >= 1) {
              camlInClause = "<Or>" + camlInClause + "</Or>";
            }
          }
          if (camlInClause.length > 0) {
            const camlQuery: ICamlQuery = {};
            camlQuery.ViewXml =
              "<View><Query><Where>" + camlInClause + "</Where></Query></View>";
            await sp.web.lists
              .getByTitle("Projects")
              .getItemsByCAMLQuery(camlQuery, "FieldValuesAsText")
              .then((prjItems) => {
                console.log("FieldValuesAsText value");
                console.log(prjItems);
                taskItems.map(
                  (tskItm) => {
                    var showRecord = false;
                    var projectItm = prjItems.filter((itm) => {
                      return itm.Id == tskItm.ProjectID;
                    });

                    if (projectItm.length > 0) {
                      if (
                        status == "Completed" &&
                        projectItm[0].Status != "Project Initiated"
                      ) {
                        showRecord = true;
                      } else if (status == "In-Progress") {
                        showRecord = true;
                      }

                      if (showRecord) {
                        // let options ={
                        //   hour: 'numeric 2-digit', minute: 'numeric 2-digit', second: 'numeric 2-digit',
                        //   day:'2-digit',month:'2-digit',
                        //   timeZoneName: 'short'
                        // };
                        // let ISTtime:any = new Date(projectItm[0].Created.split("Z")[0]+"+05:30");
                        // ISTtime = ISTtime.toLocaleString('hi-IN', { hour24: false, options });

                        activeItems.push({
                          ProjectId: projectItm[0].ID,
                          OpportunityNumber: projectItm[0].OpportunityNumber,
                          ProjectName: projectItm[0].ProjectName,
                          ClientOrganization: projectItm[0].ClientOrganization,
                          ProjectType: projectItm[0].ProjectType,
                          AssignedOn: moment(
                            tskItm.FieldValuesAsText.Created
                          ).format("DD/MM/YYYY hh:mm A"),
                          TaskLink:
                            tskItm.Status == "In-Progress"
                              ? contextUrl + "/SitePages/" + tskItm.VIew
                              : [
                                  "Technical Action Awaited",
                                  "Delivery Action Awaited",
                                  "Legal Action Awaited",
                                  "Escalated to management",
                                  "Management Action Awaited",
                                ].indexOf(projectItm[0].Status) != -1
                              ? contextUrl + "/SitePages/" + tskItm.VIew
                              : /*contextUrl +
                              "/SitePages/UpdateContract.aspx?itemid=" +
                              projectItm[0].ID*/
                                "",
                          Status:
                            tskItm.Status == "In-Progress"
                              ? tskItm.Status
                              : projectItm[0].Status,
                          ContractDetail:
                            contextUrl +
                            "/sitePages/contract-details-New.aspx?itemid=" +
                            projectItm[0].ID,
                          Created: tskItm.Created,
                          CreatedDisplay:
                            projectItm[0].FieldValuesAsText.Created.split(
                              " "
                            )[0].split("/")[1] +
                            "/" +
                            projectItm[0].FieldValuesAsText.Created.split(
                              " "
                            )[0].split("/")[0] +
                            "/" +
                            projectItm[0].FieldValuesAsText.Created.split(
                              " "
                            )[0].split("/")[2],
                        });
                      }
                    }
                    resolve(activeItems);
                  },
                  (error: any) => {
                    reject("Error Occured" + error);
                  }
                );
              });
          }
        });
      resolve(activeItems);
    });
  };

  // public getMyTaskItems = (
  //   contextUrl: string,
  //   status: string
  // ): Promise<IMyActiveTasksItems[]> => {
  //   const activeItems: IMyActiveTasksItems[] = [];

  //   return new Promise<IMyActiveTasksItems[]>(async (resolve, reject) => {
  //     await sp.web.currentUser.get().then(async (user) => {
  //       console.log(user);
  //       await sp.web.lists
  //         .getByTitle("ProjectTasks")
  //         .items.select("*,FieldValuesAsText/Created")
  //         .expand("FieldValuesAsText")
  //         .filter(
  //           "AssignedTo eq " + user.Id + " and Status eq '" + status + "'"
  //         )
  //         .orderBy("Created desc")
  //         .getAll()
  //         .then(async (taskItems) => {
  //           console.log(taskItems);

  //           var projectId = taskItems
  //             .map((taskItem) => {
  //               return taskItem.ProjectID;
  //             })
  //             .filter((value, index, self) => {
  //               return self.indexOf(value) === index;
  //             });

  //           var camlInClause = "";

  //           for (
  //             let index = 0;
  //             projectId.slice().splice(index * 5, 5).length > 0;
  //             index++
  //           ) {
  //             var inQuery = projectId
  //               .slice()
  //               .splice(index * 5, 5)
  //               .map((id) => {
  //                 return "<Value Type='Number'>" + id + "</Value>";
  //               });

  //             camlInClause +=
  //               "<In><FieldRef Name='ID' /><Values>" +
  //               inQuery.join("") +
  //               "</Values></In>";

  //             if (index >= 1) {
  //               camlInClause = "<Or>" + camlInClause + "</Or>";
  //             }
  //           }
  //           if (camlInClause.length > 0) {
  //             const camlQuery: ICamlQuery = {};
  //             camlQuery.ViewXml =
  //               "<View><Query><Where>" +
  //               camlInClause +
  //               "</Where></Query></View>";
  //             await sp.web.lists
  //               .getByTitle("Projects")
  //               .getItemsByCAMLQuery(camlQuery, "FieldValuesAsText")
  //               .then((prjItems) => {
  //                 console.log("FieldValuesAsText value");
  //                 console.log(prjItems);
  //                 taskItems.map((tskItm) => {
  //                   var showRecord = false;
  //                   var projectItm = prjItems.filter((itm) => {
  //                     return itm.Id == tskItm.ProjectID;
  //                   });

  //                   if (projectItm.length > 0) {
  //                     if (
  //                       status == "Completed" &&
  //                       projectItm[0].Status != "Project Initiated"
  //                     ) {
  //                       showRecord = true;
  //                     } else if (status == "In-Progress") {
  //                       showRecord = true;
  //                     }

  //                     if (showRecord) {
  //                       // let options ={
  //                       //   hour: 'numeric 2-digit', minute: 'numeric 2-digit', second: 'numeric 2-digit',
  //                       //   day:'2-digit',month:'2-digit',
  //                       //   timeZoneName: 'short'
  //                       // };
  //                       // let ISTtime:any = new Date(projectItm[0].Created.split("Z")[0]+"+05:30");
  //                       // ISTtime = ISTtime.toLocaleString('hi-IN', { hour24: false, options });

  //                       activeItems.push({
  //                         ProjectId: projectItm[0].ID,
  //                         OpportunityNumber: projectItm[0].OpportunityNumber,
  //                         ProjectName: projectItm[0].ProjectName,
  //                         ClientOrganization: projectItm[0].ClientOrganization,
  //                         ProjectType: projectItm[0].ProjectType,
  //                         AssignedOn: tskItm.FieldValuesAsText.Created,
  //                         TaskLink:
  //                           tskItm.Status == "In-Progress"
  //                             ? contextUrl + "/SitePages/" + tskItm.VIew
  //                             : [
  //                                 "Technical Action Awaited",
  //                                 "Delivery Action Awaited",
  //                                 "Legal Action Awaited",
  //                                 "Management Action Awaited",
  //                               ].indexOf(projectItm[0].Status) != -1
  //                             ? contextUrl + "/SitePages/" + tskItm.VIew
  //                             : /*contextUrl +
  //                             "/SitePages/UpdateContract.aspx?itemid=" +
  //                             projectItm[0].ID*/
  //                               "",
  //                         Status:
  //                           tskItm.Status == "In-Progress"
  //                             ? tskItm.Status
  //                             : projectItm[0].Status,
  //                         ContractDetail:
  //                           contextUrl +
  //                           "/sitePages/contract-details-New.aspx?itemid=" +
  //                           projectItm[0].ID,
  //                       });
  //                     }
  //                   }
  //                 });
  //                 resolve(activeItems);

  //                 /*prjItems.map((prjItem) => {
  //                   activeItems.push({
  //                     ProjectId: prjItem.ID,
  //                     OpportunityNumber: prjItem.OpportunityNumber,
  //                     ProjectName: prjItem.ProjectName,
  //                     ClientOrganization: prjItem.ClientOrganization,
  //                     ProjectType: prjItem.ProjectType,
  //                     AssignedOn: "",
  //                   });
  //                 });*/
  //               });
  //           }
  //           ,
  //         });
  //     });
  //     //resolve(activeItems);
  //   });
  // };

  public async getDialogListItems(projectId: number) {
    /*let items = sp.web.lists
      .getByTitle("ProjectTasks")
      .items.select("*,Editor/Title")
      .expand("Editor")
      .filter("ProjectID eq " + projectId)
      .get();
    return items;*/
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
                //   hour: 'numeric', minute: 'numeric', second: 'numeric',
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
                  Status: item[i].Status,
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
    return auditItems;
  }

  // public getMyAcomplishedItems(
  //   contextUrl: string
  // ): Promise<IMyActiveTasksItems[]> {
  //   const activeItems: IMyActiveTasksItems[] = [];

  //   return new Promise<IMyActiveTasksItems[]>(async (resolve, reject) => {
  //     sp.web.currentUser.get().then((user) => {
  //       console.log(user);
  //       sp.web.lists
  //         .getByTitle("ProjectTasks")
  //         .items.select("*")
  //         .filter("AssignedTo eq " + user.Id + " and Status eq 'Completed'")
  //         .orderBy("Created desc")
  //         .getAll()
  //         .then((taskItems) => {
  //           console.log(taskItems);

  //           var projectId = taskItems
  //             .map((taskItem) => {
  //               return taskItem.ProjectID;
  //             })
  //             .filter((value, index, self) => {
  //               return self.indexOf(value) === index;
  //             });

  //           var camlInClause = "";

  //           for (
  //             let index = 0;
  //             projectId.slice().splice(index * 5, 5).length > 0;
  //             index++
  //           ) {
  //             var inQuery = projectId
  //               .slice()
  //               .splice(index * 5, 5)
  //               .map((id) => {
  //                 return "<Value Type='Number'>" + id + "</Value>";
  //               });

  //             camlInClause +=
  //               "<In><FieldRef Name='ID' /><Values>" +
  //               inQuery.join("") +
  //               "</Values></In>";

  //             if (index >= 1) {
  //               camlInClause = "<Or>" + camlInClause + "</Or>";
  //             }
  //           }
  //           if (camlInClause.length > 0) {
  //             const camlQuery: ICamlQuery = {};
  //             camlQuery.ViewXml =
  //               "<View><Query><Where>" +
  //               camlInClause +
  //               "</Where></Query></View>";
  //             sp.web.lists
  //               .getByTitle("Projects")
  //               .getItemsByCAMLQuery(camlQuery)
  //               .then((prjItems) => {
  //                 console.log(prjItems);
  //                 taskItems.map((tskItm) => {
  //                   var projectItm = prjItems.filter((itm) => {
  //                     return itm.Id == tskItm.ProjectID;
  //                   });

  //                   if (projectItm.length > 0) {
  //                     activeItems.push({
  //                       ProjectId: projectItm[0].ID,
  //                       OpportunityNumber: projectItm[0].OpportunityNumber,
  //                       ProjectName: projectItm[0].ProjectName,
  //                       ClientOrganization: projectItm[0].ClientOrganization,
  //                       ProjectType: projectItm[0].ProjectType,
  //                       AssignedOn: tskItm.Created.split("T")[0],
  //                       TaskLink: "",
  //                       Status: tskItm.Status,
  //                       ContractDetail:
  //                         contextUrl +
  //                         "/sitePages/contract-details-new.aspx?itemid=" +
  //                         projectItm[0].ID,
  //                     });
  //                   }
  //                 });
  //               });
  //           }
  //         });
  //     });
  //     resolve(activeItems);
  //   });
  // }
}
