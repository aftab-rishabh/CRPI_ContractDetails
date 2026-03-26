import { sp } from "@pnp/sp/presets/all";

import { IChatboxItems } from "../components/IChatboxItems";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class SPOperations {
  constructor(context: WebPartContext) {
    sp.setup({
      sp: { baseUrl: context.pageContext.web.absoluteUrl },
    });
  }
  public async isUserSales(
    projectID: number,
    userEmail: string
  ): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
      const item: any = await sp.web.lists
        .getByTitle("Projects")
        .items.getById(projectID)
        .select("Author/EMail")
        .expand("Author")();
      console.log(item.Author.EMail);

      resolve(item.Author.EMail === userEmail);
    });
  }

  public async isUserManagementTeam(userEmail: string): Promise<boolean> {
    let isManagement: boolean = false;
    return new Promise<boolean>(async (resolve, reject) => {
      await sp.web.lists
        .getByTitle("Management Team")
        .items.select("Manager/EMail")
        .expand("Manager")
        .getAll()
        .then((items) => {
          console.log(items);
          items.forEach((item) => {
            if (item.Manager.EMail == userEmail) {
              isManagement = true;
            }
          });
        });
      resolve(isManagement);
    });
  }

  public getConversations(projectId: number): Promise<IChatboxItems[]> {
    const result: IChatboxItems[] = [];
    return new Promise<IChatboxItems[]>((resolve, reject) => {
      sp.web.lists
        .getByTitle("ContractNotes")
        .items.select(
          "ID,Notes,InitiatedBy,Created,Author/Title,Author/EMail,FieldValuesAsText/Created"
        )
        .filter("ProjectId eq " + projectId)
        .expand("Author/Id, FieldValuesAsText")
        .orderBy("Created desc")
        .getAll()
        .then((items) => {
          //console.log(items);
          items.map((item) => {
            result.push({
              Notes: item.Notes,
              AddedBy: item.Author.Title,
              AddedByEmail: item.Author.EMail,
              ID: item.ID,
              CreatedDt: item.FieldValuesAsText.Created,
              InitiatedBy: item.InitiatedBy,
            });
          });
        });
      resolve(result);
    });
  }

  public AddEditConversation(
    Id: number,
    note: string,
    projectId: number,
    initiatedBy: string
  ): Promise<string> {
    if (Id == 0) {
      return new Promise<string>(async (resolve, reject) => {
        sp.web.lists
          .getByTitle("ContractNotes")
          .items.add({
            Notes: note,
            ProjectId: projectId,
            InitiatedBy: initiatedBy,
          })
          .then(
            (result: any) => {
              resolve("Item added successfully");
            },
            (error: any) => {
              reject("error occured");
            }
          );
      });
    } else {
      return new Promise<string>(async (resolve, reject) => {
        sp.web.lists
          .getByTitle("ContractNotes")
          .items.getById(Id)
          .update({ Notes: note })
          .then(
            (result: any) => {
              resolve("Item added successfully");
            },
            (error: any) => {
              reject("error occured");
            }
          );
      });
    }
  }

  public DeleteConversation(id: number): Promise<string> {
    return new Promise<string>(async (resolve, reject) => {
      sp.web.lists
        .getByTitle("ContractNotes")
        .items.getById(id)
        .delete()
        .then(
          (result: any) => {
            resolve("Item deleted successfully");
          },
          (error: any) => {
            reject("error occured");
          }
        );
    });
  }
}
