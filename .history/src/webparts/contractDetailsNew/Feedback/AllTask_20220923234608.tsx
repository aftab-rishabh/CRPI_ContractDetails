import * as React from 'react';
import styles from './Feedback.module.scss';

import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';

import * as $ from 'jquery';
/**
 * Icon styles. Feel free to change them
 */


export interface IAllTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface IAllTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class AllTask extends React.Component<IAllTaskProps, IAllTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: IAllTaskProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {
    var UserEmail = this.context.pageContext.user.email;
    this.createTable();

  }
  public render(): React.ReactElement<IAllTaskProps> {
    return (

      <div className={styles.Feedback}>
        <div className={styles.container}>

          <div className={styles.row} style={{ paddingBottom: 0 }}>
            <div className={styles.viewdetail_block} style={{ paddingBottom: 0 }}>

              <div className={styles.sectionblock}>
                <PrimaryButton text="Delete Contract" onClick={() => this.DeleteContract()}> </PrimaryButton>
              </div>

            </div>
            <div className={styles.viewdetail_block}>

              <div className="AllTaskData">

                <div className={styles.sectionblock}>
                  {/* <PrimaryButton text="Delete Contract" onClick={()=> this.DeleteContract()}> </PrimaryButton> */}
                  <Label className={styles.headers}><u>All Task</u></Label>
                </div>
                <div id="myTable">

                </div>


              </div>

            </div>
            <div id="loader" className={styles.modal}>
                    <div className="">
                        <div className={styles.loader} style={{margin:"200px"}}></div>
                    </div>
                </div>

                <div id="ContractDeletedModal" className={styles.modal}>
                <div className={styles.modalcontent}>
                    <span className={styles.close} id="Close0">&times</span>
                    <label className={styles.header2}>Contract Deleted!</label>
                </div>
            </div>
          </div>
        </div>
      </div>

    );
  }

  private createTable() {

    var table;
    var myTableDiv = document.getElementById("myTable");  //indiv
    table = document.createElement("TABLE");   //TABLE??
    table.setAttribute("id", "data");
    table.border = '1';
    myTableDiv.appendChild(table);
    //appendChild() insert it in the document (table --> myTableDiv)

    var header = table.createTHead();

    var th0 = table.tHead.appendChild(document.createElement("th"));
    th0.innerHTML = "Activity Name";
    var th1 = table.tHead.appendChild(document.createElement("th"));
    th1.innerHTML = "Assigned To";
    var th2 = table.tHead.appendChild(document.createElement("th"));
    th2.innerHTML = "Status";


    this.appendRow(table);

  }


  private async appendRow(table) {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("ProjectTasks").items.filter('ProjectID eq ' + itemID).select('*', 'AssignedTo/Title').expand('AssignedTo/Id').get().then((item) => {
      var i;
      for (i = 0; i < item.length; i++) {


        var AName = " " + item[i].Title + " ";
        var ATo = " " + item[i].AssignedTo.Title + " ";
        var Status1 = " " + item[i].Status + " ";

        var rowCount = table.rows.length;
        var row = table.insertRow(rowCount);

        row.insertCell(0).innerHTML = AName;
        row.insertCell(1).innerHTML = ATo;
        row.insertCell(2).innerHTML = Status1;

        var tds = document.getElementById("myTable").getElementsByTagName("td");

        for (var x = 0; x < tds.length; x++) {
          tds[x].style.padding = "5px";
        }

      }
    });


  }


  private DeleteContract() {
    $("#loader").show();
    // const itemID = new URLSearchParams(window.location.search).get('itemid');
    // sp.web.lists.getByTitle("Projects")
    //   .items.getById(parseInt(itemID)).update({
    //     Status: "Deleted",
    //   }).then(async i => {
    //     let web = Web(this.props.webURL);
    //     await web.lists.getByTitle('CompleteTask').items.add({
    //       ProjectID: parseInt(itemID),
    //       TaskType: "Deleted"
    //     }).then(newListItem => {
    //       $("#loader").show();
          setTimeout(() => { $('#loader').hide(); }, 3000);
          setTimeout(() => { $('#ContractDeletedModal').show(); }, 3500);
      //   });
      // });
  }


}


