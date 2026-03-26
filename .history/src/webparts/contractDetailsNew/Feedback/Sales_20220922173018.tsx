import * as React from 'react';
import styles from './Feedback.module.scss';

import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';

import * as $ from 'jquery';
/**
 * Icon styles. Feel free to change them
 */


 export interface ISalesProps {
    defaultCollapsed?: boolean;    
    className?: string;
    webURL:string;
  }
  
  export interface ISalesState {
    expanded: boolean;
    Items: any; 
    HTML: any;
  }
  
// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Sales extends React.Component<ISalesProps, ISalesState> {
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: ISalesProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {
    this.fetchData();
    
  }
  public render(): React.ReactElement<ISalesProps> {
    return (
      
    <div className={styles.Feedback}>
        <div className={styles.container}>
            <div className={styles.row}>
                  <div className={styles.viewdetail_block}>

                    <div className="SalesData">

                        <div className={styles.sectionblock}>
                        <Label className={styles.headers}><u>Sales Feedback</u></Label>
                        </div>

                        <div id="BDMQuestionContainer">
                      
                       
                            </div> 

                    </div>

                  </div> 
            </div>
        </div>
    </div>
     
    );
  }

  public async fetchData() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    // const items: any[] = await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select("*", 'SalesManager/Title', 'AccountManager/Title').expand('SalesManager/Id', 'AccountManager/Id').get();
    // console.log(items);
    // this.setState({ Items: items });
    this.SetQuestions();
    
  }

  // private async getYesNoAndComment() {
  //   // const itemID = new URLSearchParams(window.location.search).get('itemid');
  //   let web = Web(this.props.webURL);
  //   await web.lists.getByTitle("SalesFeedback").items.filter("Team eq 'Sales'").orderBy('Sequence').get().then(async (item) => {
      
  //       this.setState({ Items: item });
       
  //   }).then(x => {
       
  //     this.setAnwers();
  //   });
  // }

 
  private async SetQuestions() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    var iLength;
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("SalesFeedback").items.orderBy('QuestionNumber').filter('ProjectID eq ' + parseInt(itemID)).get().then((item) => {
        iLength = item.length;

        for (var i = 0; i < iLength; i++) {
            var sectionblockx = styles.sectionblockx;
            var viewanswer_listing = styles.viewanswer_listing;
            var viewquestion = styles.viewquestion;
            var viewanswer = styles.viewanswer;
            var viewcomment = styles.viewcomment;
            var ID = item[i].ID;
            var html = "";


            html += '<div class="' + sectionblockx + '">'
                + '<div class="' + viewanswer_listing + '">'
                + '<ul>'
                + '<li>'
                + '<div class="' + viewquestion + '">' + item[i].Title + '</div>'
                + '<lable class="' + viewcomment + '">Answer:<lable class="' + viewanswer + '" id = "CRAnswer' + ID + '"><b>' + item[i].Answer + '</b></lable></lable>'
                + '<div class="CRdiv' + ID + '" id="CRdiv' + ID + '" style="display:block">'
                + '<lable class="' + viewcomment + '">Comment:</lable>'
                + '<lable class="' + viewanswer + '" id="crcomment' + ID + '">' + item[i].Comment + '</lable>'
                + '</div>'
                + '</li>'
                + '</ul>'
                + '</div>'
                + '</div>';

            $("#BDMQuestionContainer").append(html);

            if (item[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[i].Answer == "No") {
                document.getElementById("CRAnswer" + ID).style.color = "red";
                document.getElementById("CRdiv" + ID).style.display = "block";

            }
            else if (item[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[i].Answer == "Yes") {
                document.getElementById("CRAnswer" + ID).style.color = "red";
                document.getElementById("CRdiv" + ID).style.display = "block";
            }

            if (item[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[i].Answer == "Yes") {
                document.getElementById("CRAnswer" + ID).style.color = "green";
                document.getElementById("CRdiv" + ID).style.display = "none";
            }
            else if (item[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[i].Answer == "No") {
                document.getElementById("CRAnswer" + ID).style.color = "green";
                document.getElementById("CRdiv" + ID).style.display = "none";
            }

            if (item[i].Requirement.toLowerCase() == "textbox".toLowerCase()) {
                document.getElementById("CRdiv" + ID).style.display = "block";
            }

            if (item[i].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[i].Answer == "NA") {
                document.getElementById("CRAnswer" + ID).style.color = "green";
                document.getElementById("CRdiv" + ID).style.display = "none";
            }

            if (item[i].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[i].Answer == "NA") {
                document.getElementById("CRAnswer" + ID).style.color = "green";
                document.getElementById("CRdiv" + ID).style.display = "none";
            }



            if (item[i].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
                document.getElementById("CRAnswer" + ID).style.color = "orange";
                document.getElementById("CRdiv" + ID).style.display = "none";
            }

        }

    });
}

  
}


