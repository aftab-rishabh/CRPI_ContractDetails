import * as React from 'react';
import styles from './Feedback.module.scss';

import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
/**
 * Icon styles. Feel free to change them
 */


 export interface ITechnicalProps {
    defaultCollapsed?: boolean;    
    className?: string;
    webURL:string;
  }
  
  export interface ITechnicalState {
    expanded: boolean;
    Items: any; 
    HTML: any;
  }
  
const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Technical extends React.Component<ITechnicalProps, ITechnicalState> {
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: ITechnicalProps) {
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
  public render(): React.ReactElement<ITechnicalProps> {
    return (
      <div className={css(styles.Feedback, this.props.className)}>
        
    <div >
        <div className={styles.container}>
            <div className={styles.row}>
                  <div className={styles.viewdetail_block}>

                    <div className="TechnicalData">

                        <div className={styles.sectionblock}>
                        <Label className={styles.headers}><u>Technical Feedback</u></Label>
                        </div>

                        <div id="TECQuestionContainer">
                        {this.state.HTML}
                            </div> 

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
    // const items: any[] = await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select("*", 'DeliveryManager/Title', 'AccountManager/Title').expand('DeliveryManager/Id', 'AccountManager/Id').get();
    // console.log(items);
    // this.setState({ Items: items });
    let html = await this.getYesNoAndComment();
    this.setState({ HTML: html });
  }

  private async getYesNoAndComment() {
    // const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").orderBy('Sequence').get().then(async (item) => {
        var html = "";
        for (var i = 0; i < item.length; i++) {
        var sectionblockx = styles.sectionblock;
        var viewanswer_listing = styles.viewanswer_listing;
        var viewquestion = styles.viewquestion;
       

        html += '<div className="{' + sectionblockx + ' QuestionDiv' + item[i].QuestionNumber + '}" id="QuestionDiv' + item[i].QuestionNumber + '">'
          + '<div className="' + viewanswer_listing + '">'
          + '<ul>'
          + '<li>'
          + '<div className="{' + viewquestion + '}">' + item[i].Question + '</div>'
          + '<div className="{' + viewquestion + '}" id=' + item[i].QuestionNumber + '></div>'
          + '</li>'
          + '</ul>'
          + '</div>'
          + '</div>';

        // $("#TECQuestionContainer").append(html);
       
      }
      return html;
    }).then(x => {
       
    //   this.setAnwers();
    });
  }

  private async setAnwers() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');

    var sectionblockx = styles.sectionblockx;
    var viewanswer_listing = styles.viewanswer_listing;
    var viewanswer = styles.viewanswer;
    var viewcomment = styles.viewcomment;
    var filterStr = "ProjectID eq '" + parseInt(itemID) + "' and " + "FeedbackStatus eq 'Submit'";
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("TechnicalFeedback").items.filter(filterStr).get().then((item) => {
      for (var j = 0; j < item.length; j++) {
        var Depart = item[j].Department.split("-")[1].split(' ').join('');
        var html2 = "";
        html2 += '<div className="' + sectionblockx + '">'
          + '<div className="' + viewanswer_listing + '">'
          + '<ul>'
          + '<li>'
          + '<Label id="AnserLable' + item[j].QuestionNumber + '' + Depart + '" className="' + viewcomment + '">' + item[j].Department + '&nbspAnswer:<Label className="' + viewanswer + '" id = "Answer' + item[j].QuestionNumber + '' + Depart + '"><b>' + item[j].Answer + '</b></Label></Label>'
          + '<div className="DIV' + item[j].ID + '" id="DIV' + item[j].QuestionNumber + '' + Depart + '" style="display:block">'
          + '<Label className="' + viewcomment + '">' + item[j].Department + '&nbspComment:</Label>'
          + '<Label className="' + viewanswer + '" id="Comment' + item[j].QuestionNumber + '">' + item[j].Comment + '</Label>'
          + '</div>'
          + '</li>'
          + '</ul>'
          + '</div>'
          + '</div>';
        // var xyz = $("#" + item[j].QuestionNumber);
        // xyz.append(html2);

        if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "No") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "red";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "block";

        }
        else if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "Yes") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "red";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "block";
        }

        if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "Yes") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "green";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "none";
        } 
        else if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "No") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "green";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "none";
        }

        if (item[j].Requirement.toLowerCase() == "textbox".toLowerCase()) {
          document.getElementById("AnserLable" + item[j].QuestionNumber + Depart).style.display = "none";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "block";
        }

        if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "NA") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "green";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "none";
        }

        if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "NA") {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "green";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "none";
        }

        if (item[j].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
          document.getElementById("Answer" + item[j].QuestionNumber + Depart).style.color = "orange";
          document.getElementById("DIV" + item[j].QuestionNumber + Depart).style.display = "none";
        }


      }
    }).then(() =>{
      this.notSubmittedTechLable();
    });
  }

  private async notSubmittedTechLable(){
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    const Dpt = new URLSearchParams(window.location.search).get('dpt');
    var filterStr4 = "ProjectID eq '" + itemID + "'";
    var sectionblockx = styles.sectionblockx;
    var viewanswer_listing = styles.viewanswer_listing;
    var viewanswer = styles.viewanswer;
    var viewcomment = styles.viewcomment;
    var lablehtml = "";
 
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("TechnicalComments").items.filter(filterStr4).get().then((item) => {
      for (var j = 0; j < item.length; j++) {
        var Depart = item[j].Department.split("-")[1].split(' ').join('');
        if(item[j].IsFinished == false){
          lablehtml += '<div className="' + sectionblockx + '">'
          + '<div className="' + viewanswer_listing + '">'
          + '<ul>'
          + '<li>'
          + '<lable id="AnserLable'+ Depart + '" className="' + viewcomment + '">' + item[j].Department + '&nbspAnswer:<lable className="' + viewanswer + '" id = "Answer'+ Depart + '"></lable></lable>'
          + '<div className="DIV' + item[j].ID + '" id="DIV'+ Depart + '" style="display:block">'
          + '<lable className="' + viewcomment + '">' + item[j].Department + '&nbspComment:</lable>'
          + '<lable className="' + viewanswer + '" "></lable>'
          + '</div>'
          + '</li>'
          + '</ul>'
          + '</div>'
          + '</div>';
          
        }
      }
    }).then(async () =>{
       
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Technical'").orderBy('Sequence').get().then((item) => {
        for (var i = 0; i < item.length; i++) {
        //   var xyz = $("#" + item[i].QuestionNumber);
        //   xyz.append(lablehtml);
        }
      });  

    });
  }

  
}


