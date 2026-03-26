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


 export interface IDeliveryProps {
    defaultCollapsed?: boolean;    
    className?: string;
    webURL:string;
  }
  
  export interface IDeliveryState {
    expanded: boolean;
    Items: any; 
    HTML: any;
  }
  
// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Delivery extends React.Component<IDeliveryProps, IDeliveryState> {
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: IDeliveryProps) {
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
  public render(): React.ReactElement<IDeliveryProps> {
    return (
      
    <div className={styles.Feedback}>
        <div className={styles.container}>
            <div className={styles.row}>
                  <div className={styles.viewdetail_block}>

                    <div className="DeliveryData">

                        <div className={styles.sectionblock}>
                        <Label className={styles.headers}><u>Delivery Feedback</u></Label>
                        </div>

                        <div id="DELQuestionContainer">
                      
                        {this.state.Items.map(function(item,key){  
                              var viewquestionDiv = 'QuestionDiv' + item.QuestionNumber;  
                              var viewquestionNumber=  item.QuestionNumber; 
                              var viewquestion = styles.viewquestion; 
                          return (
                        <div className={styles.sectionblock} key={key}  id={viewquestionDiv}>
                            <div className={styles.viewanswer_listing}>
                                <ul>
                            
                            <li>
                                <div className={styles.viewquestion}>{item.Question}</div>
                        <div className={styles.viewquestion} id={viewquestionNumber}></div>
                    </li>
                   
          
        </ul>
    </div>
</div>

                   ); 
                 })}  
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
    this.getYesNoAndComment();
    
  }

  private async getYesNoAndComment() {
    // const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").orderBy('Sequence').get().then(async (item) => {
      
        this.setState({ Items: item });
       
    }).then(x => {
       
      this.setAnwers();
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
    await web.lists.getByTitle("DeliveryFeedback").items.orderBy('QuestionNumber').filter(filterStr).get().then(async (item) => {
      if(item.length !== 0){
        for (var j = 0; j < item.length; j++) {
          var html2 = "";
          html2 += '<div class="' + sectionblockx + '">'
            + '<div class="' + viewanswer_listing + '">'
            + '<ul>'
            + '<li>'
            + '<lable class="' + viewcomment + '">Delivery Answer:<lable class="' + viewanswer + '" id = "DeliveryAnswer' + item[j].ID + '"><b>' + item[j].Answer + '</b></lable></lable>'
            + '<div class="Deliverydiv' + item[j].ID + '" id="Deliverydiv' + item[j].ID + '" style="display:block">'
            + '<lable class="' + viewcomment + '">Delivery Comment:</lable>'
            + '<lable class="' + viewanswer + '" id="Deliverycomment' + item[j].ID + '">' + item[j].Comment + '</lable>'
            + '</div>'
            + '</li>'
            + '</ul>'
            + '</div>'
            + '</div>';
          var xyz = $("#" + item[j].QuestionNumber);
          xyz.append(html2);
  
          if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "No") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "red";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "block";
          }
          else if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "Yes") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "red";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "block";
          }
  
          if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "Yes") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "green";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "none";
          }
          else if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "No") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "green";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "none";
          }
  
          if (item[j].Requirement.toLowerCase() == "textbox".toLowerCase()) {
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "block";
          }
  
          if (item[j].Requirement.toLowerCase() == "IfNoComment".toLowerCase() && item[j].Answer == "NA") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "green";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "none";
          }

          if (item[j].Requirement.toLowerCase() == "IfYesComment".toLowerCase() && item[j].Answer == "NA") {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "green";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "none";
          }
 
          if (item[j].Requirement.toLowerCase() == "IfYesNoCommentYes".toLowerCase() || item[j].Requirement.toLowerCase() == "IfYesNoCommentNo".toLowerCase()) {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "red";
          } 
  
          if (item[j].Requirement.toLowerCase() == "Dropdown".toLowerCase()) {
            document.getElementById("DeliveryAnswer" + item[j].ID).style.color = "orange";
            document.getElementById("Deliverydiv" + item[j].ID).style.display = "none";
          }
        }
          
      }

      else {
        var DelQueLength; 
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("FeedbackQuestions").items.filter("Team eq 'Delivery'").orderBy('Sequence').get().then((items) => {
          DelQueLength = items.length;

          for (var k = 0; k < DelQueLength; k++) {
            var html3 = "";
            html3 += '<div class="' + sectionblockx + '">'
              + '<div class="' + viewanswer_listing + '">'
              + '<ul>'
              + '<li>'
              + '<lable class="' + viewcomment + '">Delivery Answer:<lable class="' + viewanswer + '" id = ""></lable></lable>'
              + '<lable class="' + viewcomment + '">Delivery Comment:</lable>'
              + '</li>'
              + '</ul>'
              + '</div>'
              + '</div>';

            var wxyz = $("#" + items[k].QuestionNumber);
            wxyz.append(html3);

          }
  

        });
      }
    });
  }

  
}


