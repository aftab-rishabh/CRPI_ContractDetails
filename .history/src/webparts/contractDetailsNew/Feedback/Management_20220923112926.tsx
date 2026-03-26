import * as React from 'react';
import styles from './Feedback.module.scss';

import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
// import { CurrentUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import * as $ from 'jquery';
/**
 * Icon styles. Feel free to change them
 */


 export interface IManagementProps {
    defaultCollapsed?: boolean;    
    className?: string;
    webURL:string;
  }
  
  export interface IManagementState {
    expanded: boolean;
    Items: any; 
    HTML: any;
  }
  
// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Management extends React.Component<IManagementProps, IManagementState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ManagementFlag=false;
  constructor(props: IManagementProps) {
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
  public render(): React.ReactElement<IManagementProps> {
    return (
      
      <div className={styles.Feedback}>
      <div className={styles.container}>
          <div className={styles.row}>
                <div className={styles.viewdetail_block}>

                  <div className="ManagementData">

                      <div className={styles.sectionblock}>
                      <Label className={styles.headers}><u>Management Feedback</u></Label>
                      </div>

                      <div id="managementDetailsSection" style={{display:'none'}}>
                      {this.state.Items ? (
                          this.state.Items.map(function(item,key){  
                         
                           var MgtComments = item.MgtComments;
                           if (MgtComments === null ) {
                            MgtComments = "";
                           }
                           else if(item.FeedbackStatus === "Draft"){
                            MgtComments = "";
                           }
                      
                        return (
                      <div className={styles.sectionblock} key={key} >
                          <div className={styles.viewanswer_listing}>
                              <ul>
                          
                          <li>
                              <div className={styles.viewcomment}>Management comment:</div>
                      <div className={styles.viewanswer} id="LegalComment">{MgtComments}</div>
                  </li>
                 
        
      </ul>
  </div>
</div>

                 ); 
               })
               ) :"Management feedback is not available"}  
                          </div> 

                  </div>

                </div> 
          </div>
      </div>
  </div>
    );
  }

  public async fetchData() {
    // this.getloggedInUser();
    this.GetManagementTeam("");
  }

  private async setData() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects")
    .items.getById(parseInt(itemID)).select('*', 'DeliveryManager/Title', 'AccountManager/Title').expand('DeliveryManager/Id', 'AccountManager/Id').get().then((item) => {
      
        this.setState({ Items: item });
        // if (item.WithTemplate == "Yes") {
        //   document.getElementById("managementDetailsSection").style.display = "none";
        // }
  
    }).then(x => {
       
      // this.HideNAData();
    });
  }

  private async GetManagementTeam(CurrentUserEmail) {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Management Team")
      .items.select('*', 'Manager/Title', 'Manager/EMail').expand('Manager').get().then((data) => {        
        for (var i = 0; i < data.length; i++) {
          // if(CurrentUserEmail.toLowerCase() == data[i].Manager.EMail.toLowerCase()){
          //   this.ManagementFlag = true;
          // }
        }
        // if(this.ManagementFlag){
          $("#managementDetailsSection").show();
          this.setData();
          
        // }
       
      });
  }
//Get Current User Display Name  
// private getloggedInUser(): void {  
//   let web = Web(this.props.webURL);    
//   web.currentUser.get().then((r: CurrentUser) => {  
//     console.log(r['UserPrincipalName']); 
//     // this.GetManagementTeam(r['UserPrincipalName']);
//   });  
// } 


  
}


