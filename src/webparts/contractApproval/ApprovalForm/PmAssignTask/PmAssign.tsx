// import * as React from 'react';
// import styles1 from '../../components/ContractApproval.module.scss';
// import styles from './PmAssignTaskFormWebPart.module.scss';
// import { css } from "@uifabric/utilities/lib/css";
// import { DefaultButton, FontWeights, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
// import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { sp, Web, IWeb } from "@pnp/sp/presets/all";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import { Label } from 'office-ui-fabric-react';
// import { SPComponentLoader } from '@microsoft/sp-loader';
// import * as $ from 'jquery';
// import * as moment from 'moment';
// import { Dropdown, TextField } from '@fluentui/react';
// import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

// require('.././css/jquery-ui.css');
// let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
// SPComponentLoader.loadCss(cssURL);
// SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
// require('../scripts/sp.peoplepicker.js');
// require('../scripts/app.js');
// /**
//  * Icon styles. Feel free to change them
//  */


// export interface IPmAssignProps {
//   defaultCollapsed?: boolean;
//   className?: string;
//   webURL: string;
//   context: WebPartContext;
// }

// export interface IPmAssignState {
//   expanded: boolean;
//   Items: any;
//   HTML: any;
//   selectedUsers: any[];
// }



// declare var SP: any;
// declare var SPClientPeoplePicker_InitStandaloneControlWrapper: any;
// declare var SPClientPeoplePicker: any;

// SPComponentLoader.loadCss('/_layouts/15/1033/styles/corev15.css');

// SPComponentLoader.loadScript('/_layouts/15/init.js', {
//   globalExportsName: '$_global_init'
// })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
//       globalExportsName: 'Sys'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us', {
//       globalExportsName: 'Sys'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/sp.init.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/1033/strings.js', {
//       globalExportsName: 'Strings'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/sp.ui.dialog.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/clienttemplates.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/clientforms.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/clientpeoplepicker.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/autofill.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then((): Promise<{}> => {
//     return SPComponentLoader.loadScript('/_layouts/15/sp.core.js', {
//       globalExportsName: 'SP'
//     });
//   })
//   .then(() => {
   
//   });


// // const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// // const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

// export class PmAssignWebPart extends React.Component<IPmAssignProps, IPmAssignState> {

  
//   [x: string]: any;
//   private _drawerDiv: HTMLDivElement = undefined;

// private peoplePickerContext: IPeoplePickerContext;

//   ContractStatus = "";
//   OpportunityID = "";
//   constructor(props: IPmAssignProps) {
//     super(props);
    

//     this.state = {
//       expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
//       Items: [],
//       HTML: [],
//      selectedUsers: []
//     };

//         this.peoplePickerContext = {
//       absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
//       msGraphClientFactory: this.props.context.msGraphClientFactory as any,
//       spHttpClient: this.props.context.spHttpClient as any
//     };

//   }


  

//   public async componentDidMount() {
//     SP.SOD.executeOrDelayUntilScriptLoaded(() => {
//       var schema = {};
//       schema['PrincipalAccountType'] = 'User,DL,SecGroup,SPGroup';
//       schema['SearchPrincipalSource'] = 15;
//       schema['ResolvePrincipalSource'] = 15;
//       schema['maxSelectedUsers'] = 1;
//       schema['AllowMultipleValues'] = false;
//       schema['MaximumEntitySuggestions'] = 10;
//       schema['Width'] = '280px';
//       SPClientPeoplePicker_InitStandaloneControlWrapper("_UserName", null, schema);
//     }, 'clientpeoplepicker.js');
//     this.setForm();
   

//   }
//   public render(): React.ReactElement<IPmAssignProps> {
//     return (
//       <div className={styles.PmAssign}>
//         <div className={styles.container}>
//           <div className={styles.row} style={{ paddingTop: '0px' }}>

          

//           <div id="pm-assignment-awaited-buttons" style={{ display: 'none' }}>
//               <div className={styles.sectionblockPMAssignment}>
//               <Label className={styles.headersPMAssignment} style={{ marginTop: '0px' }}><u>Select Project Manager</u></Label>
//             </div>

//             <div className="row">
//                 <div className="col-lg-12 ">
//                     <Label className={styles.lablecontrol}>Project Manager: <span className={styles.estric}>*</span></Label><br/>
//                     <div id="_UserName"></div>
// <PeoplePicker
//   context={this.peoplePickerContext}
//   titleText="Select Users"
//   personSelectionLimit={3}
//   showtooltip={true}
//   //isRequired={true}
//   disabled={false}
//   onChange={(items) => this.setState({ selectedUsers: items })}
//   //showHiddenInUI={false}
//   principalTypes={[PrincipalType.User]}
//   resolveDelay={1000}
// />
//                     <Label className={styles.errorlable} id="projectManagerddErr" style={{ display: 'none' }}
//                         >This field is required.</Label>
//                     <Label className={styles.errorlable} id="projectManagerddErr2" style={{ display: 'none' }}
//                         >PM is allocated by other Delivery Manager.<span id="DMName"></span></Label>    
//                 </div>
//             </div>

   
//             <div id="loader" className={styles.modal} >
//                   <div className="">
//                     <div className={styles.loader} style={{ margin: '200px auto' }}></div>
//                   </div>
//                 </div>


//             <div id="SubmittedModal" className={styles.modal}>
//                   <div className={styles.modalcontent}>
//                     <span className={styles.close} onClick={() => this.Close1()}>
//                       &times;
//                     </span>
//                     <Label className={styles.header2}>PM Assigned!
//                     </Label>
//                   </div>
//                 </div>

//                 <div className="row">
//                   <div className="col col-lg-12">
//                   <div className={styles.form_footer}>
//                     <PrimaryButton text="Assign" className={styles.btn} id="Assign" onClick={() => this.AssignPM()}></PrimaryButton>
               
//                   </div></div>
//             </div><br/><br/><br/><br/><br/><br/><br/>



//           </div>
//         </div >
//       </div >
//       </div >
//     );
//   }


//   private setForm() {

//     const itemID = new URLSearchParams(window.location.search).get('itemid');
//     let web = Web(this.props.webURL);
//       web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'Status', 'OpportunityID').get().then((data) => {

//       if (data.Status == "PM Assignment Awaited") {
//         document.getElementById("pm-assignment-awaited-buttons").style.display = "block";
//       }
//     });
//   }



//   private AssignPM() {

//     var isvalid = true;
//     const itemID = new URLSearchParams(window.location.search).get('itemid');
//     let web = Web(this.props.webURL);
//     web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((data) => {
//       if (data.ProjectManagerId) {
//         $("#projectManagerddErr").hide();
//         $("#projectManagerddErr2").show();
//         isvalid = false;
//         window.stop();
//       }
//     });

//     if ($('#_UserName')[0].innerText.split("x").length !== 2) {
//       isvalid = false;
//       window.stop();
//     }

//     if ($('#_UserName')[0].innerText.split("x").length == 1) {
//       $("#projectManagerddErr2").hide();
//       $("#projectManagerddErr").show();
//       isvalid = false;
//       window.stop();
//     }
//     else
//       $("#projectManagerddErr").hide();

//     if (isvalid == true) {
//       this.AssignPM2();
//     }

//   }

//   private async AssignPM2() {
//     var _this=this;
//     const itemID = new URLSearchParams(window.location.search).get('itemid');
//     var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict._UserName_TopSpan;
//     //console.log(peoplePicker.GetAllUserKeys());
//     var users = peoplePicker.GetAllUserInfo();
//     console.log(users[0].EntityData.Email);
//     let web = Web(_this.props.webURL);
//     web.ensureUser(users[0].EntityData.Email).then(function (result) {
      
//       var PM = result.data.Id;
//       let web = Web(_this.props.webURL);
//       web.lists.getByTitle("Projects")
//         .items.getById(parseInt(itemID)).update({
//           Status: `PM Action Awaited`,
//           ProjectManagerId: PM,
//         }).then(i => {
//           let web = Web(_this.props.webURL);
//       web.lists.getByTitle('CompleteTask').items.add({
//             ProjectID: parseInt(itemID),
//             TaskType: "Delivery Team"
//           }).then(newListItem => {
//             $("#loader").show();
//             setTimeout(() => { $('#loader').hide(); }, 3000);
//             setTimeout(() => { $('#SubmittedModal').show(); }, 3500);
//           });
//         });

//       // pnp.sp.web.siteUsers.getByEmail(users[0].EntityData.Email).get().then(function (result) {
//       //   console.log(result);
//       // });
//       //  var userDetails = this.GetUserId(encodeURIComponent(users[0].Key));
//     }).catch(function (err) {
//       console.log(err);
//       alert("Unable to assign PM at this time. Please try again.");
//     });
//   }

//   private fetchUserId(loginName) {
//     var context = new SP.ClientContext.get_current();
//     this.txtuser = context.get_web().ensureUser(loginName);
//     context.load(this.txtuser);
//     context.executeQueryAsync(
//       function (sender, args) { this.txtuserid = this.txtuser.get_id(); return this.txtuser.get_id(); },
//       function (sender, args) { console.log(args.get_message()); return ""; }
//     );
//   }

//   private GetUserId(userName) {
//     let web = Web(this.props.webURL);
//     var siteUrl = web;
//     console.log(siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v='" + userName + "'");

//     var call = $.ajax({
//       //url: siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",
//       url: siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v='" + userName + "'",
//       method: "GET",
//       headers: { "Accept": "application/json; odata=verbose" },
//       async: false,
//       dataType: 'json'
//     }).responseJSON;
//     return call;
//   }


//   private Close1() {
//     var modal = document.getElementById("SubmittedModal");
//     var hostUrl = this.props.webURL;
//     modal.style.display = "none";
//     $('#LegalTaskForm').fadeOut(2500);
//     setTimeout(() => { window.parent.location.href = hostUrl; }, 2000);
//   }



// }


import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './PmAssignTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, FontWeights, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import * as moment from 'moment';
import { Dropdown, TextField } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';


require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
require('../scripts/sp.peoplepicker.js');
require('../scripts/app.js');
/**
 * Icon styles. Feel free to change them
 */


export interface IPmAssignProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
  context: WebPartContext;
}

export interface IPmAssignState {
  expanded: boolean;
  Items: any;
  HTML: any;
  selectedUsers: any[];
}



// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class PmAssignWebPart extends React.Component<IPmAssignProps, IPmAssignState> {

  
  [x: string]: any;
  private _drawerDiv: HTMLDivElement = undefined;



  ContractStatus = "";
  OpportunityID = "";
  constructor(props: IPmAssignProps) {
    super(props);
    

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: [],
     selectedUsers: []
    };


  }


  

  public async componentDidMount() {
    this.setForm();
   

  }
  public render(): React.ReactElement<IPmAssignProps> {
    return (
      <div className={styles.PmAssign}>
        <div className={styles.container}>
          <div className={styles.row} style={{ paddingTop: '0px' }}>

          

          <div id="pm-assignment-awaited-buttons" style={{ display: 'none' }}>
              <div className={styles.sectionblockPMAssignment}>
              <Label className={styles.headersPMAssignment} style={{ marginTop: '0px' }}><u>Select Project Manager</u></Label>
            </div>

            <div className="row">
                <div className="col-lg-12 ">
                    <Label className={styles.lablecontrol}>Project Manager: <span className={styles.estric}>*</span></Label><br/>
                    <div id="_UserName"></div>
<PeoplePicker
  context={this.props.context as any}
  titleText="Select Users"
  personSelectionLimit={1}
  showtooltip={true}
  //isRequired={true}
  disabled={false}
  onChange={(items) => this.setState({ selectedUsers: items })}
  //showHiddenInUI={false}
  principalTypes={[PrincipalType.User]}
  resolveDelay={1000}
/>
                    <Label className={styles.errorlable} id="projectManagerddErr" style={{ display: 'none' }}
                        >This field is required.</Label>
                    <Label className={styles.errorlable} id="projectManagerddErr2" style={{ display: 'none' }}
                        >PM is allocated by other Delivery Manager.<span id="DMName"></span></Label>    
                </div>
            </div>

   
            <div id="loader" className={styles.modal} >
                  <div className="">
                    <div className={styles.loader} style={{ margin: '200px auto' }}></div>
                  </div>
                </div>


            <div id="SubmittedModal" className={styles.modal}>
                  <div className={styles.modalcontent}>
                    <span className={styles.close} onClick={() => this.Close1()}>
                      &times;
                    </span>
                    <Label className={styles.header2}>PM Assigned!
                    </Label>
                  </div>
                </div>

                <div className="row">
                  <div className="col col-lg-12">
                  <div className={styles.form_footer}>
                    <PrimaryButton text="Assign" className={styles.btn} id="Assign" onClick={() => this.AssignPM()}></PrimaryButton>
               
                  </div></div>
            </div><br/><br/><br/><br/><br/><br/><br/>



          </div>
        </div >
      </div >
      </div >
    );
  }


  private setForm() {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
      web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'Status', 'OpportunityID').get().then((data) => {

      if (data.Status == "PM Assignment Awaited") {
        document.getElementById("pm-assignment-awaited-buttons").style.display = "block";
      }
    });
  }


private AssignPM() {

    var isvalid = true;
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);

    web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((data) => {
      if (data.ProjectManagerId) {
        $("#projectManagerddErr").hide();
        $("#projectManagerddErr2").show();
        isvalid = false;
        window.stop();
      }
    });

    // ✅ Updated validation
    if (!this.state.selectedUsers || this.state.selectedUsers.length === 0) {
      $("#projectManagerddErr2").hide();
      $("#projectManagerddErr").show();
      isvalid = false;
      window.stop();
    } else {
      $("#projectManagerddErr").hide();
    }

    if (isvalid == true) {
      this.AssignPM2();
    }
  }

  private async AssignPM2() {

    const selectedUser = this.state.selectedUsers[0];
    if (!selectedUser) return;

    const email = selectedUser.secondaryText;
    console.log(email);

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);

    try {
      const result = await web.ensureUser(email);
      const PM = result.data.Id;

      await web.lists.getByTitle("Projects")
        .items.getById(parseInt(itemID))
        .update({
          Status: `PM Action Awaited`,
          ProjectManagerId: PM,
        });

      await web.lists.getByTitle('CompleteTask').items.add({
        ProjectID: parseInt(itemID),
        TaskType: "Delivery Team"
      });

      $("#loader").show();
      setTimeout(() => $('#loader').hide(), 3000);
      setTimeout(() => $('#SubmittedModal').show(), 3500);

    } catch (err) {
      console.log(err);
      alert("Unable to assign PM at this time. Please try again.");
    }
  }

  private Close1() {
    var modal = document.getElementById("SubmittedModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#LegalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 2000);
  }
}