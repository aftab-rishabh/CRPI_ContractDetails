import * as React from 'react';
import styles from './ContractApproval.module.scss';
import { IContractApprovalProps } from './IContractApprovalProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IIconProps, PrimaryButton } from '@fluentui/react';
import * as $ from "jquery";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { Web } from "@pnp/sp/webs";
import { ProjectDetails } from "../../contractDetailsNew/Feedback/ProjectDetails";
import styles1 from "../../accordion/Accordion.module.scss";
import { css } from "@uifabric/utilities/lib/css";
import { TechnicalTask } from '../ApprovalForm/TechnicalTask/TechnicalTask';
import { Technical } from '../../contractDetailsNew/Feedback/Technical';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { DeliveryTask } from '../ApprovalForm/DeliveryTask/DeliveryTask';
import { LegalTask } from '../ApprovalForm/LegalTask/LegalTask';
import { Management } from '../ApprovalForm/ManagementTask/Management';
import { ContractOwner } from '../ApprovalForm/ContractOwnerTask/ContractOwner';
import { CmDocumentUpload } from '../ApprovalForm/CmDocumentUploadTask/CmDocumentUpload';
import { PmAssignWebPart } from '../ApprovalForm/PmAssignTask/PmAssign';
import { PmActionWebPart } from '../ApprovalForm/PmActionTask/PmAction';
import { KickOffDocUpload } from '../ApprovalForm/KickOffDocTask/KickOffDocTask';
import { Sales } from '../../contractDetailsNew/Feedback/Sales';
import { Legal } from '../../contractDetailsNew/Feedback/Legal';
import { Delivery } from '../../contractDetailsNew/Feedback/Delivery';
// require('bootstrap');
// require('./css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");

const collapsedIcon: IIconProps = {
  iconName: "ChevronRight",
  className: styles1.accordionChevron,
};
const expandedIcon: IIconProps = {
  iconName: "ChevronDown",
  className: styles1.accordionChevron,
};
export interface IContractApprovalState {
  btnProjectDetails: boolean;
  btnSales: boolean;
  btnTechnicalTask: boolean;
  btnDeliveryTask: boolean;
  btnLegalTask: boolean;
  btnManagement: boolean;
  btnexpanded: boolean;
  btnContractOwner: boolean;
  btnCMDocUpload: boolean;
  btnPmAssign: boolean;
  btnPmAction: boolean;
  btnKickOffDocUpload: boolean,

  ProjectDetails: boolean;
  Sales: boolean;
  TechnicalTask: boolean;
  DeliveryTask: boolean;
  LegalTask: boolean;
  ManagementTask: boolean;
  expanded: boolean;
  ContractOwner: boolean;
  CMDocUpload: boolean;
  PmAssign: boolean;
  PmAction: boolean;
  KickOffDocUpload: boolean,
  hideDialog: boolean;

  TechnicalFeedback: boolean;
  DeliveryFeedback: boolean;
  LegalFeedback: boolean;
  ManagementFeedback:boolean;

}
export default class ContractApproval extends React.Component<IContractApprovalProps, IContractApprovalState> {
  private _drawerDiv: HTMLDivElement = undefined;
  currentuseremail: any;
  itemID:any;
  Dpt:any="";
  Action:any="";
   CreatedEmail:any="";
  Status:any="";
  constructor(props: IContractApprovalProps) {
    super(props);

    this.state = {
      btnProjectDetails: true,
      btnSales: false,
      btnTechnicalTask: false,
      btnDeliveryTask: false,
      btnLegalTask: false,
      btnManagement: false,
      btnexpanded: false,
      btnContractOwner: false,
      btnCMDocUpload: false,
      btnPmAssign: false,
      btnPmAction: false,
      btnKickOffDocUpload: false,

      ProjectDetails: true,
      Sales: false,
      TechnicalTask: false,
      DeliveryTask: false,
      LegalTask: false,
      ManagementTask: false,
      ContractOwner: false,
      CMDocUpload: false,
      PmAssign: false,
      PmAction: false,
      KickOffDocUpload: false,
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      hideDialog: true,

      TechnicalFeedback: false,
      DeliveryFeedback: false,
      LegalFeedback: false,
      ManagementFeedback:false

    };
  }
  public async componentDidMount() {

    let userDetails = await this.spLoggedInUserDetails();
    this.currentuseremail = userDetails.Email;
    this.setData();
  }


  public render(): React.ReactElement<IContractApprovalProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div className={styles.contractApproval}>
        <div className={styles.viewdetail_block}>
          <div className={styles.sectionblock}>
            <div className={styles.section_left}>
              <div className={css(styles1.accordion, this.props.className)}>
                <div>
                  <DefaultButton
                    toggle
                    checked={this.state.expanded}
                    text={"Project Details"}
                    style={{ width: "97%", marginBottom: "2px" }}
                    iconProps={
                      this.state.expanded ? expandedIcon : collapsedIcon
                    }
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        ProjectDetails: true,
                        ContractOwner: false,
                        Sales: true,
                        TechnicalTask: false,
                        DeliveryTask: false,
                        LegalTask: false,
                        ManagementTask: false,
                        TechnicalFeedback: false,
                        DeliveryFeedback: false,
                        LegalFeedback:false,
                        ManagementFeedback:false,
                        CMDocUpload: false,
                        PmAssign: false,
                        KickOffDocUpload: false,
                        // ChatBox: false,
                      });
                    }}
                    aria-expanded={this.state.expanded}
                    aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />

                  {this.state.expanded && (
                    <div
                      className={styles1.drawer}
                      ref={(el) => {
                        this._drawerDiv = el;
                      }}
                    >
                      {/* {this.state.HTML} */}
                      {/* <ContractOwner webURL={this.props.webURL}></ContractOwner> */}
                      <ProjectDetails
                        webURL={this.props.webURL}
                        context={this.props.context}
                      ></ProjectDetails>
                    </div>
                  )}
                </div>
              </div>
              {/* <Accordion title={'Project Details'} webURL={this.props.webURL}>

              </Accordion> */}
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnSales && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.Sales}
                      text={"Sales Feedback"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: false,
                          Sales: true,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback: false,
                          DeliveryFeedback: false,
                          LegalFeedback:false,
                          ManagementFeedback:false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnLegalTask && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.LegalTask}
                      text={"Legal Feedback"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        if (this.Status == "Technical Action Awaited") {
                          if(this.Action=="Legal"){
                            this.setState({ btnLegalTask: true, LegalTask:true,btnSales: true, btnTechnicalTask: true});
                          }else{
                          this.setState({ btnTechnicalTask: true, btnSales: true, TechnicalTask: true,LegalFeedback:true });
                          }
                      }
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback: false,
                          DeliveryFeedback: false,
                          LegalFeedback:false,
                          ManagementFeedback:false,
                          LegalTask: true,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnTechnicalTask && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.TechnicalTask || this.state.TechnicalFeedback}
                      text={"Technical Feedback"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        if (this.Status == "Technical Action Awaited") {
                          if(this.Action=="Legal"){
                            this.setState({ btnLegalTask: true, LegalTask:false,btnSales: true, btnTechnicalTask: true,LegalFeedback:false});
                          }else{
                          this.setState({ btnTechnicalTask: true, btnSales: true, TechnicalTask: true,LegalFeedback:false });
                          }
                        }
                        else if (this.Status == "Delivery Action Awaited") {
                          if (this.Dpt != null) {
                            this.setState({ btnDeliveryTask: false, DeliveryTask: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: true });
                          }
                          else if(this.Action=="Legal"){
                            this.setState({ btnLegalTask: true,LegalTask:false,btnManagement: false, btnSales: true,btnTechnicalTask: true,btnDeliveryTask:true,TechnicalTask: true });
                          }
                          else {
                            this.setState({ btnDeliveryTask: true, DeliveryTask: false, btnSales: true,btnTechnicalTask: true,TechnicalFeedback: true,  });
                          }
                        }else if (this.Status == "Management Action Awaited" || this.Status == "Escalated to management"){
                          if (this.Dpt != null) {
                            this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: true,btnDeliveryTask: true,DeliveryFeedback:false });
                          }else if(this.Action=="Legal"){
                            this.setState({ btnLegalTask: true,btnManagement: false, btnSales: true,btnTechnicalTask: true,btnDeliveryTask:true,TechnicalTask: true  });
                          }
                          else if (this.Action == "Delivery") {
                            this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,DeliveryFeedback:false,TechnicalFeedback:true });
                          }else{
                            this.setState({  DeliveryFeedback: false,TechnicalTask: false,TechnicalFeedback:true,btnDeliveryTask: true, DeliveryTask: false,ManagementFeedback:false,ManagementTask:false  }); 
                          }
                        }else{
                          this.setState({  DeliveryFeedback: false,TechnicalTask: false,TechnicalFeedback:true,btnDeliveryTask: true, DeliveryTask: false,ManagementFeedback:false,ManagementTask:false  }); 
                        }
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          DeliveryTask: false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnDeliveryTask && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.DeliveryTask || this.state.DeliveryFeedback}
                      text={"Delivery Feedback"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        if (this.Status == "Delivery Action Awaited") {
                            this.setState({  DeliveryTask: true, btnSales: true,btnTechnicalTask: true,  }); 
                        }
                        else if (this.Status == "Management Action Awaited" || this.Status == "Escalated to management"){
                          if (this.Dpt != null) {
                            this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: false,btnDeliveryTask: true ,DeliveryFeedback:true});
                          }
                          else if(this.Action=="Legal"){
                            this.setState({ btnLegalTask: true,LegalTask:false,DeliveryTask:true,btnManagement: false, btnSales: true,btnTechnicalTask: true,btnDeliveryTask:true  });
                          }
                          else if (this.Action == "Delivery") {
                            this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: false,TechnicalFeedback:false,btnDeliveryTask: true, DeliveryTask: true });
                          }else{
                          this.setState({  DeliveryFeedback: true,TechnicalTask: false,TechnicalFeedback:false,btnDeliveryTask: true, DeliveryTask: false,ManagementFeedback:false,ManagementTask:false  }); 
                          }
                        }else{
                          this.setState({  DeliveryFeedback: true,TechnicalTask: false,TechnicalFeedback:false,btnDeliveryTask: true, DeliveryTask: false,ManagementFeedback:false,ManagementTask:false  }); 
                        }
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnManagement && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.ManagementTask || this.state.ManagementFeedback}
                      text={"Management Feedback"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        // if (this.Status == "Management Action Awaited"){

                        //   this.setState({  DeliveryFeedback: false,TechnicalTask: false,TechnicalFeedback:false,btnDeliveryTask: true, DeliveryTask: false,ManagementFeedback:false,ManagementTask:false  }); 
                          
                        // }
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          TechnicalFeedback:false,
                          DeliveryTask: false,
                          DeliveryFeedback: false,
                          LegalFeedback:false,
                          LegalTask: false,
                          ManagementTask: true,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnContractOwner && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.ContractOwner}
                      text={"Contract Owners Form"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: true,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          LegalTask: false,
                          TechnicalFeedback:false,
                          DeliveryFeedback:false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnCMDocUpload && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.CMDocUpload}
                      text={"CM Document Upload"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback:false,
                          DeliveryFeedback:false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: true,
                          PmAssign: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnPmAssign && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.PmAssign}
                      text={"PM Assign Task"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback:false,
                          DeliveryFeedback:false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: true,
                          PmAction: false,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnPmAction && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.PmAction}
                      text={"PM Action Task"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback:false,
                          DeliveryFeedback:false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: true,
                          KickOffDocUpload: false,
                          // ChatBox: false,
                        });
                      }}
                    // aria-expanded={this.state.expanded}
                    // aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                  </div>
                )}
              </div>
              <div className={css(styles1.accordion, this.props.className)}>
                {this.state.btnKickOffDocUpload && (
                  <div>
                    <DefaultButton
                      toggle
                      checked={this.state.KickOffDocUpload}
                      text={"Kick Off Document Upload"}
                      style={{ width: "97%", marginBottom: "2px" }}
                      // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                      onClick={(e) => {
                        this.setState({
                          ContractOwner: false,
                          Sales: false,
                          TechnicalTask: false,
                          DeliveryTask: false,
                          TechnicalFeedback:false,
                          DeliveryFeedback:false,
                          LegalTask: false,
                          ManagementTask: false,
                          CMDocUpload: false,
                          PmAssign: false,
                          PmAction: false,
                          KickOffDocUpload: true,
                          // ChatBox: false,
                        });
                      }}

                    />
                  </div>
                )}
              </div>

            </div>
            <div className={styles.section_right}>
              {/* {this.state.Sales && <Sales webURL={this.props.webURL}></Sales>} */}
              {this.state.Sales && (
                <Sales webURL={this.props.webURL}></Sales>
              )}
              {this.state.TechnicalTask && (
                <TechnicalTask webURL={this.props.webURL}></TechnicalTask>
              )}
              {this.state.DeliveryTask && (
                <DeliveryTask webURL={this.props.webURL}></DeliveryTask>
              )}
              {this.state.LegalTask && (
                <LegalTask webURL={this.props.webURL}></LegalTask>)}
              {this.state.ManagementTask && (
                <Management webURL={this.props.webURL}></Management>
              )}
              {this.state.ContractOwner && (
                <ContractOwner webURL={this.props.webURL}></ContractOwner>
              )}
              {this.state.CMDocUpload && (
                <CmDocumentUpload webURL={this.props.webURL}></CmDocumentUpload>
              )}
              {this.state.PmAssign && (
                <PmAssignWebPart webURL={this.props.webURL}></PmAssignWebPart>
              )}
              {this.state.PmAction && (
                <PmActionWebPart webURL={this.props.webURL}></PmActionWebPart>
              )}
              {this.state.KickOffDocUpload && (
                <KickOffDocUpload webURL={this.props.webURL}></KickOffDocUpload>
              )}

              {this.state.TechnicalFeedback && (
                <Technical webURL={this.props.webURL}></Technical>
              )}
              {this.state.DeliveryFeedback && (
                <Delivery webURL={this.props.webURL}></Delivery>
              )}
              {this.state.LegalFeedback && (
                <Legal webURL={this.props.webURL}></Legal>
              )}

            </div>
          </div>
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

        </div>
      </div>
    );
  }

  // Get Current User Display Name
  private async spLoggedInUserDetails() {
    let web = Web(this.props.webURL);
    return await web.currentUser.get();
  }
  private async setData() {
    this.itemID = new URLSearchParams(window.location.search).get("itemid");
    this.Dpt = new URLSearchParams(window.location.search).get('dpt');
    this.Action = new URLSearchParams(window.location.search).get('action');
    let web = Web(this.props.webURL);
    await web.lists
      .getByTitle("Projects")
      .items.getById(parseInt(this.itemID))
      .select(
        "*",
        "Author/Id",
        "Author/EMail",
        "DeliveryManager/Title",
        "AccountManager/Title"
      )
      .expand("DeliveryManager/Id", "AccountManager/Id", "Author/EMail")
      .get()
      .then(async (item) => {
        this.CreatedEmail = item.Author.EMail;
        this.Status = item.Status;
        if (this.currentuseremail.toLowerCase() == this.CreatedEmail.toLowerCase()) {
          if (item.Status == "Technical Action Awaited") {
            if(this.Action=="Legal"){
              this.setState({ btnLegalTask: true, LegalTask:true,btnSales: true, TechnicalTask: true });
            }else{
            this.setState({ btnTechnicalTask: true, btnSales: true, TechnicalTask: true });
            }
          } else if (item.Status == "Delivery Action Awaited") {
            if (this.Dpt != null) {
              this.setState({ btnDeliveryTask: false, DeliveryTask: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: true });
            }else if(this.Action=="Legal"){
              this.setState({ btnLegalTask: true,LegalTask:true, btnSales: true,btnTechnicalTask: true,btnDeliveryTask:true  });
            }
            else if (this.Action == "Delivery") {
              this.setState({ btnDeliveryTask: true, DeliveryTask: true, btnSales: true,btnTechnicalTask: true,  });
             
            }
            else {
              this.setState({ btnDeliveryTask: false, DeliveryTask: false, btnSales: false,btnTechnicalTask: false,  });
            }
          } else if (item.Status == "Management Action Awaited" || item.Status == "Escalated to management") {
            if (this.Dpt != null) {
              this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true, TechnicalTask: true,btnDeliveryTask: true });
            }
            else if(this.Action=="Legal"){
              this.setState({ btnLegalTask: true,LegalTask:true,btnManagement: false, btnSales: true,btnTechnicalTask: true,btnDeliveryTask:true  });
            }
            else if (this.Action == "Delivery") {
              this.setState({ btnManagement: false, btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true, DeliveryTask: true });
            }
            else {
              this.setState({ btnManagement: true,ManagementTask:true, btnSales: true,btnDeliveryTask:true,btnTechnicalTask: true });
            }
            
          } else if (item.Status == "Approved PI Workflow Awaited") {
            this.setState({ btnContractOwner: true,ContractOwner:true ,btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,});
          } else if (item.Status == "CM Documents Upload Awaited") {
            this.setState({btnCMDocUpload: true,CMDocUpload:true,btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,});
          
          } else if (item.Status == "PM Assignment Awaited") {
          
            this.setState({btnPmAssign: true,PmAssign:true,btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,});
          
          } else if (item.Status == "PM Action Awaited") {
            this.setState({btnPmAction: true,PmAction:true,btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,});
          
          } else if (item.Status == "Kickoff Docs Awaited") {            
            this.setState({btnKickOffDocUpload: true,KickOffDocUpload:true,btnSales: true, btnTechnicalTask: true,btnDeliveryTask: true,});
          
          }
          else {
            
            this.setState({btnSales: true,Sales:true, btnTechnicalTask: true,btnDeliveryTask: true,});
          
          }

          // this.setState({
          //   ContractOwner: true,
          // });

        } else {
          this.setState({ btnProjectDetails: true, Sales: true });

        }
      })
      .then((x) => {
        // if (!this.state.IsSales) {
        //   web.lists
        //     .getByTitle("Management Team")
        //     .items.select("Manager/EMail")
        //     .expand("Manager")
        //     .getAll()
        //     .then((items) => {
        //       //console.log(items);
        //       items.forEach((item) => {
        //         if (
        //           item.Manager.EMail.toLowerCase() ==
        //           this.currentuseremail.toLowerCase()
        //         ) {
        //           this.setState({ IsManagement: true });
        //         }
        //       });
        //     });
        // }
      });
  }
  Close0(): void {
    throw new Error('Method not implemented.');
  }
}
