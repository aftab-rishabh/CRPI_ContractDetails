import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './PmActionTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, FontWeights, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import * as moment from 'moment';
import { Dropdown, TextField } from '@fluentui/react';
// require('bootstrap');
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
// require('../scripts/sp.peoplepicker.js');
// require('./scripts/app.js');
/**
 * Icon styles. Feel free to change them
 */


export interface IPmActionProps {
    defaultCollapsed?: boolean;
    className?: string;
    webURL: string;
}

export interface IPmActionState {
    expanded: boolean;
    Items: any;
    HTML: any;
}



// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class PmActionWebPart extends React.Component<IPmActionProps, IPmActionState> {
    [x: string]: any;
    private _drawerDiv: HTMLDivElement = undefined;
    ContractStatus = "";
    OpportunityID = "";
    constructor(props: IPmActionProps) {
        super(props);

        this.state = {
            expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
            Items: [],
            HTML: [],
        };
    }
    public async componentDidMount() {

        this.setForm();
        this.viewItem();

    }
    public render(): React.ReactElement<IPmActionProps> {
        return (
            <div className={styles.PmAction}>
                <div>
                    <div className={styles.pagetitle_wrap} style={{ paddingTop: '0px' }}>
                        <div id="status" className="status ${styles.pagetitle}"></div>
                        <div id="oppID" style={{ display: 'none' }} className={styles.pagesubtitle}></div>
                        <div id="oppNo" style={{ display: 'none' }} className={styles.pagesubtitle}></div>
                    </div>

                    <div id="pm-action-awaited-buttons" style={{ display: 'none' }}>
                        <div className={styles.sectionblockPMAction}>
                            <Label className={styles.headersPMAction} style={{ marginTop: '0px' }}><u>PI Note Form</u></Label>
                        </div>


                        <div id="ProjectDetails" style={{ display: 'block' }}>
                            <h3 className={styles.subheading3}>1. Project Details (For Projects Admin Group)</h3>

                            <div className="row">
                                <div className="col-md-6">
                                    <div className={styles.sectionblock}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Description:</div>
                                                        <div className={styles.viewlable_info} id="Description"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>

                                </div>
                                <div className="col-md-6">
                                    <div className={styles.sectionblock}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Project Manager:</div>
                                                        <div className={styles.viewlable_info} id="projectManager"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div className="col-md-6">
                                    <div className={styles.sectionblock} >
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Department:</div>
                                                        <div className={styles.viewlable_info} id="Department"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div className="col-md-6" >
                                    <div className={styles.sectionblock}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Engagement Type:</div>
                                                        <div className={styles.viewlable_info} id="EngagementType"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>


                                <div className="col-md-6" style={{ display: 'none' }}>
                                    <div className={styles.sectionblock}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Leaves Approved for the Year:</div>
                                                        <div className={styles.viewlable_info} id="LeavesApproved"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div className="col-md-6" style={{ display: 'none' }}>
                                    <div className={styles.sectionblock} >
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Overtime/Less Time:</div>
                                                        <div className={styles.viewlable_info} id="OvertimeLessTime"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-md-6" style={{ display: 'none' }}>
                                    <div className={styles.sectionblock}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.view_listing}>
                                                <ul>
                                                    <li>
                                                        <div className={styles.viewlable}>Notice Period for Termination:</div>
                                                        <div className={styles.viewlable_info} id="NoticePeriod"> </div>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>


                            <div className="row ${styles.sectionEnd}">
                                <div className="col-md-6 ${styles.form_group}">
                                    <Label className={styles.lablecontrol}>Sprint Duration:<span className={styles.estric}>*</span>
                                    </Label>
                                    <input className={styles.form_control} type="text" placeholder="SprintDuration" id="SprintDuration"></input>
                                    <Label className={styles.errorlable} id="SprintDurationErr" style={{ display: 'none' }}
                                    >This field is
                                        required.</Label>
                                </div>
                                <div className="col-md-6 ${styles.form_group}">
                                    <Label className={styles.lablecontrol}>Project Initiation Duration: <span
                                        className={styles.estric}>*</span></Label>
                                    <input className={styles.form_control} type="text" placeholder="ProjectInitiationDuration"
                                        id="ProjectInitiationDuration"></input>
                                    <Label className={styles.errorlable} id="ProjectInitiationDurationErr" style={{ display: 'none' }}
                                    >This field is
                                        required.</Label>
                                </div>
                            </div>

                            <div id="ClientAndBilling" style={{ display: 'block' }}>


                                <h3 className={styles.subheading3}>2 Client & Billing Details (For Projects Admin Group)</h3>
                                <h4 className={styles.subheading4} >2.1 Payment Milestone</h4>

                                <div className="row ${styles.sectionEnd}">
                                    <div className="col-lg-12" id="MilestoneTable" style={{ display: 'none' }}>
                                        <div className={styles["table-bordered"]} style={{ border: '0' }}>
                                        <table className="table table-bordered" id="dataTable" width="60%">
                                            <tr>
                                                <th className={styles.ThH}>Payment Milestone Name (Phase)</th>
                                                <th className={styles.ThH}>Milestone Date (Optional)</th>
                                                <th className={styles.ThH}>% Of Total Price</th>
                                            </tr>
                                        </table>
                                        </div>
                                    </div>



                                    <div className="row">
                                        <div className="col-md-6" id="FrequencySection" style={{ display: 'none' }}>
                                        <div className={styles.viewdetail_block}>
                                            <div className={styles.sectionblock}>
                                             
                                                    <div className={styles.view_listing}>
                                                        <ul>
                                                            <li>
                                                                <div className={styles.viewlable}>Frequency:</div>
                                                                <div className={styles.viewlable_info} id="Frequency"> </div>
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </div>

                                        </div>

                                    </div>
                                </div></div>
                            <div id="DeliveryCommitments" style={{ display: 'block' }}>

                                <h3 className={styles.subheading3}>3. Delivery Commitments</h3>

                                <h4 className={styles.subheading4}>3.1 Project Timelines By PMO By Referring Contract </h4>

                                <div className={styles.sectionEnd}>
                                    <table className="table table-bordered">
                                        <tr>
                                            <th className={styles.ThH}>Deliverable Name</th>
                                            <th className={styles.ThH}>Start Date</th>
                                            <th className={styles.ThH}>End Date</th>
                                        </tr>
                                        <tr>
                                            <td><input type="text" className={styles.form_control} id="DeliverableName1"></input>
                                            </td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate1"></input></td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate11"></input></td>
                                        </tr>
                                        <tr>
                                            <td><input type="text" className={styles.form_control} id="DeliverableName2"></input>
                                            </td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate2"></input></td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate22"></input></td>
                                        </tr>
                                        <tr>
                                            <td><input type="text" className={styles.form_control} id="DeliverableName3"></input>
                                            </td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate3"></input></td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate33"></input></td>
                                        </tr>
                                        <tr>
                                            <td><input type="text" className={styles.form_control} id="DeliverableName4">
                                            </input> </td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate4"></input></td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate44"></input></td>
                                        </tr>
                                        <tr>
                                            <td><input type="text" className={styles.form_control} id="DeliverableName5"></input>
                                            </td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate5"></input></td>
                                            <td><input type="date" placeholder="DD/MM/YYYY" className={styles.form_control}
                                                id="WeekDate55"></input></td>
                                        </tr>

                                    </table>
                                    <div><Label className={styles.errorlable} id="StartDateErr1" style={{ display: 'none' }} >Deliverable 1: Start Date Cannot Be Older Than Today's Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="StartDateErr2" style={{ display: 'none' }} >Deliverable 2: Start Date Cannot Be Older Than Today's Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="StartDateErr3" style={{ display: 'none' }} >Deliverable 3: Start Date Cannot Be Older Than Today's Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="StartDateErr4" style={{ display: 'none' }} >Deliverable 4: Start Date Cannot Be Older Than Today's Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="StartDateErr5" style={{ display: 'none' }} >Deliverable 5: Start Date Cannot Be Older Than Today's Date.</Label></div>

                                    <div><Label className={styles.errorlable} id="EndDateErr1" style={{ display: 'none' }} >Deliverable 1: End Date Cannot Be Older Than Start Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="EndDateErr2" style={{ display: 'none' }} >Deliverable 2: End Date Cannot Be Older Than Start Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="EndDateErr3" style={{ display: 'none' }} >Deliverable 3: End Date Cannot Be Older Than Start Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="EndDateErr4" style={{ display: 'none' }} >Deliverable 4: End Date Cannot Be Older Than Start Date.</Label></div>
                                    <div><Label className={styles.errorlable} id="EndDateErr5" style={{ display: 'none' }} >Deliverable 5: End Date Cannot Be Older Than Start Date.</Label></div>

                                    <Label className={styles.errorlable} id="DeliverableErr" style={{ display: 'none' }} >Atleast One Deliverable should be present.</Label>

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


                                <div id="SubmittedModal" className={styles.modal}>
                                    <div className={styles.modalcontent}>
                                        <span className={styles.close} onClick={() => this.Close1()}>
                                            &times;
                                        </span>
                                        <Label className={styles.header2}>PI Note details submitted!</Label>
                                    </div>
                                </div>


                                <h4 className={styles.subheading4}>Client Provided Documents and Artefacts:</h4>

                                <div className="row">
                                    <div className="col-md-4">
                                        <div className="form-group">
                                            <Label className={styles.lablecontrol}>Document Title:<span className={styles.estric}>*</span></Label>
                                            <input className={styles.form_control} type="text" placeholder="File Title" id="DocType"></input>
                                            <span id="DocTypeErr" style={{ display: 'none' }} className={styles.errorlable} >Enter Document
                                                type.</span>

                                        </div>
                                    </div>

                                    <div className="col-md-5">
                                        <div className="form-group">
                                            <Label className={styles.lablecontrol}>Attachment:<span className={styles.estric}>*</span></Label>
                                            <input type="file" className={styles.form_control} id="file" name="file" ></input>
                                            <span id="fileErr0" style={{ display: 'none' }} className={styles.errorlable} >Please Upload Mandatory
                                                Files.</span>
                                            <span id="fileErr1" style={{ display: 'none' }} className={styles.errorlable} >Please select file to
                                                upload.</span>
                                        </div>
                                    </div>

                                    <div className="col-md-3">
                                        <Label className={styles.lablecontrol}>&nbsp;</Label>
                                        <PrimaryButton type="button" className={styles.btn}
                                            id="UploadDocs"  onClick={() => this.UploadFiles()} >Upload</PrimaryButton>
                                    </div>

                                </div>

                                <div className="row">
                                    <div className="col-lg-12">
                                        <div className={styles.file_list}>
                                            <div>
                                                <ul className="FileList2" id="FileList2"></ul>
                                            </div>
                                        </div>
                                    </div>
                                </div><br />




                                <div className="row">
                                    <div className="col-lg-12">
                                    <div className={styles.form_footer}>
                                        <PrimaryButton text="Submit" className={styles.btn} style={{ borderRadius: "10px",marginRight:"10px" }} id="Submit" onClick={() => this.InitiateProject()}></PrimaryButton>
                                        <PrimaryButton text="Cancel" className={styles.btn} style={{ borderRadius: "10px"}} id="Cancel" onClick={() => this.CancelInitiateProject()}></PrimaryButton>
                                        </div>
                                    </div>
                                </div>

                            </div>

                            <div>

                                <div id="WCR1-1" style={{ display: 'none' }}>
                                    <div className="row">
                                        <div className="col-lg-12 ${styles.form_group}">
                                            <Label className={styles.lablecontrol}>keyPointers:</Label>
                                            <input id="keyPointers1" className={styles.form_control} ></input>

                                        </div>
                                    </div>
                                </div>

                            </div>


                        </div>
                    </div>
                </div>


            </div >
        );
    }




    private getPaymentMilestone() {
        const itemID = new URLSearchParams(window.location.search).get('itemid');
        let web = Web(this.props.webURL);
        web.lists.getByTitle("PaymentMileStoneTable").items.filter('Title eq ' + itemID).getAll().then((milestone) => {
            if (milestone.length > 0) {
                var milestonetbl
                milestonetbl += '<tr id="TR1">'
                    + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone1">' + milestone[0].PaymentMilestone1 + '</lable></td>'
                    + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate1">' + milestone[0].MilestoneDate1 + '</lable></td>'
                    + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage1">' + milestone[0].Percentage1 + '</lable></td>'
                    + '</tr>'
                if (milestone[0].PaymentMilestone2 != null && milestone[0].PaymentMilestone2 != "") {
                    milestonetbl += '<tr id="TR2">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone2">' + milestone[0].PaymentMilestone2 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate2">' + milestone[0].MilestoneDate2 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage2">' + milestone[0].Percentage2 + '</lable></td>'
                        + '</tr>'
                }
                if (milestone[0].PaymentMilestone3 != null && milestone[0].PaymentMilestone3 != "") {
                    milestonetbl += '<tr id="TR3">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone3">' + milestone[0].PaymentMilestone3 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate3">' + milestone[0].MilestoneDate3 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage3">' + milestone[0].Percentage3 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone4 != null && milestone[0].PaymentMilestone4 != "") {
                    milestonetbl += '<tr id="TR4">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone4">' + milestone[0].PaymentMilestone4 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate4">' + milestone[0].MilestoneDate4 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage4">' + milestone[0].Percentage4 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone5 != null && milestone[0].PaymentMilestone5 != "") {
                    milestonetbl += '<tr id="TR5">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone5">' + milestone[0].PaymentMilestone5 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate5">' + milestone[0].MilestoneDate5 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage5">' + milestone[0].Percentage5 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone6 != null && milestone[0].PaymentMilestone6 != "") {
                    milestonetbl += '<tr id="TR6">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone6">' + milestone[0].PaymentMilestone6 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate6">' + milestone[0].MilestoneDate6 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage6">' + milestone[0].Percentage6 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone7 != null && milestone[0].PaymentMilestone7 != "") {
                    milestonetbl += '<tr id="TR7">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone7">' + milestone[0].PaymentMilestone7 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate7">' + milestone[0].MilestoneDate7 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage7">' + milestone[0].Percentage7 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone8 != null && milestone[0].PaymentMilestone8 != "") {
                    milestonetbl += '<tr id="TR8">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone8">' + milestone[0].PaymentMilestone8 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate8">' + milestone[0].MilestoneDate8 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage8">' + milestone[0].Percentage8 + '</lable></td>'
                        + '</tr>'

                }

                if (milestone[0].PaymentMilestone9 != null && milestone[0].PaymentMilestone9 != "") {
                    milestonetbl += '<tr id="TR9">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone9">' + milestone[0].PaymentMilestone9 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate9">' + milestone[0].MilestoneDate9 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage9">' + milestone[0].Percentage9 + '</lable></td>'
                        + '</tr>'

                }
                if (milestone[0].PaymentMilestone10 != null && milestone[0].PaymentMilestone10 != "") {
                    milestonetbl += '<tr id="TR10">'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="PaymentMilestone10">' + milestone[0].PaymentMilestone10 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="MilestoneDate10">' + milestone[0].MilestoneDate10 + '</lable></td>'
                        + '<td><lable style="width: 100%; text-align:center; display:inline-block"  id="Percentage10">' + milestone[0].Percentage10 + '</lable></td>'
                        + '</tr>';

                }

            }
            $('#dataTable').append(milestonetbl);

        }).then(() => {

            for (var x = 0; x < 10; x++) {
                var PM = "PaymentMilestone" + (x + 1);
                var tr = "TR" + (x + 1);
                var NullDateId = "MilestoneDate" + (x + 1);
                var nulltxt = document.getElementById(PM).innerText;
                var nullDate = document.getElementById(NullDateId).innerText;

                if (nulltxt == "null") {
                    $("#" + tr).empty();
                }

                if (nullDate == "null") {
                    document.getElementById(NullDateId).innerText = "NA";
                }

            }

        });
    }


    private setForm() {

        const itemID = new URLSearchParams(window.location.search).get('itemid');
        let web = Web(this.props.webURL);
        web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('Id', 'ProjectType', 'Status', 'OpportunityID', 'OpportunityNumber').get().then((data) => {
            $('#oppID').text(data.OpportunityID);
            $('#oppNo').text(data.OpportunityNumber);

            if (data.Status == "PM Action Awaited") {
                document.getElementById("pm-action-awaited-buttons").style.display = "block";
            }
            if (data.ProjectType == "Fixed Price") {
                document.getElementById("ClientAndBilling").style.display = "block";
                document.getElementById("MilestoneTable").style.display = "block";
                document.getElementById("FrequencySection").style.display = "none";
            }
            else if (data.ProjectType == "Dedicated Team") {
                document.getElementById("ClientAndBilling").style.display = "none";
                document.getElementById("MilestoneTable").style.display = "none";
                document.getElementById("FrequencySection").style.display = "none";

            } else if (data.ProjectType == "Time & Material") {
                document.getElementById("ClientAndBilling").style.display = "block";
                document.getElementById("MilestoneTable").style.display = "none";
                document.getElementById("FrequencySection").style.display = "block";
            }
        }).then(fileslisting => {
            this.getFiles();
        });
    }


    private UploadFiles() {
        $("#loader").show();


        var isvalid = true;

        var OID = $("#oppID")[0].innerHTML;

        if ($("#DocType").val() == "") {
            $("#DocTypeErr").show();
            isvalid = false;
        }
        else
            $("#DocTypeErr").hide();

        if (isvalid) {
            var files = (document.getElementById('file') as HTMLInputElement).files;
            if (files.length !== 0) {
                $("#fileErr1").hide();
                let web = Web(this.props.webURL);
                // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
                web.folders.getByName('ProjectDocuments').folders.add(OID).then((data) => {
                }).then(b => {
                    var fileArr = [];
                    for (var g = 0; g < files.length; g++) {
                        var file = files[g];
                        fileArr.push(file);
                        let web = Web(this.props.webURL);
                        var hostUrl = '/sites/demo/ProjectDocuments/' + OID;
                        // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
                        web.getFolderByServerRelativeUrl(hostUrl).files.add(file.name, file, true).then(f => {
                            f.file.getItem().then(item => {
                                item.update({
                                    Title: $("#DocType").val()
                                }).then(r => {
                                    $("#DocType").val('');
                                    $("#fileErr0").hide();
                                    $("#fileErr").hide();
                                    this.getFiles();
                                    (document.getElementById('file') as HTMLInputElement).value = "";
                                });
                            });
                        });
                    }
                }).then(y => {
                    this.getFiles();
                });
            }
            else {
                $("#fileErr1").show();
            }

        }

    }


    private getFiles() {
        let attachmentfiles: string = "";
        var OID = $("#oppID").text();
        var NumberOfFiles = [];
        let web = Web(this.props.webURL);
        var hostUrl = '/sites/demo/ProjectDocuments/' + OID;
        // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
        web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
            if (files.length == 0) {
                $("#FileList2").empty();
            }
            else {
                for (var i = 0; i < files.length; i++) {
                    var title = files[i].Title;
                    if (files[i].Title != "Contract" && files[i].Title != "Estimation file" && files[i].Title != "Others" && files[i].Title != "Legal Documents File" && files[i].Title != "Project Related Document from CM" && files[i].Title != "Kick-off document" && files[i].Title !="Contract - Commercial") {
                        var Uid = "DeleteFile" + files[i].UniqueId;
                        var valu = files[i].Name;
                        NumberOfFiles.push(1);
                        attachmentfiles += `<li id="FileListItem` + i + `" value=` + valu + `><button class="CloseBtn"`
                            + `id="CloseBtn">&times;</button><lable style="display:none;">` + files[i].UniqueId + `</lable>`
                            + `<lable id="filesnames` + i + `">`
                            + `&nbsp&nbsp&nbsp&nbsp${title}&nbsp-&nbsp</lable>`
                            + `<a href="${files[i].ServerRelativeUrl}" target="_blank" id="FileName` + i + `">`
                            + `${files[i].Name}</a></li>`;
                        $("#FileList2").append(attachmentfiles);
                    }
                }
            }
        }).then(help => {
            this.DeleteFileEvent();
            if (NumberOfFiles.length == 0) {
                $("#FileList2").empty();
            }
        }).then(() => {
            setTimeout(() => { $('#loader').hide(); }, 1000);
        });
    }

    private DeleteFileEvent() {
        $("#loader").show();

        var OID = $("#oppID").text();
        let web = Web(this.props.webURL);
        var hostUrl = '/sites/demo/ProjectDocuments/' + OID;
        // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
        web.getFolderByServerRelativeUrl(hostUrl).files.get().then(files => {
            for (var i = 0; i < files.length; i++) {
                var btnClass = document.getElementsByClassName("CloseBtn")[i] as HTMLElement;
                btnClass.addEventListener('click', (f) => {
                    var xyz = f.currentTarget as HTMLElement;
                    var fName = xyz.nextElementSibling.nextElementSibling.nextElementSibling.textContent;
                    web.getFolderByServerRelativeUrl(hostUrl).files.getByName(fName).delete().then((g) => {
                    }).then(check => {
                        this.getFiles();
                    }).then(() => {
                        setTimeout(() => { $('#loader').hide(); }, 1000);
                    });
                });
            }
        });
    }



    private Close1() {
        var modal = document.getElementById("SubmittedModal");
        var hostUrl = this.props.webURL;
        modal.style.display = "none";
        $('#LegalTaskForm').fadeOut(2500);
        setTimeout(() => { window.parent.location.href = hostUrl; }, 2500);
    }



    private viewItem(): void {

        this.getPaymentMilestone();
        const itemID = new URLSearchParams(window.location.search).get('itemid');
        let web = Web(this.props.webURL);
        web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select('*', 'Description', 'ProjectType', 'ProjectManager/Title').expand('ProjectManager/Id').get().then((item) => {

            var PT = item.PrimaryTechnology;
            var pt = [];
            for (var x = 0; x < PT.length; x++) {
                pt.push(PT[x].split("-")[1]);

            }


            $('#Description').text(item.Description);
            $('#projectManager').text(item.ProjectManager.Title);
            $('#Department').text(pt.toString());
            $('#EngagementType').text(item.ProjectType);
            $('#LeavesApproved').text(item.LeavesApproved);
            $('#OvertimeLessTime').text(item.OvertimeLessTime);
            $('#NoticePeriod').text(item.NoticePeriod);
            $('#Frequency').text(item.Frequency);
        });
    }

    private InitiateProject() {
        var isvalid = true;

        if ($("#SprintDuration").val() === "") {
            $("#SprintDuration").focus();
            $("#SprintDurationErr").show();
            isvalid = false;
        }
        else
            $("#SprintDurationErr").hide();

        if ($("#ProjectInitiationDuration").val() === "") {
            $("#ProjectInitiationDuration").focus();
            $("#ProjectInitiationDurationErr").show();
            isvalid = false;
        }
        else
            $("#ProjectInitiationDurationErr").hide();

        // if ($("#LeavesApproved").val() === "") {
        //   $("#LeavesApproved").focus();
        //   $("#LeavesApprovedErr").show();
        //   isvalid = false;
        // }
        // else
        //   $("#LeavesApprovedErr").hide();

        // if ($("#OvertimeLessTime").val() === "") {
        //   $("#OvertimeLessTime").focus();
        //   $("#OvertimeLessTimeErr").show();
        //   isvalid = false;
        // }
        // else
        //   $("#OvertimeLessTimeErr").hide();

        // if ($("#NoticePeriod").val() === "") {
        //   $("#NoticePeriod").focus();
        //   $("#NoticePeriodErr").show();
        //   isvalid = false;
        // }
        // else
        //   $("#NoticePeriodErr").hide();


        if ($("#EngagementType")[0].innerHTML == "Fixed Price" && $("#DeliverableName1").val() === "") {
            $("#DeliverableName1").focus();
            $("#DeliverableErr").show();
            isvalid = false;
        }
        else if ($("#EngagementType")[0].innerHTML == "Fixed Price" && $("#WeekDate1").val() === "") {
            $("#WeekDate1").focus();
            $("#DeliverableErr").show();
            isvalid = false;
        }
        else if ($("#EngagementType")[0].innerHTML == "Fixed Price" && $("#WeekDate11").val() === "") {
            $("#WeekDate11").focus();
            $("#DeliverableErr").show();
            isvalid = false;
        }
        else if ($("#DeliverableName1").val() === "" || $("#WeekDate1").val() === "" || $("#WeekDate11").val() === "") {
            $("#DeliverableErr").hide();
            $("#DeliverableErr").hide();
            $("#DeliverableErr").hide();
        }

        var datetd = moment(new Date()).format('YYYY-MM-DD');

        var datest1 = moment($("#WeekDate1").val()).format('YYYY-MM-DD');
        var datest2 = moment($("#WeekDate2").val()).format('YYYY-MM-DD');
        var datest3 = moment($("#WeekDate3").val()).format('YYYY-MM-DD');
        var datest4 = moment($("#WeekDate4").val()).format('YYYY-MM-DD');
        var datest5 = moment($("#WeekDate5").val()).format('YYYY-MM-DD');

        var datest11 = moment($("#WeekDate11").val()).format('YYYY-MM-DD');
        var datest22 = moment($("#WeekDate22").val()).format('YYYY-MM-DD');
        var datest33 = moment($("#WeekDate33").val()).format('YYYY-MM-DD');
        var datest44 = moment($("#WeekDate44").val()).format('YYYY-MM-DD');
        var datest55 = moment($("#WeekDate55").val()).format('YYYY-MM-DD');

        if (datest1 < datetd) {
            $("#StartDateErr1").show();
            $("#WeekDate1").focus();
            isvalid = false;
        }
        else
            $("#StartDateErr1").hide();

        if (datest2 < datetd) {
            $("#StartDateErr2").show();
            $("#WeekDate2").focus();
            isvalid = false;
        }
        else
            $("#StartDateErr2").hide();

        if (datest3 < datetd) {
            $("#StartDateErr3").show();
            $("#WeekDate3").focus();
            isvalid = false;
        }
        else
            $("#StartDateErr3").hide();

        if (datest4 < datetd) {
            $("#StartDateErr4").show();
            $("#WeekDate4").focus();
            isvalid = false;
        }
        else
            $("#StartDateErr4").hide();

        if (datest5 < datetd) {
            $("#StartDateErr5").show();
            $("#WeekDate5").focus();
            isvalid = false;
        }
        else
            $("#StartDateErr5").hide();

        if (datest11 < datest1) {
            $("#EndDateErr1").show();
            $("#WeekDate1").focus();
            isvalid = false;
        }
        else
            $("#EndDateErr1").hide();

        if (datest22 < datest2) {
            $("#EndDateErr2").show();
            $("#WeekDate2").focus();
            isvalid = false;
        }
        else
            $("#EndDateErr2").hide();

        if (datest33 < datest3) {
            $("#EndDateErr3").show();
            $("#WeekDate3").focus();
            isvalid = false;
        }
        else
            $("#EndDateErr3").hide();

        if (datest44 < datest4) {
            $("#EndDateErr4").show();
            $("#WeekDate4").focus();
            isvalid = false;
        }
        else
            $("#EndDateErr4").hide();

        if (datest55 < datest5) {
            $("#EndDateErr5").show();
            $("#WeekDate5").focus();
            isvalid = false;
        }
        else
            $("#EndDateErr5").hide();


        if (isvalid) {
            this.InitiateProject2();
        }
    }

    private InitiateProject2(): void {
        const itemID = new URLSearchParams(window.location.search).get('itemid');
        var filterStr = "ItemId eq '" + itemID + "'";
        let web = Web(this.props.webURL);
        web.lists.getByTitle("PINote").items.filter(filterStr).get().then((data) => {
            if (data.length) {
                var item = data[0];
                var pinoteID = item.ID;
                const itemID1 = pinoteID;


                // =================== Mileston name column =================

                if ($('#PaymentMilestone1')[0] != undefined) {
                    var PM1 = $('#PaymentMilestone1')[0].innerHTML;
                }
                else {
                    PM1 = "";
                }

                if ($('#PaymentMilestone2')[0] != undefined) {
                    var PM2 = $('#PaymentMilestone2')[0].innerHTML;
                }
                else {
                    PM2 = "";
                }

                if ($('#PaymentMilestone3')[0] != undefined) {
                    var PM3 = $('#PaymentMilestone3')[0].innerHTML;
                }
                else {
                    PM3 = "";
                }

                if ($('#PaymentMilestone4')[0] != undefined) {
                    var PM4 = $('#PaymentMilestone4')[0].innerHTML;
                }
                else {
                    PM4 = "";
                }

                if ($('#PaymentMilestone5')[0] != undefined) {
                    var PM5 = $('#PaymentMilestone5')[0].innerHTML;
                }
                else {
                    PM5 = "";
                }

                if ($('#PaymentMilestone6')[0] != undefined) {
                    var PM6 = $('#PaymentMilestone6')[0].innerHTML;
                }
                else {
                    PM6 = "";
                }

                if ($('#PaymentMilestone7')[0] != undefined) {
                    var PM7 = $('#PaymentMilestone7')[0].innerHTML;
                }
                else {
                    PM7 = "";
                }

                if ($('#PaymentMilestone8')[0] != undefined) {
                    var PM8 = $('#PaymentMilestone8')[0].innerHTML;
                }
                else {
                    PM8 = "";
                }

                if ($('#PaymentMilestone9')[0] != undefined) {
                    var PM9 = $('#PaymentMilestone9')[0].innerHTML;
                }
                else {
                    PM9 = "";
                }

                if ($('#PaymentMilestone10')[0] != undefined) {
                    var PM10 = $('#PaymentMilestone10')[0].innerHTML;
                }
                else {
                    PM10 = "";
                }

                // ======================== Milestone date colum =========================

                if ($('#MilestoneDate1')[0] != undefined) {
                    var MD1 = $('#MilestoneDate1')[0].innerHTML;
                }
                else {
                    MD1 = "";
                }

                if ($('#MilestoneDate2')[0] != undefined) {
                    var MD2 = $('#MilestoneDate2')[0].innerHTML;
                }
                else {
                    MD2 = "";
                }

                if ($('#MilestoneDate3')[0] != undefined) {
                    var MD3 = $('#MilestoneDate3')[0].innerHTML;
                }
                else {
                    MD3 = "";
                }

                if ($('#MilestoneDate4')[0] != undefined) {
                    var MD4 = $('#MilestoneDate4')[0].innerHTML;
                }
                else {
                    MD4 = "";
                }

                if ($('#MilestoneDate5')[0] != undefined) {
                    var MD5 = $('#MilestoneDate5')[0].innerHTML;
                }
                else {
                    MD5 = "";
                }

                if ($('#MilestoneDate6')[0] != undefined) {
                    var MD6 = $('#MilestoneDate6')[0].innerHTML;
                }
                else {
                    MD6 = "";
                }

                if ($('#MilestoneDate7')[0] != undefined) {
                    var MD7 = $('#MilestoneDate7')[0].innerHTML;
                }
                else {
                    MD7 = "";
                }

                if ($('#MilestoneDate8')[0] != undefined) {
                    var MD8 = $('#MilestoneDate8')[0].innerHTML;
                }
                else {
                    MD8 = "";
                }

                if ($('#MilestoneDate9')[0] != undefined) {
                    var MD9 = $('#MilestoneDate9')[0].innerHTML;
                }
                else {
                    MD9 = "";
                }

                if ($('#MilestoneDate10')[0] != undefined) {
                    var MD10 = $('#MilestoneDate10')[0].innerHTML;
                }
                else {
                    MD10 = "";
                }

                // ============================= Percentage column ================

                if ($('#Percentage1')[0] != undefined) {
                    var PE1 = $('#Percentage1')[0].innerHTML;
                }
                else {
                    PE1 = "";
                }

                if ($('#Percentage2')[0] != undefined) {
                    var PE2 = $('#Percentage2')[0].innerHTML;
                }
                else {
                    PE2 = "";
                }

                if ($('#Percentage3')[0] != undefined) {
                    var PE3 = $('#Percentage3')[0].innerHTML;
                }
                else {
                    PE3 = "";
                }

                if ($('#Percentage4')[0] != undefined) {
                    var PE4 = $('#Percentage4')[0].innerHTML;
                }
                else {
                    PE4 = "";
                }

                if ($('#Percentage5')[0] != undefined) {
                    var PE5 = $('#Percentage5')[0].innerHTML;
                }
                else {
                    PE5 = "";
                }

                if ($('#Percentage6')[0] != undefined) {
                    var PE6 = $('#Percentage6')[0].innerHTML;
                }
                else {
                    PE6 = "";
                }

                if ($('#Percentage7')[0] != undefined) {
                    var PE7 = $('#Percentage7')[0].innerHTML;
                }
                else {
                    PE7 = "";
                }

                if ($('#Percentage8')[0] != undefined) {
                    var PE8 = $('#Percentage8')[0].innerHTML;
                }
                else {
                    PE8 = "";
                }

                if ($('#Percentage9')[0] != undefined) {
                    var PE9 = $('#Percentage9')[0].innerHTML;
                }
                else {
                    PE9 = "";
                }

                if ($('#Percentage10')[0] != undefined) {
                    var PE10 = $('#Percentage10')[0].innerHTML;
                }
                else {
                    PE10 = "";
                }


                let web = Web(this.props.webURL);
                web.lists.getByTitle("PINote").items.getById(parseInt(itemID1)).update({

                    Description: $('#Description').text(),
                    ProjectManager: $('#projectManager').text(),
                    Department: $('#Department').text(),
                    ProjectType: $('#EngagementType').text(),
                    SprintDuration: $('#SprintDuration').val(),

                    ProjectInitiationDuration: $('#ProjectInitiationDuration').val(),
                    LeavesApproved: $('#LeavesApproved').text(),
                    OvertimeLessTime: $('#OvertimeLessTime').text(),
                    NoticePeriod: $('#NoticePeriod').text(),
                    OpportunityNumber: $('#oppNo').text(),

                    PaymentMilestone1: PM1,
                    PaymentMilestone2: PM2,
                    PaymentMilestone3: PM3,
                    PaymentMilestone4: PM4,
                    PaymentMilestone5: PM5,
                    PaymentMilestone6: PM6,
                    PaymentMilestone7: PM7,
                    PaymentMilestone8: PM8,
                    PaymentMilestone9: PM9,
                    PaymentMilestone10: PM10,

                    MilestoneDate1: MD1,
                    MilestoneDate2: MD2,
                    MilestoneDate3: MD3,
                    MilestoneDate4: MD4,
                    MilestoneDate5: MD5,
                    MilestoneDate6: MD6,
                    MilestoneDate7: MD7,
                    MilestoneDate8: MD8,
                    MilestoneDate9: MD9,
                    MilestoneDate10: MD10,


                    Percentage1: PE1,
                    Percentage2: PE2,
                    Percentage3: PE3,
                    Percentage4: PE4,
                    Percentage5: PE5,
                    Percentage6: PE6,
                    Percentage7: PE7,
                    Percentage8: PE8,
                    Percentage9: PE9,
                    Percentage10: PE10,

                    DeliverableName1: $('#DeliverableName1').val(),
                    WeekDate1: $('#WeekDate1').val(),
                    WeekDate11: $('#WeekDate11').val(),
                    DeliverableName2: $('#DeliverableName2').val(),
                    WeekDate2: $('#WeekDate2').val(),
                    WeekDate22: $('#WeekDate22').val(),
                    DeliverableName3: $('#DeliverableName3').val(),
                    WeekDate3: $('#WeekDate3').val(),
                    WeekDate33: $('#WeekDate33').val(),
                    DeliverableName4: $('#DeliverableName4').val(),
                    WeekDate4: $('#WeekDate4').val(),
                    WeekDate44: $('#WeekDate44').val(),
                    DeliverableName5: $('#DeliverableName5').val(),
                    WeekDate5: $('#WeekDate5').val(),
                    WeekDate55: $('#WeekDate55').val(),
                }).then(i => {
                    let web = Web(this.props.webURL);
                    web.lists.getByTitle('CompleteTask').items.add({
                        ProjectID: parseInt(itemID),
                        TaskType: "PM"
                    }).then(j => {
                        const itemID2 = new URLSearchParams(window.location.search).get('itemid');
                        let web = Web(this.props.webURL);
                        web.lists.getByTitle("Projects")
                            .items.getById(parseInt(itemID2)).update({
                                Status: `Kickoff Docs Awaited`
                            }).then(newListItem => {
                                $("#loader").show();
                                setTimeout(() => { $('#loader').hide(); }, 3000);
                                setTimeout(() => { $('#SubmittedModal').show(); }, 3500);
                            });
                    });
                });
            }
        }).then(i => {
        });
    }

    private CancelInitiateProject() {
        var hostUrl = this.props.webURL;
        window.parent.location.href = hostUrl;
    }



}


