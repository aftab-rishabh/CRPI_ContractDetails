import * as React from 'react';
import styles from './Accordion.module.scss';
import { IAccordionProps, IAccordionState } from './index';
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
const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.accordionChevron };
const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.accordionChevron };

export class Accordion extends React.Component<IAccordionProps, IAccordionState> {
    private _drawerDiv: HTMLDivElement = undefined;
    constructor(props: IAccordionProps) {
        super(props);

        this.state = {
            expanded: props.defaultCollapsed == null ? true : !props.defaultCollapsed,
            Items: [],
            HTML: [],
            OppID: 0,
            files: []
        };
    }
    public async componentDidMount() {
        await this.fetchData();
    }
    public render(): React.ReactElement<IAccordionProps> {
        return (
            <div className={css(styles.accordion, this.props.className)}>
                <div >
                    <DefaultButton
                        toggle
                        checked={this.state.expanded}
                        text={this.props.title}
                        style={{ width: '97%', marginBottom: '2px' }}
                        iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                        onClick={(e) => {
                            this.setState({
                                expanded: !this.state.expanded
                            });
                        }}
                        aria-expanded={this.state.expanded}
                        aria-controls={this._drawerDiv && this._drawerDiv.id}
                    />
                    {this.state.expanded &&
                        <div className={styles.drawer} ref={(el) => { this._drawerDiv = el; }}>
                            {this.state.HTML}

                        </div>
                    }
                </div>
            </div>
        );
    }

    public async fetchData() {
        const itemID = new URLSearchParams(window.location.search).get('itemid');
        let web = Web(this.props.webURL);
        const items: any[] = await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).select("*", 'DeliveryManager/Title', 'AccountManager/Title').expand('DeliveryManager/Id', 'AccountManager/Id').get();
        console.log(items);
        this.setState({ Items: items });
        await this.getFiles(items);

    }

    public async getHTML(items, files) {

        var sdate = items.StartDate;
        // var sdate = moment(startd).format('DD/MM/YYYY');
        var edate = items.EndDate;
        // var edate = moment(endd).format('DD/MM/YYYY');
        var OppID = items.OpportunityID;
        if (items.Status == "Rejected to Sales") {
            OppID = items.OpportunityID + "-Rejected";
        }
        // primary technology array
        var PT = items.PrimaryTechnology;
        var pt = [];
        for (var x = 0; x < PT.length; x++) {
            pt.push(PT[x].split("-")[1]);

        }
        this.setState({ OppID: OppID });

        var OppID = items.OpportunityID;
        if (items.Status == "Rejected to Sales") {
            OppID = items.OpportunityID + "-Rejected";
        }

        // secondary technology array
        var ST = items.SecondaryTechnology;
        var st = [];
        if (items.SecondaryTechnology != null) {
            for (var y = 0; y < ST.length; y++) {
                st.push(ST[y].split("-")[1]);
            }
        }
        else {
            st.push("");

        }

        var tabledata = <div className={styles.contractDetails}>

            <div className={styles.viewdetail_block}>

                <div className={styles.sectionblock}>
                    <Label className={styles.headers}><u>Project Details</u></Label>
                </div>

                <Label id="ifRejected">{OppID}</Label>

                <div className={styles.sectionblock}>
                    <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Opportunity ID:</div>
                                    <div className={styles.viewlable_info} id="oppID" >{items.OpportunityID}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Opportunity Number:</div>
                                    <div className={styles.viewlable_info}>{items.OpportunityNumber}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Project Name:</div>
                                    <div className={styles.viewlable_info}>{items.ProjectName}</div>
                                </li>
                            </ul>
                        </div>
                    </div>
                    <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Short Name:</div>
                                    <div className={styles.viewlable_info}>{items.ShortName}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Client Organization:</div>
                                    <div className={styles.viewlable_info}>{items.ClientOrganization}</div>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>

                <div className={styles.sectionblock}>

                    <div className={styles.view_listing}>
                        <ul>
                            <li>
                                <div className={styles.viewlable}>Description:</div>
                                <div className={styles.viewlable_info}>{items.Description}</div>
                            </li>
                        </ul>
                    </div>

                </div>



                <div className={styles.sectionblock}>
                    <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Industry:</div>
                                    <div className={styles.viewlable_info} id="oppID" >{items.Industry}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Practice:</div>
                                    <div className={styles.viewlable_info}>{items.Practice}</div>
                                </li>

                            </ul>
                        </div>
                    </div>
                    <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Engagement Type:</div>
                                    <div className={styles.viewlable_info}>{items.EngagementType}</div>
                                </li>

                            </ul>
                        </div>
                    </div>
                </div>


                <div className={styles.sectionblock}>
                    <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Start Date:</div>
                                    <div className={styles.viewlable_info}>{sdate}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>End Date:</div>
                                    <div className={styles.viewlable_info}>{edate}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Project Type:</div>
                                    <div className={styles.viewlable_info}>{items.ProjectType}</div>
                                </li>
                                <li id="NoOfResources" >
                                    <div className={styles.viewlable}>No. of Resources:</div>
                                    <div className={styles.viewlable_info}>{items.NoOfResources}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Estimated Hours:</div>
                                    <div className={styles.viewlable_info}>{items.EstimatedHours}</div>
                                </li>

                            </ul>
                        </div>
                    </div>

                    <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Account Manager:</div>
                                    <div className={styles.viewlable_info}>{items.AccountManager.Title}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Business Manager:</div>
                                    <div className={styles.viewlable_info}>{items.BusinessManager0}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Delivery Manager:</div>
                                    <div className={styles.viewlable_info}>{items.DeliveryManager.Title}</div>
                                </li>
                                <li id="SelectedTeam" >
                                    <div className={styles.viewlable}>Selected Team:</div>
                                    <div className={styles.viewlable_info}>{items.SpecificTeam}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>With RSPL Template:</div>
                                    <div className={styles.viewlable_info}>{items.WithTemplate}</div>
                                </li>


                            </ul>
                        </div>
                    </div>
                </div>


                <div className={styles.sectionblock}>
                    <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Leaves Approved for the Year:</div>
                                    <div className={styles.viewlable_info}>{items.LeavesApproved}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Overtime/Less Time:</div>
                                    <div className={styles.viewlable_info}>{items.OvertimeLessTime}</div>
                                </li>

                            </ul>
                        </div>
                    </div>
                    <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                            <ul>
                                <li>
                                    <div className={styles.viewlable}>Notice Period for Termination:</div>
                                    <div className={styles.viewlable_info}>{items.NoticePeriod}</div>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>


                <div className={styles.sectionblock}>
                    <div className={styles.section_left}>
                        <div className={styles.view_listing}>
                            <ul>

                                <li>
                                    <div className={styles.viewlable}>Primary Technology:</div>
                                    <div className={styles.viewlable_info}>{pt}</div>
                                </li>
                                <li>
                                    <div className={styles.viewlable}>Secondary Technology:</div>
                                    <div className={styles.viewlable_info}>{st}</div>
                                </li>

                            </ul>
                        </div>
                    </div>
                    <div className={styles.section_right}>
                        <div className={styles.view_listing}>
                            <ul>

                                <li>
                                    <div className={styles.viewlable}>Are we using a rate card of current year i.e. Year: <span id="rateCardYear"></span></div>
                                    <div className={styles.viewlable_info}>{items.UseRateCardCurrYear}</div>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>


                <div className={styles.sectionblock}>

                    <div className={styles.view_listing}>
                        <ul>
                            <li>
                                <div className={styles.viewlable}>All Files:</div>
                            </li>
                        </ul>
                        <div id="fileList" className={styles.file_list}>
                            <ul>

                                {files.map(function (item, key) {

                                    <li>{item.Title}<a href={item.ServerRelativeUrl} target="_blank">{item.Name}</a></li>


                                })}
                            </ul>
                        </div>

                    </div>
                </div>

            </div>
        </div>
        return await tabledata;
    }


    private getFiles(items) {
        // this.context.pageContext.site.serverRelativeUrl 
        let web = Web(this.props.webURL);
        var hostUrl = '/sites/demo/ProjectDocuments/' + items.OpportunityID;
        // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
        web.getFolderByServerRelativeUrl(hostUrl).files.get().then(async files => {

            this.setState({
                files: files
            })

            let html = await this.getHTML(this.state.Items, files);
            this.setState({ HTML: html })
        });
    }


}


