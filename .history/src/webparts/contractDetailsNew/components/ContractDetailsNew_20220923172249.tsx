import * as React from 'react';
import styles from './ContractDetailsNew.module.scss';
import styles1 from '../../accordion/Accordion.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { IContractDetailsNewProps } from './IContractDetailsNewProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import { Accordion } from '../../accordion';
import { Nav, INavLink, INavStyles, INavLinkGroup } from '@fluentui/react/lib/Nav';
// import SplitPane, { Pane } from 'react-split-pane';
import { Technical } from '../Feedback/Technical';
import { Delivery } from '../Feedback/Delivery';
import { Legal } from '../Feedback/Legal';
import { Sales } from '../Feedback/Sales';
import { DefaultButton, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { Management } from '../Feedback/Management';
import { AllTask } from '../Feedback/AllTask';
import { ProjectDetails } from '../Feedback/ProjectDetails';

// const navStyles: Partial<INavStyles> = {
//   root: {
//     width: 208,
//     height: 350,
//     boxSizing: 'border-box',
//     border: '1px solid #eee',
//     overflowY: 'auto',
//   },
// };

// const navStyles: Partial<INavStyles> = { root: { width: 300 } };
const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.accordionChevron };
const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.accordionChevron };
export interface IContractDetailsNewState {
  Sales: boolean;
  Technical: boolean;
  Delivery: boolean;
  Legal: boolean;
  Management:boolean;
  expanded: boolean;

}

const navLinkGroups: INavLinkGroup[] = [
  {
    links: [
      {
        name: 'Sales Feedback',
        url: '',
        key: 'key1',
        isExpanded: true,
        target: '_blank',
      },
      {
        name: 'Legal Feedback',
        url: '',
        key: 'key1',
        disabled: true,
      },
      {
        name: 'Technical Feedback',
        url: '',
        key: 'key3',
        target: '_blank',
      },

      {
        name: 'Delivery Feedback',
        url: '',
        key: 'key4',
        target: '_blank',
      },

    ],
  },
];


export default class ContractDetailsNew extends React.Component<IContractDetailsNewProps, IContractDetailsNewState> {
  thiscontext = this;
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: IContractDetailsNewProps) {
    super(props);

    this.state = {
      Sales: false,
      Technical: false,
      Delivery: false,
      Legal: false,
      Management:false,
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
    };
  }
  public render(): React.ReactElement<IContractDetailsNewProps> {

    return (
      <div className={styles.contractDetailsNew}>

        <div className={styles.viewdetail_block}>
          <div className={styles.sectionblock}>
            <div className={styles.section_left}>
            <div className={css(styles1.accordion, this.props.className)}>
                <div >
                    <DefaultButton
                        toggle
                        checked={this.state.expanded}
                        text={"All Tasks"}
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
                        <div className={styles1.drawer} ref={(el) => { this._drawerDiv = el; }}>
                            {/* {this.state.HTML} */}
                            {/* <AllTask webURL={this.props.webURL}></AllTask> */}
                            <ProjectDetails webURL={this.props.webURL}></ProjectDetails>
                        </div>
                    }
                </div>
            </div>
              {/* <Accordion title={'Project Details'} webURL={this.props.webURL}>

              </Accordion> */}
              <div className={css(styles1.accordion, this.props.className)}>
                <div >
                  <DefaultButton
                    toggle
                    checked={this.state.Sales}
                    text={"Sales Feedback"}
                    style={{ width: '97%',marginBottom:'2px' }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        Sales: true,
                        Technical: false,
                        Delivery: false,
                        Legal: false,
                        Management:false
                      });
                    }}
                  // aria-expanded={this.state.expanded}
                  // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div></div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div >
                  <DefaultButton
                    toggle
                    checked={this.state.Legal}
                    text={"Legal Feedback"}
                    style={{ width: '97%',marginBottom:'2px' }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        Sales: false,
                        Technical: false,
                        Delivery: false,
                        Legal: true,
                        Management:false
                      });
                    }}
                  // aria-expanded={this.state.expanded}
                  // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div></div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div >
                  <DefaultButton
                    toggle
                    checked={this.state.Technical}
                    text={"Technical Feedback"}
                    style={{ width: '97%',marginBottom:'2px' }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        Sales: false,
                        Technical: true,
                        Delivery: false,
                        Legal: false,
                        Management:false
                      });
                    }}
                  // aria-expanded={this.state.expanded}
                  // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div></div>
              <div className={css(styles1.accordion, this.props.className)}>
                <div >
                  <DefaultButton
                    toggle
                    checked={this.state.Delivery}
                    text={"Delivery Feedback"}
                    style={{ width: '97%',marginBottom:'2px' }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        Sales: false,
                        Technical: false,
                        Delivery: true,
                        Legal: false,
                        Management:false
                      });
                    }}
                  // aria-expanded={this.state.expanded}
                  // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div></div>
                <div className={css(styles1.accordion, this.props.className)}>
                <div >
                  <DefaultButton
                    toggle
                    checked={this.state.Management}
                    text={"Management Feedback"}
                    style={{ width: '97%',marginBottom:'2px' }}
                    // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
                    onClick={(e) => {
                      this.setState({
                        expanded: !this.state.expanded,
                        Sales: false,
                        Technical: false,
                        Delivery: false,
                        Legal: false,
                        Management:true
                      });
                    }}
                  // aria-expanded={this.state.expanded}
                  // aria-controls={this._drawerDiv && this._drawerDiv.id}
                  />
                </div></div>
            </div>
            <div className={styles.section_right}>
              {/* <AllTask webURL={this.props.webURL}></AllTask> */}
              
              {this.state.Sales &&
                <Sales webURL={this.props.webURL}></Sales>
              }
              {this.state.Technical &&
                <Technical webURL={this.props.webURL}></Technical>
              }
              {this.state.Delivery &&
                <Delivery webURL={this.props.webURL}></Delivery>
              }
              {this.state.Legal &&
                <Legal webURL={this.props.webURL}></Legal>
              }
              {this.state.Management &&
                <Management webURL={this.props.webURL}></Management>
              }



            </div>
          </div></div></div>

    );
  }

}


