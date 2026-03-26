import * as React from 'react';
import styles from './ContractDetailsNew.module.scss';
import { IContractDetailsNewProps } from './IContractDetailsNewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Accordion } from '../../accordion';
import { Nav, INavLink, INavStyles, INavLinkGroup } from '@fluentui/react/lib/Nav';
// import SplitPane, { Pane } from 'react-split-pane';
import { Technical } from '../Feedback/Technical';
import { Delivery } from '../Feedback/Delivery';
import { Legal } from '../Feedback/Legal';
import { Sales } from '../Feedback/Sales';
import { DefaultButton, IIconProps } from 'office-ui-fabric-react';
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
  Technical:boolean;
  Delivery:boolean;
  Legal:boolean;
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
  thiscontext=this;
  constructor(props: IContractDetailsNewProps) {
    super(props);

    this.state = {
      Sales: false,
      Technical:false,
      Delivery:false,
      Legal:false,
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
    };
  }
  public render(): React.ReactElement<IContractDetailsNewProps> {
  
    return (
      <div className={styles.contractDetailsNew}>

      <div className={styles.viewdetail_block}>
      <div className={styles.sectionblock}>
                                <div className={styles.section_left}>
    <Accordion title={'Project Details'}  webURL={this.props.webURL}>

      </Accordion>
      <DefaultButton
          toggle
          checked={this.state.expanded}
          text={"Sales Feedback"}
          // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
          onClick={(e) => {
            this.setState({
              expanded: !this.state.expanded,   
              Sales: true,
              Technical:false,
              Delivery:false,
              Legal:false
            });
          }}
          aria-expanded={this.state.expanded}
          // aria-controls={this._drawerDiv && this._drawerDiv.id}
        />
         <DefaultButton
          toggle
          checked={this.state.expanded}
          text={"Technical Feedback"}
          // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
          onClick={(e) => {
            this.setState({
              expanded: !this.state.expanded,   
              Sales: false,
              Technical:true,
              Delivery:false,
              Legal:false
            });
          }}
          aria-expanded={this.state.expanded}
          // aria-controls={this._drawerDiv && this._drawerDiv.id}
        />
     <DefaultButton
          toggle
          checked={this.state.expanded}
          text={"Delivery Feedback"}
          // iconProps={this.state.expanded ? expandedIcon : collapsedIcon}
          onClick={(e) => {
            this.setState({
              expanded: !this.state.expanded,   
              Sales: false,
              Technical:false,
              Delivery:true,
              Legal:false
            });
          }}
          aria-expanded={this.state.expanded}
          // aria-controls={this._drawerDiv && this._drawerDiv.id}
        />
      </div>
      <div className={styles.section_right}>
      {this.state.Sales && 
              <Sales webURL={this.props.webURL}></Sales>
      }
      {this.state.Technical &&
      <Technical webURL={this.props.webURL}></Technical>
  }
  {this.state.Delivery &&
      <Delivery webURL={this.props.webURL}></Delivery>
  }
      {/* <Legal webURL={this.props.webURL}></Legal> */}
 
    
           </div>
          </div></div></div>

    );
  }
  
}
function _onLinkClick(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) {
  if ( item.name === 'Sales Feedback') {
   this.setState({
    Sales: true,
    Technical:false,
    Delivery:false,
    Legal:false
   })
  }
  else if ( item.name === 'Technical Feedback') {
    this.setState({
      Sales: false,
      Technical:true,
      Delivery:false,
      Legal:false
   })
  }else if ( item.name === 'Delivery Feedback') {
    this.setState({
      Sales: false,
      Technical:false,
      Delivery:true,
      Legal:false
    })
  }
}

