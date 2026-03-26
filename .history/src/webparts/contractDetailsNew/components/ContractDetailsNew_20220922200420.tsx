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

export interface IContractDetailsNewState {
  Sales: boolean;
  Technical:boolean;
  Delivery:boolean;
  Legal:boolean

}

const navLinkGroups: INavLinkGroup[] = [
  {
    links: [
      {
        name: 'Documents',
        url: 'http://example.com',
        key: 'key3',
        isExpanded: true,
        target: '_blank',
      },
      {
        name: 'Pages',
        url: 'http://msn.com',
        key: 'key4',
        target: '_blank',
      },
      {
        name: 'Notebook',
        url: 'http://msn.com',
        key: 'key5',
        disabled: true,
      },
      {
        name: 'Communication and Media',
        url: 'http://msn.com',
        key: 'key6',
        target: '_blank',
      },
    
    ],
  },
];


export default class ContractDetailsNew extends React.Component<IContractDetailsNewProps, IContractDetailsNewState> {
  constructor(props: IContractDetailsNewProps) {
    super(props);

    this.state = {
      Sales: false,
      Technical:false,
      Delivery:false,
      Legal:false
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
      <Nav
          onLinkClick={_onLinkClick}
          selectedKey="key3"
          ariaLabel="Nav basic example"
          // styles={navStyles}
          groups={navLinkGroups} />
      </div>
      <div className={styles.section_right}>
      if (this.state.Sales) {
        <Sales webURL={this.props.webURL}></Sales>
      }
      if (this.state.Technical) {
      <Technical webURL={this.props.webURL}></Technical>
  }
  if (this.state.Delivery) {
      <Delivery webURL={this.props.webURL}></Delivery>
  }
      {/* <Legal webURL={this.props.webURL}></Legal> */}
 
    
           </div>
          </div></div></div>

    );
  }
  
}
function _onLinkClick(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) {
  if (item && item.name === 'News') {
    alert('News link clicked');
  }
}

