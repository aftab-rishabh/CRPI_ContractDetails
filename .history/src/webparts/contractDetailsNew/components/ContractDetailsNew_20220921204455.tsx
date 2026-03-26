import * as React from 'react';
import styles from './ContractDetailsNew.module.scss';
import { IContractDetailsNewProps } from './IContractDetailsNewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Accordion } from '../../accordion';
import { Nav, INavLink, INavStyles, INavLinkGroup } from '@fluentui/react/lib/Nav';
import SplitPane, { Pane } from 'react-split-pane';
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


export default class ContractDetailsNew extends React.Component<IContractDetailsNewProps, {}> {
  public render(): React.ReactElement<IContractDetailsNewProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
    
      <div className="row"><div style={{ width: "30%" }}>
        <Nav
          onLinkClick={_onLinkClick}
          selectedKey="key3"
          ariaLabel="Nav basic example"
          // styles={navStyles}
          groups={navLinkGroups} />
      </div><div style={{ width: "70%" }}>
          <label htmlFor="Testing">{"Testing"}</label>
        </div></div>

    );
  }
  
}
function _onLinkClick(ev?: React.MouseEvent<HTMLElement>, item?: INavLink) {
  if (item && item.name === 'News') {
    alert('News link clicked');
  }
}

