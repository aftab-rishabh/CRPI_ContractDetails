import * as React from 'react';
import styles from './Accordion.module.scss';
import { IAccordionProps, IAccordionState } from './index';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
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
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
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
            {this.props.children}
          </div>
        }
        </div>
      </div>
    );
  }

  public async fetchData() {
       
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists.getByTitle("EmployeeDetails").items.select("*", "Employee_x0020_Name/Title").expand("Employee_x0020_Name/ID").get();
    console.log(items);
    this.setState({ Items: items });
    let html = await this.getHTML(items);
    this.setState({ HTML: html });
  }

  public async getHTML(items) {
    var tabledata = <table >
      <thead>
        <tr>
          <th>Employee Name</th>
          <th>Hire Date</th>
          <th>Job Description</th>
        </tr>
      </thead>
      <tbody>
        {items && items.map((item, i) => {
          return [
            
          ];
        })}
      </tbody>

    </table>;
    return await tabledata;
  }


  
}


