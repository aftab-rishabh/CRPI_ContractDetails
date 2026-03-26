import * as React from 'react';
import styles from './Feedback.module.scss';

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


export interface ILegalProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ILegalState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class Legal extends React.Component<ILegalProps, ILegalState> {
  private _drawerDiv: HTMLDivElement = undefined;
  constructor(props: ILegalProps) {
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
  public render(): React.ReactElement<ILegalProps> {
    return (

      <div className={styles.Feedback}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.viewdetail_block}>

              <div className="LegalData">

                <div className={styles.sectionblock}>
                  <Label className={styles.headers}><u>Legal Feedback</u></Label>
                </div>

                <div id="spListContainer" style={{ display: 'block' }}>

                  <div className={styles.sectionblock}>
                    <div className={styles.viewanswer_listing}>
                      <ul>

                        <li>
                          <div className={styles.viewcomment}>Legal Comment:</div>
                          <div className={styles.viewanswer} id="LegalComment">{this.state.Items}</div>
                        </li>


                      </ul>
                    </div>
                  </div>


                </div>

              </div>

            </div>
          </div>
        </div>
      </div>

    );
  }

  public async fetchData() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    this.getYesNoAndComment();

  }

  private async getYesNoAndComment() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((item) => {
      var LegalComments = item.LegalComments;
      if (LegalComments === null) {
        LegalComments = "";
      }
      else if (item.FeedbackStatus === "Draft") {
        LegalComments = "";
      }
      this.setState({ Items: LegalComments });
      if (item.WithTemplate == "Yes") {
        document.getElementById("spListContainer").style.display = "none";
      }

    }).then(x => {

      this.HideNAData();
    });
  }

  private HideNAData() {
    if (document.getElementById('LegalComment').innerHTML == "NA") {
      document.getElementById('LegalComment').style.display = "none";
    }
  }

}


