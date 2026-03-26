import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './managementConsentTaskFormWebPart.module.scss';
import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Label } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
require('.././css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
/**
 * Icon styles. Feel free to change them
 */


export interface ImanagementConsentTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ImanagementConsentTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class managementConsentTask extends React.Component<ImanagementConsentTaskProps, ImanagementConsentTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";

  constructor(props: ImanagementConsentTaskProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed == null ? false : !props.defaultCollapsed,
      Items: [],
      HTML: []
    };
  }
  public async componentDidMount() {

    this.setForm();
  }
  public render(): React.ReactElement<ImanagementConsentTaskProps> {
    return (
      <div className={styles.managementConsentTaskForm}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>

           



            </div>
          </div>
        </div>
      </div>
    );
  }


  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((data) => {
      if (data.FeedbackStatus == "Submit") {
        $("#LegalDraft").hide();
      }
      else {
        $("#LegalDraft").show();
      }

      $("#LegalComments").val(data.LegalComments);
      $("#LegalCommentsByMKTG").val(data.LegalCommentsByMKTG);
      $('.oppID').text(data.OpportunityID);
    }).then(i => {
      this.getFiles();
    });
  }


  private Close0() {
    var modal = document.getElementById("ExistingModal");
    modal.style.display = "none";
  }
  private Close1() {
    var modal = document.getElementById("SubmittedModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#LegalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 1000);

  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#LegalTaskForm').fadeOut(2500);
    setTimeout(() => { window.parent.location.href = hostUrl; }, 1000);
  }






}


