import * as React from 'react';
import styles1 from '../../components/ContractApproval.module.scss';
import styles from './LegalTaskFormWebPart.module.scss';
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


export interface ILegalTaskProps {
  defaultCollapsed?: boolean;
  className?: string;
  webURL: string;
}

export interface ILegalTaskState {
  expanded: boolean;
  Items: any;
  HTML: any;
}

// const collapsedIcon: IIconProps = { iconName: 'ChevronRight', className: styles.FeedbackChevron };
// const expandedIcon: IIconProps = { iconName: 'ChevronDown', className: styles.FeedbackChevron };

export class LegalTask extends React.Component<ILegalTaskProps, ILegalTaskState> {
  private _drawerDiv: HTMLDivElement = undefined;
  ContractStatus = "";

  constructor(props: ILegalTaskProps) {
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
  public render(): React.ReactElement<ILegalTaskProps> {
    return (
      <div className={styles.LegalTaskForm}>
        <div className={styles.container}>
        <div className={styles.row}>
        <div className={styles.column}>

<div id="LegalTaskForm">
<div id="oppID" style={{ display: 'none' }}  className={styles.pagesubtitle}></div>
<div className={styles.sectionblock}>
<Label className={styles.headers}><u>Legal Feedback Form</u></Label>
</div>
  <div className="row">
    <div className="col col-lg-12"> 
    <div className={styles.form_group}> 
      <Label className={styles.lablecontrol}>Comments:<span className={styles.estric}>*</span></Label> 
      <textarea placeholder="LegalComments" id="LegalComments" className={styles.form_control} required></textarea> 
      <label id="LegalErr" style={{ display: 'none',color:'red',fontSize:'x=small' }} className="error">This field is required.</label> 
      </div>
    </div>
  </div>

  <div id=""className="row">
    <div className="col col-lg-12">     
    <div className={styles.form_group}>
        <div className={styles.mb_1}><input type="file" className={styles.form_control} id="uploadFile"></input></div> 
        <span id="fileErr1" style={{ display: 'none' }} className={styles.errorlable}>Please select file to upload.</span>
        <div className={styles.mb_1}>
            <PrimaryButton id="UploadDocs" text="Upload" className={styles.btn} onClick={() => this.uploadFileFromControl()} >
                  </PrimaryButton>
        </div> 
        <div className={styles.file_list}><ul id="filesList"></ul></div>
        </div>
    </div>
  </div>

  
  
  <div id="loader" className={styles.modal}>
            <div className="">
              <div
                className={styles.loader}
               style={{ margin: '200px auto' }}
              ></div>
            </div>
          </div>

            <div id="ExistingModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close0()}>
                &times;
              </span>
                <label className={styles.header2}>Former feedback is already approved!</label>
              </div>
            </div>

            <div id="SubmittedModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close1()}>
                &times;
              </span>
                <label className={styles.header2}>Feedback Submitted!</label>
              </div>
            </div>

            <div id="DraftModal" className={styles.modal}>
              <div className={styles.modalcontent}>
              <span className={styles.close} onClick={() => this.Close2()}>
                &times;
              </span>
                <label className={styles.header2}>Draft Saved!</label>
              </div>
            </div>


            <div className="row">
              <div className="col col-lg-12">
              <div className={styles.form_footer}>
                <PrimaryButton text="Draft" className={styles.btn}  style={{ borderRadius: "10px",marginRight:"10px"}} id="LegalDraft"   onClick={() => this.LegalDraft()} >Save as Draft</PrimaryButton>
                <PrimaryButton text="Submit" className={styles.btn}  style={{ borderRadius: "10px",marginRight:"10px"}} id="LegalSubmit"   onClick={() => this.LegSubmitData()}></PrimaryButton>
                <PrimaryButton text="Cancel" className={styles.btn}  style={{ borderRadius: "10px"}} id="LegalCancel"   onClick={() => this.LegCancelForm()}></PrimaryButton>
              </div></div>
            </div>

 </div>



        </div>
      </div>
      </div>
      </div>
    );
  }

  
  private setForm() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    sp.web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((data) => {
      if(data.FeedbackStatus == "Submit"){
        $("#LegalDraft").hide();
      }
      else{
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
    setTimeout( () => { window.parent.location.href = hostUrl; }, 1000);

  }

  private Close2() {
    var modal = document.getElementById("DraftModal");
    var hostUrl = this.props.webURL;
    modal.style.display = "none";
    $('#LegalTaskForm').fadeOut(2500);
    setTimeout( () => { window.parent.location.href = hostUrl; }, 1000);
  }





  private uploadFileFromControl() {
    $("#loader").show();

    var files = (document.getElementById('uploadFile') as HTMLInputElement).files;
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl+ '/ProjectDocuments/'+$('#oppID').text();
    var fileArr = [];

    
      if (files.length !== 0) {
        $("#fileErr1").hide();
  
        for (var g = 0; g < files.length; g++) {
          var file = files[g];
          fileArr.push(file);
          let web = Web(this.props.webURL);
    var hostUrl = '/sites/demo/ProjectDocuments/' + $('#oppID').text();
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web.getFolderByServerRelativeUrl(hostUrl).files.add(file.name, file, true).then(f => {
            f.file.getItem().then(item => {
              item.update({
                Title: "Legal Documents File"
              }).then(h => {
                this.getFiles();
                $("#uploadFile").val("");
                setTimeout( ()=> {$('#loader').hide();}, 1000);
              });
            });
          });
        }
      }
      else {
        $("#fileErr1").show();
        setTimeout( ()=> {$('#loader').hide();}, 1000);
      }
  }

  private getFiles() {
    let attachmentfiles: string = "";
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl+ '/ProjectDocuments/'+$('#oppID').text();
    let web = Web(this.props.webURL);
    var hostUrl = '/sites/demo/ProjectDocuments/' + $('#oppID').text();
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web.getFolderByServerRelativeUrl(hostUrl).files.filter("Title eq 'Legal Documents File'").get().then(files => {
      if(files.length !== 0){
        for (var i = 0; i < files.length; i++) {
          var title = files[i].Title;
          var valu = files[i].Name;
          attachmentfiles += `<li id="FileListItem` + i + `" value=` + valu + `><button className="CloseBtn"`
            + `id="CloseBtn">&times;</button><Label style="display:none;">` + files[i].UniqueId + `</Label>`
            + `<Label id="filesnames` + i + `">`
            + `&nbsp&nbsp&nbsp&nbsp${title}&nbsp-&nbsp</Label>`
            + `<a href="${files[i].ServerRelativeUrl} target="_blank" id="FileName` + i + `">`
            + `${files[i].Name}</a></li>`;
            document.getElementById("#filesList").innerHTML = attachmentfiles;
        }
        this.DeleteFileEvent();  
      }
      else{
        $("#filesList").empty();
      }
    }).then(() =>{
      setTimeout( ()=> {$('#loader').hide();}, 1000);      
    });
  }

  private DeleteFileEvent() {
    // var hostUrl = this.context.pageContext.site.serverRelativeUrl+ '/ProjectDocuments/'+$('#oppID').text();
    let web = Web(this.props.webURL);
    var hostUrl = '/sites/demo/ProjectDocuments/' + $('#oppID').text();
    // var hostUrl =  this.props.webURL+ '/ProjectDocuments/' + this.state.OppID;
    web.getFolderByServerRelativeUrl(hostUrl).files.filter("Title eq 'Legal Documents File'").get().then(files => {
      for (var i = 0; i < files.length; i++) {
        var btnClass = document.getElementsByClassName("CloseBtn")[i] as HTMLElement;
        btnClass.addEventListener('click', (f) => {
          var xyz = f.currentTarget as HTMLElement;
          var fName = xyz.nextElementSibling.nextElementSibling.nextElementSibling.textContent;
          sp.web.getFolderByServerRelativeUrl(hostUrl).files.getByName(fName).delete().then( (g) => {
          }).then(check => {
            $("#loader").show();
            this.getFiles();
            setTimeout( ()=> {$('#loader').hide();}, 1000);      
          });
        });
      }
    });
  }

  private LegCancelForm() {
  var hostUrl = this.props.webURL;
  window.parent.location.href = hostUrl;
}

  private LegSubmitData() {

    var isvalid = true;

    if ($("#LegalComments").val() === "") {
      $("#LegalErr").show();
      isvalid = false;
    }
    else
      $("#LegalErr").hide();
    if (isvalid) {
      this.LegSubmitData1();
    }
  }

  private async LegSubmitData1() {
    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects").items.getById(parseInt(itemID)).get().then((Pdata) => {

      if (Pdata.LegalDone != 2) {
        this.LegSubmitData2();
      }
      else{
        var modal = document.getElementById("ExistingModal");
        modal.style.display = "block";
      }
    });
  }

  private LegSubmitData2() {


    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).select('Id', 'Status', 'DeliveryDone', 'TechnicalDone', 'LegalDone').get().then((data) => {

        var status = data.Status;
        if (data.DeliveryDone == 1) {
          status = "Management Action Awaited";
        }
        let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects")
          .items.getById(parseInt(itemID)).update({
            LegalComments: $("#LegalComments").val().toString(),
            Status: status,
            LegalDone: 1,
            FeedbackStatus: "Submit"
          }).then(i => {
            let web = Web(this.props.webURL);
    web.lists.getByTitle('CompleteTask').items.add({
              ProjectID: parseInt(itemID),
              TaskType: "Legal"
            }).then(newListItem => {
              $("#loader").show();
              setTimeout( ()=> {$('#loader').hide();}, 3000);
              setTimeout( ()=> {$('#SubmittedModal').show();}, 3500);   
            });
          });
      });

  }



  private LegalDraft() {

    var isvalid = true;

    if ($("#LegalComments").val() === "") {
      $("#LegalErr").show();
      isvalid = false;
    }
    else
      $("#LegalErr").hide();
    if (isvalid) {
      this.LegalDraft1();
    }
  }

  private async LegalDraft1() {

    const itemID = new URLSearchParams(window.location.search).get('itemid');
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Projects")
      .items.getById(parseInt(itemID)).select('Id', 'Status', 'DeliveryDone', 'TechnicalDone', 'LegalDone').get().then((data) => {

        let web = Web(this.props.webURL);
    web.lists.getByTitle("Projects")
          .items.getById(parseInt(itemID)).update({
            LegalComments: $("#LegalComments").val().toString(),
            FeedbackStatus: "Draft"
          }).then(i => {
            $("#loader").show();
            setTimeout( ()=> {$('#loader').hide();}, 3000);
            setTimeout( ()=> {$('#DraftModal').show();}, 3500);    
          });
      });
  }




}


