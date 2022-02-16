import { LightningElement, api, wire, track } from 'lwc';
import getEmbeddingDataForReport from '@salesforce/apex/PowerBiEmbedManager.getEmbeddingDataForReport';
import getReportsData from '@salesforce/apex/PowerBiEmbedManager.getReportsData';
import getWorkspaceId from '@salesforce/apex/PowerBiEmbedManager.getWorkspaceId';
import getActiveReports from '@salesforce/apex/PowerBiEmbedManager.getActiveReports';
import ExportToFileInGroup from '@salesforce/apex/PowerBiEmbedManager.ExportToFileInGroup';
import GetExportToFileStatus from '@salesforce/apex/PowerBiEmbedManager.GetExportToFileStatus';
import GetFileOfExportToFileInGroup from '@salesforce/apex/PowerBiEmbedManager.GetFileOfExportToFileInGroup';
import getFavorites from '@salesforce/apex/PowerBiEmbedManager.getFavorites';
import insertFavorite from '@salesforce/apex/PowerBiEmbedManager.insertFavorite';
import deleteFavorite from '@salesforce/apex/PowerBiEmbedManager.deleteFavorite';
import getPowerBiAccessToken from '@salesforce/apex/PowerBiEmbedManager.getPowerBiAccessToken';
import JQuery2 from '@salesforce/resourceUrl/JQuery2';
import bootstrap_css from '@salesforce/resourceUrl/bootstrap_css';
import appInsights from '@salesforce/resourceUrl/ApplicationInsights';
import download from '@salesforce/resourceUrl/download';
import powerbijs from '@salesforce/resourceUrl/powerbijs';
import { loadScript, loadStyle } from 'lightning/platformResourceLoader';
import {subscribe, unsubscribe, MessageContext} from 'lightning/messageService';
import SelectedAccountMessageChannel from '@salesforce/messageChannel/SelectedAccountMessageChannel__c';


export default class PowerBiReport extends LightningElement {

  //Variable and method declarations

  //A boolean variable to check current visible report is favorite or not
  @api current_report_favorite = false;

  //A variable to check the Reports dropdown parent is closed or opened
  @api dd_parent=0;

  //Subscription variable is used in message service to get the account id from the other LWC 
  subscription = null;


  //workspace id is to store the Workspace id related to the selected account
  //Workspace id is set when the user selected an account
  @api WorkspaceId = '';

  //Report id is to store the current visible report id
  //Report Id is set when the report button is clicked in the reports dropdown
  @api ReportId = '';

  //Account id is set when the user selected the account in accounts dropdown
  //By default it is set to default selected option in the dropdown
  @api account_id = "";

  //This variable is to identify the current visible report
  //When user clicked on any report, it will be the current report
  //It is used in favorites, export and download functionalities
  @api current_report_name;

  //Used to check if the reports data is ready to render on the UI
  @track dataReady = false;

  //Store key-value pair for report name and report id 
  //Reports in the workspace which are active in status will be stored in this array
  @track reports_details = [];

  //A map object to store the list name and type of the report
  @track active_reports_list = new Map();

  //All the reports list available in the workspace related to current user
  @track account_specific_reports = [];

  //List of custom reports
  @track custom_reports=[];

  //List of standard reports
  @track standard_reports=[];

  //List of premium reports
  @track premium_reports = [];

  //store all the favorite report names
  @track favorite_report_names = [];

  //Store all the favorite report ids
  @track favorite_report_ids = [];

  //Message context is to exchange account id between two lightning components
  @wire(MessageContext) messageContext;

  //To get the active reports data from Salesforce objects
  @wire(getActiveReports) activeReports;

  //To get the favorite reports for the current user
  @wire(getFavorites) UserFavorites;

  //Embedding method call
  @wire(getEmbeddingDataForReport, {
    WorkspaceId: "$WorkspaceId",
    ReportId: "$ReportId"
  }) report;

  //Variables and methods declarations completed  

  //Function definitions

  //Method used to hold execution for specified number of milli seconds
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  //To reset all the variables when user changed the Account 
  resetData(){
    this.favorite_report_names.length = 0;
    this.current_report_favorite = false;
    this.favorite_report_ids.length = 0;
    this.account_specific_reports.length = 0;
    this.active_reports_list.length = 0;
    this.standard_reports.length = 0;
    this.premium_reports.length = 0;
    this.custom_reports.length = 0;
    this.reports_details.length = 0; 
  }

  //Reports dropdown close and open function
  //if dd_parent is false, it means parent dropdown is closed 
  //When user clicked on the dropdown it should be opened
  //When dropdown is opened and user clicked on it, it should be closed and update dd_parent to false
  parentDropdownClicked(){
    //get caret symbol and child dropdown elements 
    let caret = this.template.querySelector(".parent-caret");
    let acc_child = this.template.querySelector(".accordion-child");

    if(!this.dd_parent){
        caret.classList.add("rotate");
        this.dd_parent = !this.dd_parent;
        acc_child.style.display="block";
      }
    else{
      caret.classList.remove("rotate");
      this.dd_parent = !this.dd_parent;
      acc_child.style.display="none";
    }
  }


  //Drop down functionality for report types in reports dropdown
  childDropdownClicked(event){
    //Check the event triggered element and set the element to button 
    //User might have clicked on caret symbol or name. To avoid ambiguity element is set to button
    //Buton will consists the name of reports type
    let element = event.target;
    if(element.tagName=="P" ){
      element = element.parentElement;
    }
    else if(element.tagName == "path"){
      element = element.parentElement.parentElement.parentElement;
    }
    else if(element.tagName == "svg"){
      element = element.parentElement.parentElement;
    }

    //Check if the dropdown is already opened
    //Show collapse classes will open the dropdown 
    //collapsed class will close the drop down
    let classlist = element.classList;
    if(element.classList.contains("show")){
      classlist.add("collapsed");
      classlist.remove("show");
      console.log(classlist);
      element.nextElementSibling.style.display = "none";
    }
    else{
      //close all active report type dropdowns and open user clicked dropdown
      let all_active = this.template.querySelectorAll(".show"); 
      if(all_active){
          for(var i=0;i<all_active.length;i++){
            all_active[i].classList.add("collapsed");
            all_active[i].classList.remove("show");
            all_active[i].nextElementSibling.style.display="none";
          }
      }
      classlist.add("collapse");
      classlist.add("show");
      element.nextElementSibling.style.display = "block";
    }
  }


  //To Insert or Delete the favorite records in Salesforce objects
  //Take name of the report and check if it is already in favorites 
  //If it is in favorites, remove it from favorites
  //else add the report to favorites
  insertDeleteFavData(name_of_report){
    console.log("insert delete is called "+name_of_report);
    let report_name;
    let report_id;
    //Get the relevant report_id
    this.reports_details.forEach(function(item) {
    if(item["key"] == name_of_report){
        report_name = item["key"];
        report_id= item["value"];
        return true;
    }
    });

    if(this.favorite_report_names.includes(report_name)){
        deleteFavorite({report_id: report_id}).then(result=>{
        console.log("Deleted: "+result);
        });
    }
    else{
        insertFavorite({report_id: report_id}).then(result=>{
        console.log("inserted: "+result);
        });
    }
  }

  //This method will get triggered when Global Add to favorites option is clicked
  //Display the favorites pop up
  //Add the report to favorite_report_names list and make current report favorite as true
  addToFav(event){
    //Display the favorites pop up
    this.template.querySelector(".embed").classList.add("blur");
    this.template.querySelector(".pop-up").style.display = "block";
    this.insertDeleteFavData(this.current_report_name);
    this.current_report_favorite=true;

    //Add to favorites list
    if(!this.favorite_report_names.includes(this.current_report_name)){
      this.favorite_report_names.push(this.current_report_name);
    }
    this.fillFavoriteIcons();
  }

  //Close the favorites pop up
  closeFav(event){
    this.template.querySelector(".pop-up").style.display="none";
    this.template.querySelector(".embed").classList.remove("blur");
  };


  //This method get triggered when the global remove as favorite button is clicked 
  remAsFav(event){
    this.insertDeleteFavData(this.current_report_name);
    this.current_report_favorite=false;   
    if(this.favorite_report_names.includes(this.current_report_name)){
      //Delete the report name from favorites list
      var index = this.favorite_report_names.indexOf(this.current_report_name);
      if (index !== -1) {
        this.favorite_report_names.splice(index, 1);
      }
    }
    this.fillFavoriteIcons();
  }

  //This method get triggered If we click on the favorites icon beside the report name in reports dropdown 
  //Stop event propogation to avoid parent button click -> To avoid report button click when user clicked on favorite icon
  //Check if the report is already in favorites -> if the report is in favorite, remove from favorite and remove the fill color for favorite icon
  //If it is not in the favorites list, then add the report to list and fill the favorites icon color
  favoriteIconClicked(event){
    event.stopPropagation();
    console.log("clicked fav icon");
    //get report name from the event 
    let report_name = event.target.dataset.id;
    //make insertion or deletion in Salesforce object
    this.insertDeleteFavData(report_name);
    //Make insertion or deletion in favorite reports list array
    if(this.favorite_report_names.includes(report_name)){
      let index = this.favorite_report_names.indexOf(report_name);
      console.log(index);
      if (index !== -1) {
        this.favorite_report_names.splice(index, 1);
        //Check if favorite icon of current report is clicked and removed from favorites, Update the global button as "add to favorites"
        if(this.current_report_name==report_name){
          this.current_report_favorite = false;
        }
      }
    }else{
      this.favorite_report_names.push(report_name);
      //Check if favorite icon of current report is clicked and added to favorites , Update the global button as "Remove from favorites"
      if(this.current_report_name==report_name){
        this.current_report_favorite = true;
      }
    }
    this.fillFavoriteIcons();
  }

  //If the report name is in favorites list, this method will fill the favorite icon beside that report name
  fillFavoriteIcons(event){
    //Fetch all the report icons as html elements
    let elements = this.template.querySelectorAll(".favorite-icon");

    //Loop over each element and check if the report is in favorites list or not
    // if the report is in favorites, add the class name fill-fav of the element classList. 
    //Else remove the class name from the element classList
    for(let i=0; i< elements.length;i++){
        if(this.favorite_report_names.includes(elements[i].dataset.id)) {
        elements[i].classList.add("fill-fav");
        //Add the opacity and stroke for button element to disable favorite icon visibility on hovering report button
        elements[i].parentElement.parentElement.style.opacity = 100;
        elements[i].parentElement.parentElement.style.stroke = "none";
        }
        else{
          if(elements[i].classList.contains("fill-fav")){
          elements[i].classList.remove("fill-fav");
          }
          //Remove the opacity and stroke for button element to enable favorite icon visibility on hovering report button
          elements[i].parentElement.parentElement.style.removeProperty("opacity");
          elements[i].parentElement.parentElement.style.removeProperty("stroke");
        } 
    }
  }

  //accountClicked function will fetch the related workspace_id for the selected account
  accountClicked(value){
    this.account_id = value.accountId;
    getWorkspaceId({
      account_id: this.account_id
    })
      .then((result) => {
        this.WorkspaceId = result;
        console.log("here is workspace id");
        console.log(this.WorkspaceId);
        // Invoke getReports function to get the list of all reports in the workspace
        //Pass the workspace id as parameter
        this.getReports(this.WorkspaceId);
      })
      .catch((error) => {
        console.log(error);
      })
      .finally(() => {
        console.log('Finally');
      })
  }

    //this function will get the report_name and report_id for all the reports in the workspace 
    getReports(WorkspaceId) {
      getReportsData({
        workspace_id: WorkspaceId
      })
        .then((result) => {
        //result contains details of all the reports in the workspace   

        //Reset all the variable, As we are fetching data from a new workspace 
        this.resetData();

        //List of report names available in the workspace/account
        for (let key in result){
            this.account_specific_reports.push(key);
        }
        
        //Fetch all the favorite report ids for current user from Salesforce object
        for(let fav_report in this.UserFavorites.data){
          let report_id = this.UserFavorites.data[fav_report].Report_ID__c;
          this.favorite_report_ids.push(report_id);
        }
        
        //store favorite report names in favorite_report_names array
        //Check if the each favorite report status is active 
        //There might be records in favorites which are not in active state. 
        //So, Consider intersection of the two lists active reports and favorite reports
    
      
        //Check if each report in active reports list is in the list of reports in workspace
          for(let item in this.activeReports.data){
            let report_id = this.activeReports.data[item].Id__c;

            let name_of_report = this.activeReports.data[item].Report_Name__c;

            let type_of_report = this.activeReports.data[item].Report_Type__c;

            if(this.favorite_report_ids.includes(report_id)){
              this.favorite_report_names.push(name_of_report);
            }
            
            if(this.account_specific_reports.includes(name_of_report)){

                if(type_of_report === "Standard"){

                  this.standard_reports.push(name_of_report);

                }
                else if(type_of_report === "Premium"){

                  this.premium_reports.push(name_of_report);

                }
                else{

                  this.custom_reports.push(name_of_report);

                }

                this.active_reports_list.set(name_of_report , type_of_report);
          }

          }   

          for(let key in result){
            if(this.active_reports_list.has(key)){
              this.reports_details.push({key:key,value:result[key]});
            }
        }
        //Boolean to check if the reports data is ready to render in the dropdown
        this.dataReady = true;

        })
        .catch((error) => {
          console.log(error);
        })
        .finally(() => {
          console.log('Finally');

          //Fill favorites icons as per favorites list
          this.fillFavoriteIcons();
        })
    }

  //When the user clicked on report button in dropdown menu, It will call this function  
  reportClicked(event){    
    //
    let report_id;
    var report_name;
    //For each loop to search details of clicked report in the reports_details (contains all the reports data in current account/workspace)
    //Get the report id from reports_details and set the ReportId variable to Embed the report in the widget
    this.reports_details.forEach(function(item) {
        report_name = item["key"];
        if(report_name==event.target.dataset.id){
          report_id= item["value"];
          return true;
        }
    });
    //Set current report name 
    this.current_report_name = event.target.dataset.id;
    //Check if the current report is favorite
    if(this.favorite_report_names.includes(this.current_report_name)){
      this.current_report_favorite = true;
    }
    else{
      this.current_report_favorite = false;
    }
    //Set the ReportId so that getEmbeddingDataForReport method will automatically triggers
    this.ReportId = report_id;
  }


  exportToPDF(){
    this.exportReport("PDF");
    //Blocking the button code here
  }

  exportToPPT(){
    this.exportReport("PPTX");
    //blocking the button code here
  }

  //Export and download functionality for the report 
  async exportReport(file_format){
    let maxNumberOfRetries = 3;
    let retryAttempt = 1;
    let status;
    //Call Power BI ExportToFileInGroup API to create the export job for the requested report
    do{
    let exportId = await ExportToFileInGroup({ workspace_id: this.WorkspaceId, report_id: this.ReportId, format: file_format });
    console.log(exportId);
      do{
        //Wait and call GetExportToFileStatus API to get the status of the export job
        await this.sleep(4000);
        status = await GetExportToFileStatus({ workspace_id: this.WorkspaceId, report_id: this.ReportId, export_id: exportId });
        console.log(status);
        
        if(status == 'Failed' || status == null){
          console.log('Job Failed');
          break;
        }
        else if(status == 'Succeeded'){
          //Get the file from API and download it the user in the requested format
          console.log('Job status success');
          GetFileOfExportToFileInGroup({ workspace_id: this.WorkspaceId, report_id: this.ReportId, export_id: exportId }).then((result)=>{
            let downloadLink = document.createElement("a");
            downloadLink.setAttribute("type", "hidden");
            downloadLink.href = "data:application/octet-stream;base64,"+result;
            downloadLink.download = this.current_report_name+'.'+file_format;
            document.body.appendChild(downloadLink);
            downloadLink.click();
            downloadLink.remove();
          });
          break;
        }

      }
      while(status=='NotStarted' || status == 'Running');

      retryAttempt++;
    }while(status!='Succeeded' && retryAttempt<=maxNumberOfRetries);
  }



  renderedCallback() {

   this.isRenderCallbackActionExecuted = true;
    // const appInsights = new ApplicationInsights({ config: {
    //   instrumentationKey: '34cc9d36-c1a1-4564-95f6-d0f4b169bfcd',
    //   autoTrackPageVisitTime: true,
    // } });
    // appInsights.loadAppInsights();
    // appInsights.trackPageView();
    // appInsights.trackEvent();

    console.log(getPowerBiAccessToken());

    //subscribing to the lightning message service to get the account selected from the jsasDashboardLWC
    if(!this.subscription){
        this.subscription = subscribe(
          this.messageContext,
          SelectedAccountMessageChannel,
          (value) =>this.accountClicked(value)
        )
    }

    //Load all static resources 
    Promise.all([loadScript(this, powerbijs),loadScript(this, appInsights),loadScript(this,JQuery2),loadScript(this,download),loadStyle(this,bootstrap_css)]).then(() => {
      var snippet = {
              config: {
                  instrumentationKey: "34cc9d36-c1a1-4564-95f6-d0f4b169bfcd"
              }
          };
          var init = new Microsoft.ApplicationInsights.ApplicationInsights(snippet);
          var appInsights = init.loadAppInsights();
          console.log(appInsights);
          appInsights.trackPageView();
          appInsights.trackEvent({name: "download eventttttt"});
          appInsights.trackEvent({
            name: 'some event',
            properties: { 
                prop1: 'string',
                prop2: 123.45,
                prop3: { nested: 'objects are okay too' }
            }
        });
        appInsights.trackPageView({ name: 'some page' });
     appInsights.trackPageViewPerformance({ name: 'some page', url: 'some url' });
     appInsights.trackException({ exception: new Error('some error') });


     var telemetryInitializer = (envelope) => {
      envelope.data.someField = 'This item passed through my telemetry initializer';
      };
      appInsights.addTelemetryInitializer(telemetryInitializer);
      appInsights.trackTrace({ message: 'This message will use a telemetry initializer' });
      appInsights.trackMetric({ name: 'some metric', average: 42 });
      appInsights.trackDependencyData({ absoluteUrl: 'some url', responseCode: 200, method: 'GET', id: 'some id' });
      appInsights.startTrackPage("pageName");
      appInsights.stopTrackPage("pageName", null, { customProp1: "some value" });
      appInsights.startTrackEvent("event");
      appInsights.stopTrackEvent("event", null, { customProp1: "some value" });
      appInsights.flush();

      if (this.report.data) {

        if (this.report.data.embedUrl && this.report.data.embedToken) {
          var reportContainer = this.template.querySelector('[data-id="embed-container"');
          var reportId = this.report.data.reportId;
          var embedUrl = this.report.data.embedUrl;
          var token = this.report.data.embedToken;
          //Create configuration object to embed the report
          var config = {
            type: 'report',
            id: reportId,
            embedUrl: embedUrl,
            accessToken: token,
            pageView: 'fitToWidth',
            // pageName: 'ReportSection5f908a2221a6b0e54dd2',
            tokenType: 1,
            settings: {
              panes: {
                filters: { expanded: false, visible: false },
                pageNavigation: { visible: false }
              }
            }
          };

          //Embed the report and display it within the div container.
          var report = powerbi.embed(reportContainer, config);
          // console.log(report.allowedEvents);
          report.on('pageChanged', event => {
            console.log("page changeddd");
            console.log(event);
        });

        
        }
        else {
          console.log('no embedUrl or embedToken');
        }

      }
      else {
        console.log('no report.data yet');
      }


    });

  }

}

