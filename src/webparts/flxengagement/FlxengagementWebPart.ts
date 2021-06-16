import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FlxengagementWebPart.module.scss';
import * as strings from 'FlxengagementWebPartStrings';

import "../../ExternalRef/css/bootstrap.css";
import "../../ExternalRef/css/style.css";

import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/js/bootstrap.js";
import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");

// var pagename = "";


 import * as $ from "jquery";
 import * as moment from "moment";  
 var Badgingdays="";
 var itemid = "";
 var listUrl = ""
 let LGUID = "";
 let SiteName = "";
 let SelectedImage = "";
 var width="",height="",widthedit="",heightedit="";
 var FilteredAdmin =[];
var currentuser = "";
export interface IFlxengagementWebPartProps {
  description: string;
}

export default class FlxengagementWebPart extends BaseClientSideWebPart<IFlxengagementWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {    
    listUrl = this.context.pageContext.web.absoluteUrl;
    currentuser = this.context.pageContext.user.email;
    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
    SiteName = listUrl.split("/")[2]
    console.log(SiteName);        

    this.domElement.innerHTML = `
    <div class="viewallannounce text-end ">
    
  <!--  <button class="btn btn-outline-theme  rounded-0"  data-bs-toggle="modal" data-bs-target="#exampleModalscrollengage">View All</button> -->
  <a href="#" class="info"  class="color-info" id="ViewAll">View All</a> 
  <a href="#" class="info"  class="color-info" id="ShowVisible">End User View</a> 
    </div>  
    <div class="loader-section" style="display:none"> 
    <div class="loader"></div>  
    </div> 
    <div class="flx-engagement-section">
             
    <!-- <div class="addiconengage"> 
    <span class="addiconflxengage"data-bs-toggle="modal" data-bs-target="#staticBackdroptwo" ></span></div> -->
    <div class="engagement-section d-flex flex-wrap" id="engageedit" >
    <!-- <div class="q-link m-2 border text-center p-2">  
    <div class="iconaddengage"> 
    <span class="editimageflxengage"data-bs-toggle="modal" data-bs-target="#staticBackdropone"></span></div>
   
    <img class="q-link-img" src="https://media-exp1.licdn.com/dms/image/C5603AQG1sCX4C8uhCg/profile-displayphoto-shrink_800_800/0/1615561085905?e=1626307200&v=beta&t=SbQf_3OwRf5JSfMqg9fXLYvIwcfYWOLKl_lmJGr4qSE" alt="img"/>
    <div class="q-link-title">Personal Coaching</div>
    </div>
    <div class="q-link m-2 border text-center p-2">
    <img class="q-link-img" src="https://media-exp1.licdn.com/dms/image/C5603AQG1sCX4C8uhCg/profile-displayphoto-shrink_800_800/0/1615561085905?e=1626307200&v=beta&t=SbQf_3OwRf5JSfMqg9fXLYvIwcfYWOLKl_lmJGr4qSE" alt="img"/>
    <div class="q-link-title">Education</div>
    </div> 
    <div class="q-link m-2 border text-center p-2">
    <img class="q-link-img" src="https://media-exp1.licdn.com/dms/image/C5603AQG1sCX4C8uhCg/profile-displayphoto-shrink_800_800/0/1615561085905?e=1626307200&v=beta&t=SbQf_3OwRf5JSfMqg9fXLYvIwcfYWOLKl_lmJGr4qSE" alt="img"/>
    <div class="q-link-title">Education</div>
    </div> 
    <div class="q-link m-2 border text-center p-2">
    <img class="q-link-img" src="https://media-exp1.licdn.com/dms/image/C5603AQG1sCX4C8uhCg/profile-displayphoto-shrink_800_800/0/1615561085905?e=1626307200&v=beta&t=SbQf_3OwRf5JSfMqg9fXLYvIwcfYWOLKl_lmJGr4qSE" alt="img"/>
    <div class="q-link-title">Personal Coaching</div>
    </div>
    <div class="q-link m-2 border text-center p-2">
    <img class="q-link-img" src="https://media-exp1.licdn.com/dms/image/C5603AQG1sCX4C8uhCg/profile-displayphoto-shrink_800_800/0/1615561085905?e=1626307200&v=beta&t=SbQf_3OwRf5JSfMqg9fXLYvIwcfYWOLKl_lmJGr4qSE" alt="img"/>
    <div class="q-link-title">Education</div>
    </div>-->
    
   
    </div>
    
    </div>  
    <!--<div class="card text-center" style="width: 9rem; height:10.5rem ;border-radius:0">
    <div class="card-body my-4">
    <span class="engage-add-icon" data-bs-toggle="modal" data-bs-target="#staticBackdroptwo"></span>
    <p class="engage-title my-2">Add Link</p>
    </div>
  </div>-->
                                             
                                                       
                                                <!-- Modal -->
 
 <div class="modal fade" id="staticBackdropone" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="staticBackdropLabel" aria-hidden="true">
   <div class="modal-dialog engagement-modal-dialog ">
     <div class="modal-content rounded-0">  
       <div class="modal-header modal-tile-header">   
         <h5 class="modal-title w-100 text-center modalengage-color" id="staticBackdropLabel"> Edit Link </h5>
      <!--   <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button> -->
       </div>
        
       
       <div class="modal-body  modalbody-flexengage">
       <div class="row align-items-center my-3"><div class="col-4 titleflxengageman">Title</div><div class="col-1">:</div><div class="col-7">
       <input type="text" class="form-control rounded-0" id="TitleFLXEngagement" aria-describedby=""></div></div>
       <div class="row align-items-center my-3"><div class="col-4 titleflxengageman">Url</div><div class="col-1">:</div><div class="col-7">
       <input type="text" class="form-control rounded-0" id="URLFLXEngagement" value="" aria-describedby=""></div></div>
 
      <!-- <div class="row align-items-center my-3"><div class="col-4">OpeningNewTab</div>
       <div class="col-1">:</div><div class="col-7">
       <div class="form-check">
   <input class="form-check-input rounded" type="checkbox" value="Yes" id="checkboxopenewtabFLXEngagement">
  
 </div></div></div>                               
       
      
       <div class="row align-items-center my-3"><div class="col-4">Visible</div><div class="col-1">:</div><div class="col-7 custom-arrow">
       <div class="form-check">
   <input class="form-check-input rounded-0" type="checkbox" value="Yes" id="checkboxvisibleFLXEngagement">
   
 </div></div></div>  -->
 
 <div class="row align-items-center my-3"><div class="col-4">Url Properties</div><div class="col-1">:</div>
 <div class="col-7">
 <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">
 
 <!--<input type="checkbox" class="btn-check" id="checkboxvisibleFLXEngagement" autocomplete="off">
 <label class="btn btn-engage-outline-theme" for="checkboxvisibleFLXEngagement">Visible</label>       
 
 <input type="checkbox" class="btn-check" id="checkboxopenewtabFLXEngagement" autocomplete="off">
 <label class="btn btn-engage-outline-theme" for="checkboxopenewtabFLXEngagement">Open a new tab</label>-->
  
 
   
   <input type="checkbox" class="btn-check" id="checkboxvisibleFLXEngagement" autocomplete="off">
   <label class="btn btn-outline-theme" for="checkboxvisibleFLXEngagement">Visible</label>
 
   <input type="checkbox" class="btn-check" id="checkboxopenewtabFLXEngagement" autocomplete="off">
   <label class="btn btn-outline-theme" for="checkboxopenewtabFLXEngagement">Open a new tab</label>
 </div> 
 </div>
 </div>
       <div class="row align-items-start my-3"><div class="col-4 titleflxengageman">Image</div><div class="col-1">:</div><div class="col-7">
        <input type="file" class="form-control-file custom-life-engage" class="mt-1" id="File1FLXengageEdit" accept="image/*">
        <div id="engagementUpdateFileEmpty"></div>
        <div id="engagementEditFile"></div>
        </div></div>
     </div>
 
 
       
     <div class="modal-footer  modal-tile-footer justify-content-between"> 
     <div class="btns-left">
     <button type="button" class="btn btn-sm btn-danger rounded-0" id="engagementDeleteModal"  style="" data-bs-toggle="modal" data-bs-target="#EngagementDeleteModal">Delete</button>
     </div>
 
       <div class="btns-right d-flex">
       <div class="addScreen">
       <button type="button" class="btn btn-sm btn-secondary rounded-0" id="btnengagementEditClose" data-bs-dismiss="modal">Close</button>
       <button type="button" class="btn btn-sm btn-theme rounded-0" id="btnmodalSubmit" style="display: none;">Submit</button>
       </div>
       <div class="viewScreen">
       <!--<button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Close</button>-->
       <button type="button" class="btn btn-sm btn-theme ms-2 rounded-0" id="btnUpdateengage" style="">Update</button>
       </div>
       </div>
     </div>
     </div>
   </div>
 </div>  
 
 
 
 
 
                                                    <!-- Modal Add -->
 
 
 <div class="modal fade" id="staticBackdroptwo" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="staticBackdropLabel" aria-hidden="true">
   <div class="modal-dialog  engagement-modal-dialog">
     <div class="modal-content rounded-0">
       <div class="modal-header  modal-tile-header ">
         <h5 class="modal-title w-100 text-center modalengage-color" id="staticBackdropLabel">Add Link</h5>
         <!-- <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button> -->
       </div>
        
       
       <div class="modal-body">
       <div class="row align-items-center my-3"><div class="col-4 titleflxengageman">Title</div><div class="col-1">:</div><div class="col-7"><input type="text" class="form-control rounded-0" id="TitleFlXengage" aria-describedby=""></div></div>
       <div class="row align-items-center my-3"><div class="col-4 titleflxengageman">Url</div><div class="col-1">:</div><div class="col-7"><input type="text" class="form-control rounded-0" id="URLFlXengage" value="" aria-describedby=""></div></div>
 
      <!-- <div class="row align-items-center my-3"><div class="col-4">OpeningNewTab</div><div class="col-1">:</div>
       <div class="col-7">
       <div class="form-check">
   <input class="form-check-input" type="checkbox" value="Yes" id="checkboxopentewtabFlXengage">
  
 </div></div></div>  -->
       
      
   <!--    <div class="row align-items-center my-3"><div class="col-4">Visible</div><div class="col-1">:</div><div class="col-7 custom-arrow">
       <div class="form-check">
   <input class="form-check-input" type="checkbox" value="Yes" id="checkboxvisibleFlXengage">
   
 </div></div></div> -->
 <div class="row align-items-center my-3"><div class="col-4">Url Properties</div><div class="col-1">:</div>
 <div class="col-7 ">
 <div class="btn-group option-checkboxes w-100" role="group" aria-label="Basic checkbox toggle button group">
 
 <!--<input type="checkbox" class="btn-check" id="checkboxvisibleFlXengage" autocomplete="off">
 <label class="btn btn-engage-outline-theme" for="checkboxvisibleFlXengage">Visible</label>       
 
 <input type="checkbox" class="btn-check" id="checkboxopentewtabFlXengage" autocomplete="off">
 <label class="btn btn-engage-outline-theme" for="checkboxopentewtabFlXengage">Open a new tab</label> -->
 
 <input type="checkbox" class="btn-check" id="checkboxvisibleFlXengage" autocomplete="off">
   <label class="btn btn-outline-theme" for="checkboxvisibleFlXengage">Visible</label>
 
   <input type="checkbox" class="btn-check" id="checkboxopentewtabFlXengage" autocomplete="off">
   <label class="btn btn-outline-theme" for="checkboxopentewtabFlXengage">Open a new tab</label>
 </div> 
 </div>
 </div> 
       <div class="row align-items-start my-3"><div class="col-4 titleflxengageman">Image</div><div class="col-1">:</div><div class="col-7"> 
       <input type="file" class="form-control-file custom-life-engage" class="mt-1" id="File1FlXengage"  accept="image/*">
       <div id="engagementAddFileEmpty"></div>
       </div></div>
     </div>
   
                    
       
     <div class="modal-footer modal-tile-footer justify-content-between"> 
     <div class="btns-left">
     <button type="button" class="btn btn-sm btn-danger rounded-0" id="" style="display: none;" data-bs-toggle="modal" data-bs-target="#deleteAlterModal">Delete</button>
     </div>
 
       <div class="btns-right d-flex">
       <div class="addScreen">
       <button type="button" class="btn btn-sm btn-secondary rounded-0" id="btnengagementAddClose" data-bs-dismiss="modal">Close</button>
       <button type="button" class="btn btn-sm btn-theme rounded-0" id="btnmodalSubmit" style="display: none;">Submit</button>
       </div>
       <div class="viewScreen">
       <!--<button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Close</button>-->
       <button type="button" class="btn btn-sm btn-theme ms-2 rounded-0" id="btnSubmitengageFLXengage" style="">Submit</button>
       </div>
       </div>
     </div>
     </div>
   </div>
 </div>  
 
 
 
 
 
 <!-- Delete Modal -->
 
     <div class="modal fade" id="EngagementDeleteModal" tabindex="-1" aria-labelledby="AnADeleteModalLabel" aria-hidden="true">
   <div class="modal-dialog engagement-delete-warning-dialog">
     <div class="modal-content rounded-0">
       <div class="modal-header">
          
         <!-- <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>-->
       </div>
       <div class="modal-body engagement-delete-warning text-center pt-5"> 
       <h5 class="modal-title" id="EngagementDeleteModallLabel">Confirmation</h5>
       <p class="mb-0">Are you sure want to Delete?</p>
       </div>
       <div class="modal-footer">
         <button type="button" id="cancelengagementDelete" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal">No</button>
         <button type="button" id="btnDeleteengage" class="btn btn-sm btn-danger rounded-0">Yes</button>
       </div>
     </div>
   </div>
 </div>
 <!-- Delete Modal -->

  
 <!---viewall popup -->

<div class="modal fade" id="exampleModalscrollengage" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
  <div class="modal-dialog   modal-dialog-scrollable"">  
    <div class="modal-content rounded-0">
      <div class="modal-header">      
        <h5 class="modal-title fw-bold w-100 text-center" id="exampleModalLabel">FLX Engagement</h5>
    <!--   <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>  -->
      </div>     
      <div class="modal-body viewallmodal"> 
      <div class="viewallanounce">
      <ul class="list-unstyled">      
                 <li class="d-flex align-items-center row mx-1 my-3">
                 <span class="addlearn col-1 addiconflxlearn"></span><a href="#" class="col-8 ">Test Data</a>
                 <span class="icon-edit-announce col-2 pencileditvis"></span>
                 </li>
                 <li class="d-flex align-items-center row mx-1 my-3">
                 <span class="addlearn col-1 addiconflxlearn"></span><a href="#" class="col-8 ">Test Data</a>
                 <span class="icon-edit-announce col-2 pencileditvis"></span>
                 </li>
                 <li class="d-flex align-items-center row mx-1 my-3">
                 <span class="addlearn col-1 addiconflxlearn"></span><a href="#" class="col-8 ">Test Data</a>
                 <span class="icon-edit-announce col-2 pencileditvis"></span>
                 </li>
                 <li class="d-flex align-items-center row mx-1 my-3">
                 <span class="addlearn col-1 addiconflxlearn"></span><a href="#" class="col-8 ">Test Data</a>
                 <span class="icon-edit-announce col-2 pencileditvis"></span>
                 </li>
                 <li class="d-flex align-items-center row mx-1 my-3">
                 <span class="addlearn col-1 addiconflxlearn"></span><a href="#" class="col-8 ">Test Data</a>
                 <span class="icon-edit-announce col-2 pencileditvis"></span>
                 </li>   
                 </ul> 
      
      </div>
      </div>
      <div class="modal-footer"> 
        <button type="button" class="btn btn-sm btn-secondary rounded-0" data-bs-dismiss="modal" id="btnclose">Close</button>
        <button type="button" class="btn btn-sm btn-theme rounded-0" id="btnsubmit">Submit</button> 
      </div>        
    </div>
  </div>
</div>

`;

getadminfromsite();
      $("#ShowVisible").hide();
      $("#ViewAll").show();
      $("#ViewAll").click(()=>{
        FetchFLXEngagementAll();
      });
      $("#ShowVisible").click(()=>{
        FetchFLXEngagement();
      });
    

$("#engagementDeleteModal").click(()=>{
  $(".engagement-modal-dialog").hide();
})
$("#cancelengagementDelete").click(()=>{
  $(".engagement-modal-dialog").show();
})
$(document).on('click','.editimageflxengage',function(e){
  e.currentTarget.getAttribute("data-id");
   itemid =e.currentTarget.getAttribute("data-id") ;
   GetFLXEngagement(); 
 
})
$("#btnUpdateengage").click(function(){
  if(mandatoryforUpdateFLXEngagement()) {
    UpdateFLXEngagement(itemid);   
  } else {
    console.log("All fileds not filled");
  }
    // UpdateFLXEngagement(itemid);
  })
  
  $("#btnDeleteengage").click(function(){
    DeleteFLXEngagement(itemid);
  })
  
  // $(document).on('click','#btnSubmitengageFLXengage',function(){
  //   AddFLXEngagement();
  // })
  //$("#btnSubmitengageFLXengage").click(()=>{AddFLXEngagement();})
  $("#btnSubmitengageFLXengage").click(function(){
    if (mandatoryforAddFLXEngagement()) {
      AddFLXEngagement();   
} else {
  console.log("All fileds not filled");
}
    // AddFLXEngagement();  
  })    

  $(document).on("change", "#File1FLXengageEdit", function () {
    var _URL = window.URL;
    var file, img;
    if ((file = this.files[0])) {
        img = new Image();
        img.onload = function () {
            //alert("Width:" + this.width + "   Height: " + this.height);
            widthedit=this.width;
            heightedit=this.height;
            };
            img.src = _URL.createObjectURL(file);
    }
    if($("#File1FLXengageEdit").prop('files').length > 0){
      $("#engagementEditFile").hide()
    }else{
      $("#engagementEditFile").show()
    }
  })

  $(document).on("change", "#File1FlXengage", function () {
    var _URL = window.URL;
    var file, img;
    if ((file = this.files[0])) {
        img = new Image();
        img.onload = function () {
            //alert("Width:" + this.width + "   Height: " + this.height);
            width=this.width;
            height=this.height;
            };
            img.src = _URL.createObjectURL(file);
    }
  })
  
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
async function getadminfromsite() {
  $(".loader-section").show();

  var bag=[];
  let listLocation  = await sp.web.getList(listUrl + "Badging").items.get(); 
  listLocation.forEach((li) => {
   bag.push(li.Days); 
   console.log(bag);
  });
  Badgingdays= bag[0];
  console.log(Badgingdays);
  var AdminInfo = [];
  await sp.web.siteGroups
    .getByName("FLX Admins")
    .users.get()
    .then(function (result) {
      for (var i = 0; i < result.length; i++) {
        AdminInfo.push({
          Title: result[i].Title,
          ID: result[i].Id,
          Email: result[i].Email,
        });
      }
      FilteredAdmin = AdminInfo.filter((admin)=>{return (admin.Email == currentuser)});
      console.log(FilteredAdmin);
      FetchFLXEngagement();
    })
    .catch(function (err) {
      alert("Group not found: " + err);
      $(".loader-section").hide();
    });
    $(".loader-section").hide();
}
function FetchFLXEngagement() {
  $(".loader-section").show();
  $("#ShowVisible").hide();
  $("#ViewAll").show();
  let list = sp.web.lists.getByTitle("FLXEngagement");
list.get().then(l => {
    console.log("List Id: " + l.Id);
    LGUID=l.Id;
}); 
  var html = "";
  
  sp.web.lists
    .getByTitle("FLXEngagement")
    .items.select("*","Title", "URL", "OpeningNewTab", "Visible", "Image").filter("Visible eq '1'").orderBy("Modified",false).get()
    .then((items: any[]) => {
      console.log(items);
      if (FilteredAdmin.length>0) 
      {
      for (var i = 0; i < items.length; i++) {
        const item = items[i];
        const itemImage = JSON.parse(item.Image) || {};
        const serverUrl = itemImage.serverUrl || "";
        const imageUrl = itemImage.serverRelativeUrl || "";
        
        if (item.OpeningNewTab === true) {
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
            <span class="editimageflxengage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id ="${item.ID}"></span></div>
            <a data-interception="off" href="${item.URL}" target="_blank"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a data-interception="off" class="" href="${item.URL}" target="_blank">
            <div class="q-link-title">${item.Title}</div></a></div></div>`
          // console.log(items)
        }    
        else { 
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
          <span class="editimageflxengage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id="${item.ID}"></span>
          </div>
            <a href="${item.URL}"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a class="" href="${item.URL}">
            <div class="q-link-title ">${item.Title}</div></a></div></div>`
        }
      } 

      if(items.length>=0){
        html+=`<div class="card text-center m-2 p-2 mt-3 flxengagecursor" style="width: 9rem; height:10.5rem ;border-radius:0">
        <div class="card-body my-4">
        <span class="engage-add-icon" data-bs-toggle="modal" data-bs-target="#staticBackdroptwo"></span>
        <p class="engage-title my-2">Add Link</p>
        </div>
      </div>`
      }   


      var element = document.getElementById("engageedit");
      element.innerHTML = html;
    }
    else{
      for (var i = 0; i < items.length; i++) {
        const item = items[i];
        const itemImage = JSON.parse(item.Image) || {};
        const serverUrl = itemImage.serverUrl || "";
        const imageUrl = itemImage.serverRelativeUrl || "";  
        
        if (item.OpeningNewTab === true) {
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
            <span class="editimage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id ="${item.ID}"></span></div>
            <a data-interception="off" href="${item.URL}" target="_blank"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a data-interception="off" class="" href="${item.URL}" target="_blank">
            <div class="q-link-title">${item.Title}</div></a></div></div>`
          // console.log(items)
        }    
        else { 
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
          <span class="editimage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id="${item.ID}"></span>
          </div>
            <a href="${item.URL}"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a class="" href="${item.URL}">
            <div class="q-link-title ">${item.Title}</div></a></div></div>`
        }
      } 
      var element = document.getElementById("engageedit");
      element.innerHTML = html;

      $("#ViewAll").hide();
      $("#ShowVisible").hide();
    }
    var count;
    for(var i=0;i<items.length;i++){
      count=i;
      var today = new Date();
      var startdate=new Date(items[i].Created);
          var sdate=new Date(items[i].Created);
          var Edate=sdate.setDate(sdate.getDate() + parseInt(Badgingdays));
    var enddate=new Date(Edate);
    var startdatemt=moment(startdate).format("YYYY-MM-DD");
    var enddatemt=moment(enddate).format("YYYY-MM-DD");
    var todaymt=moment(today).format("YYYY-MM-DD");
    
          if(todaymt >= startdatemt && todaymt < enddatemt || todaymt > startdatemt && todaymt <= enddatemt){
    
    $(".engslearn"+count).show();   
    }  
    else{    
      $(".engslearn"+count).hide(); 
   
    }
    }
   
    
    })
    $(".loader-section").hide();
} 

 function GetFLXEngagement(){
  
  $("#engagementEditFile").show();
  $(".loader-section").show();
 sp.web.lists.getByTitle("FLXEngagement").items.getById(parseInt(itemid)).get().then((items: any[]) => 
  {
    const itemImage = JSON.parse(items["Image"]) || {};
        const serverUrl = itemImage.serverUrl || "";
        const imageUrl = itemImage.serverRelativeUrl || "";
        console.log();
        SelectedImage = imageUrl.split("/").pop()
        
    $("#File1FLXengageEdit").val("")
    $("#TitleFLXEngagement").val(items['Title']);
    $("#URLFLXEngagement").val(items['URL']);
    $("#checkboxopenewtabFLXEngagement" ).prop("checked",items['OpeningNewTab']); 
    $("#checkboxvisibleFLXEngagement" ).prop( "checked",items['Visible']);
    $("#engagementEditFile").html(`<a href="#" onclick='window.open("${imageUrl}");return false;'>${imageUrl.split("/").pop()}</a>`);
    
  console.log(items);
  
 })
 $(".loader-section").hide();
}

function UpdateFLXEngagement(itemid){
  $(".loader-section").show();
  console.log(LGUID);

  if($('#File1FLXengageEdit').prop('files').length > 0){
    var Editfile =$('#File1FLXengageEdit').prop('files')[0];
    var ht=parseInt(heightedit),wt=parseInt(widthedit);
    if(ht > 500 || wt > 500)
    {
      $("#engagementUpdateFileEmpty").html(`<p class="text-danger m-0">Height and Width must not exceed 500px</p>`)
    }
else{
    sp.web.getFolderByServerRelativeUrl(`/sites/${SiteName}/SiteAssets/Lists/${LGUID}`).files
  .add(Editfile.name, Editfile, true).then((fileItem)=>{
    sp.web.lists.getByTitle("FLXEngagement").items.getById(parseInt(itemid)).update({
      Title: $("#TitleFLXEngagement").val(),
           URL: $("#URLFLXEngagement").val(),
           OpeningNewTab: $("#checkboxopenewtabFLXEngagement").is(':checked') ? true : false,
           Visible: $("#checkboxvisibleFLXEngagement").is(':checked') ? true : false,
           Image: JSON.stringify({
            "serverRelativeUrl": fileItem.data.ServerRelativeUrl
          })
    }).then((i)=>{
      $("#btnengagementEditClose").trigger('click');
      AlertMessage("<div class='alertfy-success'>Record updated successfully</div>");
    })
  })
}
  }else{
    sp.web.lists.getByTitle("FLXEngagement").items.getById(parseInt(itemid)).update({
      Title: $("#TitleFLXEngagement").val(),
       URL: $("#URLFLXEngagement").val(),
       OpeningNewTab: $("#checkboxopenewtabFLXEngagement").is(':checked') ? true : false,
       Visible: $("#checkboxvisibleFLXEngagement").is(':checked') ? true : false,
       
       
      }).then(i => {
        $("#btnengagementEditClose").trigger('click');
        AlertMessage("<div class='alertfy-success'>Record updated successfully</div>");
      });
  }
  $(".loader-section").hide();
}  
function AddFLXEngagement() {
  $(".loader-section").show();
console.log(LGUID);


if($('#File1FlXengage').prop('files').length == 0){
$("#engagementAddFileEmpty").html(`<p class="text-danger m-0">Please Choose a File</p>`)
}
else if($('#File1FlXengage').prop('files').length > 0)
{
  var ht=parseInt(height),wt=parseInt(width);
  if(ht > 500 || wt > 500)
  {
  $("#engagementAddFileEmpty").html(`<p class="text-danger m-0">Height and Width must not exceed 500px</p>`)
  }

else {
//uploadfile
var file =$('#File1FlXengage').prop('files')[0];  
sp.web.getFolderByServerRelativeUrl(`/sites/${SiteName}/SiteAssets/Lists/${LGUID}`).files
 .add(file.name, file, true)
 .then((fileItem) => { 
   console.log(fileItem);  
   sp.web.lists.getByTitle("FLXEngagement").items.add({
   Title: $("#TitleFlXengage").val(),
   URL: $("#URLFlXengage").val(),
   OpeningNewTab: $("#checkboxopentewtabFlXengage").is(':checked') ? true : false,
   Visible: $("#checkboxvisibleFlXengage").is(':checked') ? true : false,
   Image: JSON.stringify({
     "serverRelativeUrl": fileItem.data.ServerRelativeUrl
   })
 }).then(i => {
   $("#btnengagementAddClose").trigger('click');
   AlertMessage("<div class='alertfy-success'>Record created successfully</div>");
 });
});
}
}
$(".loader-section").hide();
}

function DeleteFLXEngagement(itemid){
  $(".loader-section").show();
  sp.web.lists.getByTitle("FLXEngagement").items.getById((parseInt(itemid))).delete().then(()=>{location.reload()}).catch((error)=>{alert("Error Occured");})
  // AlertMessage("Record Deleted successfully");
  $(".loader-section").hide();
}
 
 
function AlertMessage(Message) {
  alertify
    .alert()
    .setting({
      label: "OK",

      message: Message,

      onok: function () {
        window.location.href = "#";
        $(".loader-section").hide();
        location.reload();

        //window.location.href = "#";
      },
    })
    .show()
    .setHeader("<div class='fw-bold alertifyConfirmation'>Confirmation</div>")
    .set("closable", false);
}  


function mandatoryforAddFLXEngagement(){
  var isAllvalueFilled = true;
  if (!$("#TitleFlXengage").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#URLFlXengage").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  else if (!$("#File1FlXengage").val()) {
    alertify.error("Please upload file");
    isAllvalueFilled = false;  
  }   
  return isAllvalueFilled;
}

function mandatoryforUpdateFLXEngagement(){
  var isAllvalueFilled = true;
  if (!$("#TitleFLXEngagement").val()) {
    alertify.error("Please Enter the Title");
    isAllvalueFilled = false;
  } else if (!$("#URLFLXEngagement").val()) {
    alertify.error("Please Enter the url ");
    isAllvalueFilled = false;
  }
  // else if (!$("#File1FLXengageEdit").val()) {
  //   alertify.error("Please upload file");
  //   isAllvalueFilled = false;  
  // }     
  return isAllvalueFilled;
}

function FetchFLXEngagementAll() {
  $(".loader-section").show();
  $("#ShowVisible").show();
  $("#ViewAll").hide();
  let list = sp.web.lists.getByTitle("FLXEngagement");
list.get().then(l => {
    console.log("List Id: " + l.Id);
    LGUID=l.Id;
}); 
  var html = "";
  
  sp.web.lists
    .getByTitle("FLXEngagement")
    .items.select("*","Title", "URL", "OpeningNewTab", "Visible", "Image").orderBy("Modified",false).get()
    .then((items: any[]) => {
      console.log(items);
      
      for (var i = 0; i < items.length; i++) {
        const item = items[i];   
        const itemImage = JSON.parse(item.Image) || {};
        const serverUrl = itemImage.serverUrl || "";
        const imageUrl = itemImage.serverRelativeUrl || "";
        
        if (item.OpeningNewTab === true) {
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
            <span class="editimageflxengage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id ="${item.ID}"></span></div>
            <a data-interception="off" href="${item.URL}" target="_blank"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a data-interception="off" class="" href="${item.URL}" target="_blank">
            <div class="q-link-title">${item.Title}</div></a></div></div>`
          // console.log(items)
        }    
        else {   
          html += `<div class="englrn"><span class="englearn engslearn${i}">New</span><div class = "q-link m-3 text-center p-2"><div class="iconaddengage text-end py-1 px-2">
          <span class="editimageflxengage" data-bs-toggle="modal" data-bs-target="#staticBackdropone" data-id="${item.ID}"></span>
          </div>
            <a href="${item.URL}"><img class="q-link-img" src="${serverUrl}${imageUrl}" alt="img"/></a><a class="" href="${item.URL}">
            <div class="q-link-title ">${item.Title}</div></a></div></div>`
        }
      } 

      if(items.length>=0){
        html+=`<div class="card text-center m-2 p-2 mt-3 flxengagecursor" style="width: 9rem; height:10.5rem ;border-radius:0">
        <div class="card-body my-4">
        <span class="engage-add-icon" data-bs-toggle="modal" data-bs-target="#staticBackdroptwo"></span>
        <p class="engage-title my-2">Add Link</p>
        </div>
      </div>`
      }   


      var element = document.getElementById("engageedit");
      element.innerHTML = html;
      var count;
      for(var i=0;i<items.length;i++){
        count=i;
        var today = new Date();
        var startdate=new Date(items[i].Created);
            var sdate=new Date(items[i].Created);
            var Edate=sdate.setDate(sdate.getDate() + parseInt(Badgingdays));
      var enddate=new Date(Edate);
      var startdatemt=moment(startdate).format("YYYY-MM-DD");
      var enddatemt=moment(enddate).format("YYYY-MM-DD");
      var todaymt=moment(today).format("YYYY-MM-DD");
      
            if(todaymt >= startdatemt && todaymt < enddatemt || todaymt > startdatemt && todaymt <= enddatemt){
      
      $(".engslearn"+count).show();   
      }  
      else{    
        $(".engslearn"+count).hide(); 
     
      }
      }
     

    })
    $(".loader-section").hide();
} 

