const SpecialFields="System.AuthorizedDate;System.RevisedDate;System.ChangedDate;System.Rev;System.ChangedBy;System.Watermark;System.AuthorizedAs;System.PersonId"

var authHeader,hostName,projectName;
const contributors=new Map(); //hashmap to store all users that have updates


//to dynamically access properties of an object, used in HTML Template.
const getDescendantProp = (obj, path) => (
    path.split('.').reduce((acc, part) => acc && acc[part], obj)
);

//Create timestamp from date so we can sort on integers instead of strings.
function toTimestamp(strDate){
    var datum = Date.parse(strDate);
    return datum/1000;
}

//convert the json arrays returned from REST Api, into a "more flat array around updates" so it can be sorted by date
function FlattenArr(inArray){
    var outArray=[];

    try{
        for(var i=0;i<inArray.length;i++)
        {
            if(inArray[i].count>0)
            {
                for(var j=0;j<inArray[i].count;j++)
                {
                    outArray.push(inArray[i].value[j]);
                }    
            }
    
        }    
    }
    catch(Err){
        console.error("Error in FlattenCommentsArr: "+ Err);
    }
    return outArray;
}

function addContributor(contributor){

    if (typeof contributor.uniqueName !== 'undefined')
    {
        if (contributors.has(contributor.uniqueName))
        {contributors.get(contributor.uniqueName).contributions++;}
        else
        { 
            try{
                if (contributor.uniqueName.indexOf("VisualStudio")==-1)
                {
                    contributors.set(contributor.uniqueName,{uniqueName:contributor.uniqueName, image:contributor.imageUrl,contributions:1,name:contributor.displayName});
                }
            }catch(err){
                console.log("addContributor: Error : "+err);
            }
        }
    }
}

//Copy information from idsArr to the updates arrays, so each item has "complete info" 
function copyProperties(idsArr,updates)
{
    updates.forEach(function(item) {
        
        var idItem=idsArr.find(x => x.id === item.workItemId);
        if (idItem==null) 
            {console.log("Error id not found:" + item.id);}
        else {
            Object.defineProperty(item, 'title', { value: idItem.fields["System.Title"] } );  
            Object.defineProperty(item, 'workItemURL', { value: getWorkItemUrl(hostName,projectName,item.workItemId) } );  
            Object.defineProperty(item, 'numRevisions', { value: idItem.rev } );  
            
            if(item.hasOwnProperty("fields")){
                Object.defineProperty(item, 'timestamp', { value: toTimestamp(item.fields["System.ChangedDate"].newValue) } ); 
                Object.defineProperty(item, 'datePassedDesc', { value: GetDateDiffDescriptionVsNow(item.fields["System.ChangedDate"].newValue) } );
            }
            else {
                Object.defineProperty(item, 'timestamp', { value: toTimestamp(item.revisedDate) } ); 
                Object.defineProperty(item, 'datePassedDesc', { value: GetDateDiffDescriptionVsNow(item.revisedDate) } );
            }
            var fieldsChangedHTML=createfieldsChangedHTML(item);

            Object.defineProperty(item, 'fieldsChangedHTML', { value: fieldsChangedHTML.HTML } ); 
            Object.defineProperty(item, 'countNormalFields', { value: fieldsChangedHTML.countNormalFields } ); 
            Object.defineProperty(item, 'countSpecialFields', { value: fieldsChangedHTML.countSpecialFields } ); 

            if (fieldsChangedHTML.countNormalFields>0)
            {addContributor(item.revisedBy);    }
        }


    });
    
}

//Convert a hashmap to an array since toArray() function was not working
function ConvertMaptoArray(hashmap){
    var newArray=[];
    hashmap.forEach(function(item){
        newArray.push(item);
    });
    return newArray;
}
//https://stackoverflow.com/questions/208016/how-to-list-the-properties-of-a-javascript-object
//List all properties of an object
var getKeys = function(obj){
    var keys = [];
    for(var key in obj){
       keys.push(key);
    }
    return keys;
 }

 
function getObjectFields(field,item)
{
    if (field=="System.ChangedBy" || field=="System.CreatedBy" || field=="System.AuthorizedAs" || field=="System.AssignedTo")
    {
        return item["displayName"];
    }
    else{
        var html=[];
        var keys = [];
        html.push("<span>");
        for(var key in item){
           html.push("<br>"+key+"="+item[key]);
        }
        html.push("</span>");
        return html.join("");
    }

}
//Create the HTML that summarizes the changes (in fields) made to a workitem
function createfieldsChangedHTML(item)
{   
    var html=[];
    var fields=item.fields;
    var countNormalFields=0;
    var countSpecialFields=0;
    html.push("<div class=\"divTable\"><div class=\"divTableBody\">");
    html.push("<div class=\"divTableHeading\"><div class=\"divTableCell\">Field</div><div class=\"divTableCell\">Old value</div><div class=\"divTableCell\">New value</div></div>");
    var fields=getKeys(item.fields);
    fields.forEach(function(field)
    {
            if (SpecialFields.includes(field))
            {
                html.push("<div class=\"divTableRow showHideSpecialField\">");
                countSpecialFields++;
            }
            else 
            {
                html.push("<div class=\"divTableRow\">");
                countNormalFields++;
            }
            
            html.push("<div class=\"divTableCell\">"+field+"</div>");
            if(item.fields[field].hasOwnProperty("oldValue"))
                {
                    if (typeof item.fields[field].oldValue === 'object')
                    {
                        html.push("<div class=\"divTableCell\"><del>"+getObjectFields(field,item.fields[field].oldValue)+"</del></div>");
                    }else {
                        html.push("<div class=\"divTableCell\"><del>"+item.fields[field].oldValue+"</del></div>");
                    }
                }
            else {html.push("<div class=\"divTableCell\">&nbsp;</div>");}
            
            if(item.fields[field].hasOwnProperty("newValue"))
            {
                if (typeof item.fields[field].newValue === 'object')
                {
                    html.push("<div class=\"divTableCell\">"+getObjectFields(field,item.fields[field].newValue)+"</div>");
                }else {
                    html.push("<div class=\"divTableCell\">"+item.fields[field].newValue+"</div>");
                }
                
            }
            else {html.push("<div class=\"divTableCell\">&nbsp;</div>");}
            html.push("</div>");
     });
    html.push("</div></div>");

    var retObj={
        countNormalFields:countNormalFields,
        countSpecialFields:countSpecialFields,
        HTML: html.join("\n")
    };

    return retObj;
}


//Function to compare timestamps, to be used in sorting an array
function compareTimestamp( a, b ) {
    if ( a.timestamp < b.timestamp ){
      return 1;
    }
    if ( a.timestamp > b.timestamp ){
      return -1;
    }
    return 0;
  }
  function compareContributions( a, b ) {
    if ( a.contributions < b.contributions ){
      return 1;
    }
    if ( a.contributions > b.contributions ){
      return -1;
    }
    return 0;
  }

  function removeItemsBeforeDate(itemArr,_dateFilter){
    var retArr=[];
    if (_dateFilter==null){return itemArr;}

    _dateFilterTimeStamp=toTimestamp(_dateFilter);
    try{

        for(var i=0; i<itemArr.length;i++){
            element=itemArr[i];
            if(!element.hasOwnProperty("timestamp"))
            {
                if(element.hasOwnProperty("fields")&& element.fields["System.ChangedDate"]!==undefined){
                    Object.defineProperty(element, 'timestamp', { value: toTimestamp(element.fields["System.ChangedDate"].newValue) } ); 
                }
                else {
                    Object.defineProperty(element, 'timestamp', { value: toTimestamp(element.revisedDate) } ); 

                }
            }

            if(element.timestamp>_dateFilterTimeStamp)
            {retArr.push(element);}            

        }

        // itemArr.forEach(element => {
        //     if(!element.hasOwnProperty("timestamp"))
        //     {
        //         if(element.hasOwnProperty("fields")&& element.fields["System.ChangedDate"]!==undefined){
        //             Object.defineProperty(element, 'timestamp', { value: toTimestamp(element.fields["System.ChangedDate"].newValue) } ); 
        //         }
        //         else {
        //             Object.defineProperty(element, 'timestamp', { value: toTimestamp(element.revisedDate) } ); 

        //         }
        //     }

        //     if(element.timestamp>_dateFilterTimeStamp)
        //     {retArr.push(element);}            

        // });
    }
    catch(err){console.log("err "+ err);}
    return retArr;
  }

  function allProgress(proms, progress_cb) {
    let d = 0;
    progress_cb(0);
    for (const p of proms) {
      p.then(()=> {    
        d ++;
        progress_cb( (d * 100) / proms.length );
      });
    }
    return Promise.all(proms);
  }

  function updateProgress(progValue)
  {
    //console.log("updateProgress: "+progValue);
    const myBarElem = document.getElementById("myBar"); 
    myBarElem.innerHTML="Loading "+progValue + '%';
    myBarElem.style.width = progValue + '%'; 
    if (progValue>99) 
    {
        document.getElementById("myProgress").style.display="none";
    }
  }
//Call Rest API's based on array of ID's 
function fetchContent(_idsArr,_authHeader,_hostName,_projectName,_dateFilter)
{
    return new Promise(function(resolve, reject) {

        authHeader=_authHeader;
        hostName=_hostName;
        projectName=_projectName;

        try
        {

            console.log("=========fetchContent");

            //Fetch all updates from workitems with ids in the _idsArray
            var promiseArr=new Array(_idsArr.length);
            for (var i=0;i<_idsArr.length;i++){
                promiseArr[i]=get(getUpdateRestApiUrl(hostName,projectName,_idsArr[i].id));
            }

            
            allProgress(promiseArr,
                (p) => {
                    //console.log(`% Done = ${p.toFixed(2)}`);
                    updateProgress(50+Math.floor(p.toFixed(0)/2));
     
            })
            //Promise.all(promiseArr)
            .then(function(values) {
                var updates=FlattenArr(values);
                console.log("Got updates:"+updates.length);
                updates=removeItemsBeforeDate(updates,_dateFilter);
                copyProperties(_idsArr,updates);
                updates.sort(compareTimestamp);
                resolve(updates);
                
            }).catch(error => 
                reject(error));
        }
        catch(Err){
            console.error("Error:"+Err);
            reject(Err);
        }
    });
}


//Perform a XMLHttpRequest authenticated json call to an URL (Using VSS token)
function get(url) {
    return new Promise(function(resolve, reject) {
      // Do the usual XHR stuff
      var req = new XMLHttpRequest();
      
      req.open('GET', url);
      req.setRequestHeader( "Authorization", authHeader );
      req.responseType = 'json'
      req.onload = function() {
        // This is called even on 404 etc
        // so check the status
        if (req.status == 200) {
          // Resolve the promise with the response text
          resolve(req.response);
        }
        else {
          // Otherwise reject with the status text
          reject(Error(req.statusText));
        }
      };
  
      // Handle network errors
      req.onerror = function() {
        reject(Error("Network Error"));
      };
  
      // Make the request
      req.send();
    });
}


function escapeUnicode(str) {
    return str.replace(/[^\0-~]/g, function(ch) {
        return "\\u" + ("0000" + ch.charCodeAt().toString(16)).slice(-4);
    });
}

//Create the URL for VSS Work Item Updates REST API
function getUpdateRestApiUrl(organization, project, id)
{
    //https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/updates/list
    return "https://"+organization+".visualstudio.com/"+project+"/_apis/wit/workItems/"+id+"/updates/?api-version=5.1";
}

//Create the URL for VSS WorkItem Comments REST API
// function getCommentsRestApiUrl(organization, project, id)
// {
//     return "https://"+organization+".visualstudio.com/"+project+"/_apis/wit/workItems/"+id+"/comments";
// }

//Create the URL for a specific work item within an organization and project
function getWorkItemUrl(organization, project, id)
{
    //https://dev.azure.com/TASAlpineSkiHouse/MyFirstProject/_workitems/edit/1/
    return "https://"+organization+".visualstudio.com/"+project+"/_workItems/edit/"+id;
}


//Return a written for of how long it passed sice a given date, if too long return just the date.
function GetDateDiffDescriptionVsNow(_date){
    try
    {
    var theDate=new Date(_date);
    var now=new Date();
    var dateDiff=DateDiffHours(theDate,now);
    dateDiff=dateDiff*-1;
    switch(true){
        
        case (dateDiff<1):
            return "less than one hour ago";
            break;
        case (dateDiff>=1 &&  dateDiff < 24):
            return Math.floor(dateDiff)+" hours ago";
            break;
        case (dateDiff>=24 &&  dateDiff < 48):
            return "yesterday";
            break;
        case (dateDiff>=48 &&  dateDiff < 128):
                return Math.floor(dateDiff/24)+" days ago";
                break;
        
         default:
            return "on "+theDate.toLocaleDateString();
            
            break;

    }
    }catch(Err){
        console.log("Eorror in GetDateDiffDescriptionVsNow"+Err);
    }
}
function DateDiffHours(date1, date2) {
    var datediff = date1.getTime() - date2.getTime(); //store the getTime diff - or +
    return (datediff / (60*60*1000));    
}