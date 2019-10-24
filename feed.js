var authHeader,hostName,projectName;
const SpecialFields="System.AuthorizedDate;System.RevisedDate;System.ChangedDate;System.Rev;System.ChangedBy;System.Watermark"

//to dynamically access properties of an object, used in HTML Template.
const getDescendantProp = (obj, path) => (
    path.split('.').reduce((acc, part) => acc && acc[part], obj)
);


function toTimestamp(strDate){
    var datum = Date.parse(strDate);
    return datum/1000;
}
//convert the json arrays returned from REST Api, into a "more flat array" so can be sorted by date
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
                    // var jsonData={createdDate:toTimestamp(inArray[i].comments[j].createdDate),comment:inArray[i].comments[j]};
                    // //var jsonData={createdDate:i*j+i-j,comment:inArray[i].comments[j]};
                    // outArray.push(jsonData);
                }    
            }
    
        }    
    }
    catch(Err){
        console.error("Error in FlattenCommentsArr: "+ Err);
    }
    return outArray;
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
            }
            else {
                Object.defineProperty(item, 'timestamp', { value: toTimestamp(item.revisedDate) } ); 
            }
            Object.defineProperty(item, 'fieldsChangedHTML', { value: createfieldsChangedHTML(item) } ); 
        }
    });
    
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

//Create the HTML that summarizes the changes made to a workitem
function createfieldsChangedHTML(item)
{   
    var html=[];
    var fields=item.fields;
    html.push("<div class=\"divTable\"><div class=\"divTableBody\">");
    html.push("<div class=\"divTableHeading\"><div class=\"divTableCell\">Field</div><div class=\"divTableCell\">Old value</div><div class=\"divTableCell\">New value</div></div>");
    var fields=getKeys(item.fields);
    fields.forEach(function(field)
    {
            if (SpecialFields.includes(field))
            {
                html.push("<div class=\"divTableRow showHideSpecialField\">")   
            }
            else 
            {
                html.push("<div class=\"divTableRow\">")    
            }
            
            html.push("<div class=\"divTableCell\">"+field+"</div>");
            if(item.fields[field].hasOwnProperty("oldValue"))
                {html.push("<div class=\"divTableCell\"><del>"+item.fields[field].oldValue+"</del></div>");}
            else {html.push("<div class=\"divTableCell\">&nbsp;</div>");}
            
            if(item.fields[field].hasOwnProperty("newValue"))
                {html.push("<div class=\"divTableCell\">"+item.fields[field].newValue+"</div>");}
            else {html.push("<div class=\"divTableCell\">&nbsp;</div>");}
            html.push("</div>");
     });
    html.push("</div></div>");
    return html.join("\n");
}

function compare( a, b ) {
    if ( a.timestamp < b.timestamp ){
      return 1;
    }
    if ( a.timestamp > b.timestamp ){
      return -1;
    }
    return 0;
  }

//Call Rest API's based on array of ID's 
function fetchContent(_idsArr,_authHeader,_hostName,_projectName)
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
            Promise.all(promiseArr).then(function(values) {
                
                console.log("Promise.all done");
                var updates=FlattenArr(values);
                console.log("Got updates:"+updates.length);
                copyProperties(_idsArr,updates);
                //Sort by date (timestamp)
                //updates.sort((a,b) => (a.timestamp > b.timestamp) ? 0 : ((b.timestamp > a.timestamp) ? -1 : 1));
                updates.sort(compare);
                resolve(updates);
            }).catch(error => reject(error));
        }
        catch(Err){
            console.error("Error:"+Err);
            reject(Err);
        }
    });
}


function get(url) {

        // Return a new promise.
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
          // which will hopefully be a meaningful error
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

function getUpdateRestApiUrl(organization, project, id)
{
    return "https://"+organization+".visualstudio.com/"+project+"/_apis/wit/workItems/"+id+"/updates/?api-version=5.1";
}

function getCommentsRestApiUrl(organization, project, id)
{
    return "https://"+organization+".visualstudio.com/"+project+"/_apis/wit/workItems/"+id+"/comments";
}

function getWorkItemUrl(organization, project, id)
{
    //https://dev.azure.com/TASAlpineSkiHouse/MyFirstProject/_workitems/edit/1/
    return "https://"+organization+".visualstudio.com/"+project+"/_workItems/edit/"+id;
}
function getCommentsHTML(comments){
    html=[];
    comments.forEach(function(item,index){
        html.push("<br>"+item.createdDate+" "+item.comment.text);
    });
    return html.join("");;
}