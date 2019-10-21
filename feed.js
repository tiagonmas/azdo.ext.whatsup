
var authHeader,hostName,projectName;

//to dynamically access properties of an object
const getDescendantProp = (obj, path) => (
    path.split('.').reduce((acc, part) => acc && acc[part], obj)
);

function toTimestamp(strDate){
    var datum = Date.parse(strDate);
    return datum/1000;
}
//convert the json arrays returned from REST Api, into a "flat array" so can be sorted by date
function FlattenCommentsArrAndSort(inArray){
    var commentsTotal=0;
    var outArray=[];

    try{
        for(var i=0;i<inArray.length;i++)
        {
            if(inArray[i].comments.length>0)
            {
                for(var j=0;j<inArray[i].comments.length;j++)
                {
                    outArray.push(inArray[i].comments[j]);
                    // var jsonData={createdDate:toTimestamp(inArray[i].comments[j].createdDate),comment:inArray[i].comments[j]};
                    // //var jsonData={createdDate:i*j+i-j,comment:inArray[i].comments[j]};
                    // outArray.push(jsonData);
                }    
            }
    
        }

        //Sort by date (converting to timestamp before)
        outArray.sort(function(a, b) {
            return toTimestamp(parseFloat(a.createdDate)) - toTimestamp(parseFloat(b.createdDate));
        });
    
    }
    catch(Err){
        console.error("Error in FlattenCommentsArr: "+ Err);
    }
    return outArray;
}

function fetchContent(_idsArr,_authHeader,_hostName,_projectName)
{
    return new Promise(function(resolve, reject) {

        authHeader=_authHeader;
        hostName=_hostName;
        projectName=_projectName;

        try
        {

            console.log("=========fetchContent");

            //Fetch all comments from workitems with ids in the _idsArray
            var promiseArr=new Array(_idsArr.length);
            for (var i=0;i<_idsArr.length;i++){
                promiseArr[i]=get(getCommentsRestApiUrl(hostName,projectName,_idsArr[i]));
            }
            Promise.all(promiseArr).then(function(values) {
                
                console.log("Promise.all done");
                var comments=FlattenCommentsArrAndSort(values);
                console.log("Got comments:"+comments.length);
                resolve(comments);
            }).catch(error => reject(error));
                
            // get(getCommentsRestApiUrl(hostName,projectName,1)).then(function(result){
            //     console.log("I got getCommentsRestApiUrl:"+result);
            //     console.log("getCommentsRestApiUrl #items:"+result.count);
            // }, function (err) 
            //     {console.log("Err1:"+err);}
            // ).then(
            //     get(getUpdateRestApiUrl(hostName,projectName,1)).then(function(result){
            //         console.log("I got getUpdateRestApiUrl:"+result);
            //         console.log("getUpdateRestApiUrl #items:"+result.count);
            //     },function (err) 
            //     {console.log("Err2:"+err);})  
            // )

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

function getCommentsHTML(comments){
    html=[];
    comments.forEach(function(item,index){
        html.push("<br>"+item.createdDate+" "+item.comment.text);
    });
    return html.join("");;
}