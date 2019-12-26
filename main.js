var LastExecDate //Last time the extension was executed. Saved in settings.
var filterFromDate=new Date() //Date when we should start showing events from 
var filteredUser="";
var extVersion;
var sourceWits;
var queryId;
appInsights.startTrackPage("Page");


VSS.init({
    explicitNotifyLoaded: true,
    usePlatformScripts: true, 
    usePlatformStyles: true
});

VSS.ready(function() {
    

});

function SaveSetting(setting,value){
    return new Promise(function(resolve, reject) {
        try
        {
            VSS.getService(VSS.ServiceIds.ExtensionData).then(function(dataService) {
            // Set value in user scope
            dataService.setValue(setting, value, {scopeType: "User"}).then(function(value) {
                    resolve();
                });
            });
        }
        catch(Err){reject(Err);}
    });
}
function GetSetting(setting){
    return new Promise(function(resolve, reject) {
        try
        {
            // Get data service
            VSS.getService(VSS.ServiceIds.ExtensionData).then(function(dataService) {
                    // Get value in user scope
                    dataService.getValue(setting, {scopeType: "User"}).then(function(value) {
                        resolve(value);
                    });
                });
        }
        catch(Err){reject(Err);}
    });

}
VSS.require(["VSS/Service", "TFS/WorkItemTracking/RestClient","VSS/Authentication/Services"], function (VSS_Service, TFS_Wit_WebApi,VSS_Auth_Service) {
    var context = VSS.getWebContext();
    extVersion=VSS.getExtensionContext().version;
    var projectId = context.project.id;
    var projectName = context.project.name;
    var HostName = context.host.name;
    var HostUri = context.host.uri;

    $("#version").html(extVersion);

    VSS.getAccessToken().then(function(token){
            return VSS_Auth_Service.authTokenManager.getAuthorizationHeader(token);
        }).then(function(authHeader){					
            
            try
            {
                var witClient = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
                
                loadSettings().then(function(){

                var query;
                switch(dateFilter.value){
                    case 'all':
                        query = {query: "SELECT [System.Id] FROM workitems WHERE [System.Id] In (@Follows) AND [System.State] NOT IN ('Closed','Inactive','Completed') ORDER BY [System.ChangedDate] DESC" };
                        filterFromDate=null;
                        break;
                    case 'one':
                        query = {query: "SELECT [System.Id] FROM workitems WHERE [System.Id] In (@Follows) AND [System.State] NOT IN ('Closed','Inactive','Completed') AND [System.ChangedDate]>@today-1 ORDER BY [System.ChangedDate] DESC" };
                        filterFromDate.setDate(filterFromDate.getDate()-1);
                        break;
                    case 'seven':
                        query = {query: "SELECT [System.Id] FROM workitems WHERE [System.Id] In (@Follows) AND [System.State] NOT IN ('Closed','Inactive','Completed') AND [System.ChangedDate]>@today-7 ORDER BY [System.ChangedDate] DESC" };
                        filterFromDate.setDate(filterFromDate.getDate()-7);
                        break;
                    case 'month':
                            query = {query: "SELECT [System.Id] FROM workitems WHERE [System.Id] In (@Follows) AND [System.State] NOT IN ('Closed','Inactive','Completed') AND [System.ChangedDate]>@today-31 ORDER BY [System.ChangedDate] DESC" };
                            filterFromDate.setDate(filterFromDate.getDate()-31);
                            break;
        
                }
                
                var queryPromise;

                switch(sourceWits){
                    case "source_following":
                        queryPromise=witClient.queryByWiql(query, projectId);
                        console.log("DateFilter combo is "+dateFilter.value+" and query="+query.query);
                        break;
                    case "source_query":
                        queryPromise=witClient.queryById(queryId,projectId);
                        console.log("queryID="+queryId);
                        break;
                    default:
                        console.log("ERROR Invalid sourceWits");
                }

                
                queryPromise.then(
                    function(queryResult) {  
                        var idsArr;

                        if(queryResult.queryResultType==1){
                            //https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/Wiql/Query%20By%20Wiql?view=azure-devops-rest-5.1#queryresulttype
                            idsArr=new Array(queryResult.workItems.length);
                        }
                        else{
                            idsArr=new Array(queryResult.workItemRelations.length);
                        }

                        if (idsArr.length==0)
                        {
                            appInsights.trackEvent({name: "noContent"});
                            document.getElementById("nocontent").style.visibility="visible" ;
                            document.getElementById("headbox").style.visibility="visible" ;
                            
                            //Improve: break out of promise chain. No need to continue moving forward 
                        }
                        else {
                            appInsights.trackEvent({name: "Content"});
                            appInsights.trackMetric("FollowingItems",idsArr.length );
                            document.getElementById("nocontent").style.display="none" ;
                            document.getElementById("headbox").style.visibility="visible" ;
                        }


                        for (var i=0;i<idsArr.length;i++){
                            if(queryResult.queryResultType==1){
                             
                                idsArr[i]=queryResult.workItems[i].id
                            }
                            else{
                                idsArr[i]=queryResult.workItemRelations[i].target.id;
                            }
                            

                        }								

                        let removeDups = (ids) => ids.filter((v,i) => ids.indexOf(v) === i)
                        removeDups(idsArr);

                        return idsArr;
                    },function(Err){
                        console.error("error loading query "+Err);
                        showError("error loading query"+Err);
                        return null;
                    }
                    
                    ).then(function(idsArr){

                        if (idsArr!=null && idsArr.length>0)
                        {return witClient.getWorkItems(idsArr, ["System.Title"]);}
                        else {return [];}
                    }).then(function(itemsArr){
                        if (itemsArr.length>0)
                        {fetchContent(itemsArr,authHeader,HostName,projectName,filterFromDate).
                                then(function(updates){
                                    
                                    //Update UI with comments applying template
                                    var commentTpl = $('script[data-template="commentTemplate"]').text().split(/\$\{(.+?)\}/g);
                                    var fieldsTpl = $('script[data-template="fieldsTemplate"]').text().split(/\$\{(.+?)\}/g);
                                    var contributorsTpl = $('script[data-template="contributorsTemplate"]').text().split(/\$\{(.+?)\}/g);


                                    try{
                                        var contriArray=ConvertMaptoArray(contributors);
                                        $('#contributors').append(contriArray.map(function (item) {
                                            return contributorsTpl.map(render(item)).join('');
                                        }));           
                                        contriArray.sort(compareContributions);                         

                                    }catch(Err){
                                        console.error("Err:"+Err);
                                        showError("Err:"+e);
                                    }

                                    $('#list-comment-items').append(updates.map(function (item) {
                                        var myItemhtml;
                                        if (item.hasOwnProperty("fields") && item.countNormalFields>0){
                                            var myfields=item.fields;
                                            if(myfields.hasOwnProperty("System.History")){
                                                myItemhtml=commentTpl.map(render(item)).join('');
                                            }else {
                                                myItemhtml=fieldsTpl.map(render(item)).join('');
                                            }
                                        }
                                        return myItemhtml;
                                    }));
                                
                                    updateVisibility(document.getElementById("filterSelection").value);

                                    appInsights.stopTrackPage("Page");
                                },function(err) {
                                    console.error("========ERROR: "+err);
                                    showError("ERROR:"+e);
                                    appInsights.trackException(err, "ErrFinal");
                                });

                        }
                        
                    }).catch(function(e) { 
                        showError("err23"+e);
                        console.error('err23',e) 
                        
                    });
                }).catch(function(e) { 
                    showError("err24"+e);
                    console.error('err24',e) 
                });
                VSS.notifyLoadSucceeded();
            }catch(err)
            {
                console.error("error queryByWiql:"+err);
                showError("error queryByWiql"+err);
                appInsights.trackException(err, "ErrqueryByWiql");
                VSS.notifyLoadSucceeded();
            }
        });
                
});

function showError(msg){
    document.getElementById("headbox").style.visibility="visible";
    document.getElementById("loader").style.visibility="hidden";
    document.getElementById("errorDiv").innerHTML=msg;
}



function loadSettings(){
    return new Promise(function(resolve, reject) {
        //Get Settings
        var promiseSettings = Promise.all(
            [GetSetting("FilterSetting").catch(error => { 
                console.error("Error in GetSetting1"); })
            , GetSetting("LastExecDate").catch(error => { 
                console.error("Error in GetSetting2"); })
            , GetSetting("DateFilter").catch(error => { 
                console.error("Error in GetSetting3"); })
            , GetSetting("SourceSetting").catch(error => { 
                console.error("Error in GetSetting4"); })
            ,GetSetting("QueryId").catch(error => { 
                console.error("Error in GetSetting5"); })
            ]);
        promiseSettings.then(function(data) {
            //FilterSetting
            if (data[0]==null){
                _filterSetting="somefields";
            }
            console.log("FilterSetting="+data[0]);
            filterSelection.value=data[0];

            //LastExecDate
            LastExecDate=data[1];
            SaveSetting("LastExecDate",new Date());

            //DateFilter
            if (data[2]==null){
                data[2]="seven";
            }
            console.log("DateFilter="+data[2]);
            dateFilter.value=data[2];

            //SourceSetting
            if (data[3]==null){
                data[3]="source_following";
                SaveSetting("SourceSetting","source_following");
                sourceWits="source_following";
            }
            else {sourceWits=data[3];}
            console.log("SourceSetting="+data[3]);
            document.getElementById(data[3]).checked = true; 
            
            
            //QueryId
            if (data[4]!=null){
                document.getElementById("queryId").value = data[4];   
                
            }
            queryId=data[4];
            console.log("QueryId="+data[4]);
            resolve();
        })
    });
}
function changeDisplayByClassName(elementClass,newValue){
    
        Array.prototype.forEach.call(document.getElementsByClassName(elementClass),element => {	
                element.style.display = newValue;	
            });
}
function removeStyleByClassName(elementClass){
    
    Array.prototype.forEach.call(document.getElementsByClassName(elementClass),element => {	
            element.removeAttribute("style")
        });
}	

function onSaveQueryId(){
    var _queryId=document.getElementById("queryId").value;
    SaveSetting("QueryId",_queryId).then(function(){window.location.reload();});;
}
function onchangeSource(element){
    SaveSetting("SourceSetting",element.value);
}
function onchangeDateFilter(element){
    SaveSetting("DateFilter",element).then(function(){window.location.reload();});
    
    
}
function onchangeShowFilter(element){
    SaveSetting("FilterSetting",element).then(function(){updateVisibility(element);});
    
}

//Show or hide html elements based on the information we want to see (Filter drop down)
function updateVisibility(element){
    SaveSetting("FilterSetting",element);
    switch(element)
    {
        case 'comments':
            changeDisplayByClassName("showHideFields","none");
            changeDisplayByClassName("showHideSpecialField","block");
            appInsights.trackEvent({name:"FilterComments"});
            
            break;
        case 'somefields':
            changeDisplayByClassName("showHideFields","block");
            changeDisplayByClassName("showHideSpecialField","none");	
            appInsights.trackEvent({name:"FilterSomeFields"});
            break;
        case 'all':
            changeDisplayByClassName("showHideFields","block");
            removeStyleByClassName("showHideSpecialField");
            appInsights.trackEvent({name:"FilterAll"});
            break;
    }

}

function filterByContributor_click(user){
    updateVisibility(document.getElementById("filterSelection").value);
    
    document.getElementById("resetToAll").disabled = (user=='resetToAll'); //enable the Reset filter only if we're filtering by user

    contributors.forEach(function(item){
        if(user=='resetToAll' || item.uniqueName==user ){
            changeDisplayByClassName(item.uniqueName,"block");
            
        }else {
            changeDisplayByClassName(item.uniqueName,"none");
        }
    });
    
    filteredUser=user;
    
}

function pageLoaded(){
    document.getElementById("resetToAll").disabled = true; 
}


///Part of templating html 
function render(props) {
    return function(tok, i) { 
        //  return (i % 2) ? getDescendantProp(props,tok) : tok; 
        //return (i % 2) ? Object.byString(props,tok) : tok; 
        if (i % 2){
            //NEEDS TO BE FIXED. getDescendantProp is not solving for this nested type query because of the dot.
            return eval("props."+tok);
            // if (tok=="fields[System.History][newValue]") {
            // 	// //return props.fields[System.History][newValue]];
            // 	return props.fields["System.History"].newValue;
            // }else {
            // 	return getDescendantProp(props,tok);
            // }
        }
        else {return tok;}
    };
}
