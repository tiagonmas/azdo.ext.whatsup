var LastExecDate //Last time the extension was executed. Saved in settings.
var filterFromDate=new Date() //Date when we should start showing events from 
var filteredUser="";
var extVersion;
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
                
                //Get exisiting settings
                GetSetting("FilterSetting").then(function(_filterSetting){
                    if (_filterSetting==null){
                        _filterSetting="somefields";
                    }
                    console.log("FilterSetting="+_filterSetting);
                    filterSelection.value=_filterSetting;});

                    
                GetSetting("LastExecDate").then(function(_lastExecDate){
                    LastExecDate=_lastExecDate;
                    SaveSetting("LastExecDate",new Date());});

                //To do: All the getsettings should wait for completion before moving forward. promise all...
                GetSetting("DateFilter").then(function(_dateFilter){
                    if (_dateFilter==null){
                        _dateFilter="seven";
                    }
                    console.log("DateFilter="+_dateFilter);
                    dateFilter.value=_dateFilter;
                }).then(function(){

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

                console.log("DateFilter combo is "+dateFilter.value+" and query="+query.query);
                //https://docs.microsoft.com/en-us/azure/devops/extend/reference/client/api/tfs/workitemtracking/restclient/workitemtrackinghttpclient2_2?view=azure-devops#method_queryById
                //https://docs.microsoft.com/en-us/azure/devops/extend/reference/client/api/tfs/workitemtracking/restclient/workitemtrackinghttpclient2_2?view=azure-devops#method_queryById
                witClient.queryById("3397ce13-7f0f-4737-a453-820bb890c37e",projectId).then(function(foo){
                   console.log("queryById"+foo);
                },function(bar){
                    console.log("queryById rejected"+bar);
                });

                witClient.queryByWiql(query, projectId).then(
                    function(queryByWiqlResult) {  
                        var idsArr=new Array(queryByWiqlResult.workItems.length);
                        if (queryByWiqlResult.workItems.length==0)
                        {
                            appInsights.trackEvent({name: "noContent"});
                            document.getElementById("nocontent").style.visibility="visible" ;
                            document.getElementById("headbox").style.visibility="visible" ;
                            
                            //Improve: break out of promise chain. No need to continue moving forward 
                        }
                        else {
                            appInsights.trackEvent({name: "Content"});
                            appInsights.trackMetric("FollowingItems",queryByWiqlResult.workItems.length );
                            document.getElementById("nocontent").style.display="none" ;
                            document.getElementById("headbox").style.visibility="visible" ;
                        }

                        for (var i=0;i<queryByWiqlResult.workItems.length;i++){
                            idsArr[i]=queryByWiqlResult.workItems[i].id
                        }								
                        return idsArr;
                    }).then(function(idsArr){
                        if (idsArr.length>0)
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
                                        console.log("Err:"+Err);
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
                                    console.log("========ERROR: "+err);
                                    appInsights.trackException(err, "ErrFinal");
                                });

                        }
                        VSS.notifyLoadSucceeded();
                    });
                })

            }catch(err)
            {
                console.log("error queryByWiql:"+err);
                appInsights.trackException(err, "ErrqueryByWiql");
            }
        });
                
});

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
