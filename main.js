var LastExecDate //Last time the extension was executed. Saved in settings.

appInsights.startTrackPage("Page");

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

VSS.init({
    explicitNotifyLoaded: true,
    usePlatformScripts: true, 
    usePlatformStyles: true
});

VSS.ready(function() {
    
    GetSetting("FilterSetting").then(function(filterSetting){
        if (filterSetting==null){
            filterSetting="somefields"
        }
        filterSelection.value=filterSetting;});
    
    GetSetting("LastExecDate").then(function(_lastExecDate){
        LastExecDate=_lastExecDate;
        SaveSetting("LastExecDate",new Date());
    });
    
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
    var projectId = context.project.id;
    var projectName = context.project.name;
    var HostName = context.host.name;
    var HostUri = context.host.uri;

    VSS.getAccessToken().then(function(token){
            return VSS_Auth_Service.authTokenManager.getAuthorizationHeader(token);
        }).then(function(authHeader){					
            
            try
            {
                var witClient = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
                
                //Get all the workitems that we're following
                var query = {query: "SELECT [System.Id] FROM workitems WHERE [System.Id] In (@Follows) AND [System.State] NOT IN ('Closed','Inactive','Completed') ORDER BY [System.ChangedDate] DESC" };
                witClient.queryByWiql(query, projectId).then(
                    function(queryByWiqlResult) {  
                        var idsArr=new Array(queryByWiqlResult.workItems.length);
                        if (queryByWiqlResult.workItems.length==0)
                        {
                            appInsights.trackEvent({name: "noContent"});
                            document.getElementById("nocontent").style.visibility="visible" ;
                            document.getElementById("headbox").style.visibility="hidden" ;
                            
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
                        return witClient.getWorkItems(idsArr, ["System.Title"])
                    }).then(function(itemsArr){
                        fetchContent(itemsArr,authHeader,HostName,projectName).
                            then(function(updates){
                                
                                //Update UI with comments applying template
                                var commentTpl = $('script[data-template="commentTemplate"]').text().split(/\$\{(.+?)\}/g);
                                var fieldsTpl = $('script[data-template="fieldsTemplate"]').text().split(/\$\{(.+?)\}/g);
                                
                                $('#list-comment-items').append(updates.map(function (item) {
                                    var myItemhtml;
                                    if (item.hasOwnProperty("fields")){
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

                                VSS.notifyLoadSucceeded();
                                appInsights.stopTrackPage("Page");
                            },function(err) {
                                console.log("========ERROR: "+err);
                                appInsights.trackException(err, "ErrFinal");
                            });

                    });

            }catch(err)
            {
                console.log("error queryByWiql:"+err);
                appInsights.trackException(err, "ErrqueryByWiql");
            }
        });
                
});

function changeDisplay(elementClass,newValue){
    
        Array.prototype.forEach.call(document.getElementsByClassName(elementClass),element => {	
                element.style.display = newValue;	
            });
}
function removeStyle(elementClass){
    
    Array.prototype.forEach.call(document.getElementsByClassName(elementClass),element => {	
            element.removeAttribute("style")
        });
}	

function onchangeFilter(element){
    SaveSetting("FilterSetting",element);
    updateVisibility(element);
}

//Show or hide html elements based on the information we want to see (Filter drop down)
function updateVisibility(element){
    SaveSetting("FilterSetting",element);
    switch(element)
    {
        case 'comments':
            changeDisplay("showHideFields","none");
            changeDisplay("showHideSpecialField","flex");
            appInsights.trackEvent({name:"FilterComments"});
            
            break;
        case 'somefields':
            changeDisplay("showHideFields","block");
            changeDisplay("showHideSpecialField","none");	
            appInsights.trackEvent({name:"FilterSomeFields"});
            break;
        case 'all':
            changeDisplay("showHideFields","block");
            removeStyle("showHideSpecialField");
            appInsights.trackEvent({name:"FilterAll"});
            break;
    }

}