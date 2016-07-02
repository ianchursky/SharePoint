var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Main = {
    
    // e.g. _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists/getbytitle('Pages')/items"
    //      _spPageContextInfo.webServerRelativeUrl + "/_api/web/sitegroups"
    ajax: function(type, url, successCallback, errorCallback) {
        var xhr = new XMLHttpRequest();
        if(typeof type === 'undefined') {
            type = 'GET';
        }
        xhr.open(type, url, true);
        xhr.setRequestHeader("Accept", "application/json; odata=verbose");
        xhr.setRequestHeader("Content-Type", "application/json; odata=verbose");
        xhr.onreadystatechange = function() { 
            if(xhr.readyState === 4) {
                if(xhr.status === 200){
                    if(successCallback && typeof successCallback === 'function') {
                        successCallback.call(this, xhr.responseText);
                    }
                } else {
                    if(errorCallback && typeof errorCallback === 'function') {
                        errorCallback.call(this, xhr.status + ' - ' + xhr.statusText);
                    }
                }
            }
        };        
        xhr.send();        
    }
   
};