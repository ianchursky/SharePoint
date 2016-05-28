var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.CustomActions = {
    
    getAllSiteCustomActions: function(){
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_site();
            var customActions = web.get_userCustomActions();
            var customActionArray = [];
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                var customActionEnumerator = customActions.getEnumerator();
                while (customActionEnumerator.moveNext()) {
                    var currentCustomAction = customActionEnumerator.get_current();
                    var data = {
                        Title: currentCustomAction.get_title(),
                        Description: currentCustomAction.get_description(),
                        URL: currentCustomAction.get_scriptSrc()
                    };
                    console.log(data);
                    customActionArray.push(data);
                }
            });
        });                
    },       
    setSiteCustomAction: function(name, description, url, sequence){
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_site();
            var customActions = web.get_userCustomActions();
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                customAction = customActions.add();
                customAction.set_title(name);
                customAction.set_location("ScriptLink");
                customAction.set_description(description || "No Description");
                customAction.set_scriptSrc(url);
                customAction.set_sequence(sequence || 100);
                customAction.update();
                clientContext.executeQueryAsync();
            });
        })                
    },
    removeSiteCustomAction: function(name) {
        
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_site();
            var customActions = web.get_userCustomActions();

            var customAction;
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                var customActionEnumerator = customActions.getEnumerator();
                while (customActionEnumerator.moveNext()) {
                    var currentCustomAction = customActionEnumerator.get_current();
                    if (currentCustomAction.get_title() == name) {
                        customAction = currentCustomAction;
                    }
                }
                customAction.deleteObject();
                clientContext.executeQueryAsync();
            });
        });        
    },
    getAllWebCustomActions: function(){
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_web();
            var customActions = web.get_userCustomActions();
            var customActionArray = [];
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                var customActionEnumerator = customActions.getEnumerator();
                while (customActionEnumerator.moveNext()) {
                    var currentCustomAction = customActionEnumerator.get_current();
                    var data = {
                        Title: currentCustomAction.get_title(),
                        Description: currentCustomAction.get_description(),
                        URL: currentCustomAction.get_scriptSrc()
                    };
                    console.log(data);
                    customActionArray.push(data);
                }
            });
        });                
    },       
    setWebCustomAction: function(name, description, url, sequence){
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_web();
            var customActions = web.get_userCustomActions();
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                customAction = customActions.add();
                customAction.set_title(name);
                customAction.set_location("ScriptLink");
                customAction.set_description(description || "No Description");
                customAction.set_scriptSrc(url);
                customAction.set_sequence(sequence || 100);
                customAction.update();
                clientContext.executeQueryAsync();
            });
        })                
    },
    removeWebCustomAction: function(name) {
        
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_web();
            var customActions = web.get_userCustomActions();

            var customAction;
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                var customActionEnumerator = customActions.getEnumerator();
                while (customActionEnumerator.moveNext()) {
                    var currentCustomAction = customActionEnumerator.get_current();
                    if (currentCustomAction.get_title() == name) {
                        customAction = currentCustomAction;
                    }
                }
                customAction.deleteObject();
                clientContext.executeQueryAsync();
            });
        });        
    }           
        
};