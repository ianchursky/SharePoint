var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.CustomActions = {
    
    getAllCustomActions: function(name, url){
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
    setCustomAction: function(name, url){
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var clientContext = SP.ClientContext.get_current();
            var web = clientContext.get_web();
            var customActions = web.get_userCustomActions();
            clientContext.load(customActions);
            clientContext.executeQueryAsync(function () {
                customAction = customActions.add();
                customAction.set_title(name);
                customAction.set_location("ScriptLink");
                customAction.set_scriptSrc(url);
                customAction.set_sequence(100);
                customAction.update();
                clientContext.executeQueryAsync();
            });
        })                
    },
    removeCustomAction: function(name) {
        
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