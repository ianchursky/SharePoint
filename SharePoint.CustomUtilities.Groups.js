var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Groups = {

    // https://msdn.microsoft.com/en-us/library/office/jj667833.aspx
    getSiteGroups: function(){
        var clientContext = SP.ClientContext.get_current(); 
        var web = clientContext.get_web(); 
        var groups = web.get_siteGroups(); 
        clientContext.load(groups);
        clientContext.executeQueryAsync(function(sender, args) {
            var groupEnumerator = groups.getEnumerator();
            var groupArray = [];
            while (groupEnumerator.moveNext()) {
                var item = groupEnumerator.get_current();
                
                groupArray.push({
                    'ID': item.get_id(),
                    'Title': item.get_title()
                });
            }

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });          
        
    }
    
};