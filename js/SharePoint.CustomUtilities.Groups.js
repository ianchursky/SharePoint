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
            console.log(groupArray);
            return groupArray;

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });          
        
    },
    getSitePeopleAndGroups: function(){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var userList = web.get_siteUserInfoList();
        var query = new SP.CamlQuery(); // object is a required parameter for this query
        var items = userList.getItems(query);
        clientContext.load(items);

        clientContext.executeQueryAsync(function(sender, args) {
            var itemEnumerator = items.getEnumerator();
            var itemArray = [];
            while (itemEnumerator.moveNext()) {
                var item = itemEnumerator.get_current();

                //console.log(item.get_fieldValues()); // displays all the proerties of list item
                
                var data = {
                    'ID': item.get_id(),
                    'Title': item.get_item('Title'),
                    'UserName': item.get_item('UserName'),
                    'IsActive': item.get_item('IsActive'),                    
                    'IsSiteAdmin': item.get_item('IsSiteAdmin')
                };
                itemArray.push(data);            
            }
            console.log(itemArray);
            return itemArray;
                        
        }, function(sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });    
    },
    getUsersInGroupByName: function(groupName) {
        var clientContext = new SP.ClientContext.get_current();
        var groupCollection = clientContext.get_web().get_siteGroups();
        var membersGroup = groupCollection.getByName(groupName);
        var users = membersGroup.get_users();
        clientContext.load(users);
        clientContext.executeQueryAsync(function (sender, args) {
        
            var userEnumerator = users.getEnumerator();
            var userArray = [];
            while (userEnumerator.moveNext()) {
                var item = userEnumerator.get_current();
                var data = {
                    'ID': item.get_id(),
                    'Title': item.get_title(),
                    'Email': item.get_email(),
                    'LoginName': item.get_loginName()
                }
                console.log(data);
                userArray.push(data);
            }
        
        }, function (sender, args) {
            console.log(args)
        });        
    }        
    
};