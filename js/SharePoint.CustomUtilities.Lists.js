var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Lists = {
    getSiteLists: function(){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var lists = web.get_lists();
        clientContext.load(lists);
        clientContext.executeQueryAsync(function(sender, args){
            var itemEnumerator = lists.getEnumerator();
            var itemArray = [];
            while (itemEnumerator.moveNext()) {
                var data = {
                    'ID': itemEnumerator.get_current().get_id(),
                    'Title': itemEnumerator.get_current().get_title(),
                    'Description': itemEnumerator.get_current().get_description(),
                    'Created': itemEnumerator.get_current().get_created(),
                };
                itemArray.push(data);            
            }
            console.log(itemArray);
            return itemArray;          
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                
    },
    getAllListFields: function(name) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(name);   
        var fields = list.get_fields(); // get fields / column names
        clientContext.load(fields);
        clientContext.executeQueryAsync(function(sender, args){
            var fieldArray = [];
            var fieldEnumerator = fields.getEnumerator();
            while (fieldEnumerator.moveNext()) {
                fieldArray.push(fieldEnumerator.get_current().get_title())
            }
            console.log(fieldArray);
            return fieldArray;
                              
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                     
    },
    getAllListItems: function(name) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(name);
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml("<Query><OrderBy><FieldRef Name='ID' /></OrderBy></Query>"); // generic query to get all
        var items = list.getItems(camlQuery);
        clientContext.load(items);
        clientContext.executeQueryAsync(function(sender, args){
            var itemArray = [];
            var itemEnumerator = items.getEnumerator();
            while (itemEnumerator.moveNext()) {
                itemArray.push(itemEnumerator.get_current().get_fieldValues())
            }
            console.log(itemArray);
            return itemArray;
                              
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });         
    }   
};