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
                // https://msdn.microsoft.com/en-us/library/ee553810.aspx
                fieldArray.push({
                    'ID': fieldEnumerator.get_current().get_id(),
                    'Group': fieldEnumerator.get_current().get_group(),
                    'Title': fieldEnumerator.get_current().get_title(),
                    'Static Name': fieldEnumerator.get_current().get_staticName(),
                    'Internal Name': fieldEnumerator.get_current().get_internalName(),
                    'Field Type Kind': fieldEnumerator.get_current().get_fieldTypeKind()
                });
                
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
    },
    getAllListViews: function(name) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(name);   
        var views = list.get_views(); // get views
        
        clientContext.load(views);
        clientContext.executeQueryAsync(function(sender, args){
            var viewArray = [];
            var viewEnumerator = views.getEnumerator();
            while (viewEnumerator.moveNext()) {
                viewArray.push(viewEnumerator.get_current().get_title())
            }
            console.log(viewArray);
            return viewArray;
                              
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                  
    },
    getListItemsInView: function(listName, viewName) {
        var self = this;
        var clientContext = SP.ClientContext.get_current();
        var list = clientContext.get_web().get_lists().getByTitle(listName);
        var view = list.get_views().getByTitle(viewName);
        clientContext.load(view,'ViewQuery');
        clientContext.executeQueryAsync(function(sender, args){
             var viewQuery = "<View><Query>" + view.get_viewQuery() + "</Query></View>";
             self.getAllListItems(listName, viewQuery);                      
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                 
    },
    deleteListItem: function(listName, itemName) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(listName);   
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml("<View><Query><Where><Eq><FieldRef Name='Name' /><Value Type='Text'>"+ itemName +"</Value></Eq></Where></Query></View>");
        var items = list.getItems(camlQuery);
        clientContext.load(items);
        clientContext.executeQueryAsync(function(sender, args){
            var itemEnumerator = items.getEnumerator();
            var itemArray = []; // Sys.InvalidOperationException is thrown if list is modified while iterating
            while (itemEnumerator.moveNext()) {
                var item = itemEnumerator.get_current();
                itemArray.push(item);
            }
 
            itemArray.forEach(function(item){
                item.deleteObject(); 
            });

            clientContext.executeQueryAsync(function(sender, args){
                console.log('Success');
            }, function(sender, args){
                console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            });
                                
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                 
    },       
};


