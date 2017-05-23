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
    createListItem: function(siteUrl, listTitle, valuesArray) {
        var clientContext = new SP.ClientContext(siteUrl);
        var list = clientContext.get_web().get_lists().getByTitle(listTitle);
        var itemCreateInfo = new SP.ListItemCreationInformation();
        var listItem = list.addItem(itemCreateInfo);
        valuesArray.forEach(function(item, index){
            listItem.set_item(item["Title"], item["Value"]);
        });
        listItem.update();
        clientContext.load(listItem);            
        clientContext.executeQueryAsync(function(sender, args){
            console.log('Success');
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

    createList: function(listName, templateType) {

        var def = $.Deferred();
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title(listName);
        listCreationInfo.set_templateType(templateType || SP.ListTemplateType.genericList);
        var list = web.get_lists().add(listCreationInfo);
        var defaultView = list.get_defaultView();
        defaultView.get_viewFields().add("Created");
        defaultView.update();

        list.get_fields().addFieldAsXml("<Field Type='URL' Name='SiteUrl' DisplayName='Site Url' StaticName='SiteUrl' Required='TRUE' Format='Hyperlink' />", true, SP.AddFieldOptions.addFieldInternalNameHint);
        list.get_fields().addFieldAsXml("<Field Type='Choice' Name='SiteTemplate' DisplayName='Site Template' StaticName='SiteTemplate' Required='TRUE' Format='Dropdown'>"
        + "<Default>BLANKINTERNETCONTAINER#0</Default>"
        + "<CHOICES>"
        + "    <CHOICE>BLANKINTERNETCONTAINER#0</CHOICE>"
        + "    <CHOICE>STS#0</CHOICE>"
        + "    <CHOICE>DEV#0</CHOICE>"
        + "</CHOICES>"
        + "</Field>", true, SP.AddFieldOptions.addFieldInternalNameHint);

        clientContext.load(list);
        clientContext.executeQueryAsync(function () { }, function () { });
    },    

    deleteList: function (listName) {
        var def = $.Deferred();
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(listName);
        list.deleteObject();
        clientContext.executeQueryAsync(function () {
        }, function () {
        });
    },

    getImageRenditions: function(){
        var clientContext = new SP.ClientContext.get_current();
        var renditions = SP.Publishing.SiteImageRenditions.getRenditions(clientContext);
        clientContext.executeQueryAsync(function(sender, args){
            var itemArray = [];
            
            renditions.forEach(function(item, index){
                itemArray.push({
                    'ID': item.get_id(),
                    'TypeID': item.get_typeId(), 
                    'Version': item.get_version(), 
                    'Name': item.get_name(),
                    'Group': item.get_group(),     
                    'Width': item.get_width(),        
                    'Height': item.get_height(), 
                });
            })

            console.log(itemArray);

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
    },

    addEventReceiversToList: function(listName, name, url) {
        var context = new SP.ClientContext();
        var web = context.get_web();
        var list = web.get_lists().getByTitle(listName);
        var receiver = new SP.EventReceiverDefinitionCreationInformation();
        receiver.set_receiverName(name)
        receiver.set_sequenceNumber(100);
        receiver.set_receiverUrl(url);
        receiver.set_eventType(SP.EventReceiverType.itemAdded); // SP.EventReceiverType.itemAdded = 10001 = Item added
        receiver.set_synchronization(SP.EventReceiverSynchronization.synchronous);
        list.get_eventReceivers().add(receiver);
        context.load(list);
        context.executeQueryAsync(function(sender, args) {
            console.log('Event receiver added...')
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });

    },

    getEventReceiversForList: function(listName) {
        var context = new SP.ClientContext();
        var web = context.get_web();
        var list = web.get_lists().getByTitle(listName);
        var eventReceivers = list.get_eventReceivers();
        context.load(eventReceivers);

        context.executeQueryAsync(function(sender, args) {
            var eventReceiverArray = [];
            for(var i=0; i < eventReceivers.get_count(); i++) {
                var eventReceiver = eventReceivers.get_item(i);
                var item = {
                    "ID": eventReceiver.get_receiverId(),
                    "Name": eventReceiver.get_receiverName(),
                    "Path": eventReceiver.get_path(),
                    "EventType": eventReceiver.get_eventType(),
                    "Url": eventReceiver.get_receiverUrl(),
                    "Synchronization": eventReceiver.get_synchronization(),
                    "Class": eventReceiver.get_receiverClass(),
                    "Assembly": eventReceiver.get_receiverAssembly(),
                    "SequenceNumber": eventReceiver.get_sequenceNumber()  
                };
                eventReceiverArray.push(item);
            };

            console.log(eventReceiverArray);

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });
    }           
};


