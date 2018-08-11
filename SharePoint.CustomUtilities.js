var SharePoint = SharePoint || {};
SharePoint.ContentTypes = SharePoint.ContentTypes || {};

SharePoint.CustomUtilities.ContentTypes = {

    getContentType: function(contentTypeId){
        var clientContext = SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var contentTypeCollection = web.get_contentTypes();
        var contentType = contentTypeCollection.getById(contentTypeId);      
        contentType.get_name();
    },

    getFieldsOnContentType: function(contentTypeId){
        var clientContext = SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var contentTypeCollection = web.get_contentTypes();
        var contentType = contentTypeCollection.getById(contentTypeId)
        var fields = contentType.get_fields();
        clientContext.load(fields);
        clientContext.executeQueryAsync(function(){
            var fieldEnumerator = fields.getEnumerator();
            while (fieldEnumerator.moveNext()) {
                var item = fieldEnumerator.get_current();
                var data = {
                    'ID': item.get_id().toString(),
                    'Title': item.get_title(),
                }
                console.log(data);
            }
        },
        function(sender,args){ 
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
    },

    getSiteFields: function(){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var fields = web.get_fields();
        clientContext.load(fields);
        clientContext.executeQueryAsync(function(){
            var fieldEnumerator = fields.getEnumerator();
            while (fieldEnumerator.moveNext()) {
                var item = fieldEnumerator.get_current();
                var data = {
                    'ID': item.get_id().toString(),
                    'Title': item.get_title(),
                }
                console.log(data);
            }            
        },
        function(sender,args){ 
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                 
    },

    setSiteFieldToHidden: function(title){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var fields = web.get_fields();
        var field = fields.getByTitle(title);
        field.set_hidden(true);
        clientContext.load(field);                
        clientContext.executeQueryAsync(function(){  
            console.log("Field " + title + " set to hidden...")   
        },
        function(sender,args){ 
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });                 
    }    

}
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
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Files = {

    uploadFile: function(libraryName, fileName, fileContent){

        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(libraryName);
        var fileCreateInfo = new SP.FileCreationInformation();
        fileCreateInfo.set_url(fileName);
        fileCreateInfo.set_content(new SP.Base64EncodedByteArray());

        for (var i = 0; i < fileContent.length; i++) {
            fileCreateInfo.get_content().append(fileContent.charCodeAt(i));
        }

        var newFile = list.get_rootFolder().get_files().add(fileCreateInfo);
        clientContext.load(newFile);

        clientContext.executeQueryAsync(function(sender, args){
            console.log('File uploaded successfully!')
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });
       
    },
    deleteFile: function(filePath){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        clientContext.load(web);
        clientContext.executeQueryAsync(function(sender, args){
            var fileUrl = web.get_serverRelativeUrl() + filePath;
            var fileToDelete = web.getFileByServerRelativeUrl(fileUrl); 
            fileToDelete.deleteObject();
            clientContext.executeQueryAsync(function(sender, args){
                console.log('File deleted successfully');
            }, function(sender, args){
                console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            });

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
        
    }

};
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Groups = {

    // https://msdn.microsoft.com/en-us/library/office/jj667833.aspx
    // API Endpoint: _spPageContextInfo.webServerRelativeUrl + "/_api/web/sitegroups"
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
            console.error(args)
        });        
    },
    createSiteGroup: function(name, description) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var groupCollection = web.get_siteGroups();
        
        var group = new SP.GroupCreationInformation();
        group.set_title(name);
        group.set_description(description);

        var newGroup = groupCollection.add(group);
        var roleDefinition = web.get_roleDefinitions().getByType(SP.RoleType.editor);
        var roleDefinitionCollection = SP.RoleDefinitionBindingCollection.newObject(clientContext);  
        roleDefinitionCollection.add(roleDefinition);   
        var roleAssignments = web.get_roleAssignments();  
        roleAssignments.add(newGroup, roleDefinitionCollection);           
        newGroup.set_allowMembersEditMembership(true);  
        newGroup.set_onlyAllowMembersViewMembership(false);

        clientContext.executeQueryAsync(function (sender, args) {
            console.log("Group " + name + " created...")
        }, function (sender, args) {
            console.error(args)
        });        
    },

    
};
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
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.ManagedMetadata = {

    getTermStores: function(){
        var context = new SP.ClientContext.get_current();
        var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStores = session.get_termStores();
        context.load(session);
        context.load(termStores);
        context.executeQueryAsync(function (sender, args) {
            var termStoresEnumerator = termStores.getEnumerator();
            var termStoresArray = [];
    
            while (termStoresEnumerator.moveNext()) {
                var item = termStoresEnumerator.get_current();
                termStoresArray.push({
                    'ID': item.get_id(),
                    'Name': item.get_name()  
                });
            }

            console.log(termStoresArray);
                        
        }, function (sender, args) {
            console.error('Error: ' + args.get_message() + '\n' + args.get_stackTrace());
        });
    },
    getTermGroups: function(){

        var context = new SP.ClientContext.get_current();
        var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
        context.load(termStore);
        context.executeQueryAsync(function (sender, args) {
            var termGroups = termStore.get_groups();
            context.load(termGroups);
            context.executeQueryAsync(function (sender, args) { 
                var termGroupsEnumerator = termGroups.getEnumerator();
                while (termGroupsEnumerator.moveNext()) {
                    var termGroup = termGroupsEnumerator.get_current();
                    console.log({
                        'id': termGroup.get_id(),
                        'name': termGroup.get_name()
                    });
                }
            }, function(sender, args){
                console.error('Unable to get term groups');
            }); 
        }, function(sender, args){
            console.error('Unable to load term store');
        });

    },    
    getTermSets: function(){

        var context = new SP.ClientContext.get_current();
        var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
        context.load(termStore);
        context.executeQueryAsync(function (sender, args) {
            var termGroups = termStore.get_groups();
            context.load(termGroups);
            context.executeQueryAsync(function (sender, args) { 
                var termGroupsEnumerator = termGroups.getEnumerator();
                while (termGroupsEnumerator.moveNext()) {
                    var termGroup = termGroupsEnumerator.get_current();
                    context.load(termGroup);
                    context.executeQueryAsync(function (sender, args) {
                        var termSets = termGroup.get_termSets();
                        context.load(termSets);
                        context.executeQueryAsync(function (sender, args) {
                            var termSetsEnumerator = termSets.getEnumerator();
                            var termSetArray = [];
                            while(termSetsEnumerator.moveNext()) { 
                                var termSet = termSetsEnumerator.get_current();
                                termSetArray.push({
                                    'id': termSet.get_id(),
                                    'name': termSet.get_name()
                                });
                            }

                            console.log(termSetArray);

                        }, function(sender, args){

                        });
                    }, function(sender, args){
                        console.error('Unable to get term sets in group');
                    });
                }
            }, function(sender, args){
                console.error('Unable to get term groups');
            }); 
        }, function(sender, args){
            console.error('Unable to load term store');
        });

    },     
    getDefaultSiteCollectionTermGroup: function () {

        var context = SP.ClientContext.get_current();
        var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
        var termGroup = termStore.getSiteCollectionGroup(context.get_site(), true);
        context.load(termGroup);
        context.executeQueryAsync(function () {
            var defaultGroup = termGroup.get_name();
            console.log(defaultGroup)
        }, function (sender, args) {
            console.error('Error: ' + args.get_message() + '\n' + args.get_stackTrace());
        });

    },      
    getTermsByTermSetName: function(termSetName){
        var context = SP.ClientContext.get_current();
        var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
        var termSets = termStore.getTermSetsByName(termSetName, 1033);
        var termSet = termSets.getByName(termSetName);
        var terms = termSet.getAllTerms();
        context.load(terms);
        context.executeQueryAsync(function () {

            var termEnumerator = terms.getEnumerator();
            var termArray = [];
            while (termEnumerator.moveNext()) {
                var term = termEnumerator.get_current();
                termArray.push({
                    'id': term.get_id(),
                    'name': term.get_name()
                });
            }
            console.log(termArray);

        }, function (sender, args) {
            console.error('Error: ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
    },
    getTermsByTermSetId(termsetId){

        var context = SP.ClientContext.get_current();
        var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
        var termSet = termStore.getTermSet(termsetId);
        var terms = termSet.getAllTerms();
        context.load(terms);
        context.executeQueryAsync(function(){
            var termEnumerator = terms.getEnumerator();
            var termArray = [];
            while (termEnumerator.moveNext()) {
                var term = termEnumerator.get_current();
                termArray.push({
                    'id': term.get_id(),
                    'name': term.get_name()
                });
            }
            console.log(termArray);

        }, function(sender,args){
            console.log(args.get_message());
        });

    }    
}
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Properties = {

    getAllSPWebProperties: function(){
        var clientContext = new SP.ClientContext();
        var web = clientContext.get_web();
        clientContext.load(web);
        var webProperties = web.get_allProperties();
        clientContext.load(webProperties);
        clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
            var properties = webProperties.get_fieldValues();
            for (var property in properties)
            {
                var propertyName = property;
                var propertyValue = properties[property];
                console.log(propertyName + " : " + propertyValue);
            }
        }),
        function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
    },

    getSPWebProperty: function(url, name){
        var clientContext = url ? new SP.ClientContext(url) : new SP.ClientContext().get_current();
        var webProperties = clientContext.get_web().get_allProperties();
        clientContext.load(webProperties);
        clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
            var prop = webProperties.get_fieldValues()[name];
        }),
        function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });	
    },	
	
    setSPWebProperty: function(url, name, value){

        var clientContext = url ? new SP.ClientContext(url) : new SP.ClientContext().get_current();
        var web = clientContext.get_web();
        clientContext.load(web);
        var webProperties = web.get_allProperties();
        clientContext.load(webProperties);
        webProperties.set_item(name, value);
        web.update();
        clientContext.executeQueryAsync(function (sender, args) {
            var prop = webProperties.get_fieldValues()[name];
        },function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });
    }
};
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};
SharePoint.CustomUtilities.Main = SharePoint.CustomUtilities.Main || {};

SharePoint.CustomUtilities.Search = {
    
    search: function(query, scope, contentType, limit){
        
        if(typeof scope === "undefined") {
            scope = _spPageContextInfo.webAbsoluteUrl;
        }
        
        // query e.g. (owstaxIdTaxKeyword:Seattle OR owstaxIdTaxKeyword:Seahawks)
        // contentType e.g. 0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D0014107EC48E4A4F77B204AD0DA6EDE3B1
        var baseUrl =  + _spPageContextInfo.webServerRelativeUrl + "_api/search/query?querytext='";
        var filters = "Path:"+ scope +" "+ contentType +"* AND ("+ query +")" + "'";
        var props = "&trimduplicates=false&rowlimit="+ limit + "&selectproperties='LastModifiedTime%2cTitle%2cPath'";
        
        SharePoint.CustomUtilities.Main.ajax("GET", baseUrl + filters + props, function(result){
            console.log(result);
        }, function(error){
            console.error(error)
        });
    },
    getPageSearchInfo: function(propertyArray){ // Credit: Ronnie B.
        var context = SP.ClientContext.get_current();
        var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(context);
        keywordQuery.set_queryText("Path:" + window.location.href);
        var properties = keywordQuery.get_selectProperties();
        
        for(var i =0; i < propertyArray.length; i++) {
            properties.add(propertyArray[i]);
        }
        
        var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(context);
        var results = searchExecutor.executeQuery(keywordQuery);
        context.executeQueryAsync(function()  {
            
            if (results.m_value.ResultTables) {
                $.each(results.m_value.ResultTables, function(index, table) {  
                    if(table.TableType == "RelevantResults") {
                        $.each(results.m_value.ResultTables[index].ResultRows, function () {  
                            console.log(this);
                        })  
                    }
                });  
            }           
        });
    },

    getAllUsers: function(searchTerm) {
 
        var clientContext = new SP.ClientContext.get_current();
        var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);
        keywordQuery.set_queryText(searchTerm);
        keywordQuery.set_sourceId("B09A7990-05EA-4AF9-81EF-EDFAB16C4E31");
        keywordQuery.set_rowLimit(500);
        keywordQuery.set_trimDuplicates(false); 
        var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
        var results = searchExecutor.executeQuery(keywordQuery);
         
        clientContext.executeQueryAsync(function(sender, args){
            console.log(results)
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        }); 
    },
    getSearchResults: function (queryResponse) {
        var results = { };
        results.elapsedTime = queryResponse.ElapsedTime;
        results.suggestion = queryResponse.SpellingSuggestion;
        results.resultsCount = queryResponse.PrimaryQueryResult.RelevantResults.RowCount;
        results.totalResults = queryResponse.PrimaryQueryResult.RelevantResults.TotalRows;
        results.totalResultsIncludingDuplicates = queryResponse.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates;
        results.items = this.convertSearchRowsToObjects(queryResponse.PrimaryQueryResult.RelevantResults.Table.Rows.results);
        return results;
    },    
    convertSearchRowsToObjects: function (itemRows) {
        var items = [];
        for (var i = 0; i < itemRows.length; i++) {
            var row = itemRows[i], item = {};
            for (var j = 0; j < row.Cells.results.length; j++) {
                item[row.Cells.results[j].Key] = row.Cells.results[j].Value;
            }
            items.push(item);
        }
        return items;
    }
};
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Sites = {

    loopSubWebs: function() {
        var clientContext = new SP.ClientContext();
        var web = clientContext.get_web();
        var webs = web.get_webs()
        clientContext.load(webs);
        clientContext.executeQueryAsync(function (sender, args) {
            var webEnumerator = webs.getEnumerator();
            var subwebArray = [];
            while (webEnumerator.moveNext()) {
                subwebArray.push({
                    "title": webEnumerator.get_current().get_title(),
                    "url": webEnumerator.get_current().get_url()
                });
            }

            console.log(subwebArray);
            return subwebArray;

        }, function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });
    }
        
};
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Themes = {

    getThemeGalleryFolder: function(){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var catalog = web.getCatalog(123);
        var rootFolder = catalog.get_rootFolder();
        var folders = rootFolder.get_folders();
        var folder15;
        clientContext.load(folders);
        clientContext.executeQueryAsync(function(sender, args){
            var folderEnumerator = folders.getEnumerator();
            while (folderEnumerator.moveNext()) { 
                var currentFolder = folderEnumerator.get_current();
                if (currentFolder.get_name() === '15') {
                    folder15 = currentFolder;
                }
            }

            if (!folder15) {
                console.error('Could not find the 15 folder in the themes directory');
            }            
            
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        

    },
    addFileToThemeGallery: function (fileName, fileContent) {
        var folder15;

        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var catalog = web.getCatalog(123);
        var rootFolder = catalog.get_rootFolder();
        var folders = rootFolder.get_folders();
        clientContext.load(folders);
        clientContext.executeQueryAsync(function (sender, args) {
            var folderEnumerator = folders.getEnumerator();
            while (folderEnumerator.moveNext()) {
                var currentFolder = folderEnumerator.get_current();
                if (currentFolder.get_name() === '15') {
                    folder15 = currentFolder;
                }
            }

            if (!folder15) {
                helper.showNotification('Error', 'Could not fild the 15 folder in the themes directory', 'error');
            }

            var fileCreateInfo = new SP.FileCreationInformation();
            fileCreateInfo.set_url(fileName);
            fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
            for (var i = 0; i < fileContent.length; i++) {
                fileCreateInfo.get_content().append(fileContent.charCodeAt(i));
            }
            var newFile = folder15.get_files().add(fileCreateInfo);
            clientContext.load(newFile);
            clientContext.executeQueryAsync(function (sender, args) {
                helper.showNotification('Success', 'File uploaded successfully', 'success');
            }, function (sender, args) {
                helper.showNotification('Error', args.get_message(), 'error');
            });

        },
        function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });

    },
    getComposedLook: function(name) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.getCatalog(124);
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml("<View><Query><Where><Eq><FieldRef Name='Name' /><Value Type='Text'>"+ name +"</Value></Eq></Where></Query></View>");
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
    createComposedLook: function(themeName, spcolorName, spfontName, masterPageName){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.getCatalog(124);
        var listItemCreateInfo = new SP.ListItemCreationInformation();
        var item = list.addItem(listItemCreateInfo);
        item.set_item('Name', themeName);
        item.set_item('Title', themeName);
        item.set_item('ThemeUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spcolorName +".spcolor");
        item.set_item('FontSchemeUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spfontName +".spfont");
        item.set_item('ImageUrl', null);
        item.set_item('MasterPageUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/"+ masterPageName +".master");
        item.update();
        clientContext.load(item);
        clientContext.executeQueryAsync(function(sender, args){
            console.log('Composed look created successfully!');
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });          

    },
    applyTheme: function(spcolorName, spfontName) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var colorPaletteUrl = _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spcolorName +".spcolor";
        var fontSchemeUrl = _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spfontName +".spfont";
        var backgroundImageUrl = null;
        var shareGenerated = true;
        web.applyTheme(colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated);
        web.update();
        clientContext.executeQueryAsync(function (sender, args) {
            helper.showNotification('Success', 'Theme applied successfully', 'success');
        }, function (sender, args) {
            helper.showNotification('Error', 'Error applying theme', 'error');
            console.log("Error: " + args.get_message());
        });
    }

}
var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Users = {

    // https://msdn.microsoft.com/en-us/library/office/jj712733.aspx
    getUserProfileProperties: function(name){
        SP.SOD.executeFunc('personproperties', 'SP.UserProfiles', function () {
            var clientContext = new SP.ClientContext.get_current();
            var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
            var personProperties = peopleManager.getPropertiesFor(name); /* string: 'domain\\name' for SharePoint on prem and 'i:0#.f|membership|admin@domainname.onmicrosoft.com' for Office 365 */
            clientContext.load(personProperties);
            clientContext.executeQueryAsync(function(sender, args) {
                
                console.log(personProperties.get_userProfileProperties());
                
                return {
                    'AccountName': personProperties.get_accountName(),
                    'PersonalUrl': personProperties.get_personalUrl(),
                    'PictureUrl': personProperties.get_pictureUrl(),
                    'Properties': personProperties.get_userProfileProperties()
                };

            }, function(sender, args){
                console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            });
        });                
    },  
    getMyProfileProperties: function(){

        SP.SOD.executeFunc('personproperties', 'SP.UserProfiles', function () {
            var clientContext = SP.ClientContext.get_current();
            var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
            var personProperties = peopleManager.getMyProperties();
            clientContext.load(personProperties);
            clientContext.executeQueryAsync(function(sender, args) {
                
                console.log(personProperties.get_userProfileProperties());
                
                return {
                    'AccountName': personProperties.get_accountName(),
                    'PersonalUrl': personProperties.get_personalUrl(),
                    'PictureUrl': personProperties.get_pictureUrl(),
                    'Properties': personProperties.get_userProfileProperties()
                };
                
            }, function(sender, args){
                console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            });           
            
        });
    },
    setUserProfileProperties: function(){
        var clientContext = SP.ClientContext.get_current();
        var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
        var userProfileProperties = peopleManager.getMyProperties();
        clientContext.load(userProfileProperties);
        clientContext.executeQueryAsync(function () {
    
            var currentUserAccountName = userProfileProperties.get_accountName();
            peopleManager.setSingleValueProfileProperty(currentUserAccountName, "Office", "Seattle");
            peopleManager.setSingleValueProfileProperty(currentUserAccountName, "Department", "Sales");
            peopleManager.setSingleValueProfileProperty(currentUserAccountName, "SPS-MUILanguages", "en-GB,en-US");
            peopleManager.setSingleValueProfileProperty(currentUserAccountName, "SPS-ContentLanguages", "en-GB");
            peopleManager.setSingleValueProfileProperty(currentUserAccountName, "SPS-Locale", "1033"); 

            //Set a multivalue property 
            var projects = ["SharePoint", "Office 365", "Architecture"];
            peopleManager.setMultiValuedProfileProperty(currentUserAccountName, "SPS-PastProjects", projects);
    
            clientContext.executeQueryAsync(function () {
                console.log("User profile properties changed...");
            }, function (sender, args) {
                console.log(args.get_message());
            });
    
        }, function (sender, args) {
            console.log(args.get_message());
        });        
    },      
    getMyUserInfoForSite: function(){
        var context = new SP.ClientContext.get_current();
        var website = context.get_web();
        var currentUser = website.get_currentUser();
        context.load(currentUser);
        context.executeQueryAsync(function(sender, args){
			
            return {
                'ID': currentUser.get_id(),
                'Title': currentUser.get_title(),
                'UserId': currentUser.get_userId(),
                'LoginName': currentUser.get_loginName(),
                'Email': currentUser.get_email(),
                'Path': currentUser.get_path(),
                'IsSiteAdmin': currentUser.get_isSiteAdmin(),
                'Groups': currentUser.get_groups()
            }; 
			           
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
        
    },
    // https://msdn.microsoft.com/en-us/library/office/jj667833.aspx
    getMyPersonalSiteUrl: function(){

        // does the current user have a personal site?
        SP.SOD.executeFunc('userprofile', 'SP.UserProfiles', function () {
            var clientContext = SP.ClientContext.get_current();
            var profileLoader = new SP.UserProfiles.ProfileLoader.getProfileLoader(clientContext);
            var userProfile = profileLoader.getUserProfile();
            clientContext.load(userProfile);
            clientContext.executeQueryAsync(function(sender, args) {
                var personalSite = userProfile.get_personalSite();
                clientContext.load(personalSite);
                clientContext.executeQueryAsync(function (sender, args) {
                    console.log(personalSite.get_url());
                }, function(sender, args){
                    console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
                });
            });
        });
    },    
    getAllSiteUsers: function(){
        var context = new SP.ClientContext.get_current();
        var website = context.get_web();
        var userList = website.get_siteUsers();
        context.load(userList);
        context.executeQueryAsync(function(sender, args){
            
            var userEnumerator = userList.getEnumerator();
            var userArray = [];
            while (userEnumerator.moveNext()) {
                var user = userEnumerator.get_current();
                var data = {
                    'ID': user.get_id(),
                    'Title': user.get_title(),
                    'Email': user.get_email(),
                    'LoginName': user.get_loginName()
                }
                
                userArray.push(data)
            }   
            
            console.log(userArray);
            return userArray;
            
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });           
        
    }    
};