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
    }   
};



var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Main = {
    
    // e.g. _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists/getbytitle('Pages')/items"
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
                
                var data = {
                    'ID': item.get_id(),
                    'Name': item.get_name()  
                };
                
                termStoresArray.push(data);
            }
                        
        }, function (sender, args) {
            console.error('Error: ' + args.get_message() + '\n' + args.get_stackTrace());
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
        var clientContext = new SP.ClientContext(url);
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

        var clientContext = new SP.ClientContext(url);
        var web = clientContext.get_web();
        clientContext.load(web);
        var webProperties = web.get_allProperties();
        clientContext.load(webProperties);
        webProperties.set_item(name, value);
        web.update();
        clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
            var prop = webProperties.get_fieldValues()[name];
        }),
        function (sender, args) {
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
    }
};
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