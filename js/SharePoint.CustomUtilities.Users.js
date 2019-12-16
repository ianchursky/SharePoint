var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Users = {

    // https://msdn.microsoft.com/en-us/library/office/jj712733.aspx
    getUserProfileProperties: function(name){

        SP.SOD.registerSod('sp.userprofiles.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.userprofiles.js'));
        SP.SOD.loadMultiple(["sp.userprofiles.js"], function () {
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

        SP.SOD.registerSod('sp.userprofiles.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.userprofiles.js'));
        SP.SOD.loadMultiple(["sp.userprofiles.js"], function () {
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

        SP.SOD.registerSod('sp.userprofiles.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.userprofiles.js'));
        SP.SOD.loadMultiple(["sp.userprofiles.js"], function () {

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