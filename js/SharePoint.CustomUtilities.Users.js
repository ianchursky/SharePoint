var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Users = {

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
    
    // https://msdn.microsoft.com/en-us/library/office/jj712733.aspx
    getMyProfileProperties: function(){

        SP.SOD.executeFunc('personproperties', 'SP.UserProfiles', function () {
            var clientContext = SP.ClientContext.get_current();
            var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
            var personProperties = peopleManager.getMyProperties();
            clientContext.load(personProperties);
            clientContext.executeQueryAsync(function(sender, args) {
                
                return {
                    'AccountName': personProperties.get_accountName(),
                    'PersonalUrl': personProperties.get_personalUrl(),
                    'PictureUrl': personProperties.get_pictureUrl()
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
                userArray.push(data);
        
            }
        
        }, function (sender, args) {
            console.log(args)
        });        
    }    
};