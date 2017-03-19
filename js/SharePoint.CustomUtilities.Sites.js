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