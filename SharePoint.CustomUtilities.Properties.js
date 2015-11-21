var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Properties = {

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