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