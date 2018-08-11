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