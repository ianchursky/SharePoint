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