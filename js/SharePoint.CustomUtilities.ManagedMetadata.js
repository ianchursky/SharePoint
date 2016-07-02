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