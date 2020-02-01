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
    },
    getManagedProperties: function(query) {
        
        // query is a string e.g... `ContentTypeId:0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D0014107EC48E4A4F77B204AD0DA6EDE3B1*`

        fetch('/_api/contextInfo', {
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            }
        }).then(res => res.json()).then(contextData => {
            fetch(`/_api/search/query?queryText='${query}'&rowlimit=1&refiners='managedproperties(filter%3d600%2f0%2f*)'`, {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": contextData.d.GetContextWebInformation.FormDigestValue
                }
            }).then(res => res.json()).then(data => {
                let managedProperties = data.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results.map(item => item.RefinementName);
                let searchResults = data && data.d && data.d.query;
                let results = searchResults.PrimaryQueryResult && searchResults.PrimaryQueryResult.RelevantResults;
                let items = [];
                results.Table.Rows.results.forEach(row => {
                    items.push(row.Cells.results.reduce((rowData, cell) => {
                        rowData[cell.Key] = cell.Value;
                        return rowData;
                    }, {}));
                    return row;
                }, []);
                console.log('All Managed Properties for item', items[0], managedProperties);
                return managedProperties;
            });
        });
    }
};