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
    }
        
};