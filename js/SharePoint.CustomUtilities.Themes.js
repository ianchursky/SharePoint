var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Themes = {

    getThemeGalleryFolder: function(){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var catalog = web.getCatalog(123);
        var rootFolder = catalog.get_rootFolder();
        var folders = rootFolder.get_folders();
        var folder15;
        clientContext.load(folders);
        clientContext.executeQueryAsync(function(sender, args){
            var folderEnumerator = folders.getEnumerator();
            while (folderEnumerator.moveNext()) { 
                var currentFolder = folderEnumerator.get_current();
                if (currentFolder.get_name() === '15') {
                    folder15 = currentFolder;
                }
            }

            if (!folder15) {
                console.error('Could not find the 15 folder in the themes directory');
            }            
            
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        

    },
    addFileToThemeGallery: function (fileName, fileContent) {
        var folder15;

        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var catalog = web.getCatalog(123);
        var rootFolder = catalog.get_rootFolder();
        var folders = rootFolder.get_folders();
        clientContext.load(folders);
        clientContext.executeQueryAsync(function (sender, args) {
            var folderEnumerator = folders.getEnumerator();
            while (folderEnumerator.moveNext()) {
                var currentFolder = folderEnumerator.get_current();
                if (currentFolder.get_name() === '15') {
                    folder15 = currentFolder;
                }
            }

            if (!folder15) {
                helper.showNotification('Error', 'Could not fild the 15 folder in the themes directory', 'error');
            }

            var fileCreateInfo = new SP.FileCreationInformation();
            fileCreateInfo.set_url(fileName);
            fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
            for (var i = 0; i < fileContent.length; i++) {
                fileCreateInfo.get_content().append(fileContent.charCodeAt(i));
            }
            var newFile = folder15.get_files().add(fileCreateInfo);
            clientContext.load(newFile);
            clientContext.executeQueryAsync(function (sender, args) {
                helper.showNotification('Success', 'File uploaded successfully', 'success');
            }, function (sender, args) {
                helper.showNotification('Error', args.get_message(), 'error');
            });

        },
        function (sender, args) {
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });

    },
    getComposedLook: function(name) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.getCatalog(124);
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml("<View><Query><Where><Eq><FieldRef Name='Name' /><Value Type='Text'>"+ name +"</Value></Eq></Where></Query></View>");
        var items = list.getItems(camlQuery);        
        clientContext.load(items);
        clientContext.executeQueryAsync(function(sender, args){
            var itemArray = [];
            var itemEnumerator = items.getEnumerator();
            while (itemEnumerator.moveNext()) {
                itemArray.push(itemEnumerator.get_current().get_fieldValues())
            }
            console.log(itemArray);
            return itemArray;
                              
        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        }); 
    },
    createComposedLook: function(themeName, spcolorName, spfontName, masterPageName){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.getCatalog(124);
        var listItemCreateInfo = new SP.ListItemCreationInformation();
        var item = list.addItem(listItemCreateInfo);
        item.set_item('Name', themeName);
        item.set_item('Title', themeName);
        item.set_item('ThemeUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spcolorName +".spcolor");
        item.set_item('FontSchemeUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spfontName +".spfont");
        item.set_item('ImageUrl', null);
        item.set_item('MasterPageUrl', _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/masterpage/"+ masterPageName +".master");
        item.update();
        clientContext.load(item);
        clientContext.executeQueryAsync(function(sender, args){
            console.log('Composed look created successfully!');
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });          

    },
    applyTheme: function(spcolorName, spfontName) {
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var colorPaletteUrl = _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spcolorName +".spcolor";
        var fontSchemeUrl = _spPageContextInfo.siteServerRelativeUrl + "/_catalogs/theme/15/"+ spfontName +".spfont";
        var backgroundImageUrl = null;
        var shareGenerated = true;
        web.applyTheme(colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated);
        web.update();
        clientContext.executeQueryAsync(function (sender, args) {
            helper.showNotification('Success', 'Theme applied successfully', 'success');
        }, function (sender, args) {
            helper.showNotification('Error', 'Error applying theme', 'error');
            console.log("Error: " + args.get_message());
        });
    }

}