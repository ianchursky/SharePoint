var SharePoint = SharePoint || {};
SharePoint.CustomUtilities = SharePoint.CustomUtilities || {};

SharePoint.CustomUtilities.Files = {

    uploadFile: function(libraryName, fileName, fileContent){

        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var list = web.get_lists().getByTitle(libraryName);
        var fileCreateInfo = new SP.FileCreationInformation();
        fileCreateInfo.set_url(fileName);
        fileCreateInfo.set_content(new SP.Base64EncodedByteArray());

        for (var i = 0; i < fileContent.length; i++) {
            fileCreateInfo.get_content().append(fileContent.charCodeAt(i));
        }

        var newFile = list.get_rootFolder().get_files().add(fileCreateInfo);
        clientContext.load(newFile);

        clientContext.executeQueryAsync(function(sender, args){
            console.log('File uploaded successfully!')
        },
        function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });
       
    },
    deleteFile: function(filePath){
        var clientContext = new SP.ClientContext.get_current();
        var web = clientContext.get_web();
        clientContext.load(web);
        clientContext.executeQueryAsync(function(sender, args){
            var fileUrl = web.get_serverRelativeUrl() + filePath;
            var fileToDelete = web.getFileByServerRelativeUrl(fileUrl); 
            fileToDelete.deleteObject();
            clientContext.executeQueryAsync(function(sender, args){
                console.log('File deleted successfully');
            }, function(sender, args){
                console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
            });

        }, function(sender, args){
            console.error('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        });        
        
    }

};