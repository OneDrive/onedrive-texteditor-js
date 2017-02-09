// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

var markdownEditor = {
    applicationId: "02590d81-45de-4e3d-a722-4f558244f068",
    redirectUrl: "http://localhost:9999/index.html",
    defaultFileName: "textfile1.txt",

    /************ Open *************/

    // Open a new file by invoking the picker to select a file from OneDrive.
    openFile: function () {
        var parent = this;

        // Build the options that we pass to the picker
        var options = {
            // Specify the options for the picker (we want download links and only one file).
            clientId: parent.applicationId,
            action: "download",
            multiSelect: false,
            advanced: {
                filter: ".md,.mdown,.txt",
                queryParameters: "select=*,name,size"
            },
            success: function (files) {
                // Get the first selected file (since we're not doing multi-select) and open it
                var selectedFile = files.value[0];
                parent.openItemInEditor(selectedFile);

                // Update our state to remember how we can use the API to write back to this file.
                parent.accessToken = { accessToken: files.accessToken };
            },
            error: function (e) { window.alert("Error picking a file: " + e); },
        };
        OneDrive.open(options);
    },

    // Method used to open the picked file into the editor. Resets local state and downloads the file from OneDrive.
    openItemInEditor: function (fileItem) {
        this.lastSelectedFile = fileItem;

        // Retrieve the contents of the file and load it into our editor
        var downloadLink = fileItem["@microsoft.graph.downloadUrl"];
        var parent = this;
        $.ajax(downloadLink, {
            success: function (data, status, xhr) {
                parent.setEditorBody(xhr.responseText);
                parent.setFilename(fileItem.name);
                $("#canvas").attr("disabled", false);
                parent.openFileID = fileItem.id;
            }
        });
    },

    /************ Save As *************/

    // Save the contents of the editor back to the server using a new filename. Invokes
    // the picker to allow the user to select a folder where the file should be saved.
    saveAsFile: function () {
        var parent = this;
        var filename = "";
        if (this.lastSelectedFile)
            filename = this.lastSelectedFile.name;
        if (!filename || filename == "")
            filename = this.defaultFileName;

        // Build the picker options to query for a folder where the file should be saved.
        var options = {
            clientId: parent.applicationId,
            action: "query",
            advanced: {
                queryParameters: "select=id,name,parentReference"
            },
            success: function (selection) {
                // The return here is the folder where we need to upload the item
                var folder = selection.value[0]; 
                
                // Update the lastSelectedFile item with the details of the new destination folder
                if (!parent.lastSelectedFile) {
                    parent.lastSelectedFile = { 
                        id: null,
                        name: filename,
                        parentReference: {
                            driveId: folder.parentReference.driveId,
                            id: folder.id
                        }
                    }
                } else {
                    parent.lastSelectedFile.parentReference.driveId = folder.parentReference.driveId;
                    parent.lastSelectedFile.parentReference.id = folder.id
                }
                
                // Store the access token from the file picker, so we don't need to get a new one
                parent.accessToken = { accessToken: selection.accessToken };
                // Call the saveFileWithToken method, which will write this file back with Microsoft Graph
                parent.ensureAccessToken(parent.user, parent.saveFileWithToken, 
                { parent: parent, 
                  alert: "File was saved to '" + folder.name + "' successfully.",
                  uploadIntoParentFolder: true });
            },
            error: function (e) { window.alert("An error occured while saving the file: " + e);
            }
        };

        // Launch the picker
        OneDrive.save(options);
    },

    /************ Save *************/

    // Save the contents of the editor back to the file that was opened. If no file was
    // currently open, the saveAsFile method is invoked.
    saveFile: function () {
        // Check to see if we know about an open file. If not, revert to the save as flow.
        if (this.openFileID == "") {
            this.saveAsFile();
            return;
        }

        // Since we're using the API to write back changes, make sure we have a valid
        // access_token and then invoke saveFileWithToken.
        this.ensureAccessToken(this.user, this.saveFileWithToken, {parent: this});
    },

    // Save the contents of the editor back to a target drive item. Requires an access token
    // and a parent reference in the state object.
    saveFileWithToken: function (token, state) {
        if (token == null) {
            // An error occured and we don't have a token available to save.
            window.alert("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }

        if (state == null) {
            window.alert("The state parameter is required for this method.");
            return;
        }

        var parent = state.parent;
        
        // For SAVE so we don't invoke the picker, we're going to use the REST API directly
        // using some values that we stored from the picker when we opened the item.
        var url = parent.generateGraphUrl(parent.lastSelectedFile, (state && state.uploadIntoParentFolder) ? true : false, true);

        // Create a new XMLHttpRequest() and execute it.
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4) {
                // Need to update parent.lastSelectedFile with the response from this 
                // request so future saves go to the right place
                var uploadedItem = JSON.parse(xhr.responseText);
                if (uploadedItem && uploadedItem.name && uploadedItem.parentReference )
                    parent.lastSelectedFile = uploadedItem;

                if (state.alert) {
                    window.alert(state.alert);
                } else {
                    window.alert("File saved successfully.");
                }
            }
        }
        xhr.onerror = function() {
            window.alert("Error occured saving file.");
        }

        xhr.open("PUT", url, true);
        xhr.setRequestHeader("Content-type", "application/octet-stream");
        xhr.setRequestHeader("Authorization", "Bearer " + parent.accessToken.accessToken)

        // Get the body text and encode the line breaks to Windows-style.
        var bodyContent = $("#canvas").val();
        var bodyContentLineBreaks = bodyContent.replace(/\r\n|\r|\n/g, "\r\n");

        xhr.send(bodyContentLineBreaks);
    },

   

    /************ Rename File *************/

    // Rename the currently open file by providing a new name for the file via an input
    // dialog
    renameFile: function () {
        var oldFilename = (this.lastSelectedFile && this.lastSelectedFile.name) ? this.lastSelectedFile.name : this.defaultFileName;
        var newFilename = window.prompt("Rename file", oldFilename);
        if (!newFilename) return;
        
        this.setFilename(newFilename);

        if (this.lastSelectedFile && this.lastSelectedFile.id) {
            // Patch the file to rename it in real time.
            this.lastSelectedFile.name = newFilename;
            this.ensureAccessToken(this.user, this.patchDriveItemWithToken, 
                { parent: this, 
                  alert: "File was renamed successfully.", 
                  driveItem: this.lastSelectedFile, 
                  propertyList: [ "name" ]});
        }
    },

    // Patch a DriveItem via the Microsoft Graph. Expects an access token and the state
    // parameter to have a driveItem, propertyList, and parent (reference to this object) properties.
    // driveItem is the item to update, which needs to have its id and parentReference.driveId property set to the item to patch.
    // propertyList is an array of properties on the driveItem that will be submitted as the patch. These must be top-level properties of the driveItem
    patchDriveItemWithToken: function(token, state) {
        if (token == null) {
            // An error occured and we don't have a token available to save.
            window.alert("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }
        
        if (state == null) {
            window.alert("The state parameter is required for this method.");
            return;
        }

        var item = state.driveItem;
        var propList = state.propertyList;
        var parent = state.parent;
        
        // For PATCH we don't invoke the picker, we're going to use the REST API directly
        var url = parent.generateGraphUrl(item, false, false);

        // Create a new XMLHttpRequest() and execute it.
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4) {
                // Need to update parent.lastSelectedFile with the response from this 
                // request so future saves go to the right place
                var uploadedItem = JSON.parse(xhr.responseText);
                if (uploadedItem && uploadedItem.name && uploadedItem.parentReference )
                    parent.lastSelectedFile = uploadedItem;
                if (state.alert) {
                    window.alert(state.alert);
                } else {
                    window.alert("File saved successfully.");
                }
            }
        }
        xhr.onerror = function() {
            window.alert("Error occured patching file metadata.");
        }

        xhr.open("PATCH", url, true);
        xhr.setRequestHeader("Content-type", "application/json");
        xhr.setRequestHeader("Authorization", "Bearer " + parent.accessToken.accessToken)

        // Copy the values into patchData from the driveItem, based on names of properties in propertyList
        var patchData = { };
        for(var i=0, len = propList.length; i < len; i++)
        {
            patchData[propList[i]] = item[propList[i]];
        }
        xhr.send(JSON.stringify(patchData));
    },

    // Used to generate the Microsoft Graph URL for a target item, with a few parameters
    // uploadIntoParentFolder: bool, indicates that we should be targeting the parent folder + filename instead of the item itself
    // targetContentStream: bool, indicates we should append /content to the item URL
    generateGraphUrl: function(driveItem, uploadIntoParentFolder, targetContentStream) {
        var url = "https://graph.microsoft.com/v1.0/";
        if (uploadIntoParentFolder)
        {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" +driveItem.parentReference.id + "/children/" + driveItem.name;
        } else {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" + driveItem.id;
        }

        if (targetContentStream)
            url += "/content";

        return url;
    },


    /************ Create New File *************/

    // Clear the current editor buffer and reset any local state so we don't
    // overwrite an existing file by mistake.
    createNewFile: function () {
        this.lastSelectedFile = null;
        this.setFilename(this.defaultFileName);
        $("#canvas").attr("disabled", false);
        this.setEditorBody("");
    },

   
    /************ Utility functions *************/
    
    // Update the state of the editor with a new filename.
    setFilename: function (filename) {
        var btnRename = this.buttons["rename"];
        if (btnRename) {
            $(btnRename).text(filename);
        }
    },

    // Set the contents of the editor to a new value.
    setEditorBody: function (text) {
        $("#canvas").val(text);
    },

    // State and function to connect elements in the HTML page to actions in the markdown editor.
    buttons: {},
    wireUpCommandButton: function(element, cmd)
    {
        this.buttons[cmd] = element;
        switch(cmd) {
            case "new": 
                element.onclick = function () { markdownEditor.createNewFile(); return false; }
                break;
            case "open": 
                element.onclick = function () { markdownEditor.openFile(); return false; }
                break;
            case "save":
                element.onclick = function () { markdownEditor.saveFile(); return false; }
                break;
            case "saveAs":
                element.onclick = function () { markdownEditor.saveAsFile(); return false; }
                break;
            case "rename":
                element.onclick = function () { markdownEditor.renameFile(); return false; }
                break;
        }
    },

    // Uses a hidden iframe to request a new access token for Microsoft Graph, so we can make
    // Microsoft Graph calls outside of the picker context.
    // action is a function that takes two argument, the access token and a state value
    ensureAccessToken: function(user, action, state) {
        
        if (this.accessToken) {
            action(this.accessToken, state);
            return;
        }

        // NOTE: The following code is not used in this example currently, but would allow
        // JavaScript to generate new tokens as long as the user is still signed into AAD/MSA.

        // var parent = this;
        // var authorizeEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        // var authorizeParameters = "?client_id=" + this.applicationId + 
        //     "&response_type=token" + 
        //     "&scope=Files.ReadWrite" + 
        //     "&redirect_uri=" + encodeURIComponent(this.redirectUrl) + 
        //     "&prompt=none" + 
        //     "&domain_hint=" + encodeURIComponent(user.domain) + 
        //     "&login_hint=" + encodeURIComponent(user.login);

        // var iframe = $("<iframe />").attr({
        //     width: 1,
        //     height: 1,
        //     src: authorizeEndpoint + authorizeParameters
        // })
        // $(document.body).append(iframe);
        // iframe.on("load", function(iframeData) {
        //     parent.parseAccessTokenFromIFrame(iframeData, action, state);
        // });
    },

    // NOTE: The following code is not used in this example currently, but would allow
    // JavaScript to generate new tokens as long as the user is still signed into AAD/MSA.
    // Parse the token values from the iframe result
    // parseAccessTokenFromIFrame: function(iframeData, action, state) {
    //     var frameHref = "";
    //     try {
    //         // this will throw for any issues other than succes
    //         frameHref = iframeData.currentTarget.contentWindow.location.href;
    //     }
    //     catch (error) {
    //         action(null, state);
    //         return;
    //     }

    //     // parse the iframe query string
    //     var accessToken = this.getQueryStringParameterByName("access_token", frameHref);
    //     var expiresInSeconds = this.getQueryStringParameterByName("expires_in", frameHref);

    //     if (accessToken != null) {
    //         this.accessToken = { accessToken: accessToken, expiresInSeconds: expiresInSeconds };
    //     }
    
    //     var iframe = $(iframeData.currentTarget);
    //     iframe.remove();

    //     action(this.accessToken, state);
    // },

    // Parse query string parmaeters and return the value of a parameter
    // getQueryStringParameterByName: function (name, url) {
	// 		name = name.replace(/[\[\]]/g, "\\$&");
	// 		var regex = new RegExp("[?&#]" + name + "(=([^&#]*)|&|#|$)");
	// 		var results = regex.exec(url);
	// 		if (!results) return null;
	// 		if (!results[2]) return '';
	// 		return decodeURIComponent(results[2].replace(/\+/g, " "));
    // },

    // An object representing the currently selected file
    lastSelectedFile: null,

    // The access_token returned from the picker so we can make API calls again.
    accessToken: null,

    user: {
        id: "nouser@contoso.com",
        domain: "organizations"
    }
}

$("#canvas").attr("disabled", true);