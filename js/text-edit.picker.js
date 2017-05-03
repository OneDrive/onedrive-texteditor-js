// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

var markdownEditor = {
    applicationId: "5fc58006-b25f-4eaf-8f30-527c6fa7e5f5",
    defaultFileName: "textfile1.txt",
    microsoftGraphApiRoot: "https://graph.microsoft.com/v1.0/",

    /************ Open *************/

    // Open a new file by invoking the picker to select a file from OneDrive.
    openFile: function () {
        var editor = this;

        // Build the options that we pass to the picker
        var options = {
            // Specify the options for the picker (we want download links and only one file).
            clientId: editor.applicationId,
            action: "download",
            multiSelect: false,
            advanced: {
                /* Filter the files that are available for selection */
                filter: ".md,.mdown,.txt",
                /* Request a few additional properties */
                queryParameters: "select=*,name,size",
                /* Request a read-write scope for the access token */
                scopes: ["Files.ReadWrite"]
            },
            success: function (files) {
                // Update our state to remember how we can use the API to write back to this file.
                editor.accessToken = files.accessToken;
                
                // Get the first selected file (since we're not doing multi-select) and open it
                var selectedFile = files.value[0];
                editor.openItemInEditor(selectedFile);
            },
            error: function (e) { editor.showError("Error occurred while picking a file: " + e, e); },
        };
        OneDrive.open(options);
    },

    // Method used to open the picked file into the editor. Resets local state and downloads the file from OneDrive.
    openItemInEditor: function (fileItem) {
        var editor = this;
        editor.lastSelectedFile = fileItem;
        // Retrieve the contents of the file and load it into our editor
        var downloadLink = fileItem["@microsoft.graph.downloadUrl"];
        
        // Use JQuery AJAX to download the file
        $.ajax(downloadLink, {
            success: function (data, status, xhr) {
                
                // load the file into the editor
                editor.setEditorBody(xhr.responseText);
                editor.setFilename(fileItem.name);
                $("#canvas").attr("disabled", false);
                editor.openFileID = fileItem.id;
            }, 
            error: function(xhr, status, err) {
                editor.showError(err);
            }
        });
    },

    /************ Save As *************/

    // Save the contents of the editor back to the server using a new filename. Invokes
    // the picker to allow the user to select a folder where the file should be saved.
    saveAsFile: function () {
        var editor = this;
        var filename = "";

        // Ensure we have a valid filename for the file
        if (editor.lastSelectedFile) {
            filename = editor.lastSelectedFile.name;
        }
        if (!filename || filename == "") {
            filename = editor.defaultFileName;
        }

        // Build the picker options to query for a folder where the file should be saved.
        var options = {
            clientId: editor.applicationId,
            action: "query",
            advanced: {
                // Request additional parameters when we save the file
                queryParameters: "select=id,name,parentReference"
            },
            success: function (selection) {
                // The return here is the folder where we need to upload the item
                var folder = selection.value[0]; 
                
                
                if (!editor.lastSelectedFile) {
                    // Remember the details of the file we just created                    
                    editor.lastSelectedFile = { 
                        id: null,
                        name: filename,
                        parentReference: {
                            driveId: folder.parentReference.driveId,
                            id: folder.id
                        }
                    }
                } else {
                    // Update the lastSelectedFile item with the details of the new destination folder
                    editor.lastSelectedFile.parentReference.driveId = folder.parentReference.driveId;
                    editor.lastSelectedFile.parentReference.id = folder.id
                }
                
                // Store the access token from the file picker, so we don't need to get a new one
                editor.accessToken = { accessToken: selection.accessToken };

                // Call the saveFileWithToken method, which will write this file back with Microsoft Graph
                editor.saveFileWithAPI( { uploadIntoParentFolder: true });
            },
            error: function (e) { editor.showError("An error occurred while saving the file: " + e, e); }
        };

        // Launch the picker
        OneDrive.save(options);
    },

    /************ Save *************/

    // Save the contents of the editor back to the file that was opened. If no file was
    // currently open, the saveAsFile method is invoked.
    saveFile: function () {
        var editor = this;
        // Check to see if we know about an open file. If not, revert to the save as flow.
        if (editor.openFileID == "") {
            editor.saveAsFile();
            return;
        }

        // Since we're using the API to write back changes, make sure we have a valid
        // access_token and then invoke saveFileWithToken.
        editor.saveFileWithAPI();
    },

    // Save the contents of the editor back to a target drive item. Requires an access token
    // and a parent reference in the state object.
    saveFileWithAPI: function (state) {
        var editor = this;
        if (editor.accessToken == null) {
            // An error occurred and we don't have a token available to save.
            editor.showError("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }

        // This method uses the REST API directly instead of the file picker,
        // using some values that we stored from the picker when we opened the item.
        var url = editor.generateGraphUrl(editor.lastSelectedFile, (state && state.uploadIntoParentFolder) ? true : false, "/content");

        var bodyContent = $("#canvas").val().replace(/\r\n|\r|\n/g, "\r\n");

        // Call the REST API to PUT the text value to the contents of the file.
        $.ajax(url, {
            method: "PUT",
            contentType: "application/octet-stream",
            data: bodyContent,
            processData: false,
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.name && data.parentReference) {
                    editor.lastSelectedFile = data;
                    editor.showSuccessMessage("File was saved.");
                }
            },
            error: function(xhr, status, err) {
                editor.showError(err);
            }
        });
    },
   

    /************ Rename File *************/

    // Rename the currently open file by providing a new name for the file via an input
    // dialog
    renameFile: function () {
        var editor = this;

        var oldFilename = (editor.lastSelectedFile && editor.lastSelectedFile.name) ? editor.lastSelectedFile.name : editor.defaultFileName;
        var newFilename = window.prompt("Rename file", oldFilename);
        if (!newFilename) return;
        
        editor.setFilename(newFilename);

        if (editor.lastSelectedFile && editor.lastSelectedFile.id) {
            // Patch the file to rename it in real time.
            editor.lastSelectedFile.name = newFilename;
            editor.patchDriveItemWithAPI({ propertyList: [ "name" ]} );
        } else {
            // The file hasn't been saved yet, so the rename is just local
        }
    },

    // Patch a DriveItem via the Microsoft Graph. Expects an access token and the state
    // parameter to have a driveItem, propertyList, and parent (reference to this object) properties.
    // driveItem is the item to update, which needs to have its id and parentReference.driveId property set to the item to patch.
    // propertyList is an array of properties on the driveItem that will be submitted as the patch. These must be top-level properties of the driveItem
    patchDriveItemWithAPI: function(state) {
        var editor = this;
        if (editor.accessToken == null) {
            // An error occurred and we don't have a token available to save.
            editor.showError("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }
        
        if (state == null) {
            editor.showError("The state parameter is required for this method.");
            return;
        }

        var item = editor.lastSelectedFile;
        var propList = state.propertyList;
       
        // For PATCH we don't invoke the picker, we're going to use the REST API directly
        var url = editor.generateGraphUrl(item, false, null);

        // Copy the values into patchData from the driveItem, based on names of properties in propertyList
        var patchData = { };
        for(var i=0, len = propList.length; i < len; i++)
        {
            patchData[propList[i]] = item[propList[i]];
        }

        $.ajax(url, {
            method: "PATCH",
            contentType: "application/json; charset=UTF-8",
            data: JSON.stringify(patchData),
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.name && data.parentReference) {
                    editor.showSuccessMessage("File was updated successfully.");
                    editor.lastSelectedFile = data;
                }
            },
            error: function(xhr, status, err) {
                editor.showError("Unable to patch file metadata: " + err);
            }
        });
    },
    
    /***************** Share File ***************/
    shareFile: function () {
        var editor = this;
        // Make a request to Microsoft Graph to retrieve a default sharing URL for the file.
        if (!editor.lastSelectedFile || !editor.lastSelectedFile.id)
        {
            editor.showError("You need to save the file first before you can share it.");
            return;
        }

        editor.getSharingLinkWithAPI();
    },

    getSharingLinkWithAPI: function() {
        var editor = this;
        var driveItem = editor.lastSelectedFile;

        var url = editor.generateGraphUrl(driveItem, false, "/createLink");
        var requestBody = { type: "view" };

        $.ajax(url, {
            method: "POST",
            contentType: "application/json; charset=UTF-8",
            data: JSON.stringify(patchData),
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.link && data.link.webUrl) {
                    window.prompt("View-only sharing link", data.link.webUrl);
                } else {
                    editor.showError("Unable to retrieve a sharing link for this file.");
                }
            },
            error: function(xhr, status, err) {
                editor.showError("Unable to retrieve a sharing link for this file.");
            }
        });
    },

    // Used to generate the Microsoft Graph URL for a target item, with a few parameters
    // uploadIntoParentFolder: bool, indicates that we should be targeting the parent folder + filename instead of the item itself
    // targetContentStream: bool, indicates we should append /content to the item URL
    generateGraphUrl: function(driveItem, targetParentFolder, itemRelativeApiPath) {
        var url = this.microsoftGraphApiRoot;
        if (targetParentFolder)
        {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" +driveItem.parentReference.id + "/children/" + driveItem.name;
        } else {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" + driveItem.id;
        }

        if (itemRelativeApiPath) {
            url += itemRelativeApiPath;
        }
        return url;
    },


    /************ Create New File *************/

    // Clear the current editor buffer and reset any local state so we don't
    // overwrite an existing file by mistake.
    createNewFile: function () {
        this.lastSelectedFile = null;
		this.openFileID = "";
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
            case "share":
                element.onclick = function () { markdownEditor.shareFile(); return false; }
                break;
        }
    },

    // Handle displaying errors to the user
    showError: function (msg, e) {
        window.alert(msg);
    },

    showSuccessMessage: function(msg) {
        window.alert(msg);
    },

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
