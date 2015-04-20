//! Copyright (c) Microsoft Corporation. All rights reserved.
// WL.JS Version 5.5.8816.3000

(function() {
    if (!window.WL && !window.OneDrive) {
// Expose public OneDrive picker / saver SDK interface.
window.OneDrive = {};

OneDrive.Constants = {
    WebViewLink: ONEDRIVE_PARAM_LINKTYPE_WEBVIEW,
    DownloadLink: ONEDRIVE_PARAM_LINKTYPE_DOWNLOAD
};

OneDrive.open = function(options) {
    /// <summary>
    /// Opens file(s) with temporary access from the OneDrive picker.
    /// </summary>
    /// <param name="options" type="Object">
    /// Required. A JSON object that includes the following properties:
    /// &#10; multiSelect:  Optional. Boolean with value set to 'true' for multi-select (allow user to select multiple files)
    /// and 'false' (default) for single-select.
    /// &#10; linkType:  Optional. Type of link to the file, "webViewLink" (default) for a sharing link, and "downloadLink" for a
    /// link to download the file.
    /// &#10; success:  Required. A callback function invoked after the request has succeeded. Argument: [fileObject,...] where
    /// fileObject contains properties: fileName, link, linkType, size, and thumbnails.
    /// &#10; cancel:  Optional. A callback function invoked when the user cancels the file picker operation. Arguments: none
    /// </param>

    var clonedOptions = cloneObject(options);
    var app = new OneDriveApp(clonedOptions, IMETHOD_ONEDRIVE_OPEN);

    try {
        app.initialize();

        app.validateOpenParameters();
        app.executeOpenOperation();
    }
    catch (error) {
        app.processError(error, ERROR_DESC_OPERATION_UNHANDLED_EXCEPTION);
    }
};

OneDrive.save = function(options) {
    /// <summary>
    /// Saves a file to a folder in a user's OneDrive.
    /// </summary>
    /// <param name="options" type="Object">
    /// Required. A JSON object that includes the following properties:
    /// &#10; fileName:  Optional. The name for the file as it will appear in the user's OneDrive. If this is not supplied, the
    /// file name will either be inferred from the URL (e.g., http://mydomain.com/file_name.ext - file name = file_name.ext )or
    /// from the name attribute of the input element.
    /// &#10; file:  Required. Either a URL for the file to upload, or the id of a form file input element.
    /// &#10; success:  Optional. A callback function invoked after the request has succeeded. Arguments: none.
    /// &#10; progress: Optional. A callback function that is periodically invoked during the file upload. Arguments: float with
    /// value ranging from 0.0 to 100.0 representing the percentage of the file upload operation that has been completed. Will be
    /// called at least once and the final call will be passed 100 indicating completion.
    /// &#10; cancel:  Optional. A callback function invoked when the user cancels the folder picker operation. Arguments: none
    /// &#10; error:  Optional. A callback function invoked when an error occurs. Argument: errorString
    /// </param>

    var clonedOptions = cloneObject(options);
    var app = new OneDriveApp(clonedOptions, IMETHOD_ONEDRIVE_SAVE);

    try {
        app.initialize();

        app.validateSaveParameters();
        app.executeSaveOperation();
    }
    catch (error) {
        app.processErrorCallback(error, ERROR_DESC_OPERATION_UNHANDLED_EXCEPTION);
    }
};

OneDrive.createOpenButton = function (options) {
    /// <summary>
    /// Create a button that opens file(s) with temporary access from the OneDrive picker.
    /// </summary>
    /// <param name="options" type="Object">
    /// Required. A JSON object that includes the following properties:
    /// &#10; theme: Optional. The color theme of the button, either "blue" (default) or "white".
    /// &#10; multiSelect:  Optional. Boolean with value set to 'true' for multi-select (allow user to select multiple files)
    /// and 'false' (default) for single-select.
    /// &#10; linkType:  Optional. Type of link to the file, "webViewLink" (default) for a sharing link, and "downloadLink" for a
    /// link to download the file.
    /// &#10; success:  Required. A callback function invoked after the request has succeeded. Argument: [fileObject,...] where
    /// fileObject contains properties: fileName, link, linkType, size, and thumbnails.
    /// &#10; cancel:  Optional. A callback function invoked when the user cancels the file picker operation. Arguments: none
    /// </param>

    var clonedOptions = cloneObject(options);
    var app = new OneDriveApp(clonedOptions, IMETHOD_ONEDRIVE_CREATEBUTTON_OPEN);

    try {
        app.initialize();

        app.validateOpenParameters();
        app.validateButtonParameters();

        var openButton = app.createButtonElement();
        attachDOMEvent(openButton, DOM_EVENT_CLICK, function (event) {
            app.executeOpenOperation();
        });

        return openButton;
    }
    catch (error) {
        app.processError(error, ERROR_DESC_OPERATION_UNHANDLED_EXCEPTION);
        return null;
    }
};

OneDrive.createSaveButton = function(options) {
    /// <summary>
    /// Create a button that saves a file to a folder in a user's OneDrive.
    /// </summary>
    /// <param name="options" type="Object">
    /// Required. A JSON object that includes the following properties:
    /// &#10; theme: Optional. The color theme of the button, either "blue" (default) or "white".
    /// &#10; fileName:  Optional. The name for the file as it will appear in the user's OneDrive. If this is not supplied, the
    /// file name will either be inferred from the URL (e.g., http://mydomain.com/file_name.ext - file name = file_name.ext )or
    /// from the name attribute of the input element.
    /// &#10; file:  Required. Either a URL for the file to upload, or the id of a form file input element.
    /// &#10; success:  Optional. A callback function invoked after the request has succeeded. Arguments: none.
    /// &#10; progress: Optional. A callback function that is periodically invoked during the file upload. Arguments: float with
    /// value ranging from 0.0 to 100.0 representing the percentage of the file upload operation that has been completed. Will be
    /// called at least once and the final call will be passed 100 indicating completion.
    /// &#10; cancel:  Optional. A callback function invoked when the user cancels the folder picker operation. Arguments: none
    /// &#10; error:  Optional. A callback function invoked when an error occurs. Argument: errorString
    /// </param>

    var clonedOptions = cloneObject(options);
    var app = new OneDriveApp(clonedOptions, IMETHOD_ONEDRIVE_CREATEBUTTON_SAVE);

    try {
        app.initialize();

        app.validateSaveParameters();
        app.validateButtonParameters();

        var saveButton = app.createButtonElement();
        attachDOMEvent(saveButton, DOM_EVENT_CLICK, function (event) {
            app.executeSaveOperation();
        });

        return saveButton;
    }
    catch (error) {
        app.processErrorCallback(error, ERROR_DESC_OPERATION_UNHANDLED_EXCEPTION);
        return null;
    }
};

function OneDriveApp(options, method) {
    /// <summary>
    /// OneDriveApp constructor.
    /// </summary>

    var that = this;
    var internalApp = options[ONEDRIVE_PARAM_INTERNAL];

    that._internalApp = (WL.UnitTests && typeof (internalApp) === TYPE_OBJECT) ? internalApp : wl_app;

    that._options = options;
    that._method = method;
}

OneDriveApp.onloadInit = function() {
    /// <summary>
    /// Called when the script is loaded.
    /// </summary>

    checkDocumentReady(function() {
        // Turn links into OneDrive save buttons.
        var links = document.querySelectorAll(DOM_CLASS_ONEDRIVE_SAVEBUTTON);

        for (var i = 0; i < links.length; i++) {
            var link = links[i];

            // Modify the link HTML and style.
            OneDriveApp.createButtonElement(link, FILEDIALOG_PARAM_MODE_SAVE, UI_SIGNIN_THEME_BLUE);

            // Attach click event to execute save.
            var url = link.href;
            var app = new OneDriveApp({ file: url }, IMETHOD_ONEDRIVE_CREATEBUTTON_SAVE_FROMLINK);

            link.href = "#";
            attachDOMEvent(link, DOM_EVENT_CLICK, function (event) {
                app.initialize();
                app.executeSaveOperation();
                return false;
            });
        }
    });
};

OneDriveApp.createButtonElement = function(element, mode, theme) {
    /// <summary>
    /// Create the button HTML and styling.
    /// </summary>

    var properties = {};

    properties[FILEDIALOG_PARAM_MODE] = mode;
    properties[UI_PARAM_THEME] = theme;

    // Build button html.
    var buttonHtml = buildSkyDrivePickerControlInnerHtml(properties, true /* new picker */);
    element.innerHTML = buttonHtml.innerHTML;
    element.id = buttonHtml.buttonId;
    element.title = buttonHtml.buttonTitle;

    // Add button style.
    buildSkyDrivePickerControlStyle(properties, element);

    return element;
};

OneDriveApp.prototype = {
    initialize: function () {
        /// <summary>
        /// Initialize the SDK with the client_id supplied as the "client_id" attribute on the DOM element
        /// with id "onedrive_js".
        /// </summary>

        var that = this;
        if (that._internalApp._status !== APP_STATUS_INITIALIZED) {
            // Get the client_id from the DOM.
            var id = getClientIdFromDOM();
            if (!id) {
                throw new Error(ERROR_DESC_CLIENTID_MISSING.replace(METHOD, that._method));
            }

            that._internalApp.appInit({
                client_id: id,
                interface_method: that._method
            });

            logText("Picker SDK initialized with client_id: " + id, ONEDRIVE_PREFIX);
        }
    },

    executeOpenOperation: function () {
        /// <summary>
        /// Let the user open one or more files from their OneDrive and pass back the results to the
        /// third party app.
        /// </summary>

        var that = this;
        var internalApp = that._internalApp;
        var method = that._method;

        // Option parameters.
        var options = that._options;
        var linkType = options[ONEDRIVE_PARAM_LINKTYPE];
        var success = options[ONEDRIVE_PARAM_SUCCESS];
        var cancel = options[ONEDRIVE_PARAM_CANCEL];

        // Set file dialog properties.
        var fileDialogProperties = {
            mode: FILEDIALOG_PARAM_MODE_READ,
            resourceType: FILEDIALOG_PARAM_RESOURCETYPE_FILE,
            select: options[ONEDRIVE_PARAM_SELECT] ? FILEDIALOG_PARAM_SELECT_MULTI : FILEDIALOG_PARAM_SELECT_SINGLE,
            linkType: linkType,
            interface_method: method
        };

        // Open operation.
        internalApp.fileDialog(fileDialogProperties).then(
            // Success callback.
            function (fileDialogResponse) {
                // Filter API response.
                var apiResponse = fileDialogResponse.apiResponse;
                var isDownloadLinkType = linkType === ONEDRIVE_PARAM_LINKTYPE_DOWNLOAD;

                // Data returned from SDK to third party app success callback.
                var files = {
                    link: isDownloadLinkType ? null : apiResponse.webUrl,
                    values: []
                };

                var pickerFiles = isDownloadLinkType?
                    apiResponse.data :
                    (apiResponse.children && apiResponse.children.length > 0) ? apiResponse.children : [apiResponse];

                // Filter API response files.
                for (var i = 0; i < pickerFiles.length; i++) {
                    var file = pickerFiles[i];
                    var thumbnails = [];

                    // Filter file thumbnails.
                    var fileThumbnails = isDownloadLinkType ? file.images : file.thumbnails && file.thumbnails[0];
                    if (fileThumbnails) {
                        if (isDownloadLinkType) {
                            for (var j = 0; j < fileThumbnails.length; j++) {
                                thumbnails.push(fileThumbnails[j].source);
                            }
                        }
                        else {
                            for (var j = 0; j < VROOM_THUMBNAIL_SIZES.length; j++) {
                                thumbnails.push(fileThumbnails[VROOM_THUMBNAIL_SIZES[j]].url);
                            }
                        }
                    }

                    // File data returned to the app.
                    files.values.push({
                        fileName: file.name,
                        link: isDownloadLinkType ? file.source : file.webUrl,
                        linkType: linkType,
                        size: file.size,
                        thumbnails: thumbnails
                    });
                }

                success(files);
            },
            // Error callback.
            function (fileDialogError) {
                // If the error was access denied (user cancelled the file dialog), call the cancel callback,
                // otherwise process the error as normal.
                fileDialogError.error === ERROR_ACCESS_DENIED ?
                    invokeCallbackSynchronous(cancel) : that.processError(fileDialogError, ERROR_DESC_OPERATION_PICKER);
            }
        );
    },

    executeSaveOperation: function () {
        /// <summary>
        /// Let the user save a file to a folder in thier OneDrive.
        /// </summary>

        var that = this;
        var internalApp = that._internalApp;
        var options = that._options;
        var method = that._method;

        // Determine upload type and validate properties. Due to the URL detection,
        // the form element id can not start with "http://" or "https://", or it will
        // be interpreted as a URL.
        var file = options[ONEDRIVE_PARAM_FILE];
        var fileName = options[ONEDRIVE_PARAM_FILENAME];
        var uploadType = UPLOADTYPE_FORM;

        if (isPathFullUrl(file)) {
            // Upload from URL scenario.
            uploadType = UPLOADTYPE_URL;

            // If the file name is not supplied to the SDK, try to parse it from the URL.
            fileName = fileName || getFileNameFromUrl(file);
        }

        // Callbacks
        var success = options[ONEDRIVE_PARAM_SUCCESS];
        var progress = options[ONEDRIVE_PARAM_PROGRESS];
        var cancel = options[ONEDRIVE_PARAM_CANCEL];

        // Set file dialog properties.
        var fileDialogProperties = {
            mode: FILEDIALOG_PARAM_MODE_READWRITE,
            resourceType: FILEDIALOG_PARAM_RESOURCETYPE_FOLDER,
            select: FILEDIALOG_PARAM_SELECT_SINGLE,
            interface_method: method
        };

        // Save operation.
        internalApp.fileDialog(fileDialogProperties).then(
            // Success callback.
            function (fileDialogResponse) {
                var pickerResponse = fileDialogResponse.pickerResponse;
                var apiResponse = fileDialogResponse.apiResponse;

                var folderId = apiResponse.data && apiResponse.data[0].id;

                switch (uploadType) {
                    case UPLOADTYPE_URL:
                        // Folder ID comes in the format from LiveConnect: folder.{cid}.{cid}!{itemId}
                        // -> for vroom we only want {cid}!{itemId}. If the folder is the root, then the
                        // id will be folder.{cid} which is also invalid for vroom. In that case, we just
                        // want to use the string "root" for the folder ID.
                        var folderIdSplit = folderId.split(".");
                        var vroomFolderId = folderIdSplit.length > 2 ? folderIdSplit[2] : "root";

                        var accessToken = internalApp.getAccessTokenForApi();
                        var urlUploadProperties = {
                            path: "drives/" + pickerResponse.owner_cid + "/items/" + vroomFolderId + "/children",
                            method: HTTP_METHOD_POST,
                            use_vroom_api: true,
                            request_headers: [{ name: API_PARAM_PREFER, value: API_PARAM_RESPOND_ASYNC }, { name: API_PARAM_AUTH, value: "bearer " + accessToken }],
                            response_headers: [API_PARAM_LOCATION],
                            json_body: {
                                "@content.sourceUrl": file,
                                "name": fileName,
                                "file": {},
                                "@name.conflictBehavior": "overwrite"
                            },
                            interface_method: method
                        };

                        // Upload opeartion.
                        internalApp.api(urlUploadProperties).then(
                            // Success callback.
                            function (urlUploadResponse) {
                                // URL to poll for status on remote upload.
                                var location = urlUploadResponse[API_PARAM_LOCATION];

                                // Started remote upload.
                                (urlUploadResponse[API_PARAM_STATUS_HTTP] === API_STATUS_HTTP_ACCEPTED && !stringIsNullOrEmpty(location)) ?
                                    that.beginPolling(success, progress, location, accessToken) :
                                    that.processErrorCallback(urlUploadResponse, ERROR_DESC_OPERATION_UPLOAD);
                            },
                            // Error callback.
                            function (urlUploadError) {
                                that.processErrorCallback(urlUploadError, ERROR_DESC_OPERATION_UPLOAD);
                            }
                        );

                        break;
                    case UPLOADTYPE_FORM:
                        var formUploadProperties = {
                            path: folderId,
                            element: file,
                            overwrite: API_PARAM_OVERWRITE_RENAME,
                            file_name: fileName,
                            interface_method: method
                        };

                        // Upload operation.
                        internalApp.upload(formUploadProperties).then(
                            // Success callback.
                            function (formUploadResponse) {
                                invokeCallbackSynchronous(success);
                            },
                            // Error callback.
                            function (formUploadError) {
                                that.processErrorCallback(formUploadError, ERROR_DESC_OPERATION_UPLOAD);
                            },
                            // Progress callback.
                            function (uploadProgress) {
                                invokeCallbackSynchronous(progress, uploadProgress.progressPercentage);
                            }
                        );

                        break;
                    default:
                        throw new Error(ERROR_DESC_UPLOADTYPE_NOTIMPLEMENTED.replace(METHOD, method));
                }
            },
            // Error callback.
            function (fileDialogError) {
                // If the error was access denied (user cancelled the file dialog), call the cancel callback,
                // otherwise process the error as normal.
                fileDialogError.error === ERROR_ACCESS_DENIED ?
                    invokeCallbackSynchronous(cancel) : that.processError(fileDialogError, ERROR_DESC_OPERATION_PICKER);
            }
        );
    },

    beginPolling: function (success, progress, location, accessToken) {
        /// <summary>
        ///  Begin polling for remote upload completion.
        /// </summary>

        var that = this;
        var pollingInterval = POLLING_INTERVAL;
        var pollCount = POLLING_COUNTER;

        var progressApiProperties = {
            path: appendUrlParameters(location, { access_token: accessToken }),
            method: HTTP_METHOD_GET,
            use_vroom_api: true,
            response_headers: [API_PARAM_LOCATION],
            interface_method: that._method
        };

        var pollForProgress = function () {
            // Query for upload progress.
            that._internalApp.api(progressApiProperties).then(
                // Success callback.
                function(apiResponse) {
                    switch (apiResponse[API_PARAM_STATUS_HTTP]) {
                        case API_STATUS_HTTP_ACCEPTED:
                            invokeCallbackSynchronous(progress, apiResponse[API_PARAM_PERCENT_COMPLETE]);

                            // Exponential backoff on polling to prevent DOSing the service.
                            if (!pollCount--) {
                                pollingInterval *= 2;
                                pollCount = POLLING_COUNTER;
                            }

                            // Upload not yet completed, so continue polling.
                            delayInvoke(pollForProgress, pollingInterval);
                            break;
                        case API_STATUS_HTTP_OK:
                            // Call final progress update. This guarantees that we call
                            // the progress callback at least once.
                            invokeCallbackSynchronous(progress, 100.0);

                            invokeCallbackSynchronous(success);
                            break;
                        default:
                            that.processErrorCallback(apiResponse, ERROR_DESC_OPERATION_UPLOAD_POLLING);
                    }
                },
                // Error callback.
                function(apiError) {
                    that.processErrorCallback(apiError, ERROR_DESC_OPERATION_UPLOAD_POLLING);
                }
            );
        };

        // Begin polling.
        delayInvoke(pollForProgress, pollingInterval);
    },

    validateOpenParameters: function () {
        /// <summary>
        /// Validate the options passed in from the third party app
        /// for the open operation.
        /// </summary>

        validateProperties(
            this._options,
            [
                {
                    name: ONEDRIVE_PARAM_LINKTYPE,
                    type: TYPE_STRING,
                    optional: true,
                    defaultValue: ONEDRIVE_PARAM_LINKTYPE_WEBVIEW,
                    allowedValues: [ONEDRIVE_PARAM_LINKTYPE_DOWNLOAD, ONEDRIVE_PARAM_LINKTYPE_WEBVIEW]
                },
                { name: ONEDRIVE_PARAM_SELECT, type: TYPE_BOOLEAN, optional: true, defaultValue: false },
                { name: ONEDRIVE_PARAM_SUCCESS, type: TYPE_FUNCTION, optional: false },
                { name: ONEDRIVE_PARAM_CANCEL, type: TYPE_FUNCTION, optional: true }
            ],
            this._method
        );
    },

    validateSaveParameters: function () {
        /// <summary>
        /// Validate the options passed in from the third party app
        /// for the save operation.
        /// </summary>

        validateProperties(
            this._options,
            [
                { name: ONEDRIVE_PARAM_FILENAME, type: TYPE_STRING, optional: true },
                { name: ONEDRIVE_PARAM_FILE, type: TYPE_STRING, optional: false },
                { name: ONEDRIVE_PARAM_SUCCESS, type: TYPE_FUNCTION, optional: true },
                { name: ONEDRIVE_PARAM_PROGRESS, type: TYPE_FUNCTION, optional: true },
                { name: ONEDRIVE_PARAM_CANCEL, type: TYPE_FUNCTION, optional: true },
                { name: ONEDRIVE_PARAM_ERROR, type: TYPE_FUNCTION, optional: true }
            ],
            this._method
        );
    },

    validateButtonParameters: function () {
        /// <summary>
        /// Validate the options passed in from the third party app
        /// for the button UI styling.
        /// </summary>

        validateProperties(
            this._options,
            [
                {
                    name: ONEDRIVE_PARAM_THEME,
                    type: TYPE_STRING,
                    optional: true,
                    defaultValue: ONEDRIVE_PARAM_THEME_BLUE,
                    allowedValues: [ONEDRIVE_PARAM_THEME_BLUE, ONEDRIVE_PARAM_THEME_WHITE]
                }
            ],
            this._method
        );
    },

    createButtonElement: function () {
        /// <summary>
        /// Create the HTML button element for the open / save buttons.
        /// </summary>

        var buttonElement = document.createElement("button");
        var mode = this._method === IMETHOD_ONEDRIVE_CREATEBUTTON_OPEN ? FILEDIALOG_PARAM_MODE_OPEN : FILEDIALOG_PARAM_MODE_SAVE;
        var theme = this._options[ONEDRIVE_PARAM_THEME] === ONEDRIVE_PARAM_THEME_BLUE ? UI_SIGNIN_THEME_BLUE : UI_SIGNIN_THEME_WHITE;

        return OneDriveApp.createButtonElement(buttonElement, mode, theme);
    },

    processError: function (error, operation) {
        /// <summary>
        /// Log error to console.
        /// </summary>

        var errorMessage = stringFormat(
            ERROR_DESC_GENERAL,
            this._method,
            operation,
            JSON.stringify(error));

        logError(errorMessage, ONEDRIVE_PREFIX);

        return errorMessage;
    },

    processErrorCallback: function(error, operation) {
        /// <summary>
        /// Log error and execute error callback from the third party app.
        /// </summary>

        invokeCallbackSynchronous(this._options[ONEDRIVE_PARAM_ERROR], this.processError(error, operation));
    }
};


/**
 * API constants.
 */
var API_DOWNLOAD = "download",
    API_INTERFACE_METHOD = "interface_method",
    API_JSONP_CALLBACK_NAMESPACE_PREFIX = "WL.Internal.jsonp.",
    API_JSONP_URL_LIMIT = 2000,
    API_PARAM_AUTH = "Authorization",
    API_PARAM_BODY = "body",
    API_PARAM_BODY_JSON = "json_body",
    API_PARAM_CALLBACK = "callback",
    API_PARAM_CODE = "code",
    API_PARAM_CONTENTTYPE = "Content-Type",
    API_PARAM_ELEMENT = "element",
    API_PARAM_ERROR = "error",
    API_PARAM_ERROR_DESC = "error_description",
    API_PARAM_FILENAME = "file_name",
    API_PARAM_FILEINPUT = "file_input",
    API_PARAM_FILEOUTPUT = "file_output",
    API_PARAM_HEADERS_REQUEST = "request_headers",
    API_PARAM_HEADERS_RESPONSE = "response_headers",
    API_PARAM_LOCATION = "Location",
    API_PARAM_LOGGING = "logging",
    API_PARAM_MESSAGE = "message",
    API_PARAM_METHOD = "method",
    API_PARAM_OVERWRITE = "overwrite",
    API_PARAM_OVERWRITE_RENAME = "rename",
    API_PARAM_PATH = "path",
    API_PARAM_PERCENT_COMPLETE = "percentageComplete",
    API_PARAM_PREFER = "Prefer",
    API_PARAM_PRETTY = "pretty",
    API_PARAM_RESPOND_ASYNC = "respond-async",
    API_PARAM_RESULT = "result",
    API_PARAM_STATUS = "status",
    API_PARAM_STATUS_HTTP = "http_status",
    API_PARAM_SSLRESOURCE = "return_ssl_resources",
    API_PARAM_STREAMINPUT = "stream_input",
    API_PARAM_TRACING = "tracing",
    API_PARAM_VROOMAPI = "use_vroom_api",
    API_STATUS_ERROR = "error",
    API_STATUS_HTTP_ACCEPTED = 202,
    API_STATUS_HTTP_OK = 200,
    API_STATUS_HTTP_SERVERERROR = 500,
    API_STATUS_SUCCESS = "success",
    API_SUPPRESS_REDIRECTS = "suppress_redirects",
    API_SUPPRESS_RESPONSE_CODES = "suppress_response_codes",
    API_X_HTTP_LIVE_LIBRARY = "x_http_live_library";

/**
 * Application status values indicating whether the app has invoked WL.init(...).
 */
var APP_STATUS_NONE = 0,
    APP_STATUS_INITIALIZED = 1;

/**
 * Auth parameter key values used in multiple occassions: redirect_url parameter, auth cookie sub-key, auth response properties.
 */
var AK_ACCESS_TOKEN = "access_token",
    AK_APPSTATE = "appstate",
    AK_AUTH_KEY = "authentication_key"
    AK_AUTH_TOKEN = "authentication_token",
    AK_CLIENT_ID = "client_id",
    AK_DISPLAY = "display",
    AK_CODE = "code",
    AK_ERROR = "error",
    AK_ERROR_DESC = "error_description",
    AK_EXPIRES = "expires",
    AK_EXPIRES_IN = "expires_in",
    AK_ITEMID = "item_id",
    AK_LOCALE = "locale",
    AK_OWNER_CID = "owner_cid",
    AK_REDIRECT_URI = "redirect_uri",
    AK_RESPONSE = "response",
    AK_RESPONSE_TYPE = "response_type",
    AK_REQUEST_TS = "request_ts",
    AK_RESOURCEID = "resource_id",
    AK_SCOPE = "scope",
    AK_SESSION = "session",
    AK_SECURE_COOKIE = "secure_cookie",
    AK_STATE = "state",
    AK_STATUS = "status";

var AK_COOKIE_KEYS = [AK_ACCESS_TOKEN, AK_AUTH_TOKEN, AK_SCOPE, AK_EXPIRES_IN, AK_EXPIRES];

/**
 * Auth session status.
 */
var AS_CONNECTED = "connected", // The user is connected and signed in.
    AS_NOTCONNECTED = "notConnected", // The user is not connected.
    AS_UNCHECKED = "unchecked",   // We haven't checked the status yet.
    AS_UNKNOWN = "unknown",   // The user is unknown.
    AS_EXPIRING = "expiring", // The token will expire soon.
    AS_EXPIRED = "expired"; // The token is expired.

var BT_GROUP_UPLOAD = "live-sdk-upload",
    BT_GROUP_DOWNLOAD = "live-sdk-download";

/**
 * Compatible parameter keys(names).
 */
var CK_APPID = "appId",
    CK_CHANNELURL = "channelUrl";

/**
* Cookie names.
*/
var COOKIE_AUTH = "wl_auth",  // This cookie stores the Auth information.
    COOKIE_UPLOAD = "wl_upload";

/**
* Display types.
*/
var DISPLAY_PAGE = "page",
    DISPLAY_TOUCH = "touch",
    DISPLAY_NONE = "none";

var DOM_DISPLAY_NONE = "none";

/**
 * Event types.
 */
var EVENT_AUTH_LOGIN = "auth.login",
    EVENT_AUTH_LOGOUT = "auth.logout",
    EVENT_AUTH_SESSIONCHANGE = "auth.sessionChange",
    EVENT_AUTH_STATUSCHANGE = "auth.statusChange",
    EVENT_LOG = "wl.log";

/**
 * Error strings.
 */
var ERROR_ACCESS_DENIED = "access_denied",
    ERROR_CONNECTION_FAILED = "connection_failed",
    ERROR_COOKIE_ERROR = "invalid_cookie",
    ERROR_INVALID_REQUEST = "invalid_request",
    ERROR_REQ_CANCEL = "request_canceled",
    ERROR_REQUEST_FAILED = "request_failed",
    ERROR_TIMEDOUT = "timed_out",
    ERROR_UNKNOWN_USER = "unknown_user",
    ERROR_USER_CANCELED = "user_canceled",
    ERROR_DESC_ACCESS_DENIED = "METHOD: Failed to get the required user permission to perform this operation.",
    ERROR_DESC_BROWSER_ISSUE = "The request could not be completed due to browser issues.",
    ERROR_DESC_BROWSER_LIMIT = "The request could not be completed due to browser limitations.",
    ERROR_DESC_CANCEL = "METHOD: The operation has been canceled.",
    ERROR_DESC_COOKIE_INVALID = "The 'wl_auth' cookie is not valid.",
    ERROR_DESC_COOKIE_OVERWRITE = "The 'wl_auth' cookie has been modified incorrectly. Ensure that the redirect URI only modifies sub-keys for values received from the OAuth endpoint.",
    ERROR_DESC_COOKIE_MULTIPLEVALUE = "The 'wl_auth' cookie has multiple values. Ensure that the redirect URI specifies a cookie domain and path when setting cookies.",
    ERROR_DESC_DOM_INVALID = "METHOD: The input property 'PARAM' does not reference a valid DOM element.",
    ERROR_DESC_EXCEPTION = "METHOD: An exception was received for EVENT. Detail: MESSAGE",
    ERROR_DESC_ENSURE_INIT = "METHOD: The WL object must be initialized with WL.init() prior to invoking this method.",
    ERROR_DESC_FAIL_CONNECT = "A connection to the server could not be established.",
    ERROR_DESC_FAIL_IDENTIFY_USER = "The user could not be identified.",
    ERROR_DESC_FAIL_UPLOAD = "METHOD: Failed to get upload_location of the resource.",
    ERROR_DESC_LOGIN_CANCEL = "The pending login request has been canceled.",
    ERROR_DESC_LOGOUT_NOTSUPPORTED = "Logging out the user is not supported in current session because the user is logged in with a Microsoft account on this computer. To logout, the user may quit the app or log out from the computer.",
    ERROR_DESC_PARAM_INVALID = "METHOD: The input value for parameter/property 'PARAM' is not valid.",
    ERROR_DESC_PARAM_MISSING = "METHOD: The input parameter/property 'PARAM' must be included.",
    ERROR_DESC_PARAM_TYPE_INVALID = "METHOD: The type of the provided value for the input parameter/property 'PARAM' is not valid.",
    ERROR_DESC_PENDING_CALL_CONFLICT = "METHOD: There is a pending METHOD request, the current call will be ignored.",
    ERROR_DESC_PENDING_LOGIN_CONFLICT = ERROR_DESC_PENDING_CALL_CONFLICT.replace(/METHOD/g, IMETHOD_WL_LOGIN),
    ERROR_DESC_PENDING_FILEDIALOG_CONFLICT = ERROR_DESC_PENDING_CALL_CONFLICT.replace(/METHOD/g, IMETHOD_FILEDIALOG),
    ERROR_DESC_PENDING_UPLOAD_IGNORED = ERROR_DESC_PENDING_CALL_CONFLICT.replace(/METHOD/g, IMETHOD_WL_UPLOAD),
    ERROR_DESC_REDIRECTURI_MISSING = "METHOD: The input property 'redirect_uri' is required if the value of the 'response_type' property is 'code'.",
    ERROR_DESC_REDIRECTURI_INVALID_WWA = "WL.init: The redirect_uri value should be the same as the value of 'Redirect Domain' of your registered app. It must begin with 'http://' or 'https://'.",
    ERROR_DESC_UNSUPPORTED_API_CALL = "METHOD: The api call is not supported on this platform.",
    ERROR_DESC_UNSUPPORTED_RESPONSE_TYPE_CODE = "WL.init: The response_type value 'code' is not supported on this platform.",
    ERROR_DESC_URL_SSL = "METHOD: The input property 'redirect_uri' must use https: to match the scheme of the current page.",
    ERROR_TRACE_AUTH_TIMEOUT = "The auth request is timed out.",
    ERROR_TRACE_AUTH_CLOSE = "The popup is closed without receiving consent.";

/**
 * Flash initialization status.
 */
var FLASH_STATUS_NONE = 0,
    FLASH_STATUS_INITIALIZING = 1,
    FLASH_STATUS_INITIALIZED = 2,
    FLASH_STATUS_ERROR = 3;

/**
* Http method names.
*/
var HTTP_METHOD_GET = "GET",
    HTTP_METHOD_POST = "POST",
    HTTP_METHOD_PUT = "PUT",
    HTTP_METHOD_DELETE = "DELETE",
    HTTP_METHOD_COPY = "COPY",
    HTTP_METHOD_MOVE = "MOVE";

/**
 * WL methods.
 */
var IMETHOD_FILEDIALOG = "WL.fileDialog",
    IMETHOD_WL_API = "WL.api",
    IMETHOD_WL_DOWNLOAD = "WL.download",
    IMETHOD_WL_INIT = "WL.init",
    IMETHOD_WL_LOGIN = "WL.login",
    IMETHOD_WL_UI = "WL.ui",
    IMETHOD_WL_UPLOAD = "WL.upload";

/**
 * The maximum time in milliseconds to expire a getLoginStatus() request.
 */
var MAX_GETLOGINSTATUS_TIME = 30000;

var METHOD = "METHOD",
    PARAM = "PARAM";

/**
 * Promise class event names.
 */
var PROMISE_EVENT_ONSUCCESS = "onSuccess",
    PROMISE_EVENT_ONERROR = "onError",
    PROMISE_EVENT_ONPROGRESS = "onProgress";

/**
 * Used to detect the type of redirect.
 * redirect_type is the name of the parameter.
 * auth is a value of the parameter and is used for authorization redirects.
 * upload is a value of the parameter and is used for WL.upload Form POST redirects.
 */
var REDIRECT_TYPE = "redirect_type",
    REDIRECT_TYPE_AUTH = "auth",
    REDIRECT_TYPE_UPLOAD = "upload";

/**
 * Response type values.
 */
var RESPONSE_TYPE_CODE = "code",
    RESPONSE_TYPE_TOKEN = "token";

/**
* Url scheme values.
*/
var SCHEME_HTTPS = "https:",
    SCHEME_HTTP = "http:";

/**
 * Scope deliminators.
 */
var SCOPE_SIGNIN = "wl.signin",
    SCOPE_SKYDRIVE = "wl.skydrive",
    SCOPE_SKYDRIVE_UPDATE = "wl.skydrive_update",
    SCOPE_DELIMINATOR = /\s|,/;

/**
 * Type names.
 */
var TYPE_BOOLEAN = "boolean",
    TYPE_DOM = "dom",
    TYPE_FUNCTION = "function",
    TYPE_NUMBER = "number",
    TYPE_OBJECT = "object",
    TYPE_PROPERTIES = "properties",
    TYPE_STRING = "string",
    TYPE_STRINGORARRAY = "string_or_array",
    TYPE_UNDEFINED = "undefined",
    TYPE_URL = "url";

/**
 * UI constants.
 */
var UI_PARAM_NAME = "name",
    UI_PARAM_ELEMENT = "element",
    UI_PARAM_BRAND = "brand",
    UI_PARAM_TYPE = "type",
    UI_PARAM_SIGN_IN_TEXT = "sign_in_text",
    UI_PARAM_SIGN_OUT_TEXT = "sign_out_text",
    UI_PARAM_THEME = "theme",
    UI_PARAM_ONLOGGEDIN = "onloggedin",
    UI_PARAM_ONLOGGEDOUT = "onloggedout",
    UI_PARAM_ONERROR = "onerror";

var UI_BRAND_MESSENGER = "messenger",
    UI_BRAND_HOTMAIL = "hotmail",
    UI_BRAND_SKYDRIVE = "skydrive",
    UI_BRAND_WINDOWS = "windows",
    UI_BRAND_WINDOWSLIVE = "windowslive",
    UI_BRAND_NONE = "none";

var UI_SIGNIN = "signin",
    UI_SIGNIN_TYPE_SIGNIN = UI_SIGNIN,
    UI_SIGNIN_TYPE_LOGIN = "login",
    UI_SIGNIN_TYPE_CONNECT = "connect",
    UI_SIGNIN_TYPE_CUSTOM = "custom";

var UI_SIGNIN_THEME_BLUE = "blue",
    UI_SIGNIN_THEME_WHITE = "white";

/**
 * Names of parameters used in an upload request's state.
 */
var UPLOAD_STATE_ID = "id";

var WL_ONEDRIVE_API = "onedrive_api",
    WL_AUTH_SERVER = "auth_server",
    WL_APISERVICE_URI = "apiservice_uri",
    WL_SKYDRIVE_URI = "skydrive_uri",
    WL_SDK_ROOT = "sdk_root",
    WL_TRACE = "wl_trace";

window.WL = {

    getSession: function () {
        /// <summary>
        /// A synchronous function that gets the current session object, if it exists.
        /// </summary>
        /// <returns type="Object" >The current session object.</returns>

        try {
            return wl_app.getSession();
        }
        catch (e) {
            logError(e.message);
        }
    },

    getLoginStatus: function (callback, force) {
        /// <summary>
        /// Returns the status of the current user. If the user is signed in and
        /// connected to your application, it returns the session object.
        /// This is an asynchronous function that returns the user's status by contacting the Windows Live
        /// OAuth server. If the user status is already known, the library may return what is cached.
        /// However, you can force the library to retrieve up-to-date status by setting the "force"
        /// parameter to true. This is an async method that returns a Promise object that allows you to
        /// attach events to handle succeeded and failed situations.
        /// </summary>
        /// <param name="callback" type="Function">Optional. The callback function that is invoked when the user's login status is retrieved.</param>
        /// <param name="force" type="Boolean">Optional. If set to false (default), the function may return an existing user status, if it exists.
        /// Otherwise, if set to true, the function contacts the server to determine the user's status.</param>
        /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
        /// situations.</returns>

        try {
            return wl_app.getLoginStatus(
            {
                callback: findArgumentByType(arguments, TYPE_FUNCTION, 2),
                internal: false
            },
            findArgumentByType(arguments, TYPE_BOOLEAN, 2));
        }
        catch (e) {
            return handleAsyncCallingError("WL.getLoginStatus", e);
        }
    },

    logout: function (callback) {
        /// <summary>
        /// Logs the user out of Windows Live and clears any user state that is maintained
        /// by the JavaScript library, such as cookies. This is an async method that returns a Promise object that
        /// allows you to attach events to handle succeeded and failed situations.
        /// </summary>
        /// <param name="callback" type="Function">Optional. Specifies a callback function that is invoked when logout is complete.</param>
        /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
        /// situations.</returns>

        try {
            validateParams(callback, expectedCallback_Optional, "WL.logout");
            return wl_app.logout({ callback: callback });
        }
        catch (e) {
            return handleAsyncCallingError("WL.logout", e);
        }
    },

    canLogout: function () {
        /// <summary>
        /// Returns if the app can log the user out.
        /// </summary>
        /// <returns type="boolean" >Whether the app can logout.</returns>

        return wl_app.canLogout();
    },

    api: function (properties, callback) {
        /// <summary>
        /// Makes a call to the Windows Live REST API. This is an async method that returns a Promise object that allows you to
        /// attach events to handle succeeded and failed situations.
        /// </summary>
        /// <param name="properties" type="Object">Required. A JSON object containing the properties for making the API call:
        /// &#10; path: Required. The path to the REST API object.
        /// &#10; method: The HTTP method. Supported values include "GET" (default), "PUT", "POST", "DELETE", "MOVE", and "COPY".
        /// &#10; body: A JSON object containing all necessary properties for making the REST API request.
        /// </param>
        /// <param name="callback" type="Function">Required. A callback function that is invoked when the REST API call is complete.</param>
        /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
        /// situations.</returns>

        try {
            var args = normalizeApiArguments(arguments);

            // Validate parameters
            validateProperties(args,
                [{ name: API_PARAM_PATH, type: TYPE_STRING, optional: false },
                { name: API_PARAM_METHOD, type: TYPE_STRING, optional: true },
                    expectedCallback_Optional],
                IMETHOD_WL_API);

            return wl_app.api(args);
        }
        catch (e) {
            return handleAsyncCallingError(IMETHOD_WL_API, e);
        }
    }
};

var allowedEvents = [EVENT_AUTH_LOGIN, EVENT_AUTH_LOGOUT, EVENT_AUTH_SESSIONCHANGE, EVENT_AUTH_STATUSCHANGE, EVENT_LOG];
WL.Event = {

    subscribe: function (event, callback) {
        /// <summary>
        /// Adds a handler to an event.
        /// </summary>
        /// <param name="event" type="String">Required. The name of the event to add a handler to.
        /// Available events are: "auth.login", "auth.logout", "auth.sessionChange",
        /// "auth.statusChange", and "wl.log".</param>
        /// <param name="callback" type="Function">Required. The event handler function to be added to the event.</param>

        try {
            // Validate parameters
            validateParams(
                [event, callback],
                [{ name: "event", type: TYPE_STRING, allowedValues: allowedEvents, caseSensitive: true, optional: false },
                    expectedCallback_Required],
                "WL.Event.subscribe");

            wl_event.subscribe(event, callback);
        }
        catch (e) {
            logError(e.message);
        }
    },

    unsubscribe: function (event, callback) {
        /// <summary>
        /// Removes a handler from an event.
        /// </summary>
        /// <param name="event" type="String">Required. The name of the event from which to remove a handler.</param>
        /// <param name="callback" type="Function">Optional. Removes the callback from the event. If this parameter is omitted, all
        /// callback functions registered to the event are removed.</param>

        try {
            // Validate parameters
            validateParams([event, callback],
                [{ name: "event", type: TYPE_STRING, allowedValues: allowedEvents, caseSensitive: true, optional: false },
                expectedCallback_Optional],
                "WL.Event.unsubscribe");
            wl_event.unsubscribe(event, callback);
        }
        catch (e) {
            logError(e.message);
        }
    }
};

WL.Internal = {};

var wl_event = {
    subscribe: function (event, callback) {
        trace("Subscribe " + event);

        var handlers = wl_event.getHandlers(event);
        handlers.push(callback);
    },

    unsubscribe: function (event, callback) {
        trace("Unsubscribe " + event);

        var oldHandlers = wl_event.getHandlers(event);
        var newHandlers = [];

        // Constructs a new list with one callback removed.
        // If callback is not available, we remove all.
        if (callback != null) {
            var found = false;
            for (var i = 0; i < oldHandlers.length; i++) {
                if (found || oldHandlers[i] != callback) {
                    newHandlers.push(oldHandlers[i]);
                } else {
                    found = true;
                }
            }
        }

        wl_event._eHandlers[event] = newHandlers;
    },

    getHandlers: function (event) {

        if (!wl_event._eHandlers) {
            wl_event._eHandlers = {};
        }

        var eHandlers = wl_event._eHandlers[event];

        if (eHandlers == null) {
            wl_event._eHandlers[event] = eHandlers = [];
        }

        return eHandlers;
    },

    notify: function (event, data) {
        trace("Notify " + event)

        var handlers = wl_event.getHandlers(event);

        for (var i = 0; i < handlers.length; i++) {
            handlers[i](data);
        }
    }
};

/**
 * The wl_app type encapsulates the implementation of all inteface methods.
 */
var wl_app = { _status: APP_STATUS_NONE, _statusRequests: [], _rpsAuth: false };

/**
 * The implementation of WL.init().
 */
wl_app.appInit = function (properties) {
    var method = properties[API_INTERFACE_METHOD] || IMETHOD_WL_INIT;

    // If app has already invoked WL.init(), ignore this call.
    if (wl_app._status == APP_STATUS_INITIALIZED) {
        var status = wl_app._session.getNormalStatus();
        return createCompletePromise(method, true/*succeeded*/, properties.callback, status);
    }

    var sdkRoot = WL[WL_SDK_ROOT];
    if (sdkRoot) {
        if (sdkRoot.charAt(sdkRoot.length - 1) !== "/") {
            sdkRoot += "/";
        }

        wl_app[WL_SDK_ROOT] = sdkRoot;
    }

    var logging = properties[API_PARAM_LOGGING];
    if (logging === false) {
        wl_app._logEnabled = logging;
    }

    if (wl_app.testInit) {
        wl_app.testInit(properties);
    }

    wl_app._status = APP_STATUS_INITIALIZED;
    return appInitPlatformSpecific(properties);
};

/**
 * This is the very first method invoked after the script is loaded.
 */
wl_app.onloadInit = function () {
    detectBrowsers();
    handlePageLoad();
};

function ensureAppInited(method) {
    if (wl_app._status === APP_STATUS_NONE) {
        throw new Error(ERROR_DESC_ENSURE_INIT.replace(METHOD, method));
    }
}

function getCoreApp() {
    return WL.Internal.tApp || wl_app;
}

wl_app.api = function (properties) {

    ensureAppInited(IMETHOD_WL_API);

    var body = properties[API_PARAM_BODY];
    if (body) {
        properties = cloneObject(flattenApiBody(body), properties);
        delete properties[API_PARAM_BODY];
    }

    var method = properties[API_PARAM_METHOD];
    properties[API_PARAM_METHOD] = ((method != null) ? stringTrim(method) : HTTP_METHOD_GET).toUpperCase();

    return new APIRequest(properties).execute();
};

wl_app.getAccessTokenForApi = function () {
    var token = null;
    if (!wl_app._rpsAuth) {
        var status = getCoreApp()._session.getStatus();

        if (status.status === AS_EXPIRING || status.status === AS_CONNECTED) {
            token = status.session[AK_ACCESS_TOKEN];
        }
    }

    return token;
}

var generateApiRequestId = function () {
    var ticketNumber = wl_app.api.lastId,
        id;
    ticketNumber = (ticketNumber === undefined) ? 1 : ticketNumber + 1;
    id = "WLAPI_REQ_" + ticketNumber + "_" + (new Date().getTime());
    wl_app.api.lastId = ticketNumber;

    return id;
};

var APIRequest = function (properties) {
    var request = this;
    request._properties = properties;
    request._completed = false;
    request._id = generateApiRequestId();
    properties[API_PARAM_PRETTY] = false;
    properties[API_PARAM_SSLRESOURCE] = wl_app._isHttps;
    properties[API_X_HTTP_LIVE_LIBRARY] = wl_app[API_X_HTTP_LIVE_LIBRARY];

    var path = properties[API_PARAM_PATH],
        useVroomApi = properties[API_PARAM_VROOMAPI];
    request._url = isPathFullUrl(path) ?
        path :
        getApiServiceUrl(useVroomApi) + (path.charAt(0) === "/" ? path.substring(1) : path);
    request._promise = new Promise(IMETHOD_WL_API, null, null);
};

APIRequest.prototype = {
    execute: function () {
        executeApiRequest(this);
        return this._promise;
    },

    onCompleted: function (response) {
        if (this._completed) {
            return;
        }

        this._completed = true;
        invokeCallback(this._properties.callback, response, true/*synchronous*/);

        if (response[AK_ERROR]) {
            this._promise[PROMISE_EVENT_ONERROR](response);
        }
        else {
            this._promise[PROMISE_EVENT_ONSUCCESS](response);
        }
    }
};

function processXDRResponse(request, status, responseText, errorDescription, responseHeaders) {

    responseText = responseText ? stringTrim(responseText) : "";

    // Deserialize response string.
    var response = (responseText !== "") ? deserializeJSON(responseText) : null;
    if (response === null) {
        response = {};
        if (Math.floor(status / 100) !== 2) {
            response[API_PARAM_ERROR] = createErrorObject(status, errorDescription);
        }
    }

    // Add requested response headers to response.
    if (responseHeaders) {
        responseHeaders.forEach(function(header)
        {
            response[header.name] = header.value;
        });
    }

    // Add request status.
    response[API_PARAM_STATUS_HTTP] = status;

    request.onCompleted(response);
}

function createErrorObject(status, errorDescription) {
    var errorObj = {};
    errorObj[API_PARAM_CODE] = ERROR_REQUEST_FAILED;
    errorObj[API_PARAM_MESSAGE] = (errorDescription || ERROR_DESC_FAIL_CONNECT);

    return errorObj;
}

function flattenApiBody(body) {
    // If the WL.api body parameter is a nested JSON object, we convert it into a flattened dictionary that has one layer
    // and maps each leaf node value on the original JSON tree hierarchy with a key value joining each sub key on the
    // path with a dot character. E.g. { contact { name: "Lin" } } will be converted into: {"contact.name" : "Lin"}
    // If array is used in the structure, the array index value will be part of the key.
    // E.g. { employmentHistory: [ { employer: "Microsoft", period: "2007-2011"} ] } will output the following entries:
    //  {"employmentHistory.0.employer" : "Microsoft", "employmentHistory.0.period" : "2007-2011" }

    var dict = {};
    for (var key in body) {
        var value = body[key],
            type = typeof(value);

        if (value instanceof Array) {
            for (var i = 0; i < value.length; i++) {
                // Note: we shouldn't have immediate nested array cases.
                var elementValue = value[i],
                    elementValueType = typeof (elementValue);
                if (type == TYPE_OBJECT && !(value instanceof Date)) {
                    var elementDict = flattenApiBody(elementValue);
                    for (var elementSubKey in elementDict) {
                        dict[key + "." + i + "." + elementSubKey] = elementDict[elementSubKey];
                    }
                }
                else {
                    dict[key + "." + i] = elementValue;
                }
            }
        }
        else if (type == TYPE_OBJECT && !(value instanceof Date)) {
            var vDic = flattenApiBody(value);
            for (var subKey in vDic) {
                dict[key + "." + subKey] = vDic[subKey];
            }
        }
        else {
            dict[key] = value;
        }
    }

    return dict;
}

function sendAPIRequestViaXHR(request) {
    if (!canDoXHR()) {
        return false;
    }

    var xdrParams = prepareXDRRequest(request);
    var xdr = new XMLHttpRequest();

    xdr.open(xdrParams.method, xdrParams.url, true);

    // Set request headers.
    xdrParams.requestHeaders.forEach(function (header) {
        xdr.setRequestHeader(header.name, header.value);
    });

    xdr.onreadystatechange = function () {
        if (xdr.readyState === 4) {
            // Pass back the requested response headers.
            var responseHeaders = [];
            xdrParams.responseHeaders.forEach(function (headerName) {
                responseHeaders.push({ name: headerName, value: xdr.getResponseHeader(headerName) });
            });

            processXDRResponse(request, xdr.status, xdr.responseText, null, responseHeaders);
        }
    };

    xdr.send(xdrParams.body);

    return true;
}

function prepareXDRRequest(request) {
    var url = request._url;
    var method = request._properties[API_PARAM_METHOD];
    var token = wl_app.getAccessTokenForApi();
    var requestBody = null;

    var params = cloneObjectExcept(
        request._properties,
        null,
        [API_PARAM_CALLBACK, API_PARAM_PATH, API_PARAM_METHOD]);
    var requestHeaders = params[API_PARAM_HEADERS_REQUEST] || [];
    var responseHeaders = params[API_PARAM_HEADERS_RESPONSE] || [];
    var jsonBody = params[API_PARAM_BODY_JSON];
    var useVroomApi = params[API_PARAM_VROOMAPI];

    params[API_SUPPRESS_REDIRECTS] = "true";

    if (token) {
        params[AK_ACCESS_TOKEN] = token;
    }

    if (!useVroomApi) {
        appendUrlParameters(url, { ts: (new Date().getTime()), method: method });
    }

    switch (method) {
        case HTTP_METHOD_GET:
        case HTTP_METHOD_DELETE:
            if (!useVroomApi) {
                appendUrlParameters(url, params);
            }
            break;
        default:
            requestBody = useVroomApi ? JSON.stringify(jsonBody) : serializeParameters(params);
            requestHeaders.push({ name: API_PARAM_CONTENTTYPE, value: "application/" + (jsonBody ? "json" : "x-www-form-urlencoded") });
    }

    return {
        url: url,
        method: method,
        requestHeaders: requestHeaders,
        responseHeaders: responseHeaders,
        body: requestBody
    };
}

// Common shared wl.download method code.
// See wl.app.download.wwa.js or wl.app.download.web.js for platform specific
// details.

wl_app.download = function (properties) {
    validateDownloadProperties(properties);

    ensureAppInited(IMETHOD_WL_DOWNLOAD);

    return new DownloadOperation(properties).execute();
};

function buildFilePathUrlString(path, extra_params) {
    var params = extra_params || {},
        baseUrl = getApiServiceUrl();

    if (!isPathFullUrl(path)) {
        path = baseUrl + (path.charAt(0) === "/" ? path.substring(1) : path);
    }

    var token = wl_app.getAccessTokenForApi();
    if (token) {
        params[AK_ACCESS_TOKEN] = token;
    }

    params[API_X_HTTP_LIVE_LIBRARY] = wl_app[API_X_HTTP_LIVE_LIBRARY];

    return appendUrlParameters(path, params);
}

function handleDownloadErrorResponse(errorMessage, op) {
    op.downloadComplete(false, createErrorResponse(ERROR_REQUEST_FAILED, IMETHOD_WL_DOWNLOAD + ": " + errorMessage));
}

var DOWNLOAD_OPSTATE_NOTSTARTED = "notStarted",
    DOWNLOAD_OPSTATE_READY = "ready",
    DOWNLOAD_OPSTATE_DOWNLOADCOMPLETED = "downloadCompleted",
    DOWNLOAD_OPSTATE_DOWNLOADFAILED = "downloadFailed",
    DOWNLOAD_OPSTATE_CANCELED = "canceled",
    DOWNLOAD_OPSTATE_COMPLETED = "completed";

// DownloadOperation type.
function DownloadOperation(properties) {
    this._properties = properties;
    this._status = DOWNLOAD_OPSTATE_NOTSTARTED;
}

DownloadOperation.prototype = {
    execute: function () {
        this._promise = new Promise(IMETHOD_WL_DOWNLOAD, this, null);
        this._process();
        return this._promise;
    },

    cancel: function () {
        this._status = DOWNLOAD_OPSTATE_CANCELED;
        if (this._cancel) {
            try {
                this._cancel();
            }
            catch (ex) {
            }
        }
        else {
            this._result = createErrorResponse(ERROR_REQ_CANCEL, ERROR_DESC_CANCEL.replace(METHOD, IMETHOD_WL_DOWNLOAD));
            this._process();
        }
    },

    downloadComplete: function (succeeded, result) {
        var op = this;
        op._result = result;
        op._status = succeeded ? DOWNLOAD_OPSTATE_DOWNLOADCOMPLETED : DOWNLOAD_OPSTATE_DOWNLOADFAILED;
        op._process();
    },

    downloadProgress: function (progress) {
        this._promise[PROMISE_EVENT_ONPROGRESS](progress);
    },

    _process: function () {
        switch (this._status) {
            case DOWNLOAD_OPSTATE_NOTSTARTED:
                this._start();
                break;
            case DOWNLOAD_OPSTATE_READY:
                this._download();
                break;
            case DOWNLOAD_OPSTATE_DOWNLOADCOMPLETED:
            case DOWNLOAD_OPSTATE_DOWNLOADFAILED:
            case DOWNLOAD_OPSTATE_CANCELED:
                this._complete();
                break;
        }
    },

    _start: function () {
        var op = this;

        wl_app.getLoginStatus({
            internal: true,
            callback: function () {
                    op._status = DOWNLOAD_OPSTATE_READY;
                    op._process()
                }
        });
    },

    _download: function () {
        var op = this;

        // executeDownload must
        // be defined in both
        // wl.app.download.web.js and wl.app.download.wwa.js
        executeDownload(op);
    },

    _complete: function () {
        var op = this,
            result = op._result,
            promiseEvent = (op._status === DOWNLOAD_OPSTATE_DOWNLOADCOMPLETED) ?
                           PROMISE_EVENT_ONSUCCESS :
                           PROMISE_EVENT_ONERROR;

        op._status = DOWNLOAD_OPSTATE_COMPLETED;

        var callback = op._properties[API_PARAM_CALLBACK];
        if (callback) {
            callback(result);
        }

        op._promise[promiseEvent](result);
    }
};

/**
 * The implementation of WL.login() method.
 */
wl_app.login = function (properties, internal, isExternalConsentRequest) {

    ensureAppInited(IMETHOD_WL_LOGIN);

    normalizeLoginScope(properties);

    if (!handlePendingLogin(internal)) {
        return createCompletePromise(IMETHOD_WL_LOGIN,
                                     false/*succeeded*/,
                                     null,
                                     createErrorResponse(ERROR_REQUEST_FAILED, ERROR_DESC_PENDING_LOGIN_CONFLICT));
    }

    var response = wl_app._session.tryGetResponse(properties.normalizedScope);
    if (response != null) {
        return createCompletePromise(IMETHOD_WL_LOGIN, true/*succeeded*/, properties.callback, response);
    }

    wl_app._pendingLogin = createLoginRequest(properties, onAuthRequestCompleted, isExternalConsentRequest);
    return wl_app._pendingLogin.execute();
}

function onAuthRequestCompleted(requestProperties, response) {
    wl_app._pendingLogin = null;

    var error = response[AK_ERROR];
    if (error) {
        log(IMETHOD_WL_LOGIN + ": " + response[AK_ERROR_DESC]);
    }
    else {
        invokeCallback(requestProperties.callback, response, true/*synchronous*/);
    }
}

function normalizeScopeValue(scopeValue) {

    var scope = scopeValue || "";
    if (scope instanceof Array) {
        scope = scope.join(" ");
    }

    return stringTrim(scope);
}

/**
 * The implementation of WL.getSession() method.
 */
wl_app.getSession = function () {
    ensureAppInited("WL.getSession");
    return wl_app._session.getStatus()[AK_SESSION];
};

/**
 * The implementation of WL.getLoginStatus() method.
 */
wl_app.getLoginStatus = function (properties, force) {

    ensureAppInited("WL.getLoginStatus");

    properties = properties || {};

    if (!force) {
        var response = wl_app._session.tryGetResponse();
        if (response) {
            return createCompletePromise("WL.getLoginStatus", true/*succeeded*/, properties.callback, response);
        }
    }

    trace("wl_app:getLoginStatus");

    var pendingQueue = wl_app._statusRequests,
        request = null;

    if (!wl_app._pendingStatusRequest) {
        request = createLoginStatusRequest(properties, onGetLoginStatusCompleted);
        wl_app._pendingStatusRequest = request;
    }

    pendingQueue.push(properties);

    if (request != null) {
        request.execute();
    }

    return wl_app._pendingStatusRequest._promise;
}

function onGetLoginStatusCompleted(requestProperties, response) {
    var pendingQueue = wl_app._statusRequests;
    wl_app._pendingStatusRequest = null;

    trace("wl_app:onGetLoginStatusCompleted");

    var error = response[AK_ERROR],
        hasAppRequest = false;

    while (pendingQueue.length > 0) {
        var reqProperties = pendingQueue.shift(),
            responseForCallback = cloneObject(response);
        if (!error || reqProperties.internal) {
            invokeCallback(reqProperties.callback, responseForCallback, true/*synchronous*/);
        }

        if (!reqProperties.internal) {
            hasAppRequest = true;
        }
    }

    if (error) {
        if (hasAppRequest && error !== ERROR_TIMEDOUT) {
            log("WL.getLoginStatus: " + response[AK_ERROR_DESC]);
        }
        else {
            trace("wl_app-onGetLoginStatusCompleted: " + response[AK_ERROR_DESC]);
        }
    }
}


/**
 * The implementation of WL.logout() method.
 */
wl_app.logout = function (properties) {
    var methodName = "WL.logout";
    ensureAppInited(methodName);

    var promise = new Promise(methodName, null, null),
        logoutCallback = function (error) {
            // Ensure that the callback is asynchronous.
            delayInvoke(function () {
                var resp,
                    event = PROMISE_EVENT_ONSUCCESS;
                if (error) {
                    logError(error.message);
                    event = PROMISE_EVENT_ONERROR;
                    resp = createExceptionResponse(methodName, methodName, error);
                }
                else {
                    resp = wl_app._session.getNormalStatus();
                }

                invokeCallback(properties.callback, resp, false/*synchronous*/);
                promise[event](resp);
            });
        },
        logout = function () {
            var authSession = wl_app._session;
            if (authSession.isSignedIn()) {
                if (wl_app.canLogout()) {
                    authSession.updateStatus(AS_UNKNOWN);
                    logoutWindowsLive(logoutCallback);
                }
                else {
                    logoutCallback(new Error(ERROR_DESC_LOGOUT_NOTSUPPORTED));
                }
            } else {
                logoutCallback();
            }
        };

    if (wl_app._pendingStatusRequest != null) {
        // If we have a pending getLoginStatus request, let's wait for the status call to complete
        // before invoking logout.
        wl_app.getLoginStatus({ internal: true, callback: logout }, false/*force*/);
    }
    else{
        logout();
    }

    return promise;
};

// wl.app.upload.js
// Common WL.upload methods.

wl_app.upload = function (properties) {
    var method = properties[API_INTERFACE_METHOD];
    ensureAppInited(method);

    validateProperties(
        properties,
        [{ name: API_PARAM_PATH, type: TYPE_STRING, optional: false },
         expectedCallback_Optional],
        method);

    validateUploadOverwriteProperty(properties);

    normalizeUploadFileName(properties);

    return new UploadOperation(properties).execute();
}

function normalizeUploadFileName(properties) {
    var file = properties[API_PARAM_FILEINPUT],
        fileName = properties[API_PARAM_FILENAME];
    if (file) {
        properties[API_PARAM_FILENAME] = fileName || file.name;
    }
}

function buildUploadToFolderUrlString(location, file_name, overwrite) {
    var queryStringIndex = location.indexOf("?");
    var hasQueryString = queryStringIndex !== -1;
    var queryString = "";
    // Since we might be appending the fileName on to the URI, we want
    // to break the string apart at the query string.
    if (hasQueryString) {
        queryString = location.substring(queryStringIndex + 1);
        location = location.substring(0, queryStringIndex);
    }

    var hasFileName = typeof(file_name) !== TYPE_UNDEFINED;
    var hasTrailingSlash = location.charAt(location.length-1) === "/";
    if (hasFileName && !hasTrailingSlash) {
        location += "/";
    }

    var path = location,
        params = {};

    // the file_name for Form multipart uploads are undefined, and should NOT be included.
    if (hasFileName) {
        path += encodeURIComponent(file_name);
    }

    // We only apply "overwrite" parameter when upload to a folder.
    if (overwrite === API_PARAM_OVERWRITE_RENAME) {
        // the API Service's rename value is choosenewname
        params[API_PARAM_OVERWRITE] = "choosenewname";
    } else {  // if overwrite is true or false
        params[API_PARAM_OVERWRITE] = overwrite;
    }

    // If we broke apart the string, let's put it back together.
    if (hasQueryString) {
        path = appendQueryString(path, queryString);
    }

    return buildUploadFileUrlString(path, params);
}

function isFilePath(path) {
    return /^(file|\/file)/.test(path.toLowerCase());
}

function buildUploadFileUrlString(path, params) {
    params = params || {};
    params[API_SUPPRESS_RESPONSE_CODES] = "true";

    return buildFilePathUrlString(path, params);
}

function validateUploadOverwriteProperty(properties) {
    // Overwrite is an optional parameter, that can be a boolean or a string with values:
    // "true", "false", or "rename".
    // "rename" will rename the file with a suffix if the file already exists. (e.g.,
    // uploading foo.txt when foo.txt already exists will be renamed to foo(1).txt).
    if (API_PARAM_OVERWRITE in properties) {
        var interfaceMethod = properties[API_INTERFACE_METHOD],
            overwrite = properties[API_PARAM_OVERWRITE],
            type = typeof(overwrite),
            isBoolean = type === TYPE_BOOLEAN,
            isString = type === TYPE_STRING;
        if (!(isBoolean || isString)) {
            throw createParamTypeError(API_PARAM_OVERWRITE, interfaceMethod);
        }

        if (isString) {
            // the overwrite parameter can be "true", "false", or "rename".
            var hasValidValue = (/^(true|false|rename)$/i).test(overwrite);
            if (!hasValidValue) {
                throw createInvalidParamValue(API_PARAM_OVERWRITE, interfaceMethod);
            }
        }
    } else {
        // if it does not have overwrite, the default value for it is false.
        // This is to be consistent with other platforms.
        properties[API_PARAM_OVERWRITE] = false;
    }
}

var UPLOAD_OPSTATE_NOTSTARTED = 0,
    UPLOAD_OPSTATE_AUTHREADY = 1,
    UPLOAD_OPSTATE_UPLOADREADY = 2,
    UPLOAD_OPSTATE_UPLOADCOMPLETED = 3,
    UPLOAD_OPSTATE_UPLOADFAILED = 4,
    UPLOAD_OPSTATE_CANCELED = 5,
    UPLOAD_OPSTATE_COMPLETED = 6;

function UploadOperation(props) {
    this._props = props;
    this._status = UPLOAD_OPSTATE_NOTSTARTED;
}

UploadOperation.prototype = {
    execute: function () {
        var self = this;
        self._strategy = self._getStrategy(self._props);
        self._promise = new Promise(self._props[API_INTERFACE_METHOD], self, null);
        self._process();
        return self._promise;
    },

    cancel: function () {
        var self = this;
        self._status = UPLOAD_OPSTATE_CANCELED;

        if (self._cancel) {
            try {
                self._cancel();
            }
            catch (ex) {
            }
        }
        else {
            var errorDescription = ERROR_DESC_CANCEL.replace(METHOD, self._props[API_INTERFACE_METHOD]);
            self._result = createErrorResponse(ERROR_REQ_CANCEL, errorDescription);
            self._process();
        }
    },

    uploadProgress: function (progress) {
        this._promise[PROMISE_EVENT_ONPROGRESS](progress);
    },

    uploadComplete: function (succeeded, result) {
        var self = this;
        self._result = result;
        self._status = succeeded ? UPLOAD_OPSTATE_UPLOADCOMPLETED : UPLOAD_OPSTATE_UPLOADFAILED;
        self._process();
    },

    onErr: function (errorMessage) {
        var errorDescription = this._props[API_INTERFACE_METHOD] + ":" + errorMessage,
            errorResponse = createErrorResponse(ERROR_REQUEST_FAILED, errorDescription);
        this.uploadComplete(false, errorResponse);
    },

    onResp: function (responseText) {
        responseText = responseText ? stringTrim(responseText) : "";
        var response = (responseText !== "") ? deserializeJSON(responseText) : null;
        response = response || {};

        this.uploadComplete((response.error == null), response);
    },

    setFileName: function (fileName) {
        this._props[API_PARAM_FILENAME] = fileName;
    },

    _process: function () {
        switch (this._status) {
            case UPLOAD_OPSTATE_NOTSTARTED:
                this._start();
                break;
            case UPLOAD_OPSTATE_AUTHREADY:
                this._getUploadPath();
                break;
            case UPLOAD_OPSTATE_UPLOADREADY:
                this._upload();
                break;
            case UPLOAD_OPSTATE_UPLOADCOMPLETED:
            case UPLOAD_OPSTATE_UPLOADFAILED:
            case UPLOAD_OPSTATE_CANCELED:
                this._complete();
                break;
        }
    },

    _start: function () {
        var op = this;
        getCoreApp().getLoginStatus({
            internal: true,
            callback: function () {
                op._status = UPLOAD_OPSTATE_AUTHREADY;
                op._process();
            }
        });
    },

    _getUploadPath: function () {
        var op = this,
            props = op._props,
            path = props[API_PARAM_PATH];

        if (isPathFullUrl(path)) {
            op._uploadPath = buildUploadFileUrlString(path);
            op._status = UPLOAD_OPSTATE_UPLOADREADY;
            op._process();
            return;
        }

        getCoreApp().api({ path: path }).then(
            function (response) {
                var location = response.upload_location;

                if (location) {
                    // Make sure the query path sent into WL.upload is sent to the upload POST/PUT request.
                    // Not just the GET upload_location request.
                    var queryString = parseQueryString(path);
                    location = appendQueryString(location, queryString);

                    var file_name = props[API_PARAM_FILENAME],
                        overwrite = props[API_PARAM_OVERWRITE];
                    if (isFilePath(path)) {
                        op._uploadPath = buildUploadFileUrlString(location);
                    }
                    else {
                        op._uploadPath = buildUploadToFolderUrlString(location, file_name, overwrite);
                    }

                    op._status = UPLOAD_OPSTATE_UPLOADREADY;
                }
                else {
                    op._result = createErrorResponse(ERROR_REQUEST_FAILED, ERROR_DESC_FAIL_UPLOAD.replace(METHOD, props[API_INTERFACE_METHOD]));
                    op._status = UPLOAD_OPSTATE_UPLOADFAILED;
                }

                op._process();
            },
            function (error) {
                op._result = error;
                op._status = UPLOAD_OPSTATE_UPLOADFAILED;
                op._process();
            }
        );
    },

    _upload: function () {
        this._strategy.upload(this._uploadPath);
    },

    _complete: function () {
        var op = this,
            result = op._result,
            promiseEvent = (op._status === UPLOAD_OPSTATE_UPLOADCOMPLETED) ?
                            PROMISE_EVENT_ONSUCCESS : PROMISE_EVENT_ONERROR;

        op._status = UPLOAD_OPSTATE_COMPLETED;

        var callback = op._props[API_PARAM_CALLBACK];
        if (callback) {
            callback(result);
        }

        op._promise[promiseEvent](result);
    }
};

/**
 * The pattern for the scope returned at the end of an external consent flow.
 * Is of the form PERMISSION_TYPE.ACCESS_LEVEL:PICKER_TYPE|SELECTION_TYPE|LINK_TYPE:RESOURCE_ID!AUTH_KEY
 * Where RESOURCE_ID is the id of the item that the user has granted the application access to.
 */
var scopeResponsePattern = new RegExp("^\\w+\\.\\w+:\\w+[\\|\\w+]+:([\\w]+\\!\\d+)(?:\\!(.+))*");

function stringTrim(value) {
    return value.replace(/^\s+|\s+$/g, "");
}

function stringsAreEqualIgnoreCase(str1, str2) {
    return str1 && str2 ? str1.toLowerCase() === str2.toLowerCase() : str1 === str2;
}

function stringIsNullOrEmpty(s) {
    return s == null || s === "";
}

/**
 * C# style string format.
 */
function stringFormat() {
    var args = arguments,
        original = args[0];

    function replaceFunc(match) {
        return args[parseInt(match.replace(/[\{\}]/g, "")) + 1] || '';
    }

    return original.replace(/\{\d+\}/g, replaceFunc);
}

/**
 * Aggressively encodes a string to be displayed in the browser. All non-letter characters are converted
 * to their Unicode entity ref, e.g. &#65;, period, space, comma, and dash are left un-encoded as well.
 * Usage: _divElement.innerHTML =_someValue.encodeHtml());
 */
function encodeHtml(text)
{
    var charCodeResult = {
        c: 0, // Code
        s: -1 // Next skip index
    };

    // Encode if a character not matches with [a-zA-Z0-9_{space}.,-].
    return text.replace(/[^\w .,-]/g, function (match, ind, s)
    {
        if (extendedCharCodeAt(s, ind, charCodeResult)) {
            return ["&#", charCodeResult.c, ";"].join("");
        }

        // If extendedCharCodeAt returns false that means this index is the low surrogate,
        // which has already been processed, so we remove it by returning an empty string.
        return "";
    });
}

/**
 * Gets the char code from str at idx.
 * Supports Secondary-Multilingual-Plane Unicode characters (SMP), e.g. codes above 0x10000
 */
function extendedCharCodeAt(str, idx, result)
{
    var skip = (result.s === idx);
    if (!skip)
    {
        idx = idx || 0;
        var code = str.charCodeAt(idx);
        var hi, low;
        result.s = -1;
        if (code < 0xD800 || code > 0xDFFF) {
            // Main case, Basic-Multilingual-Plane (BMP) code points.
            result.c = code;
        }
        else if (code <= 0xDBFF) {
            // High surrogate of SMP
            hi = code;
            low = str.charCodeAt(idx + 1);
            result.c = ((hi - 0xD800) * 0x400) + (low - 0xDC00) + 0x10000;
            result.s = idx + 1;
        }
        else {
            // Low surrogate of SMP, 0xDC00 <= code && code <= 0xDFFF
            // Shouldn't really ever come in here, previous call to this method would set skip index in result
            // in high surrogate case, which is short-circuited at the start of this function.
            result.c = -1;
            skip = true;
        }
    }

    return !skip;
}

/**
 * Create a cloned object.(one level clone)
 */
function cloneObject(obj, target) {
    var clonedObj = target || {};
    if (obj != null) {
        for (var key in obj) {
            clonedObj[key] = obj[key];
        }
    }
    return clonedObj;
}

/**
 * Create a cloned object and remove some properties.
 */
function cloneObjectExcept(obj, target, exceptionlist) {
    var clonedObject = cloneObject(obj, target);
    for (var i = 0; i < exceptionlist.length; i++) {
        delete clonedObject[exceptionlist[i]];
    }

    return clonedObject;
}

/**
 * Checks if an array contains a given object.
 */
function arrayContains(arr, obj) {
    var i;
    for (i = 0; i < arr.length; i++) {
        if (arr[i] === obj) {
            return true;
        }
    }

    return false;
}

/**
 * Merge two arrays into one and avoid duplicate.
 */
function arrayMerge(arr1, arr2) {
    var arr = [], arrMap = {};

    var addToArray = function (elements) {
        for (var i = 0; i < elements.length; i++) {
            arrElm = elements[i];
            if (arrElm != "" && !arrMap[arrElm]) {
                arrMap[arrElm] = true;
                arr.push(arrElm);
            }
        }
    };

    addToArray(arr1);
    addToArray(arr2);

    return arr;
}

/**
 * Create a cloned array.
 */
function cloneArray(array) {
    return Array.prototype.slice.call(array);
}

/**
 * Create a delegate for a instance method.
 */
function createDelegate(instance, method) {
    return function () {
        if (typeof (method) === TYPE_FUNCTION) {
            return method.apply(instance, arguments);
        }
    }
}

/**
 * Log message into console
 */
function writeLog(text, type, prefix)
{
    prefix = prefix || "[WL]";
    text = prefix + text;

    var w = window;
    if (w.debugLogger) {
        w.debugLogger.push(text);
    }

    var c = w.console;
    if (c && c.log) {
        switch (type) {
            case "warning":
                c.warn(text);
                break;
            case "error":
                c.error(text);
                break;
            default:
                c.log(text);
        }
    }

    var o = w.opera;
    if (o) {
        o.postError(text);
    }

    var d = w.debugService;
    if (d) {
        d.trace(text);
    }
}

function isPathFullUrl(path) {
    return path.indexOf("https://") === 0 || path.indexOf("http://") === 0 || path.indexOf("data:") === 0;
}

function trace(text) {
    if (wl_app._traceEnabled) {
        writeLog(text);
    }
}

function log(text, type, prefix){
    if (wl_app._logEnabled || wl_app._traceEnabled) {
        writeLog(text, type, prefix);
    }

    wl_event.notify(EVENT_LOG, text);
}

if (window.WL && WL.Internal) {
    WL.Internal.trace = trace;
    WL.Internal.log = log;
}

function logText(text, prefix) {
    log(text, "text", prefix);
}

function logError(text, prefix) {
    log(text, "error", prefix);
}

function createHiddenIframe(url, id) {
    var iframe = createHiddenElement("iframe");
    iframe.id = id;
    iframe.src = url;
    document.body.appendChild(iframe);

    return iframe;
}

function createHiddenElement(tagName) {
    var element = document.createElement(tagName);
    element.style.position = "absolute";
    element.style.top = "-1000px";
    element.style.width = "300px";
    element.style.height = "300px";
    element.style.visibility = "hidden";

    return element;
}

function createUniqueElementId() {

    var id = null;

    while (id == null) {
        id = "wl_" + Math.floor(Math.random() * 1024 * 1024);

        if (getElementById(id) != null) {
            id = null;
        }
    }

    return id;
}

function getElementById(id) {
    return document.getElementById(id);
}

function setElementText(element, text) {
    if (element) {
        if (element.innerText) {
            element.innerText = text;
        }
        else {
            // Firefox does not have innerText property. So, we do it differently.
            var textNode = document.createTextNode(text);
            element.innerHTML = '';
            element.appendChild(textNode);
        }
    }
}

/**
 * Takes in a url string object and returns the query string portion.
 * (e.g., given foo.com?k1=v1, this function returns k1=v1).
 */
function parseQueryString(url) {
    var queryStringIndex = url.indexOf("?");
    if (queryStringIndex === -1) {
        return "";
    }

    var fragmentIndex = url.indexOf("#", queryStringIndex + 1);
    if (fragmentIndex !== -1) {
        return url.substring(queryStringIndex + 1, fragmentIndex);
    }

    return url.substring(queryStringIndex + 1);
}

/**
 * Takes in a base url and a query string and appends the query string to the end of the base url.
 * (e.g., given foo.com?k1=v1 and k2=v2, this function returns foo.com?k1=v1&k2=v2).
 * Returns the unmodified baseUrl if the queryString is empty string, null, or undefined.
 */
function appendQueryString(baseUrl, queryString) {
    if (typeof(queryString) === TYPE_UNDEFINED || queryString === null || queryString === "") {
        return baseUrl;
    }

    var queryStringIndex = baseUrl.indexOf("?");
    if (queryStringIndex === -1) {
        return baseUrl + "?" + queryString;
    }

    if (baseUrl.charAt(baseUrl.length - 1) !== "&") {
        baseUrl += "&";
    }

    return baseUrl + queryString;
}

function Uri(url) {
    cloneObject(parseUri(url), this);
}

Uri.prototype = {
    toString: function () {
        var uri = this,
        s = (uri.scheme != "" ? uri.scheme + "//" : "") + uri.host + (uri.port != "" ? ":" + uri.port : "") + uri.path;
        return s;
    },

    resolve: function () {
        var uri = this;

        if (uri.scheme == "") {
            var port = window.location.port,
                host = window.location.host;

            uri.scheme = window.location.protocol;
            uri.host = host.split(":")[0];
            uri.port = port != null ? port : "";

            if (uri.path.charAt(0) != "/") {
                uri.path = resolveRelativePath(uri.host, window.location.href, uri.path);
            }
        }
    }
};

function parseUri(url) {
    // Assume url never be null or empty.

    var scheme = (url.indexOf(SCHEME_HTTPS) == 0) ? SCHEME_HTTPS : (url.indexOf(SCHEME_HTTP) == 0) ? SCHEME_HTTP : "",
        host = "",
        port = "",
        path;

    if (scheme != "") {
        var urlPart = url.substring(scheme.length + 2),
            firstSlash = coalesceToInfinity(urlPart.indexOf("/")),
            backSlash = coalesceToInfinity(urlPart.indexOf("\\")),
            queryStart = coalesceToInfinity(urlPart.indexOf("?")),
            fragStart = coalesceToInfinity(urlPart.indexOf("#")),
            hostEnd = coalesceToNegative(Math.min(firstSlash, backSlash, queryStart, fragStart)),
            hostPortStr = (hostEnd > 0) ? urlPart.substring(0, hostEnd) : urlPart,
            hostport = hostPortStr.split(":"),
            host = hostport[0],
            port = (hostport.length > 1) ? hostport[1] : "",
            path = (hostEnd > 0) ? urlPart.substring(hostEnd) : "";
    }
    else {
        path = url;
    }

    return { scheme: scheme, host: host, port: port, path: path };
}

function coalesceToInfinity(value)
{
    if (value == -1) {
        return Number.POSITIVE_INFINITY;
    }

    return value;
}

function coalesceToNegative(value)
{
    if (value == Number.POSITIVE_INFINITY) {
        return -1;
    }

    return value;
}

function getDomainName(url) {
    return parseUri(url.toLowerCase()).host;
}

function resolveRelativePath(hostname, href, url) {
    var trimRight = function (url, char) {
        charIdx = href.indexOf(char);
        url = (charIdx > 0) ? url.substring(0, charIdx) : url;
        return url;
    };

    href = trimRight(trimRight(href, "?"), "#");

    var hostIndex = href.indexOf(hostname),
        path = href.substring(hostIndex + hostname.length),
        pathIdx = path.indexOf("/"),
        folderIndex = path.lastIndexOf('/');
    path = (folderIndex >= 0) ? path.substring(pathIdx, folderIndex) : path;

    return path + "/" + url;
}

function trimUrlQuery(url) {
    var queryStart = url.indexOf("?");
    if (queryStart > 0) {
        url = url.substring(0, queryStart);
    }

    queryStart = url.indexOf("#");
    if (queryStart > 0) {
        url = url.substring(0, queryStart);
    }

    return url;
}

function getFileNameFromUrl(url) {
    var trimmedUrl = trimUrlQuery(url),
        index = trimmedUrl.lastIndexOf("/") + 1;

    return trimmedUrl.substr(index);
}

function invokeCallbackSynchronous(callback, resp) {
    if (typeof (callback) == TYPE_FUNCTION) {
        (resp !== undefined) ? callback(resp) : callback();
    }
}

function invokeCallback(callback, resp, synchronous, state) {
    if (typeof (callback) == TYPE_FUNCTION) {
        if (state) {
            resp[AK_STATE] = state;
        }

        if (synchronous) {
            callback(resp);
        }
        else {
            delayInvoke(
                function () {
                    callback(resp);
                }
            );
        }
    }
}

function deserializeJSON(text) {
    return JSON.parse(text);
}

function getCurrentSeconds() {
    // Get current timestamp in seconds.
    return Math.floor(new Date().getTime() / 1000);
}

function foreach(elementList, processElement) {
    var count = elementList.length;
    for (var i = 0; i < count; i++) {
        processElement(elementList[i]);
    }
}

function createAuthError(error, description) {
    var errorObj = {};
    errorObj[AK_ERROR] = error;
    errorObj[AK_ERROR_DESC] = description;
    return errorObj;
}

function createErrorResponse(code, message) {
    var errorObj = {}, errorResponse = {};

    errorObj[API_PARAM_CODE] = code,
    errorObj[API_PARAM_MESSAGE] = message;
    errorResponse[API_PARAM_ERROR] = errorObj;

    return errorResponse;
}

function createExceptionResponse(opName, event, err) {
    return createErrorResponse(
        ERROR_REQUEST_FAILED,
        ERROR_DESC_EXCEPTION.replace("METHOD", opName).replace("EVENT", event).replace("MESSAGE", err.message)
        );
}

function trimVersionBuildNumber(version) {
    var versionArr = version.split(".");
    return versionArr[0] + "." + versionArr[1];
}

function delayInvoke(callback, delay) {
    window.wlUnitTests ? wlUnitTests.delayInvoke(callback) : window.setTimeout(callback, delay || 1);
}

function detectBrowsers() {
    var browser = getBrowserInfo(navigator.userAgent, document.documentMode),
        libraryValue = wl_app[API_X_HTTP_LIVE_LIBRARY];
    wl_app._browser = browser;
    wl_app[API_X_HTTP_LIVE_LIBRARY] = libraryValue.replace("DEVICE", browser.device);
}

function getBrowserInfo(ua, documentMode) {
    ua = ua.toLowerCase();
    var device = "other",
        browser = {
        "firefox": /firefox/.test(ua),
        "firefox1.5": /firefox\/1\.5/.test(ua),
        "firefox2": /firefox\/2/.test(ua),
        "firefox3": /firefox\/3/.test(ua),
        "firefox4": /firefox\/4/.test(ua),
        "ie": (/msie/.test(ua) || /trident/.test(ua)) && !/opera/.test(ua),
        "ie6": false,
        "ie7": false,
        "ie8": false,
        "ie9": false,
        "ie10": false,
        "ie11": false,
        "opera": /opera/.test(ua),
        "webkit": /webkit/.test(ua),
        "chrome": /chrome/.test(ua),
        "mobile": /mobile/.test(ua) || /phone/.test(ua)
    };

    if (browser["ie"]) {
        // detect the rendering engine IE is using.
        // if documentMode is defined, we rely on its value to determine the rendering engine.
        var engine = 0;

        if (documentMode) {
            engine = documentMode;
        }
        else {
            // if we're in a browser that doesn't support documentMode (IE6, IE7) we need to do some more sniffing.
            if (/msie 7/.test(ua)) {
                engine = 7;
            }
        }

        // clamp the engine on 6,11
        engine = Math.min(11, Math.max(engine, 6));
        device = "ie" + engine;

        browser[device] = true;
    }
    else {
        if (browser.firefox) {
            device = "firefox";
        }
        else if (browser.chrome) {
            device = "chrome";
        }
        else if (browser.webkit) {
            device = "webkit";
        }
        else if (browser.opera) {
            device = "opera";
        }
    }

    if (browser.mobile) {
        device += "mobile";
    }

    browser.device = device;
    return browser;
}

/**
 * Deserializes name/value pair parameters from a string into a dictionary object.
 */
function deserializeParameters(value, existingDict) {
    var dict = (existingDict != null) ? existingDict : {};

    if (value != null) {
        var properties = value.split('&');
        for (var i = 0; i < properties.length; i++) {
            var property = properties[i].split('=');
            if (property.length == 2) {
                dict[decodeURIComponent(property[0])] = decodeURIComponent(property[1]);
            }
        }
    }

    return dict;
}

/**
 * Serializes a dictionary object into a string value in a format n1=v1&n2=v2&...
 */
function serializeParameters(dict) {
    var serialized = "";
    if (dict != null) {
        for (var key in dict) {
            if (dict.hasOwnProperty(key)) {
                var separator = serialized.length ? "&" : "";
                var value = dict[key];
                serialized += separator + encodeURIComponent(key) + "=" + encodeURIComponent(stringifyParamValue(value));
            }
        }
    }

    return serialized;
}

/**
 * Serializes a value into string.
 */
function stringifyParamValue(v) {

    if (v instanceof Date) {
        var padding = function (n, c) {
            switch (c) {
                case 2:
                    return n < 10 ? '0' + n : n;
                case 3:
                    return (n < 10 ? '00' : (n < 100 ? '0' : '')) + n;
            }
        };

        return v.getUTCFullYear() + '-' +
            padding(v.getUTCMonth() + 1, 2) + '-' +
            padding(v.getUTCDate(), 2) + 'T' +
            padding(v.getUTCHours(), 2) + ':' +
            padding(v.getUTCMinutes(), 2) + ':' +
            padding(v.getUTCSeconds(), 2) + '.' +
            padding(v.getUTCMilliseconds(), 3) + 'Z';
    }

    return "" + v;
}

/**
 * Read Url parameters.
 */
function readUrlParameters(url) {

    var queryStart = url.indexOf('?') + 1,
        hashStart = url.indexOf('#') + 1,
        dict = {};

    if (queryStart > 0) {
        var queryEnd = (hashStart > queryStart) ? (hashStart - 1) : url.length;
        dict = deserializeParameters(url.substring(queryStart, queryEnd), dict);
    }

    if (hashStart > 0) {
        dict = deserializeParameters(url.substring(hashStart), dict);
    }

    return dict;
}

/**
 * Appends an object of parameters to the base url.
 * The function checks the base url for existing params
 * and appropriately appends a ? or an & to the base url
 * before appending the parameters
 */
function appendUrlParameters(path, params) {
    return path + ((path.indexOf("?") < 0) ? "?" : "&") + serializeParameters(params);
}

/**
 * Normalize a parameter value into boolean type.
 */
function normalizeBooleanValue(value) {
    switch (typeof (value)) {
        case TYPE_BOOLEAN:
            return value;
        case TYPE_NUMBER:
            return !!value;
        case TYPE_STRING:
            return value.toLowerCase() === "true";
        default:
            return false;
    }
}

var expectedCallback_Optional = {
    name: API_PARAM_CALLBACK,
    type: TYPE_FUNCTION,
    optional: true
};

var expectedCallback_Required = {
    name: API_PARAM_CALLBACK,
    type: TYPE_FUNCTION,
    optional: false
};

function validateParams(params, expectedParams, method) {
    if (params instanceof Array) {
        for (var i = 0; i < params.length; i++) {
            validateParam(params[i], expectedParams[i], method);
        }
    }
    else {
        validateParam(params, expectedParams, method);
    }
}

function validateParam(param, expectedParam, method) {
    validateParamType(param, expectedParam, method);

    if (expectedParam.type === TYPE_PROPERTIES) {
        validateProperties(param, expectedParam.properties, method);
    }
}

function validateParamType(param, expected, method, defaultAssignFunc) {
    var paramName = expected.name,
        paramType = typeof (param),
        expectedType = expected.type;

    if (paramType === TYPE_UNDEFINED || param == null) {
        if (expected.optional){
            defaultAssignFunc && defaultAssignFunc();
            return;
        }
        else {
            throw createMissingParamError(paramName, method);
        }
    }

    switch (expectedType) {
        case TYPE_STRING:
            {
                if (paramType != TYPE_STRING) {
                    throw createParamTypeError(paramName, method);
                }
                if (!expected.optional && stringTrim(param) === "") {
                    throw createMissingParamError(paramName, method);
                }
            }
            break;
        case TYPE_URL:
            {
                if (paramType != TYPE_STRING || !isPathFullUrl(param)) {
                    throw createParamTypeError(paramType, method);
                }
            }
            break;
        case TYPE_PROPERTIES:
            {
                if (paramType != TYPE_OBJECT) {
                    throw createParamTypeError(paramName, method);
                }
            }
            break;
        case TYPE_FUNCTION:
            {
                if (paramType != TYPE_FUNCTION) {
                    throw createParamTypeError(paramType, method);
                }
            }
            break;
        case TYPE_DOM:
            {
                if (paramType == TYPE_STRING) {
                    if (getElementById(param) == null) {
                        throw new Error(ERROR_DESC_DOM_INVALID.replace(METHOD, method).replace(PARAM, paramName));
                    }
                }
                else if (paramType != TYPE_OBJECT) {
                    throw createParamTypeError(paramName, method);
                }
            }
            break;
        case TYPE_STRINGORARRAY:
            {
                if (paramType != TYPE_STRING && !(param instanceof Array)) {
                    throw createParamTypeError(paramType, method);
                }
            }
            break;
        default:
            if (paramType != expectedType) {
                throw createParamTypeError(paramName, method);
            }
            break;
    }

    if (expected.allowedValues != null) {
         validateAllowedValue(param, expected.allowedValues, expected.caseSensitive, paramName, method);
    }
}

function validateProperties(param, expectedProperties, method) {
    var properties = param || {};
    for (var i = 0; i < expectedProperties.length; i++) {
        var expectedProperty = expectedProperties[i],
            actualProperty = properties[expectedProperty.name] || properties[expectedProperty.altName];
        validateParamType(actualProperty, expectedProperty, method, function()
        {
            // Apply default value.
            if (typeof (expectedProperty.defaultValue) != TYPE_UNDEFINED) {
                properties[expectedProperty.name] = expectedProperty.defaultValue;
            }
        });
    }
}

function validateAllowedValue(value, allowedValues, caseSensitive, paramName, method) {
    var isString = typeof (allowedValues[0]) === TYPE_STRING;
    for (var i = 0; i < allowedValues.length; i++) {
        if (isString && !caseSensitive) {
            if (value.toLowerCase() === allowedValues[i].toLowerCase()) {
                return;
            }
        }
        else if (value === allowedValues[i]) {
            return;
        }
    }

    throw createInvalidParamValue(paramName, method);
}

function createParamTypeError(paramName, method) {
    return new Error(ERROR_DESC_PARAM_TYPE_INVALID.replace(METHOD, method).replace(PARAM, paramName));
}

function createMissingParamError(paramName, method) {
    return new Error(ERROR_DESC_PARAM_MISSING.replace(METHOD, method).replace(PARAM, paramName));
}

function createInvalidParamValue(paramName, method, optionalMessage) {
    var message = ERROR_DESC_PARAM_INVALID.replace(METHOD, method).replace(PARAM, paramName);
    if (typeof(optionalMessage) !== TYPE_UNDEFINED) {
        message += " " + optionalMessage;
    }

    return new Error(message);
}

function findArgumentByType(args, type, maxToRead) {
    if (args) {
        for (var i = 0; i < maxToRead && i < args.length; i++) {
            if (type === typeof args[i]) {
                return args[i];
            }
        }
    }

    return undefined;
}

function normalizeArguments(args, methodName) {
    var receivedArgs = cloneArray(args),
        properties = null,
        callback = null;

    for (var i = 0; i < receivedArgs.length; i++) {
        var arg = receivedArgs[i],
            argType = typeof(arg);

        if (argType === TYPE_OBJECT && properties === null) {
            properties = cloneObject(arg);
        }
        else if (argType === TYPE_FUNCTION && callback === null) {
            callback = arg;
        }
    }

    properties = properties || {};

    if (callback) {
        properties.callback = callback;
    }

    properties[API_INTERFACE_METHOD] = methodName;

    return properties;
}

function normalizeApiArguments(args) {
    var receivedArgs = cloneArray(args),
        path = null,
        method = null;

    if (typeof receivedArgs[0] === TYPE_STRING) {
        // Read path
        path = receivedArgs.shift();

        if (typeof receivedArgs[0] === TYPE_STRING) {
            // Read method
            method = receivedArgs.shift();
        }
    }

    normalizedArgs = normalizeArguments(receivedArgs);

    if (path !== null) {
        normalizedArgs[API_PARAM_PATH] = path;

        if (method != null) {
            normalizedArgs[API_PARAM_METHOD] = method;
        }
    }

    return normalizedArgs;
}

function handleAsyncCallingError(name, err) {
    var error = createExceptionResponse(name, name, err);
    logError(err.message);
    return createCompletePromise(name, false, null, error);
}

var Promise = function (opName, op, uplinkPromise) {
    this._name = opName;
    this._op = op;
    this._uplinkPromise = uplinkPromise;
    this._isCompleted = false;
    this._listeners = [];
};

Promise.prototype = {
    then: function (onSuccess, onError, onProgress) {
        var chainedPromise = new Promise(null, null, this),
            listener = {};
        listener[PROMISE_EVENT_ONSUCCESS] = onSuccess;
        listener[PROMISE_EVENT_ONERROR] = onError;
        listener[PROMISE_EVENT_ONPROGRESS] = onProgress;
        listener.chainedPromise = chainedPromise;

        this._listeners.push(listener);

        return chainedPromise;
    },

    cancel: function () {
        if (this._isCompleted) return;

        if (this._uplinkPromise && !this._uplinkPromise._isCompleted) {
            // If there is incomplete uplink promise, we cancel that one and let the flow propagate to this one.
            this._uplinkPromise.cancel();
        }
        else {
            // We need to cancel the current one, if we can.
            var opCancel = (this._op) ? this._op.cancel : null;
            if (typeof (opCancel) === TYPE_FUNCTION) {
                this._op.cancel();
            }
            else {
                this.onError(
                    createErrorResponse(ERROR_REQ_CANCEL, ERROR_DESC_CANCEL.replace("METHOD", this._getName()))
                );
            }
        }
    },

    _getName: function () {

        if (this._name) {
            return this._name;
        }

        if (this._op && typeof (this._op._getName) === TYPE_FUNCTION) {
            return this._op._getName();
        }

        if (this._uplinkPromise) {
            return this._uplinkPromise._getName();
        }

        return "";
    },

    _onEvent: function (args, name) {
        if (this._isCompleted) return;
        this._isCompleted = (name !== PROMISE_EVENT_ONPROGRESS);

        this._notify(args, name);
    },

    _notify: function (args, event) {
        var currentPromise = this;
        foreach(this._listeners, function (listener) {
            var callback = listener[event],
                chainedPromise = listener.chainedPromise,
                isPromiseCompleted = (event !== PROMISE_EVENT_ONPROGRESS);

            if (callback) {
                try {
                    var chainedPromiseOrigin = callback.apply(listener, args);
                    if (isPromiseCompleted && chainedPromiseOrigin && chainedPromiseOrigin.then) {
                        // We need to link and fulfill the chained promise with the one returned from callback
                        // if this is onSuccess or onError.
                        // Also, set the new promise as the _op of the chained promise in case cancel is invoked.
                        chainedPromise._op = chainedPromiseOrigin;
                        chainedPromiseOrigin.then(
                            function (result) {
                                chainedPromise[PROMISE_EVENT_ONSUCCESS](result);
                            },
                            function (error) {
                                chainedPromise[PROMISE_EVENT_ONERROR](error);
                            },
                            function (progress) {
                                chainedPromise[PROMISE_EVENT_ONPROGRESS](progress);
                            }
                        );
                    }
                }
                catch (err) {
                    if (isPromiseCompleted) {
                        // The the callback throws an error, that should be forwarded to the chained promise.
                        chainedPromise.onError(
                            createExceptionResponse(currentPromise._getName(), event, err)
                        );
                    }
                }
            }
            else {
                if (isPromiseCompleted) {
                    // If no onSuccess/onError is handled, we forward event to the chained promise.
                    chainedPromise[event].apply(chainedPromise, args);
                }
            }
        });
    }
};

Promise.prototype[PROMISE_EVENT_ONSUCCESS] = function () {
    this._onEvent(arguments, PROMISE_EVENT_ONSUCCESS);
};

Promise.prototype[PROMISE_EVENT_ONERROR] = function () {
    this._onEvent(arguments, PROMISE_EVENT_ONERROR);
};

Promise.prototype[PROMISE_EVENT_ONPROGRESS] = function () {
    this._onEvent(arguments, PROMISE_EVENT_ONPROGRESS);
};

function createCompletePromise(opName, succeeded, callback, result) {
    var promise = new Promise(opName, null, null),
        completeEvent = succeeded ? PROMISE_EVENT_ONSUCCESS : PROMISE_EVENT_ONERROR;

    if (typeof (callback) === TYPE_FUNCTION) {
        promise.then(function (rs) {
            callback(rs);
        });
    }

    delayInvoke(
        function () {
            promise[completeEvent](result);
        }
    );

    return promise;
}


var AK_COOKIE_KEYS = [AK_ACCESS_TOKEN, AK_AUTH_TOKEN, AK_SCOPE, AK_EXPIRES_IN, AK_EXPIRES, AK_REQUEST_TS, AK_ERROR, AK_ERROR_DESC];

var AK_REFRESH_TYPE = "refresh_type",
    AK_REFRESH_TYPE_APP = "app",
    AK_REFRESH_TYPE_MS = "ms";

var AK_RESPONSE_METHOD = "response_method",
    AK_RESPONSE_METHOD_URL = "url",
    AK_RESPONSE_METHOD_COOKIE = "cookie";

var ANALYTICS_LISTENER = "onanalytics";

/**
 * Auth request types.
 */
var AUTH_REQUEST_LOGIN = "login",
    AUTH_REQUEST_LOGINSTATUS = "loginstatus";

var CHANNEL_NAME_FILEDIALOG = "file_dialog";

var EVENT_AUTH_RESPONSE = "auth.response";

/**
 * DOM strings.
 */
var DOM_ATTR_CLIENTID = "client-id",
    DOM_CLASS_ONEDRIVE_SAVEBUTTON = ".OneDriveSaveButton",
    DOM_FILE = "file",
    DOM_EVENT_CLICK = "click",
    DOM_ID_SDK = "onedrive-js";

/**
 * HTML constants.
 */
var HTML_BUTTONTEXT_MARGIN = "2px",
    HTML_BUTTONTEXT_MARGIN_NONE = "0px",
    HTML_BUTTON_PADDING = "4px";

/**
 * Error strings.
 */
var ERROR_DESC_CLIENTID_MISSING = "METHOD: Failed to initialize due to missing 'client-id'.",
    ERROR_DESC_PICKER_TIMEOUT = "Loading OneDrive picker is timed out.";

/**
 * Interface parameters.
 */
var FILEDIALOG_PARAM_MODE = "mode",
    FILEDIALOG_PARAM_MODE_OPEN = "open",
    FILEDIALOG_PARAM_MODE_SAVE = "save",
    FILEDIALOG_PARAM_MODE_READ = "read",
    FILEDIALOG_PARAM_MODE_READWRITE = "readwrite",
    FILEDIALOG_PARAM_LINKTYPE = "linkType",
    FILEDIALOG_PARAM_RESOURCETYPE = "resourceType",
    FILEDIALOG_PARAM_RESOURCETYPE_FILE = "file",
    FILEDIALOG_PARAM_RESOURCETYPE_FOLDER = "folder",
    FILEDIALOG_PARAM_SELECT = "select",
    FILEDIALOG_PARAM_SELECT_SINGLE = "single",
    FILEDIALOG_PARAM_SELECT_MULTI = "multi",
    FILEDIALOG_PARAM_PERMISSION = "permission",
    FILEDIALOG_PARAM_PERMISSION_ONETIME = "onetime",
    FILEDIALOG_PARAM_LIGHTBOX = "lightbox",
    FILEDIALOG_PARAM_LIGHTBOX_GREY = "grey",
    FILEDIALOG_PARAM_LIGHTBOX_TRANSPARENT = "transparent",
    FILEDIALOG_PARAM_LIGHTBOX_WHITE = "white",
    FILEDIALOG_PARAM_LOADING_TIMEOUT = "loading_timeout",
    FILEDIALOG_PARAM_ONSELECTED = "onselected";

/**
 * OneDrive interface parameters.
 */
var ONEDRIVE_PARAM_CANCEL = "cancel",
    ONEDRIVE_PARAM_ERROR = "error",
    ONEDRIVE_PARAM_FILE = "file",
    ONEDRIVE_PARAM_FILENAME = "fileName",
    ONEDRIVE_PARAM_INTERNAL = "internal_app",
    ONEDRIVE_PARAM_LINKTYPE = FILEDIALOG_PARAM_LINKTYPE,
    ONEDRIVE_PARAM_LINKTYPE_DOWNLOAD = "downloadLink",
    ONEDRIVE_PARAM_LINKTYPE_WEBVIEW = "webViewLink",
    ONEDRIVE_PARAM_PROGRESS = "progress",
    ONEDRIVE_PARAM_SELECT = "multiSelect",
    ONEDRIVE_PARAM_SUCCESS = "success",
    ONEDRIVE_PARAM_THEME = "theme",
    ONEDRIVE_PARAM_THEME_BLUE = "blue",
    ONEDRIVE_PARAM_THEME_WHITE = "white";

/**
 * Upload operations.
 */
var UPLOADTYPE_URL = "from_url",
    UPLOADTYPE_FORM = "form";

/**
 * OneDrive error strings.
 */
var ERROR_DESC_UPLOADTYPE_NOTIMPLEMENTED = "METHOD: This upload method is not implemented.",
    ERROR_DESC_GENERAL = "{0}: Operation: '{1}' Error message: {2}",
    ERROR_DESC_OPERATION_API = "API call",
    ERROR_DESC_OPERATION_PICKER = "invoke picker",
    ERROR_DESC_OPERATION_UNHANDLED_EXCEPTION = "unhandled exception",
    ERROR_DESC_OPERATION_UPLOAD = "upload",
    ERROR_DESC_OPERATION_UPLOAD_POLLING = "upload: poll for completion";

/**
 * OneDrive methods.
 */
var IMETHOD_ONEDRIVE_OPEN = "OneDrive.open",
    IMETHOD_ONEDRIVE_SAVE = "OneDrive.save",
    IMETHOD_ONEDRIVE_CREATEBUTTON_OPEN = "OneDrive.createOpenButton",
    IMETHOD_ONEDRIVE_CREATEBUTTON_SAVE = "OneDrive.createSaveButton",
    IMETHOD_ONEDRIVE_CREATEBUTTON_SAVE_FROMLINK = "OneDriveApp.onloadInit",
    IMETHOD_ONEDRIVE_INITIALIZE = "OneDriveApp.initialize";

/**
 * SkyDrive url parameters and values.
 */
var FILEDIALOG_PARAM_AUTH = "auth",
    FILEDIALOG_PARAM_AUTH_RPS = "rps",
    FILEDIALOG_PARAM_AUTH_OAUTH = "oauth",
    FILEDIALOG_PARAM_VIEWTYPE = "v",
    FILEDIALOG_PARAM_VIEWTYPE_FOLDERPICKER = "1",
    FILEDIALOG_PARAM_VIEWTYPE_FILEPICKER = "2",
    FILEDIALOG_PARAM_DOMAIN = "domain",
    FILEDIALOG_PARAM_LIVESDK = "livesdk",
    FILEDIALOG_PARAM_MKT = "mkt";
    FILEDIALOG_PARAM_PICKER_SCRIPT = "pickerscript";
    FILEDIALOG_CHCMD_ONCOMPLETE = "onPickerComplete",
    FILEDIALOG_CHCMD_UPDATETOKEN = "updateToken";

/**
 * Miscellaneous
 */
var FORM_UPLOAD_SIZE_LIMIT = 104857600 /* 100 MB in bytes */,
    KEYCODE_ESC = 27,
    POLLING_INTERVAL = 1000 /* 1 second in milliseconds */,
    POLLING_COUNTER = 5;
    ONEDRIVE_PREFIX = "[OneDrive]",
    UI_SKYDRIVEPICKER = "skydrivepicker",
    VROOM_THUMBNAIL_SIZES = ["large", "medium", "small"];

WL.init = function (properties) {
    /// <summary>
    /// Initializes the JavaScript library. An application must call this function before making other function calls to
    /// the library except for subscribing/unsubscribing to events.
    /// </summary>
    /// <param name="properties" type="Object">
    /// Required. A JSON object that includes the following properties:
    /// &#10; client_id:  Required. The OAuth client ID of your application.
    /// &#10; scope:  Optional. The scope values used to determine if the user has logged in.
    /// &#10; redirect_uri:  Optional. The default redirect URI used for OAuth authentication. The OAuth server redirects
    /// to this URI during the OAuth flow. The redirect_uri value must match the redirect domain of the registered app.
    /// &#10; response_type:  Optional. The OAuth response_type value. It can be either "code" or "token". If set to
    /// "token" (default), the client will receive the access token directly. If set to "code" the OAuth server will return
    /// an authorization code and the application server that serves the redirect_uri page should handle retrieving the
    /// access token from the OAuth server using the authorization code and client secret.
    /// &#10; refresh_type:  Optional. Indicates the way on how to check user's login status and retrieve a new access token
    /// if the user already consented the scopes required by the app. Checking login status and retrieving a new access token
    /// happens in the following scenarios: i) The library is initialized via WL.init(...); ii) WL.getLoginStatus() is invoked;
    /// iii) Sign-in control is initiaized via WL.ui(....); iv) The access token is expired. If set to 'app', the library will
    /// send a request to the app server with address value specified in the redirect_uri parameter in WL.init(...). The app
    /// server should authenticate the user and retrieve the corresponding access token via the persisted refresh token of the
    /// user. If not specified, by default, the library will send a request to the OAuth server to perform login status checking
    /// and access token retrieving task. Note: to retrieve a new access token using default approach requires that i) the user
    /// must have already signed in with Microsoft account in the current web session and ii) the user must have already consented
    /// 'wl.signin' scope to the app.
    /// &#10; logging: Optional. If set to true (default), the library logs error information to the JavaScript console and
    /// notifies the application through "wl.log" event.
    /// &#10; status: Optional. If set to true (default), the library attempts to get the login status of the user after
    /// WL.init is invoked.
    /// &#10; secure_cookie: Optional. If set to true (default), the library will specify secure attribute when writing the
    /// cookie on an https page.
    /// </param>
    /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
    /// situations.</returns>

    try {
        var clonedProperties = cloneObject(properties),
            method = IMETHOD_WL_INIT;

        // Validate parameters
        validateParams(
            clonedProperties,
            {
                name: TYPE_PROPERTIES,
                type: TYPE_PROPERTIES,
                optional: false,
                properties: [
                    { name: AK_CLIENT_ID, altName: CK_APPID, type: TYPE_STRING, optional: false },
                    { name: AK_SCOPE, type: TYPE_STRINGORARRAY, optional: true },
                    { name: AK_REDIRECT_URI, altName: CK_CHANNELURL, type: TYPE_STRING, optional: true },
                    { name: AK_RESPONSE_TYPE, type: TYPE_STRING, allowedValues: [RESPONSE_TYPE_CODE, RESPONSE_TYPE_TOKEN], optional: true },
                    { name: AK_REFRESH_TYPE, type: TYPE_STRING, allowedValues: [AK_REFRESH_TYPE_APP, AK_REFRESH_TYPE_MS], optional: true },
                    { name: API_PARAM_LOGGING, type: TYPE_BOOLEAN, optional: true },
                    { name: AK_STATUS, type: TYPE_BOOLEAN, optional: true }
                ]
            },
            method);

        if (!clonedProperties[AK_REDIRECT_URI] && clonedProperties[AK_RESPONSE_TYPE] === AK_CODE) {
            throw new Error(ERROR_DESC_REDIRECTURI_MISSING.replace(METHOD, method));
        }

        if (clonedProperties[AK_STATUS] == null) {
            clonedProperties[AK_STATUS] = true;
        }

        return wl_app.appInit(clonedProperties);
    }
    catch (e) {
        return handleAsyncCallingError(method, e);
    }
};

WL.login = function (properties, callback) {
    /// <summary>
    /// Signs the user in or expands the user's permission set. This function can result in launching the consent
    /// page popup. Therefore, it should only be called in response to a user action such as clicking a button.
    /// Otherwise, the web browser may block the popup. This is an async method that returns a Promise object
    /// that allows you to attach events to handle succeeded or failed situations.
    /// </summary>
    /// <param name="properties" type="Object">
    /// Required. A JSON object with the following properties:
    /// &#10; redirect_uri: Optional. By default, the redirect_uri parameter supplied to WL.init is used.
    /// An application can override it for specific scenarios with this parameter.
    /// &#10; scope: Required. The scopes for the user to authorize. It can be an array
    /// of scope string values or a string value of multiple scopes delimited by a space character.
    /// &#10; state: Optional. This parameter can be used to track the caller state at the app server side if you
    /// choose to implement a server flow authentication.
    /// </param>
    /// <param name="callback" type="Function" >Optional. A function that is invoked when login is completed.</param>
    /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
    /// situations.</returns>

    try {
        var args = normalizeArguments(arguments),
            method = IMETHOD_WL_LOGIN;

        // Validate parameters
        validateProperties(
            args,
            [
                { name: AK_SCOPE, type: TYPE_STRINGORARRAY, optional: true },
                { name: AK_REDIRECT_URI, type: TYPE_STRING, optional: true },
                { name: AK_STATE, type: TYPE_STRING, optional: true },
                expectedCallback_Optional
            ],
            method);

        return wl_app.login(args);
    }
    catch (e) {
        return handleAsyncCallingError(method, e);
    }
};

WL.download = function (properties, callback) {
    /// <summary>
    /// Makes a call to download a file from SkyDrive. This is an async method that returns a Promise object that
    /// allows you to attach events to handle succeeded and failed situations.
    /// </summary>
    /// <param name="properties" type="Object">Required. A JSON object containing the properties for downloading a file:
    /// &#10; path: Required. The path to the file to download.
    /// </param>
    /// <param name="callback" type="Function">Optional. A callback function that is invoked when the download call is complete.</param>
    /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
    /// situations.</returns>

    try {
        var method = IMETHOD_WL_DOWNLOAD,
            args = normalizeArguments(arguments, method);
        return wl_app.download(args);
    }
    catch (e) {
        return handleAsyncCallingError(method, e);
    }
};

WL.upload = function (properties, callback) {
    /// <summary>
    /// Makes a call to upload a file to SkyDrive.
    /// This is an async method that returns a Promise object that allows you to attach events to handle succeeded and failed situations.
    /// </summary>
    /// <param name="properties" type="Object">Required. A JSON object containing the properties for uploading a file:
    /// &#10; path: Required. The path to the file to download.
    /// &#10; element: Required. The DOM element or "id" value of a file input tag.
    /// &#10; overwrite: Indicates if the uploading action should overwrite a file that already exists. This only applies to when
    /// uploading to a folder. Suported values include "true", "false", "rename", true, false.
    /// </param>
    /// <param name="callback" type="Function">Optional. A callback function that is invoked when the upload call is complete.</param>
    /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded and failed
    /// situations.</returns>

    try {
        var method = IMETHOD_WL_UPLOAD,
            args = normalizeArguments(arguments, method);

        return wl_app.upload(args);
    } catch (e) {
        return handleAsyncCallingError(method, e);
    }
};

WL.ui = function (properties, callback) {
    /// <summary>
    /// Creates a user interface control on the current page.
    /// </summary>
    /// <param name="properties" type="Object">Required. A JSON object containing properties for creating the user interface element.
    /// &#10; name: Required. Specifies the name of the UI element to create. For the sign-in control, it is "signin". For SkyDrive Picker,
    /// it is "skydrivepicker".
    /// &#10; Sign-in control properties:
    /// &#10; element: Required. The DOM element to attach to the UI element.
    /// &#10; brand: Optional. Defines the brand, or type of icon, used with the signin control. It can be one of the following
    /// values: "hotmail", "messenger", "windows"(default), "skydrive", or "none".
    /// &#10; theme: Optional. The options are "blue" (default) and "white".
    /// &#10; type: Optional. Defines the type of the sign-in control. It can be one of the following values: "signin" (default),
    /// "login", "connect", or "custom".
    /// &#10; sign_in_text: If the type value is "custom", defines the signin text displayed in the sign-in control.
    /// &#10; sign_out_text: If the type value is "custom", defines the sign out text displayed in the sign-in control.
    /// &#10; state: Optional. This parameter can be used to track the caller state at the app server side if you
    /// choose to implement a server flow authentication.
    /// &#10; onloggedin: Optional. A callback function that will be invoked when the user is logged in.
    /// &#10; onloggedout: Optional. A callback function that will be invoked when the user is logged out.
    /// &#10; onerror: Optional. A callback funtion that will be invoked when there is error during logging in.
    /// &#10;
    /// &#10; SkyDrive picker properties:
    /// &#10; theme: Optional. The options are "blue" and "white" (default).
    /// &#10; mode: Required. Specify the mode of the dialog to open. It can be either "open" or "save". If it is "open", the dialog
    /// will be an open picker dialog that allows the user to select file(s). If it is "save", the dialog will be a save picker dialog
    /// that allows the user to select a folder.
    /// &#10; select: Optional. This is only used when mode value is "open" to specify if multiple files are allowed to be selected.
    /// It can be either "single" (default) or "multi".
    /// &#10; lightbox: Optional. Specifies the dialog lightbox color. It can be either "white" (default), "grey" or "transparent".
    /// &#10; onselected: Required. A callback function that will be invoked when the user has selected SkyDrive items successfully.
    /// &#10; onerror: Optional. A callback function that will be invoked when the picker dialog is closed with no SkyDrive selected.
    /// &#10; loading_timeout: Optional. Specifies number of seconds as a timeout value for loading the SkyDrive picker. If the specified has time passed and
    /// the picker dialog has not been loaded properly, the picker will be disposed and the failure event callback will be invoked. If this parameter is not
    /// specified, the timeout behavior will be disabled.
    /// </param>
    /// <param name="callback" type="Function">Optional. A callback function that is invoked when the UI element is rendered.</param>

    try {
        var args = normalizeArguments(arguments);

        // Validate parameters
        validateProperties(
                args,
                [
                    { name: UI_PARAM_NAME, type: TYPE_STRING, allowedValues: [UI_SIGNIN, UI_SKYDRIVEPICKER], optional: false },
                    expectedCallback_Optional
                ],
                IMETHOD_WL_UI);
        wl_app.ui(args);
    }
    catch (e) {
        handleUIParameterError(args, e);
    }
};

WL.fileDialog = function (properties, callback) {
    /// <summary>
    /// Shows a picker dialog to allow users to pick items from their SkyDrive account.
    /// This is an async method that returns a Promise object that allows you to attach events to handle succeeded or failed situations.
    /// </summary>
    /// <param name="properties" type="Object" optional="false">Required. A JSON object containing properties for showing the dialog.
    /// &#10; mode: Required. Specify the mode of the dialog to open. It can be either "open" or "save". If it is "open", the dialog will be an open picker
    /// &#10; dialog that allows the user to select file(s). If it is "save", the dialog will be a save picker dialog that allows the user to select a folder.
    /// &#10; select: Optional. This is only used when mode value is "open" to specify if multiple files are allowed to be selected. It can be either "single"
    /// &#10; (default) or "multi".
    /// &#10; lightbox: Optional. Specifies the dialog lightbox color. It can be either "white" (default), "grey" or "transparent".
    /// &#10; loading_timeout: Optional. Specifies number of seconds as a timeout value for loading the SkyDrive picker. If the specified time has passed and
    /// the picker dialog has not been loaded properly, the picker will be disposed and the failure event callback will be invoked. If this parameter is not
    /// specified, the timeout behavior will be disabled.
    /// </param>
    /// <param name="callback" type="Function" optional="true">Optional. A callback function that is invoked when the file dialog is closed.</param>
    /// <returns type="Promise" mayBeNull="false" >The Promise object that allows you to attach events to handle succeeded or failed situations.</returns>

    try {
        var method = IMETHOD_FILEDIALOG,
            args = normalizeArguments(arguments, method);

        validateFileDialogCall(args, method);

        return wl_app.fileDialog(args);
    }
    catch (e) {
        return handleAsyncCallingError(method, e);
    }
};

function validateFileDialogCall(args, method) {
    // Validate WL.fileDialog parameters
    validateProperties(
        args,
        [{
            name: FILEDIALOG_PARAM_MODE, type: TYPE_STRING,
            allowedValues: [FILEDIALOG_PARAM_MODE_OPEN, FILEDIALOG_PARAM_MODE_SAVE, FILEDIALOG_PARAM_MODE_READ, FILEDIALOG_PARAM_MODE_READWRITE], optional: false
         },
         {
            name: FILEDIALOG_PARAM_RESOURCETYPE, type: TYPE_STRING,
            allowedValues: [FILEDIALOG_PARAM_RESOURCETYPE_FILE, FILEDIALOG_PARAM_RESOURCETYPE_FOLDER], optional: true
         },
         {
            name: FILEDIALOG_PARAM_SELECT, type: TYPE_STRING,
            allowedValues: [FILEDIALOG_PARAM_SELECT_SINGLE, FILEDIALOG_PARAM_SELECT_MULTI], optional: true },
         {
            name: FILEDIALOG_PARAM_LIGHTBOX, type: TYPE_STRING,
            allowedValues: [FILEDIALOG_PARAM_LIGHTBOX_GREY, FILEDIALOG_PARAM_LIGHTBOX_TRANSPARENT, FILEDIALOG_PARAM_LIGHTBOX_WHITE], optional: true
         },
         {
            name: FILEDIALOG_PARAM_LOADING_TIMEOUT, type: TYPE_NUMBER, optional: true
         },
         {
            name: FILEDIALOG_PARAM_PERMISSION, type: TYPE_STRING,
            allowedValues: [FILEDIALOG_PARAM_PERMISSION_ONETIME], optional: true
         },
         expectedCallback_Optional
        ],
         method);

    // Validate additional WL.ui-skydrivepicker parameters
    if (method !== IMETHOD_FILEDIALOG) {
        validateProperties(
            args,
            [{ name: UI_PARAM_THEME, allowedValues: [UI_SIGNIN_THEME_BLUE, UI_SIGNIN_THEME_WHITE], type: TYPE_STRING, optional: true },
             { name: FILEDIALOG_PARAM_ONSELECTED, type: TYPE_FUNCTION, optional: false },
             { name: UI_PARAM_ONERROR, type: TYPE_FUNCTION, optional: true }],
            method);
    }

    if (!ChannelManager.isSupported() || !window.JSON || wl_app._browser.mobile) {
        throw new Error(ERROR_DESC_BROWSER_LIMIT);
    }
}

/**
 * Read server response via Url parameters
 */
function processUrlParameters() {
    // Parse url parameters
    wl_app._urlParams = readUrlParameters(window.location.href);

    // Deserialize state parameters
    wl_app._pageState = deserializeParameters(wl_app._urlParams[AK_STATE]);
}

function saveServerResponse() {
    var cookieState = new CookieState(COOKIE_AUTH);
    cookieState.load();
    var urlParams = wl_app._urlParams,
        pageState = wl_app._pageState,
        shouldWriteCookie = true,
        requestTs = pageState[AK_REQUEST_TS];

    if (requestTs) {
        if (requestTs != cookieState.get(AK_REQUEST_TS)) {
            cookieState.set(AK_REQUEST_TS, pageState[AK_REQUEST_TS]);
        }
        else {
            // RequestTs has already been written, assuming by the app server.
            shouldWriteCookie = false;
        }
    }

    // As long as we have a token, the status should always be connected.
    var hasResponseToken = (urlParams[AK_ACCESS_TOKEN] != null),
        hasToken = (cookieState.get(AK_ACCESS_TOKEN) != null) || hasResponseToken,
        status = hasToken ? AS_CONNECTED : AS_UNKNOWN,
        currentTs = getCurrentSeconds();

    if (pageState[AK_RESPONSE_METHOD] === AK_RESPONSE_METHOD_URL) {

        for (var i = 0; i < AK_COOKIE_KEYS.length; i++) {
            var authKey = AK_COOKIE_KEYS[i];
            if (urlParams[authKey]) {
                cookieState.set(authKey, urlParams[authKey]);
            }
        }

        if (hasResponseToken) {
            cookieState.set(AK_EXPIRES, currentTs + parseInt(urlParams[AK_EXPIRES_IN]));
            cookieState.remove(AK_ERROR);
            cookieState.remove(AK_ERROR_DESC)
        }
        else if (!hasToken) {
            if (urlParams[AK_ERROR] === ERROR_ACCESS_DENIED) {
                status = AS_NOTCONNECTED;
            }
        }
    }
    else {
        // We are in cookie mode.
        if (shouldWriteCookie) {
            var cookieErrorMsg = diagnoseAuthCookieState(cookieState);
            if (cookieErrorMsg) {
                cookieState.set(AK_ERROR, ERROR_COOKIE_ERROR);
                cookieState.set(AK_ERROR_DESC, cookieErrorMsg);
            }
        }
        else {
            // The cookie has already been written, so skip writting again.
            return;
        }
    }

    cookieState.set(AK_STATUS, status);
    cookieState.save();
}

function handlePageAction() {
    var pageState = wl_app._pageState,
        redirectType = pageState[REDIRECT_TYPE];

    if (redirectType === REDIRECT_TYPE_UPLOAD) {
        var id = pageState[UPLOAD_STATE_ID];
        var result = wl_app._urlParams[API_PARAM_RESULT];
        handleUploadRedirect(id, result);
        return;
    }

    var display = pageState[AK_DISPLAY],
        secureCookie = (pageState[AK_SECURE_COOKIE] === "true");

    wl_app._logEnabled = true;
    wl_app._traceEnabled = pageState[WL_TRACE] || wl_app._urlParams[WL_TRACE];
    wl_app._secureCookie = secureCookie;

    detectSecureConnection();

    if (display === DISPLAY_PAGE || display === DISPLAY_TOUCH) {
        saveServerResponse();

        if (display === DISPLAY_TOUCH && wl_app._browser.ie) {
            // For mobile IE, we do navigation.
            var redirLocation = pageState[AK_REDIRECT_URI];
            validateRedirectUrl(redirLocation);
            document.location = redirLocation;
        } else {
            // For popup window, we close it.
            window.close();
        }
    }
    else if (display === DISPLAY_NONE) {
        saveServerResponse();
    }
    else {
        // Neither popup, nor hidden, this is the app page.
        checkDocumentReady(onDocumentReady);

        // Invoke wlAsyncInit, if it is defined.
        var appInit = window.wlAsyncInit;
        if (appInit && (typeof (appInit) === TYPE_FUNCTION)) {
            appInit.call();
        }
    }
}

// ensure the url is an absolute http(s) url with
// a matching domain to the current one, or a sub-domain.
function validateRedirectUrl(url) {
    if (url != null) {
        var uri = new Uri(url);
        if (uri.scheme != "") {
            var currentHost = window.location.host;
            var redirHost = uri.host;
            if (redirHost == currentHost) {
                // the hosts match exactly
                return;
            }

            // for non-matching hosts, only allow it if the redirect url is a subdomain
            currentHost = '.' + currentHost;
            if (redirHost.indexOf(currentHost, redirHost.length - currentHost.length) !== -1) {
                return;
            }
        }

        // all other cases fail
        throw new Error(ERROR_DESC_REDIRECTURI_INVALID_WWA.replace("WL.init", "WL.login"));
    }
}

function normalizeRedirectUrl(url, method) {
    if (!url) {
        // We use current page as redirect_uri if not provided
        url = trimUrlQuery(window.location.href);
    }

    return normalizeRedirectUrlAndValidateHost(url, window.location.hostname, method);
}

function normalizeRedirectUrlAndValidateHost(url, host, method) {
    var uri = new Uri(url);
    // Resolves relative Url, if it is.
    uri.resolve();
    var host = host.split(":")[0].toLowerCase(),
        redirectHost = uri.host.toLowerCase();

    wl_app._domain = wl_app._domain || redirectHost;

    if (wl_app._isHttps && uri.scheme == SCHEME_HTTP) {
        throw new Error(ERROR_DESC_URL_SSL.replace("METHOD", method));
    }

    return uri.toString();
}

function diagnoseAuthCookieState(cookieState) {
    var hasToken = cookieState.get(AK_ACCESS_TOKEN) != null,
        hasError = cookieState.get(AK_ERROR) != null,
        hasScope = cookieState.get(AK_SCOPE) != null,
        hasExpiresIn = cookieState.get(AK_EXPIRES_IN) != null,
        hasClientId = cookieState.get(AK_CLIENT_ID) != null,
        error = null;

    if (!(hasToken && hasScope && hasExpiresIn) && !hasError) {
        logError(ERROR_DESC_COOKIE_INVALID);
        error = ERROR_DESC_COOKIE_INVALID;
    }

    if (!hasClientId) {
        logError(ERROR_DESC_COOKIE_OVERWRITE);
        error = ERROR_DESC_COOKIE_OVERWRITE;
    }

    return error;
}

/**
 * The Web version of handlePageLoad() method.
 */
function handlePageLoad() {

    API_JSONP_URL_LIMIT = wl_app._browser.ie ? 2000 : 4000;

    processUrlParameters();
    handlePageAction();
}

function handleUploadRedirect(id, result) {
    var uploadCookie = new CookieState(COOKIE_UPLOAD);
    uploadCookie.load();
    uploadCookie.set(id, result);
    uploadCookie.save();
}

/**
 * The Web version of appInitPlatformSpecific() method.
 */
function appInitPlatformSpecific(properties) {
    wl_app._authScope = normalizeScopeValue(properties[AK_SCOPE]);
    wl_app._secureCookie = normalizeBooleanValue(properties[AK_SECURE_COOKIE]);
    wl_app._redirect_uri = normalizeRedirectUrl(properties[AK_REDIRECT_URI], IMETHOD_WL_INIT);
    wl_app._response_type = (properties[AK_RESPONSE_TYPE] || RESPONSE_TYPE_TOKEN).toLowerCase();
    wl_app._appId = properties[AK_CLIENT_ID];
    wl_app._refreshType = (properties[AK_REFRESH_TYPE] || AK_REFRESH_TYPE_MS).toLowerCase();

    var authSession = new AuthSession(properties[AK_CLIENT_ID], COOKIE_AUTH);
    wl_app._session = authSession;

    var sessionStatus = authSession.getNormalStatus(),
        status = sessionStatus[AK_STATUS],
        promise,
        promisOpName = IMETHOD_WL_INIT;

    if (status == AS_CONNECTED) {
        wl_event.notify(EVENT_AUTH_SESSIONCHANGE, sessionStatus);
        wl_event.notify(EVENT_AUTH_STATUSCHANGE, sessionStatus);
        wl_event.notify(EVENT_AUTH_LOGIN, sessionStatus);

        promise = createCompletePromise(promisOpName, true/*succeeded*/, properties.callback, sessionStatus);
    }
    else if (properties[AK_STATUS]) {
        promise = new Promise(promisOpName, null, null);
        wl_app.getLoginStatus(
            {
                internal: true,
                callback: function (resp) {
                    var eventType = !!(resp.error) ? PROMISE_EVENT_ONERROR : PROMISE_EVENT_ONSUCCESS;
                    promise[eventType](resp);
                }
            }, true/*force*/);
    }

    return promise;
}


/**
 * The Web version of handlePendingLogin() method.
 */
function handlePendingLogin(internal) {
    var pendingRequest = wl_app._pendingLogin;
    if (pendingRequest != null) {
        pendingRequest.cancel();
        wl_app._pendingLogin = null;
    }

    return true;
}

/**
 * Normalize login scope.
 */
function normalizeLoginScope(properties) {
    var scope = normalizeScopeValue(properties[AK_SCOPE]);
    if (scope === "") {
        scope = wl_app._authScope;
    }

    if (!scope || scope === "") {
        throw createMissingParamError(AK_SCOPE, IMETHOD_WL_LOGIN);
    }

    properties.normalizedScope = scope;
}

/**
 * The Web version of createLoginRequest() method.
 */
function createLoginRequest(properties, onAuthRequestCompleted, isExternalConsentRequest) {
    return new AuthRequest(AUTH_REQUEST_LOGIN, properties, onAuthRequestCompleted, isExternalConsentRequest);
}

/**
 * The Web version of createLoginStatusRequest() method.
 */
function createLoginStatusRequest(properties, onGetLoginStatusCompleted) {
    return new AuthRequest(AUTH_REQUEST_LOGINSTATUS, properties, onGetLoginStatusCompleted);
}

/**
 * This method will do:
 *  1) show consent UI if the current session does not have required scope.
 *  2) refresh the access_token, if its valid period is smaller than what's required.
 */
wl_app.ensurePermission = function (scope, validTime, method, callback) {
    var permissionError = createErrorResponse(ERROR_ACCESS_DENIED,
                                              ERROR_DESC_ACCESS_DENIED.replace("METHOD", method));
    wl_app.login({ scope: scope }).then(
        function (resp) {
            if (resp.session[AK_EXPIRES_IN] < validTime) {
                // Needs to extend the ticket
                wl_app.getLoginStatus({ internal: true }, true/*force*/).then(
                    function (resp) {
                        // The ticket may have changed, re-check the scope
                        wl_app.login({ scope: scope }).then(
                            function (resp) {
                                // Success
                                callback(resp);
                            },
                            function (resp) {
                                // Failed to acquire user permission
                                callback(permissionError);
                            });
                    },
                    function (resp) {
                        // failed to extend the ticket
                        callback(permissionError);
                    }
                );
            }
            else {
                // Success
                callback(resp);
            }
        },
        function (resp) {
            // Failed to acquire user permission
            callback(permissionError);
        }
    );
};


wl_app.canLogout = function () {
    return true;
};

/**
 * The Web version of logoutWindowsLive() method.
 */
function logoutWindowsLive(callback) {

    cleanLogoutFrame();

    var logoutFrame = createHiddenElement("iframe"),
        authServer = getAuthServerName(),
        path = "/oauth20_logout.srf?ts=";
    logoutFrame.src = "//" + authServer + path + new Date().getTime();
    document.body.appendChild(logoutFrame);
    wl_app.logoutFrame = logoutFrame;

    // Clean logout iframe in 30s.
    window.setTimeout(function () {
        cleanLogoutFrame();
        callback();
    }, 30000);
}

function cleanLogoutFrame() {
    if (wl_app.logoutFrame != null) {
        document.body.removeChild(wl_app.logoutFrame);
        wl_app.logoutFrame = null;
    }
}

function handleUIParameterError(properties, err) {
    logError(err.message);
    var onerror = properties[UI_PARAM_ONERROR];
    if (onerror) {
        delayInvoke(function () {
            error = createExceptionResponse(IMETHOD_WL_UI, IMETHOD_WL_UI, err),
            onerror(error);
        });
    }
}

function getSDKRootPath() {
    return wl_app[WL_SDK_ROOT];
}

function getImagePath() {
    return getSDKRootPath() + "images";
}

var SignInControl = function (properties) {

    var control = this;

    control._properties = properties;

    var signInControlInit = createDelegate(control, control.init);

    checkDocumentReady(signInControlInit);
};

SignInControl.prototype = {
    init: function () {
        var control = this,
            properties = control._properties;

        if (control._inited === true) {
            return;
        }

        control._inited = true;

        try {
            control.validate();
            var element = properties[UI_PARAM_ELEMENT],
                type = properties[UI_PARAM_TYPE],
                callback = properties[API_PARAM_CALLBACK],
                signinText = properties[UI_PARAM_SIGN_IN_TEXT],
                signoutText = properties[UI_PARAM_SIGN_OUT_TEXT];

            normalizeSignInControlScope(properties);

            element = (typeof (element) === TYPE_STRING) ? getElementById(properties[UI_PARAM_ELEMENT]) : element;
            control._element = element;

            type = type != null ? type : UI_SIGNIN_TYPE_SIGNIN;
            if (type == UI_SIGNIN_TYPE_SIGNIN) {
                signinText = WLText.signIn;
                signoutText = WLText.signOut;
            }
            else if (type == UI_SIGNIN_TYPE_LOGIN) {
                signinText = WLText.login;
                signoutText = WLText.logout;
            }
            else if (type == UI_SIGNIN_TYPE_CONNECT) {
                signinText = WLText.connect;
                signoutText = WLText.signOut;
            }

            control[UI_PARAM_SIGN_IN_TEXT] = signinText;
            control[UI_PARAM_SIGN_OUT_TEXT] = signoutText;

            setInnerHtml(element, buildSignInControlHtml(properties));

            var isSignedIn = wl_app._session.isSignedIn(),
                buttonText = isSignedIn ? signoutText : signinText;
            control.updateUI(buttonText, isSignedIn);

            attachSignInControlMouseEvents(control, element.childNodes[0]);

            wl_event.subscribe(EVENT_AUTH_LOGIN, createDelegate(control, control.onLoggedIn));
            wl_event.subscribe(EVENT_AUTH_LOGOUT, createDelegate(control, control.onLoggedOut));

            wl_app.getLoginStatus({ internal: true });

            // The callback should only be invoked once for rendering complete.
            // To avoid conflict with the login callback parameter, remove it here.
            delete properties[API_PARAM_CALLBACK];

            invokeCallback(callback, properties, false/*synchronous*/);
        }
        catch (e) {
            handleUIParameterError(properties, e);
        }
    },

    validate: function () {
        var properties = this._properties;
        validateProperties(
            properties,
            [{
                name: UI_PARAM_ELEMENT,
                type: TYPE_DOM,
                optional: false
            },
             {
                 name: UI_PARAM_TYPE,
                 allowedValues: [UI_SIGNIN_TYPE_SIGNIN, UI_SIGNIN_TYPE_LOGIN, UI_SIGNIN_TYPE_CONNECT, UI_SIGNIN_TYPE_CUSTOM],
                 type: TYPE_STRING,
                 optional: true
             },
             { name: AK_SCOPE, type: TYPE_STRINGORARRAY, optional: true },
             { name: AK_STATE, type: TYPE_STRING, optional: true },
             { name: UI_PARAM_ONLOGGEDIN, type: TYPE_FUNCTION, optional: true },
             { name: UI_PARAM_ONLOGGEDOUT, type: TYPE_FUNCTION, optional: true },
             { name: UI_PARAM_ONERROR, type: TYPE_FUNCTION, optional: true }
            ],
            "WL.ui-signin");

        validateSignInControlPlatformSpecificParameters(properties);

        // Validate custom sign-in control text values
        var type = properties[UI_PARAM_TYPE];
        if (type == UI_SIGNIN_TYPE_CUSTOM) {
            validateProperties(
                properties,
                [{
                    name: UI_PARAM_SIGN_IN_TEXT,
                    type: TYPE_STRING,
                    optional: false
                },
                 {
                     name: UI_PARAM_SIGN_OUT_TEXT,
                     type: TYPE_STRING,
                     optional: false
                 }
                ],
                "WL.ui-signin");
        }
    },

    fireEvent: function (eventName, args) {
        var callback = this._properties[eventName];
        if (callback) {
            callback(args);
        }
    },

    onClick: function () {
        var ctrl = this;
        if (ctrl._element.childNodes.length == 0) {
            // The button has been cleared.
            detachSignInControlMouseEvents(ctrl);
            return false;
        }

        if (wl_app._session.isSignedIn()) {
            if (wl_app.canLogout()) {
                wl_app.logout({});
            }
        }
        else {
            wl_app.login(ctrl._properties, true/*internal*/).then(
                function (result) { },
                function (result) {
                    ctrl.fireEvent(UI_PARAM_ONERROR, result);
                });
        }

        return false;
    },

    onLoggedIn: function (e) {
        this.updateUI(this[UI_PARAM_SIGN_OUT_TEXT], true/*isSignedIn*/);
        this.fireEvent(UI_PARAM_ONLOGGEDIN, e);
    },

    onLoggedOut: function (e) {
        this.updateUI(this[UI_PARAM_SIGN_IN_TEXT], false/*isSignedIn*/);
        this.fireEvent(UI_PARAM_ONLOGGEDOUT, e);
    }
};

function normalizeSignInControlScope(properties) {
    if (wl_app._authScope && wl_app._authScope !== "") {
        // We use the scope values passed in from WL.init.
        // If it isn't available, the scope value from SignInControl will be used for backward compatibility.
        properties[AK_SCOPE] = wl_app._authScope;
    }

    if (!properties[AK_SCOPE]) {
        // If no scope is available, we use wl.signin as default for auth UI flow.
        properties[AK_SCOPE] = SCOPE_SIGNIN;
    }
}

function createSignInControlEventHandler(name, control, callback) {
    control._handlers = control._handlers || {};
    var handler = createDelegate(control, callback);
    control._handlers[name] = handler;
    return handler;
}

function getSignInControlEventHandler(name, control) {
    return control._handlers[name];
}

wl_app.ui = function (properties) {

    ensureAppInited(IMETHOD_WL_UI);

    switch (properties.name) {
        case UI_SIGNIN:
            new SignInControl(properties);
            break;
        case UI_SKYDRIVEPICKER:
            new SkyDrivePickerControl(properties);
            break;
    }
}

function validateSignInControlPlatformSpecificParameters(properties) {
    validateProperties(
        properties,
        [{
            name: UI_PARAM_THEME,
            allowedValues: [UI_SIGNIN_THEME_BLUE, UI_SIGNIN_THEME_WHITE],
            type: TYPE_STRING,
            optional: true
        },
        {
            name: UI_PARAM_BRAND,
            allowedValues: [UI_BRAND_MESSENGER, UI_BRAND_HOTMAIL, UI_BRAND_SKYDRIVE, UI_BRAND_WINDOWS, UI_BRAND_WINDOWSLIVE, UI_BRAND_NONE],
            type: TYPE_STRING,
            optional: true
        }],
        "WL.ui-signin");

    properties[UI_PARAM_THEME] = properties[UI_PARAM_THEME] || UI_SIGNIN_THEME_BLUE;
    properties[UI_PARAM_BRAND] = properties[UI_PARAM_BRAND] || UI_BRAND_WINDOWS;
}


function buildSignInControlHtml(properties) {

    var brand = (properties[UI_PARAM_BRAND]),
        theme = (properties[UI_PARAM_THEME]),
        locale = wl_app._locale,
        direction = (locale.indexOf("ar") > -1 || locale.indexOf("he") > -1) ? "RTL" : "LTR",
        buttonStyle = "cursor:pointer;background-color:transparent;border:solid 0px;display:inline-block;overflow:hidden;white-space:nowrap;padding:0px;width:auto;",
        buttonComponentStyle = "margin:0px;padding:0px;border-width:0px;vertical-align:middle;background-attachment:scroll;display:inline-block;white-space:nowrap;",
        leftStyle = getSignInImageStyle(brand, direction, theme, "left") + buttonComponentStyle,
        centerStyle = getSignInImageStyle(brand, direction, theme, "center") + buttonComponentStyle,
        rightStyle = getSignInImageStyle(brand, direction, theme, "right") + buttonComponentStyle;

    return "<button style=\"" + buttonStyle +
           "\"><span style=\"" + leftStyle +
           "\"></span><span style=\"" + centerStyle +
           "\"><span style=\"" + getSignInControlTextStyle(direction) +
           "\"></span></span><span style=\"" + rightStyle +
           "\"></span></button>";
}

function getSignInControlTextStyle(direction) {
    var b = wl_app._browser,
        ie67 = b.ie6 || b.ie7,
        ie89 = b.ie8 || b.ie9;
    var textPosition = (b.chrome || b.safari) ? "padding:2px 3px;margin:0px;" : "padding:1px 3px;margin:0px;",
        fontStyle = "font-family: Segoe UI, Verdana, Tahoma, Helvetica, Arial, sans-serif;",
        textDirection = "direction:" + direction.toLowerCase() + ";",
        textDecoration = "text-decoration:none;color:#3975a0;display:inline-block;",
        lineHeight = "150";

    switch (wl_app._locale) {
        case "ar-ploc-sa":
            if (!ie67) {
                lineHeight = "170";
            }
            break;
        case "te":
        case "ja-ploc-jp":
            if (b.ie) {
                lineHeight = "190";
            }
            break;
    }

    var textSize = "height:18px;font-size:9pt;font-weight:bold;line-height:" + lineHeight + "%;";

    return textPosition + textDirection + textDecoration + fontStyle + textSize;
}

function getSignInImageStyle(brand, direction, theme, position) {
    brand = (brand == UI_BRAND_WINDOWS) ? UI_BRAND_WINDOWSLIVE : brand;

    var imgName = brand + "_" + direction + "_" + theme,
        width,
        v_pos,
        repeat,
        signInBackGround = "background: url({imgpath}/signincontrol/{image}.png) scroll {repeat} 0px {vpos}; height: 22px; width: {width};";
    switch (position) {
        case "left":
            width = (brand === UI_BRAND_NONE || direction === "RTL") ? "3px" : "25px";
            v_pos = (direction === "LTR") ? "0px" : "-44px";
            repeat = "no-repeat";
            break;
        case "center":
            width = "auto";
            v_pos = "-22px";
            repeat = "repeat-x";
            break;
        case "right":
            width = (brand === UI_BRAND_NONE || direction === "LTR") ? "3px" : "25px";
            v_pos = (direction === "LTR") ? "-44px" : "0px";
            repeat = "no-repeat";
            break;
    }

    return signInBackGround.replace("{imgpath}", getImagePath()).replace("{image}", imgName).replace("{vpos}", v_pos).replace("{width}", width).replace("{repeat}", repeat);
}


SignInControl.prototype.updateUI = function (text, isSignedIn) {
    if (this._element.childNodes.length == 0) {
        // The button has been cleared.
        detachSignInControlMouseEvents(this);
        return;
    }

    if (text != this._text) {
        var browser = wl_app._browser,
                button = this._element.childNodes[0],
                textContainer = button.childNodes[1];

        setElementText(textContainer.childNodes[0], text);
        this._text = text;
        if (browser.ie6 || browser.ie7) {
            // Set width auto. IE6 does not shrink if text becomes shorter.
            textContainer.style.width = "auto";
            button.style.width = "auto";

            var leftWidth = button.childNodes[0].clientWidth,
                    middleWidth = button.childNodes[1].clientWidth,
                    rightWidth = button.childNodes[2].clientWidth,
                    buttonWidth = leftWidth + middleWidth + rightWidth;

            button.style.width = buttonWidth + "px";

            if (browser.ie6) {
                button.childNodes[0].style.width = leftWidth + "px";
                button.childNodes[1].style.width = middleWidth + "px";
                button.childNodes[2].style.width = rightWidth + "px";
            }
        }
    }
};

function attachSignInControlMouseEvents(control, button) {
    control._button = button;
    attachDOMEvent(button, "click", createSignInControlEventHandler("click", control, control.onClick));
}

function detachSignInControlMouseEvents(control) {
    var button = control._button;
    if (button) {
        detachDOMEvent(button, "click", getSignInControlEventHandler("click", control));
        delete control._button;
    }
}

var SkyDrivePickerControl = function (properties) {
    validateFileDialogCall(properties);

    var ctrl = this;
    ctrl._props = properties;

    checkDocumentReady(createDelegate(ctrl, ctrl.init));
};

SkyDrivePickerControl.prototype = {
    init: function () {
        var ctrl = this;
        if (ctrl._inited === true) {
            return;
        }

        ctrl._inited = true;

        try {
            var properties = ctrl._props,
                element = properties[UI_PARAM_ELEMENT],
                callback = properties[API_PARAM_CALLBACK];
            properties[API_INTERFACE_METHOD] = "WL.ui-" + UI_SKYDRIVEPICKER;
            ctrl.validate();

            properties[UI_PARAM_THEME] = properties[UI_PARAM_THEME] || UI_SIGNIN_THEME_WHITE;

            element = (typeof (element) === TYPE_STRING) ? getElementById(properties[UI_PARAM_ELEMENT]) : element;
            ctrl._element = element;

            var content = buildSkyDrivePickerControlHtml(properties);
            setInnerHtml(element, content);
            attachControlMouseEvents(element, createDelegate(ctrl, ctrl.onClick), createDelegate(ctrl, ctrl.onRender));

            ctrl.onRender(false, false);

            invokeCallback(callback, properties, false/*synchronous*/);
        }
        catch (e) {
            logError(e.message);
        }
    },

    validate: function () {
        var properties = this._props;
        validateProperties(
            properties,
            [{
                name: UI_PARAM_ELEMENT,
                type: TYPE_DOM,
                optional: false
            }],
            properties[API_INTERFACE_METHOD]);
    },

    onClick: function () {
        try {
            return wl_app.fileDialog(this._props);
        }
        catch (err) {
            // TODO: review error handling here
            logError(err.message);
        }

        return false;
    },

    onRender: function (mouseDown, mouseIn)
    {
        buildSkyDrivePickerControlStyle(this._props, this._element.childNodes[0]);
    }
};

function buildSkyDrivePickerControlStyle(properties, button)
{
    var themeWhite = properties[UI_PARAM_THEME] === UI_SIGNIN_THEME_WHITE,
        locale = wl_app._locale,
        rtl = (locale.indexOf("ar") > -1 || locale.indexOf("he") > -1),
        direction = rtl ? "RTL" : "LTR",
        img = button.childNodes[0],
        text = button.childNodes[1],
        blueCode = "#094AB2",
        whiteCode = "#ffffff",
        buttonStyle = button.style,
        imageStyle = img.style,
        textStyle = text.style;

    buttonStyle.direction = direction;
    buttonStyle.backgroundColor = themeWhite ? whiteCode : blueCode;
    buttonStyle.border = "solid 1px";
    buttonStyle.borderColor = blueCode;
    buttonStyle.height = "20px";
    buttonStyle.paddingLeft = HTML_BUTTON_PADDING;
    buttonStyle.paddingRight = HTML_BUTTON_PADDING;
    buttonStyle.textAlign = "center";
    buttonStyle.cursor = "pointer";

    // Link specific styling.
    if (button instanceof HTMLAnchorElement) {
        buttonStyle.lineHeight = "20px";
        buttonStyle.display = "inline-block";
        buttonStyle.textDecoration = "none";
    }

    imageStyle.verticalAlign = "middle";
    imageStyle.height = "16px";

    textStyle.fontFamily = "'Segoe UI', 'Segoe UI Web Regular', 'Helvetica Neue', 'BBAlpha Sans', 'S60 Sans', Arial, 'sans-serif'";
    textStyle.fontSize = "12px";
    textStyle.fontWeight = "bold";
    textStyle.color = themeWhite ? blueCode : whiteCode;
    textStyle.textAlign = "center";
    textStyle.verticalAlign = "middle";
    textStyle.marginLeft = rtl ? HTML_BUTTONTEXT_MARGIN_NONE : HTML_BUTTONTEXT_MARGIN;
    textStyle.marginRight = rtl ? HTML_BUTTONTEXT_MARGIN : HTML_BUTTONTEXT_MARGIN_NONE;
}

function buildSkyDrivePickerControlHtml(properties) {
    var button = buildSkyDrivePickerControlInnerHtml(properties, false /* old picker */);
    return "<button id='" + encodeHtml(button.buttonId) + "' title='" + encodeHtml(button.buttonTitle) + "'>" + button.innerHTML + "</button>";
}

function buildSkyDrivePickerControlInnerHtml(properties, newPicker) {
    var mode = properties[FILEDIALOG_PARAM_MODE],
        theme = properties[UI_PARAM_THEME],
        isOpenPicker = (mode === FILEDIALOG_PARAM_MODE_OPEN),
        buttonText = isOpenPicker ? WLText.skyDriveOpenPickerButtonText : WLText.skyDriveSavePickerButtonText,
        imgName = theme === UI_SIGNIN_THEME_BLUE ? "SkyDriveIcon_white.png" : "SkyDriveIcon_blue.png",
        imgHtml = "<img alt='' src='" + getImagePath() + "/SkyDrivePicker/" + imgName + "'>",
        textHtml = "<span>" + encodeHtml(buttonText) + "</span>",
        buttonTitle = isOpenPicker ? WLText.skyDriveOpenPickerButtonTooltip : WLText.skyDriveSavePickerButtonTooltip,
        openId = newPicker ? "onedriveopenpickerbutton" : "skydriveopenpickerbutton",
        saveId = newPicker ? "onedrivesavepickerbutton" : "skydrivesavepickerbutton",
        buttonId = isOpenPicker ? openId : saveId;

    return {
        innerHTML: imgHtml + textHtml,
        buttonTitle: buttonTitle,
        buttonId: buttonId
    };
}

function attachControlMouseEvents(container, onClick, onRender) {
    var button = container.childNodes[0];
    if (onClick) {
        var onMouseClick = function (e) {
            if (container.childNodes.length == 0) {
                detachDOMEvent(button, DOM_EVENT_CLICK, onMouseClick);
                return;
            }
            onClick(e);
        };
        attachDOMEvent(button, DOM_EVENT_CLICK, onMouseClick);
    }

    if (onRender) {
        var mouseDown = false, mouseIn = false;
        var onMouseEnter = function (e) {
            mouseIn = true;
            onMouseRender(e);
        };
        var onMouseLeave = function (e) {
            mouseIn = false;
            onMouseRender(e);
        };
        var onMouseDown = function (e) {
            mouseDown = true;
            onMouseRender(e);
        };
        var onMouseUp = function (e) {
            mouseDown = false;
            onMouseRender(e);
        };

        var onMouseRender = function (e) {
            if (container.childNodes.length == 0) {
                detachDOMEvent(button, "mouseenter", onMouseEnter);
                detachDOMEvent(button, "mouseleave", onMouseLeave);
                detachDOMEvent(document, "mousedown", onMouseDown);
                detachDOMEvent(document, "mouseup", onMouseUp);
                return;
            }
            onRender(mouseDown, mouseIn);
        };
        attachDOMEvent(button, "mouseenter", onMouseEnter);
        attachDOMEvent(button, "mouseleave", onMouseLeave);
        attachDOMEvent(document, "mousedown", onMouseDown);
        attachDOMEvent(document, "mouseup", onMouseUp);
    }
}

function getCookie(name) {
    var cookies = document.cookie;

    // Look for 'name='
    name += "=";

    var start = cookies.indexOf(name);
    if (start >= 0) {
        start += name.length;

        var end = cookies.indexOf(';', start);
        if (end < 0) {
            end = cookies.length;
        }
        else {
            postCookie = cookies.substring(end);
            if (postCookie.indexOf(name) >= 0) {
                throw new Error(ERROR_DESC_COOKIE_MULTIPLEVALUE);
            }
        }

        var value = cookies.substring(start, end);

        return value;
    }

    return "";
}

function setCookie(name, value, domain, secondsToExpiry) {
    value = value ? value : '';
    var cookie = name + "=" + value + "; path=/";
    if (domain && domain != "localhost") {
        cookie += "; domain=" + encodeURIComponent(domain);
    }

    if (secondsToExpiry != null) {
        var expires = new Date(0);

        if (secondsToExpiry > 0) {
            expires = new Date();
            expires.setTime(expires.getTime() + secondsToExpiry * 1000);
        }

        cookie += ";expires=" + expires.toUTCString();
    }

    if (wl_app._isHttps && wl_app._secureCookie) {
        cookie += ";secure";
    }

    document.cookie = cookie;
}

var CookieState = function (cookieName, properties) {
    this._cookieName = cookieName;
    this._states = properties || {};
    this._listeners = [];
    this._cookieWatcher = null;
};

CookieState.prototype = {

    getStates: function () {
        return this._states;
    },

    get: function (key) {
        return this._states[key];
    },

    set: function (key, value) {
        this._states[key] = value;
    },

    remove: function (key) {
        if (this._states[key]) {
            delete this._states[key];
        }
    },

    load: function () {
        try {
            var rawValue = getCookie(this._cookieName);
            if (this._rawValue != rawValue) {

                trace("Cookie changed: " + this._cookieName + "=" + rawValue);

                this._rawValue = rawValue;
                this._states = deserializeParameters(rawValue);
                for (var i = 0; i < this._listeners.length; i++) {
                    this._listeners[i]();
                }
            }
        }
        catch (error) {
            logError(error.message);
            this.stopMonitor();
        }
    },

    flush: function (data) {
        this._states = data;
        this.save();
    },

    populate: function (data) {
        cloneObject(data, this._states);
        this.save();
    },

    save: function () {
        setCookie(
            this._cookieName,
            serializeParameters(this._states),
            getCookieDomain(),  // store in the right domain
            null); // secondsToExpiry: set as session cookie.
    },

    clear: function () {
        this._states = {};
        setCookie(
            this._cookieName,
            "",
            getCookieDomain(),
            0);
    },

    addCookieChanged: function (onChanged) {
        this._listeners.push(onChanged);
        this.startMonitor();
    },

    startMonitor: function () {

        this._refreshInterval = 300;

        if (!this._cookieWatcher) {
            trace("Started monitoring cookie: " + this._cookieName);

            this._cookieWatcher = window.setInterval(createDelegate(this, this.load), this._refreshInterval);
        }
    },

    stopMonitor: function () {
        if (this._cookieWatcher) {
            trace("Stopped monitoring cookie: " + this._cookieName);

            window.clearInterval(this._cookieWatcher);
            this._cookieWatcher = null;
        }
    },

    isBeingMonitored: function() {
        return this._cookieWatcher !== null;
    }
};

/**
 * The Web version of getCookieDomain() method.
 */
function getCookieDomain() {
    var cookie_domain = wl_app._domain || window.location.hostname.split(":")[0];
    return cookie_domain;
}


var AuthRequest = function (method, properties, callback) {
    var request = this;
    request._method = method;
    request._completed = false;
    request._requestFired = false;
    request._properties = properties;
    request._callback = callback;
    request._authMonitor = createDelegate(request, request._onAuthChanged);

    request.execute = (request._method == AUTH_REQUEST_LOGIN) ? request._login : request._getLoginStatus;
};

AuthRequest.prototype = {

    cancel: function () {
        this._onComplete({ error: ERROR_REQ_CANCEL, error_description: ERROR_DESC_LOGIN_CANCEL });
    },

    _login: function () {
        var request = this;
        request._requestTs = new Date().getTime();
        var mobile = wl_app._browser.mobile,
            navigate = (mobile && wl_app._browser.ie),
            display = mobile ? DISPLAY_TOUCH : DISPLAY_PAGE,
            url = buildAuthUrl(display, request._properties.normalizedScope, getAuthRedirectUri(request._properties), request._requestTs, navigate, request._properties.state);

        if (navigate) {
            document.location = url;
        }
        else {
            request._popup = openPopUp(url, request._properties.external_consent);

            trace("AuthRequest-login: popup is opened. " + request._popup);

            request._popupWatcher = window.setInterval(createDelegate(request, request._checkPopup), 3000);

            wl_event.subscribe(EVENT_AUTH_RESPONSE, request._authMonitor);
        }

        request._promise = new Promise(IMETHOD_WL_LOGIN, null, null);
        return request._promise;
    },

    _getLoginStatus: function () {
        onCreateIframeReady(createDelegate(this, this._fireStatusRequest));
        this._timeout = window.setTimeout(createDelegate(this, this._onTimedOut), MAX_GETLOGINSTATUS_TIME);

        this._promise = new Promise("WL.getLoginStatus", null, null);
        return this._promise;
    },

    _fireStatusRequest: function () {
        var req = this;
        if (!req._requestFired) {
            req._requestTs = new Date().getTime();
            var url = (wl_app._refreshType === AK_REFRESH_TYPE_MS) ?
                buildAuthUrl(DISPLAY_NONE, wl_app._authScope, getAuthRedirectUri(), req._requestTs, false/*navigate*/) :
                buildAppRefreshUrl(getAuthRedirectUri(), wl_app._authScope, req._requestTs);

            req._statusFrame = createHiddenIframe(url);
            wl_event.subscribe(EVENT_AUTH_RESPONSE, req._authMonitor);
            req._requestFired = true;
        }
    },

    _onAuthChanged: function () {
        var response = wl_app._session.tryGetResponse(this._properties.normalizedScope, this._requestTs, this._properties.external_consent);
        if (response != null) {
            this._onComplete(this._normalizeResp(response));
        }
    },

    _normalizeResp: function (resp) {
        if (this._method === AUTH_REQUEST_LOGINSTATUS &&
            resp.error === ERROR_ACCESS_DENIED) {
            // Access_denied is not considered an error for GetLoginStatus() call.
            return wl_app._session.getNormalStatus();
        }

        return resp;
    },

    _onTimedOut: function () {
        this._onComplete({ error: ERROR_TIMEDOUT, error_description: ERROR_TRACE_AUTH_TIMEOUT });
    },

    _checkPopup: function () {
        try {
            if (this._popup === null) {
                this._onComplete({ error: ERROR_ACCESS_DENIED, error_description: ERROR_TRACE_AUTH_CLOSE });
            }
            else if (this._popup.closed) {
                // Give another round to wait for the cookie monitor to pickup.
                this._popup = null;
            }
        }
        catch (error) {
            trace("AuthRequest-checkPopup-error: " + error);
        }
    },

    _onComplete: function (response) {
        if (!this._completed) {
            this._completed = true;
            this._dispose();

            this._callback(this._properties, response);

            if (response[AK_ERROR]) {
                this._promise[PROMISE_EVENT_ONERROR](response);
            }
            else {
                this._promise[PROMISE_EVENT_ONSUCCESS](response);
            }
        }
    },

    _dispose: function () {

        trace("AuthRequest: dispose " + this._method);

        if (this._timeout) {
            window.clearTimeout(this._timeout);
            this._timeout = null;
        }

        if (this._statusFrame != null) {
            document.body.removeChild(this._statusFrame)
            this._statusFrame = null;
        }

        if (this._popupWatcher) {
            window.clearInterval(this._popupWatcher);
            this._popupWatcher = null;
        }

        wl_event.unsubscribe(EVENT_AUTH_RESPONSE, this._authMonitor);
    }
};

function openPopUp(url, isExternalConsent) {
    var width = 525, height = 525, top, left;

    if (isExternalConsent) {
        width = 800;
        height = 650;
    }

    if (wl_app._browser.ie) {
        var screenLeft = window.screenLeft,
            screenTop = window.screenTop,
            docElement = document.documentElement,
            titleHeight = 30;

        left = screenLeft + (Math.max(docElement.clientWidth - width, 0)) / 2;
        top = screenTop + (Math.max(docElement.clientHeight - height, 0)) / 2 - titleHeight;
    }
    else {
        var screenX = window.screenX,
            screenY = window.screenY,
            outerWidth = window.outerWidth,
            outerHeight = window.outerHeight;

        left = screenX + Math.max(outerWidth - width, 0) / 2;
        top = screenY + Math.max(outerHeight - height, 0) / 2;
    }

    var features = [
                "width=" + width,
                "height=" + height,
                "top=" + top,
                "left=" + left,
                "status=no",
                "resizable=yes",
                "toolbar=no",
                "menubar=no",
                "scrollbars=yes"];

    var popup = window.open(url, "oauth", features.join(","));
        popup.focus();

    return popup;
}

function buildAppRefreshUrl(baseUrl, scope, requestTs) {
    // Configure pageState.
    var pageState = {};
    pageState[REDIRECT_TYPE] = REDIRECT_TYPE_AUTH;
    pageState[AK_REQUEST_TS] = requestTs;
    pageState[AK_SECURE_COOKIE] = wl_app._secureCookie;
    pageState[AK_DISPLAY] = DISPLAY_NONE;
    pageState[AK_RESPONSE_METHOD] = AK_RESPONSE_METHOD_COOKIE; // The app page will write cookie response.

    if (wl_app.trace) {
        pageState[WL_TRACE] = true;
    }

    var state = serializeParameters(pageState),
        params = {};
    params[AK_CLIENT_ID] = wl_app._session.get(AK_CLIENT_ID);
    params[AK_RESPONSE_TYPE] = RESPONSE_TYPE_TOKEN;
    params[AK_SCOPE] = scope;
    params[AK_STATE] = state;

    var hasQuery = baseUrl.indexOf('?') > 0,
        concatChar = hasQuery ? '&' : '?',
        url = baseUrl + concatChar + serializeParameters(params);

    return url;
}

function buildAuthUrl(display, scope, redirectUrl, requestTs, navigate, appState) {
    // Configure pageState.
    var pageState = {};
    pageState[REDIRECT_TYPE] = REDIRECT_TYPE_AUTH;
    pageState[AK_DISPLAY] = display;
    pageState[AK_REQUEST_TS] = requestTs;
    if (navigate) {
        pageState[AK_REDIRECT_URI] = window.location.href;
    }

    if (wl_app.trace) {
        pageState[WL_TRACE] = true;
    }

    if (appState) {
        pageState[AK_APPSTATE] = appState;
    }

    // For the response_type=code case, the app server will transmit back the response via wl_auth cookie.
    // For the response_type=token case, the response will come from redirect uri parameters.
    var responseType = (display === DISPLAY_NONE) ? RESPONSE_TYPE_TOKEN : wl_app._response_type,
        responseMethod = (responseType === RESPONSE_TYPE_TOKEN) ? AK_RESPONSE_METHOD_URL : AK_RESPONSE_METHOD_COOKIE;

    pageState[AK_RESPONSE_METHOD] = responseMethod;
    pageState[AK_SECURE_COOKIE] = wl_app._secureCookie;

    var state = serializeParameters(pageState),
        params = {};
    params[AK_CLIENT_ID] = wl_app._session.get(AK_CLIENT_ID);
    params[AK_DISPLAY] = display;
    params[AK_LOCALE] = wl_app._locale;
    params[AK_REDIRECT_URI] = redirectUrl;
    params[AK_RESPONSE_TYPE] = responseType;
    params[AK_SCOPE] = scope;
    params[AK_STATE] = state;

    var authServer = getAuthServerName(),
        url = "https://" + authServer + "/oauth20_authorize.srf?" + serializeParameters(params);

    return url;
}

function getAuthRedirectUri(properties) {
    var redirect_uri = (properties != null) ? properties[AK_REDIRECT_URI] : null;
    return (redirect_uri != null && redirect_uri != "") ? redirect_uri : wl_app._redirect_uri;
}

var AuthSession = function (client_id, cookieName) {
    this._state = {};
    this._state[AK_CLIENT_ID] = client_id;
    this._state[AK_STATUS] = AS_UNCHECKED;
    this._cookieName = cookieName;
    this.init();
};

AuthSession.prototype = {
    get: function (key) {
        return this._state[key];
    },

    save: function () {
        if (this._stateDirty) {
            this._cookie.flush(this._state);
            this._stateDirty = false;
        }
    },

    init: function () {
        var cookieState = new CookieState(this._cookieName);

        cookieState.load();
        this._cookie = cookieState;

        if (cookieState.get(AK_CLIENT_ID) != this._state[AK_CLIENT_ID]) {
            cookieState.clear();
            cookieState.flush(this._state);
        }
        else {
            this._state = cookieState.getStates();
        }

        // If the cookie indicates there is a user ticket, we mark statusChecked as true
        // so that we may skip sending getLoginStatus request to the oauth server.
        // There are two cases on the web scenario: 1) we never checked before 2) the user status is unknown from previous page.
        var status = this._state[AK_STATUS];
        this._statusChecked = (status !== AS_UNKNOWN && status !== AS_UNCHECKED);
        this._cookie.addCookieChanged(createDelegate(this, this.onCookieChanged));
    },

    refresh: function () {
        wl_app.getLoginStatus({ internal: true }, true/*force*/);
        this._refresh = undefined;
    },

    scheduleRefresh: function () {
        this.cancelRefresh();
        var refresh_in = (this.tokenExpiresIn() - 600) * 1000;
        if (refresh_in > 0) {
            this._refresh = window.setTimeout(createDelegate(this, this.refresh), refresh_in);
        }
    },

    cancelRefresh: function () {
        if (this._refresh) {
            window.clearTimeout(this._refresh);
            this._refresh = undefined;
        }
    },

    updateStatus: function (status) {
        var oldStatus = this._state[AK_STATUS],
            oldToken = this._state[AK_ACCESS_TOKEN];
        if (oldStatus != status) {
            this._state[AK_STATUS] = status;
            this._stateDirty = true;
            this.onStatusChanged(oldStatus, status);
            this.save();

            if (oldToken != this._state[AK_ACCESS_TOKEN]) {
                wl_event.notify(EVENT_AUTH_SESSIONCHANGE, this.getNormalStatus());
            }
        }
    },

    onStatusChanged: function (oldStatus, newStatus) {

        trace("AuthSession: Auth status changed: " + oldStatus + "=>" + newStatus);

        if (oldStatus != newStatus) {
            var wasSignedin = (oldStatus == AS_CONNECTED),
            isSignedIn = (newStatus == AS_CONNECTED);

            if (!isSignedIn) {
                // Clear the session data, if the user is signed out.
                for (var i = 0; i < AK_COOKIE_KEYS.length; i++) {
                    var authKey = AK_COOKIE_KEYS[i];
                    if (this._state[authKey]) {
                        delete this._state[authKey];
                    }
                }

                this._stateDirty = true;
                this.save();
            }

            if (normalizeResponseStatus(oldStatus) != normalizeResponseStatus(newStatus)) {
                wl_event.notify(EVENT_AUTH_STATUSCHANGE, this.getNormalStatus());
            }

            if (isSignedIn != wasSignedin) {
                if (isSignedIn) {
                    wl_event.notify(EVENT_AUTH_LOGIN, this.getNormalStatus());
                }
                else {
                    wl_event.notify(EVENT_AUTH_LOGOUT, this.getNormalStatus());
                }
            }
        }
    },

    isSignedIn: function () {
        return this._state[AK_STATUS] === AS_CONNECTED;
    },

    getNormalStatus: function () {
        var authStatus = this.getStatus();
        authStatus[AK_STATUS] = normalizeResponseStatus(authStatus[AK_STATUS]);
        return authStatus;
    },

    tokenExpiresIn: function () {
        var state = this._state,
            status = state[AK_STATUS],
            expires = parseInt(state[AK_EXPIRES]);

        if (status === AS_CONNECTED) {
            return expires - getCurrentSeconds();
        }

        return -1;
    },

    onCookieChanged: function () {
        var oldState = this._state,
            newState = this._cookie.getStates();
        this._state = newState;

        trace("AuthSession: cookie changed. Has token: " + (newState[AK_ACCESS_TOKEN] != null));

        this._statusChecked = (newState[AK_STATUS] !== AS_UNCHECKED);

        if (oldState[AK_ACCESS_TOKEN] != newState[AK_ACCESS_TOKEN] ||
            oldState[AK_ERROR] != newState[AK_ERROR] ||
            oldState[AK_REQUEST_TS] != newState[AK_REQUEST_TS]) {

            wl_event.notify(EVENT_AUTH_RESPONSE);

            // We have notified auth.response events, so clean up errors from the cookie.
            delete newState[AK_ERROR];
            delete newState[AK_ERROR_DESC];
            this._stateDirty = true;
        }

        // Process state change
        if (oldState[AK_STATUS] != newState[AK_STATUS]) {
            this.onStatusChanged(oldState[AK_STATUS], newState[AK_STATUS]);
        }

        if (oldState[AK_ACCESS_TOKEN] != newState[AK_ACCESS_TOKEN]) {
            wl_event.notify(EVENT_AUTH_SESSIONCHANGE, this.getNormalStatus());

            if (newState[AK_ACCESS_TOKEN]) {
                this.scheduleRefresh();
            }
            else {
                this.cancelRefresh();
            }
        }

        this.save();
    },

    getStatus: function () {
        var session = null,
            status = this._state[AK_STATUS],
            resourceId = null,
            ownerCid = null,
            itemId = null,
            authKey = null;

        if (status === AS_CONNECTED) {
            var expiresIn = this.tokenExpiresIn();
            // We leave 10s for expiring state enough to make a call for api
            if (expiresIn <= 10) {
                // Token expired, update the status
                status = this._statusChecked ? AS_UNKNOWN : AS_UNCHECKED;
                this.updateStatus(status);

                // In a case we weren't able to refresh token, schedule once more.
                window.setTimeout(function () {
                    wl_app.getLoginStatus({ internal: true }, true/*force*/);
                }, 30000);
            }
            else {
                if (expiresIn < 60) {
                    status = AS_EXPIRING;
                }

                session = {};
                session[AK_ACCESS_TOKEN] = this._state[AK_ACCESS_TOKEN];
                session[AK_AUTH_TOKEN] = this._state[AK_AUTH_TOKEN];

                var scopes = this._state[AK_SCOPE].split(" ");
                session[AK_SCOPE] = [];

                for (var i = 0; i < scopes.length; i++) {
                    var result = scopeResponsePattern.exec(scopes[i]);
                    if (result)
                    {
                        // This is a response to an external consent request.
                        // Don't store this scope in session.
                        var rawResult = result[1].split("_"),
                            selectionString = rawResult[0];

                        itemId = rawResult[1];
                        ownerCid = itemId.split("!")[0];
                        resourceId = [selectionString, ownerCid, itemId].join(".");
                        authKey = result[2];
                    }
                    else {
                        session[AK_SCOPE].push(scopes[i]);
                    }
                }

                this._state[AK_SCOPE] = session[AK_SCOPE].join(" ");
                session[AK_EXPIRES_IN] = expiresIn;
                session[AK_EXPIRES] = this._state[AK_EXPIRES];
            }
        }
        else {
            if (!this._statusChecked) {
                status = AS_UNCHECKED;
            }
        }

        return { status: status, session: session, resource_id: resourceId, owner_cid: ownerCid, item_id: itemId, authentication_key: authKey };
    },

    tryGetResponse: function (scope, requestTs, isExternalConsentRequest) {

        trace("AuthSession.tryGetResponse: requestTs = " + requestTs + " scope = " + scope);

        var sessionStatus = this.getStatus(),
            status = sessionStatus[AK_STATUS],
            resourceId = sessionStatus[AK_RESOURCEID],
            session = sessionStatus[AK_SESSION];

        if (status == AS_UNCHECKED || status == AS_EXPIRING) {
            // We haven't checked yet or the ticket is going to expire.
            return null;
        }

        if (requestTs === undefined) {
            // Without requestTs, this is initial try for login()/getLoginStatus() in order to determine
            // if we need to send a request to server.
            if (scope) {
                // Assuming scope is always needed for login() call.
                return (session && isScopeSatisfied(session[AK_SCOPE], scope)) ? sessionStatus : null;
            }
            else {
                // For loginStatus, just return whatever we have, since it is not expiring or unchecked.
                return sessionStatus;
            }
        }
        // With requestTs, it means we are checking a server response. We need to compare the requestTs.
        var state = this._state,
            lastReqTs = parseInt(state[AK_REQUEST_TS]),
            errorCode = state[AK_ERROR],
            errorMsg = state[AK_ERROR_DESC];

        if (lastReqTs >= requestTs) {

            // lastReqTs indicates that we already have a response.
            if (session && ((!errorCode && isExternalConsentRequest) || isScopeSatisfied(session[AK_SCOPE], scope))) {
                return sessionStatus;
            }

            if (errorCode) {
                return createAuthError(errorCode, errorMsg);
            }

            if (!scope) {
                return sessionStatus;
            }
        }

        return null;
    }
};

function isScopeSatisfied(existingScopeValue, requestingScopeValue) {
    if (requestingScopeValue == null || stringTrim(requestingScopeValue) == "") {
        return true;
    }

    var requestingScopes = requestingScopeValue.split(SCOPE_DELIMINATOR);

    for (var i = 0; i < requestingScopes.length; i++) {
        var requestingScope = stringTrim(requestingScopes[i]);
        if (requestingScope != "" && !arrayContains(existingScopeValue, requestingScope)) {
            return false;
        }
    }

    return true;
}

function normalizeResponseStatus(status) {
    return (status === AS_UNCHECKED) ? AS_UNKNOWN : ((status === AS_EXPIRING) ? AS_CONNECTED : status);
}



/**
 * The Web version of executeApiRequest() method.
 */
function executeApiRequest(request) {
    if (!request._properties[API_PARAM_VROOMAPI] && sendAPIRequestViaJSONP(request)) {
        return;
    }

    if (sendAPIRequestViaXHR(request)) {
        return;
    }

    if (sendAPIRequestViaFlash(request)) {
        return;
    }

    var errorObj = {};
    errorObj[API_PARAM_CODE] = ERROR_REQUEST_FAILED;
    errorObj[API_PARAM_MESSAGE] = ERROR_DESC_BROWSER_LIMIT;
    request.onCompleted(errorObj);
}

/**
 * The Web version of canDoXHR() method.
 */
function canDoXHR() {
    return (window.XMLHttpRequest && (new XMLHttpRequest()).withCredentials !== undefined);
}

/**
 * The Web version of getAuthServerName() method.
 */
function getAuthServerName() {
    return wl_app[WL_AUTH_SERVER];
}

/**
 * The Web version of getApiServiceUrl() method.
 */
function getApiServiceUrl(useVroomApi) {
    return useVroomApi ? wl_app[WL_ONEDRIVE_API] : wl_app[WL_APISERVICE_URI];
}


// wl.app.download.web.js contains the implementation details
// for the Web WL.download method.

function validateDownloadProperties(properties) {
    validateProperties(
        properties,
        [{ name: API_PARAM_PATH, type: TYPE_STRING, optional: false }],
        properties[API_INTERFACE_METHOD]);
}

function executeDownload(op) {
    var props = op._properties,
        path = props[API_PARAM_PATH];

    startDownloadViaIFrame(path, op);
}

// Anonymous function, because we do not want downloadIFrame to be in the global
// namespace.
var startDownloadViaIFrame = (function() {
    // Holds on to the created iframe for downloading.
    // This will be re-used by ALL download calls.
    var downloadIFrame = null;
    var API_DOWNLOAD_ENABLED = 1;

    return function(path, op) {
        var params = {};
        params[API_DOWNLOAD] = API_DOWNLOAD_ENABLED;
        path = buildFilePathUrlString(path, params);

        if (downloadIFrame === null) {
            var iframeId = createUniqueElementId();
            downloadIFrame = createHiddenIframe(path, iframeId);
        } else {
            downloadIFrame.src = path;
        }

        // Since we cannot tell when the download finished, we do NOT call
        // the onSuccess callback.
    };
})();

wl_app.jsonp = {};
WL.Internal.jsonp = wl_app.jsonp;

function sendAPIRequestViaJSONP(request) {
    var scriptParent = document.getElementsByTagName("HEAD")[0],
        scriptTag = document.createElement("SCRIPT"),
        params = cloneObjectExcept(request._properties, params, [API_PARAM_CALLBACK, API_PARAM_PATH]),
        callbackName = request._id,
        token = wl_app.getAccessTokenForApi();

    if (token != null) {
        params[AK_ACCESS_TOKEN] = token;
    }

    params[API_PARAM_CALLBACK] = API_JSONP_CALLBACK_NAMESPACE_PREFIX + callbackName;
    params[API_SUPPRESS_REDIRECTS] = "true";

    var url = appendUrlParameters(request._url, params);
    if (url.length > API_JSONP_URL_LIMIT) {
        return false;
    }

    request.scriptTag = scriptTag;

    wl_app.jsonp[callbackName] = function (json) {
        cleanupJSONPRequest(callbackName, scriptTag);
        request.onCompleted(json);
    };

    attachScriptEvents(scriptTag, request);

    scriptTag.setAttribute("async", "async");
    scriptTag.type = "text/javascript";
    scriptTag.src = url;

    scriptParent.appendChild(scriptTag);

    window.setTimeout(function () {
        if (request._completed) {
            return;
        }

        cleanupJSONPRequest(callbackName, scriptTag);

    }, 30000);

    return true;
}

function attachScriptEvents(element, request) {
    if (wl_app._browser.ie && element.attachEvent) {
        element.attachEvent("onreadystatechange", function (e) {
            onScriptLoaded(e, request);
        });
    }
    else {
        element.readyState = "complete";
        element.addEventListener(
            "load",
            function (e) {
                onScriptLoaded(e, request);
            },
            false);
        element.addEventListener(
            "error",
            function (e) {
                onScriptLoaded(e, request);
            },
            false);
    }
}

function onScriptLoaded(e, request) {
    if (request._completed) {
        return;
    }

    var element = e.srcElement || e.currentTarget;
    if (!element.readyState) {
        element = e.currentTarget;
    }

    if ((element.readyState != "complete") &&
        (element.readyState != "loaded")) {
        return;
    }

    var callbackName = request._id;
    failure = (e.type == "error") || (wl_app.jsonp[callbackName] != null);
    if (failure) {
        // clean up request method and tag
        cleanupJSONPRequest(callbackName, request.scriptTag);

        // callback with error object
        var errorObj = {};
        errorObj[API_PARAM_CODE] = ERROR_CONNECTION_FAILED;
        errorObj[API_PARAM_MESSAGE] = ERROR_DESC_FAIL_CONNECT;
        request.onCompleted({ error: errorObj });
    }
}

function cleanupJSONPRequest(callbackName, scriptTag) {
    delete wl_app.jsonp[callbackName];
    document.getElementsByTagName("HEAD")[0].removeChild(scriptTag);
}

function sendAPIRequestViaFlash(request) {
    detectFlash();

    if (wl_app._browser.flashVersion < 9)
        return false;

    wl_app.xdrFlash.send(request);
    return true;
}

/**
 * The xdrFlash singleton instance handles cross domain Http request using flash.
 * Flash 9 is required.
 */
wl_app.xdrFlash = {
    _id: "",
    _status: FLASH_STATUS_NONE,
    _flashObject: null,
    _requests: {},
    _pending: [],

    init: function () {
        if (this._status != FLASH_STATUS_NONE)
            return;

        this._status = FLASH_STATUS_INITIALIZING;

        var container = createHiddenElement("div");
        container.id = "wl_xdr_container";
        document.body.appendChild(container);

        this._id = createUniqueElementId();
        var markup = flashObjectHtmlTemplate,
            xdrFlashUrl = wl_app[WL_SDK_ROOT] + "XDR.swf";

        markup = markup.replace(/{url}/g, xdrFlashUrl);
        markup = markup.replace(/{id}/g, this._id);
        markup = markup.replace(/{variables}/g, "domain=" + document.domain);

        container.innerHTML = markup;
    },

    /**
    * This is designed as callback function invoked by the flash code when the flash is ready to use.
    */
    onReady: function (success) {
        if (success) {
            if (wl_app._browser.firefox) {
                this._flashObject = document.embeds[this._id];
            }
            else {
                this._flashObject = getElementById(this._id);
            }

            this._status = FLASH_STATUS_INITIALIZED;
        }
        else {
            this._status = FLASH_STATUS_ERROR;
        }

        // Process pending requests.
        while (this._pending.length > 0) {
            this.send(this._pending.shift());
        }
    },

    onRequestCompleted: function (id, status, responseText, error) {
        var request = wl_app.xdrFlash._requests[id];
        delete wl_app.xdrFlash._requests[id];

        // Adobe Flash decodes line break characters and this breaks JSON parsing.
        // We replace those characters to work around the parsing error.
        processXDRResponse(request, status, encodeLineBreak(responseText), error);
    },

    send: function (request) {
        if (this._status < FLASH_STATUS_INITIALIZED) {
            this._pending.push(request);

            if (this._status == FLASH_STATUS_NONE)
                checkDocumentReady(createDelegate(this, this.init));
            return;
        }

        if (this._flashObject != null) {
            this._requests[request._id] = request;
            var xdrParams = prepareXDRRequest(request);

            this.invoke(
                "send",
                [request._id, xdrParams.url, xdrParams.method, xdrParams.body]);
        }
        else {
            processXDRResponse(request, 0, null, ERROR_DESC_BROWSER_ISSUE);
        }
    },

    invoke: function (name, parameters) {
        parameters = parameters || [];
        // We invoke the underlying ExternalInterface APIs directly since timing
        // issues exist where the Flash-created proxies don't always exist.
        var xmlParameters = window.__flash__argumentsToXML(parameters, 0),
            xml = "<invoke name=\"" + name + "\" returntype=\"javascript\">" + xmlParameters + "</invoke>";

        var result = this._flashObject.CallFunction(xml);
        if (result == null) {
            return null;
        }

        return eval(result);
    }
};

WL.Internal.xdrFlash = wl_app.xdrFlash;

function encodeLineBreak(text) {
    return text ? text.replace(/\r/g, " ").replace(/\n/g, " ") : text;
}

var flashObjectHtmlTemplate =
"<object classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' codebase='https://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0' width='300' height='300' id='{id}' name='{id}' type='application/x-shockwave-flash' data='{url}'>" +
"<param name='movie' value='{url}'></param>" +
"<param name='allowScriptAccess' value='always'></param>" +
"<param name='FlashVars' value='{variables}'></param>" +
"<embed play='true' menu='false' swLiveConnect='true' allowScriptAccess='always' type='application/x-shockwave-flash' FlashVars='{variables}' src='{url}' width='300' height='300' name='{id}'></embed>" +
"</object>";

var FilePickerOperation = null;
(function () {
    wl_app.fileDialog = function (args) {
        ensureAppInited(args[API_INTERFACE_METHOD]);

        if (wl_app._pendingPickerOp != null) {
            throw new Error(ERROR_DESC_PENDING_FILEDIALOG_CONFLICT);
        }

        return new FilePickerOperation(args).execute();
    };

    var FILE_PICKER_OP_NOTSTARTED = 0,
        FILE_PICKER_OP_AUTHREADY = 1,
        FILE_PICKER_OP_PICKING_COMPLETED = 2,
        FILE_PICKER_OP_COMPLETED = 3;

    FilePickerOperation = function(properties) {
        var op = this;
        op._props = properties;
        op._startTs = new Date().getTime();

        // Set default values.
        properties[FILEDIALOG_PARAM_LIGHTBOX] = properties[FILEDIALOG_PARAM_LIGHTBOX] || FILEDIALOG_PARAM_LIGHTBOX_WHITE;
        properties[FILEDIALOG_PARAM_SELECT] = properties[FILEDIALOG_PARAM_SELECT] || FILEDIALOG_PARAM_SELECT_SINGLE;

        op._state = FILE_PICKER_OP_NOTSTARTED;
        op._authDelegate = createDelegate(op, op._onAuthResp);
        wl_app._pendingPickerOp = op;
    };

    FilePickerOperation.prototype = {
        execute: function() {
            var op = this,
                promise = new Promise(op._props[API_INTERFACE_METHOD], op, null);
            op._promise = promise;
            op._process();
            return promise;
        },

        cancel: function(error) {
            // cancel
            var op = this;

            if (op._state === FILE_PICKER_OP_COMPLETED) {
                return;
            }

            if (!error) {
                error = createErrorResponse(
                    ERROR_REQ_CANCEL,
                    ERROR_DESC_CANCEL.replace(METHOD, op._props[API_INTERFACE_METHOD]));
            }

            op._onComplete(error);
        },

        _process: function() {
            var op = this;
            switch (op._state) {
                case FILE_PICKER_OP_NOTSTARTED:
                    op._validateAuth();
                    break;
                case FILE_PICKER_OP_AUTHREADY:
                    op._initPicker();
                    break;
                case FILE_PICKER_OP_PICKING_COMPLETED:
                    op._complete();
                    break;
            }
        },

        _changeState: function(state, result) {
            var op = this;
            if (op._state !== FILE_PICKER_OP_COMPLETED && op._state !== state) {
                op._state = state;

                if (result) {
                    op._result = result;
                }

                op._process();
            }
        },

        _onComplete: function(result) {
            this._changeState(FILE_PICKER_OP_PICKING_COMPLETED, result);
        },

        _validateAuth: function()
        {
            var op = this,
                isExternalFlow = false;
            if (wl_app._rpsAuth) {
                op._changeState(FILE_PICKER_OP_AUTHREADY);
            }
            else {
                var scope;
                switch (op._props[FILEDIALOG_PARAM_MODE]) {
                    case FILEDIALOG_PARAM_MODE_OPEN:
                        scope = SCOPE_SKYDRIVE;
                        break;
                    case FILEDIALOG_PARAM_MODE_SAVE:
                        scope = SCOPE_SKYDRIVE_UPDATE;
                        break;
                    case FILEDIALOG_PARAM_MODE_READ:
                    case FILEDIALOG_PARAM_MODE_READWRITE:
                        scope = op._buildExternalScope();
                        isExternalFlow = true;
                        break;
                    default:
                        var message = ERROR_DESC_PARAM_INVALID.replace(METHOD, IMETHOD_FILEDIALOG).replace(PARAM, FILEDIALOG_PARAM_MODE);
                        op._onComplete(createErrorResponse(ERROR_INVALID_REQUEST, message));
                        return;
                }

                var interfaceMethod = op._props[API_INTERFACE_METHOD],
                    TICKET_VALID_PERIOD = 650; // seconds

                if (isExternalFlow) {
                    wl_app.login({
                        scope: scope,
                        external_consent: true
                    }).then(
                        function(resp) {
                            // Success
                            op._onExternalConsentComplete(resp);
                        },
                        function(resp) {
                            // Failure
                            op._onComplete(resp);
                        });
                }
                else {
                    scope += " " + SCOPE_SIGNIN;
                    wl_app.ensurePermission(scope, TICKET_VALID_PERIOD, interfaceMethod, op._authDelegate);
                }
            }
        },

        _buildExternalScope: function () {
            var op = this;
            var scope = "onedrive_onetime.access:";
            switch (op._props[FILEDIALOG_PARAM_MODE]) {
                case FILEDIALOG_PARAM_MODE_READ:
                    scope += FILEDIALOG_PARAM_MODE_READ;
                    break;
                case FILEDIALOG_PARAM_MODE_READWRITE:
                    scope += FILEDIALOG_PARAM_MODE_READWRITE;
                    break;
            }

            switch (op._props[FILEDIALOG_PARAM_RESOURCETYPE]) {
                case FILEDIALOG_PARAM_RESOURCETYPE_FILE:
                    scope += "file|";
                    break;
                case FILEDIALOG_PARAM_RESOURCETYPE_FOLDER:
                    scope += "folder|";
                    break;
                default:
                    scope += "file|";
                    break;
            }

            scope += (op._props[FILEDIALOG_PARAM_SELECT] === FILEDIALOG_PARAM_SELECT_SINGLE) ? FILEDIALOG_PARAM_SELECT_SINGLE : FILEDIALOG_PARAM_SELECT_MULTI;

            if (op._props[FILEDIALOG_PARAM_LINKTYPE]) {
                scope += ("|" + op._props[FILEDIALOG_PARAM_LINKTYPE]);
            }

            return scope;
        },

        _onAuthResp: function(resp) {
            var op = this;
            if (resp.error) {
                if (!op._channel) {
                    // If we haven't show the picker, let the operation complete.
                    // Otherwise, let the picker continue on its own fake.
                    op._onComplete(resp);
                }
            }
            else {
                var token = resp.session[AK_ACCESS_TOKEN];
                op._props[AK_ACCESS_TOKEN] = token;

                if (op._channel) {
                    // The channel/picker is already set up. We just need to update the token.
                    op._channel.send(FILEDIALOG_CHCMD_UPDATETOKEN, token);
                }
                else {
                    op._changeState(FILE_PICKER_OP_AUTHREADY);
                }
            }
        },

        _initPicker: function() {
            // Create an iframe and start channel
            var op = this,
                props = op._props;

            showPicker(props);

            op._channel = wl_app.channel.registerOuterChannel(
                  CHANNEL_NAME_FILEDIALOG,
                  parseUri(props.url).host,
                  props.frame.contentWindow,
                  props.url,
                  createDelegate(op, op._onMessage));

            // Reply on the session auto update mechanism to update picker token.
            wl_event.subscribe(EVENT_AUTH_SESSIONCHANGE, op._authDelegate);

            var keydownHandler = function(e) {
                if (e.keyCode === KEYCODE_ESC) {
                    // When ESC is hit, we close the picker.
                    op.cancel();
                }
            };

            props.keydownHandler = keydownHandler;
            attachDOMEvent(window, "keydown", keydownHandler);

            op._initTimeout();
        },

        _initTimeout: function() {
            var op = this;
            timeoutSeconds = op._props[FILEDIALOG_PARAM_LOADING_TIMEOUT];
            // We only start a timeout behavior if it is set to do so.
            if (timeoutSeconds && timeoutSeconds > 0) {
                op._timeout = window.setTimeout(createDelegate(op, op._onTimeout), timeoutSeconds * 1000);
            }
        },

        _onTimeout: function() {
            // We use the channel connection status to determine if the Picker has been loaded properly.
            var op = this,
                connected = op._channel._connected;
            if (!connected) {
                op.cancel(createErrorResponse(ERROR_TIMEDOUT, ERROR_DESC_PICKER_TIMEOUT));
            }

            op._clearTimeout();
        },

        _clearTimeout: function() {
            var op = this;
            if (op._timeout) {
                window.clearTimeout(op._timeout);
                op._timeout = undefined;
            }
        },

        _complete: function() {
            var op = this,
                result = op._result,
                promiseEvent = (result.error) ? PROMISE_EVENT_ONERROR : PROMISE_EVENT_ONSUCCESS;

            op._state = FILE_PICKER_OP_COMPLETED;
            op._cleanup();

            op._normalizeResp();

            if (op._props[API_INTERFACE_METHOD] === IMETHOD_FILEDIALOG) {
                invokeCallback(op._props[API_PARAM_CALLBACK], result, true/*synchronous*/);
            }
            else {
                if (result.data) {
                    invokeCallback(op._props[FILEDIALOG_PARAM_ONSELECTED], result, true/*synchronous*/);
                }
                else {
                    invokeCallback(op._props[UI_PARAM_ONERROR], result, true/*synchronous*/);
                }
            }

            op._promise[promiseEvent](result);

            if (!wl_app._rpsAuth) {
                // We don't send report for 1st party apps, instead we rely on BICI report on SkyDrive side.
                // Add a delay to workaround an issue that hitting ESC key on the background area on IE9
                // will cause the JSONP request not to be sent.
                delayInvoke(function(e) {
                    op._report();
                });
            }
        },

        _report: function() {
            var props = this._props,
                result = this._result,
                duration = (new Date().getTime() - this._startTs) / 1000,
                selected_type = "none",
                foldersCount = 0,
                filesCount = 0;

            if (result.data) {
                if (result.data.folders != null) {
                    foldersCount = result.data.folders.length;
                }

                if (result.data.files != null) {
                    filesCount = result.data.files.length;
                }
            }

            selected_type = (foldersCount > 0 && filesCount > 0) ? "both" :
                (foldersCount > 0 ? FILEDIALOG_PARAM_RESOURCETYPE_FOLDER :
                    (filesCount > 0 ? FILEDIALOG_PARAM_RESOURCETYPE_FILE : "none"));

            var selected_count = foldersCount + filesCount,
                analyticsData = {
                    client: UI_SKYDRIVEPICKER,
                    api: props[API_INTERFACE_METHOD],
                    mode: props[FILEDIALOG_PARAM_MODE],
                    select: props[FILEDIALOG_PARAM_SELECT],
                    result: result.error ? result.error.code : "success",
                    duration: duration,
                    object_selected: selected_type,
                    selected_count: selected_count
                },
                analyticsListener = wl_app[ANALYTICS_LISTENER];

            wl_app.api({
                path: "web_analytics",
                method: HTTP_METHOD_POST,
                body: analyticsData
            });

            if (analyticsListener) {
                analyticsListener(analyticsData);
            }
        },

        _onExternalConsentComplete: function(response) {
            var op = this,
                generateSharingLinks = op._props[FILEDIALOG_PARAM_LINKTYPE] === ONEDRIVE_PARAM_LINKTYPE_WEBVIEW;

            if (response.error) {
                op._onComplete(response);
                return;
            }

            var ownerCid = response[AK_OWNER_CID],
                resourceId = response[AK_RESOURCEID],
                itemId = response[AK_ITEMID],
                authKey = response[AK_AUTH_KEY];

            if (!resourceId) {
                op._onComplete(createErrorResponse(ERROR_REQUEST_FAILED, "Did not get expected resource id."));
                return;
            }

            if (!authKey && generateSharingLinks) {
                op._onComplete(createErrorResponse(ERROR_REQUEST_FAILED, "Did not get expected auth key for the resource."));
                return;
            }

            var getItemProperties = {
                path: generateSharingLinks ?
                    "drives/" + ownerCid + "/items/" + itemId + "?$expand=thumbnails,children($expand=thumbnails)&authkey=" + authKey : resourceId + "/files",
                method: HTTP_METHOD_GET,
                use_vroom_api: generateSharingLinks,
                interface_method: op._props[API_INTERFACE_METHOD]
            };

            // The file dialog will pass back an id to the sharing bundle
            // representing the user's selection. To get the contents
            // of this bundle, we now have to do a GET request against the
            // sharing bundle.
            wl_app.api(getItemProperties).then(
                function (result) {
                    var fileDialogResponse = {
                        pickerResponse: response,
                        apiResponse: result
                    };

                    op._onComplete(fileDialogResponse);
                },
                function (error) {
                    op._onComplete(error);
                });
        },

        _onMessage: function(command, args) {
            log(command);
            switch (command) {
                case FILEDIALOG_CHCMD_ONCOMPLETE:
                    this._onComplete(args);
                    break;
            }
        },

        _normalizeResp: function() {
            var op = this,
                data = op._result.data,
                error = op._result.error,
                updateSelectedItem = function(item) {
                    var uploadLocation = item.upload_location;
                    if (uploadLocation) {
                        item.upload_location = uploadLocation.replace("WL_APISERVICE_URL", wl_app[WL_APISERVICE_URI]);
                    }
                };

            if (data) {
                if (data.folders) {
                    for (var i = 0; i < data.folders.length; i++) {
                        updateSelectedItem(data.folders[i]);
                    }
                }

                if (data.files) {
                    for (var i = 0; i < data.files.length; i++) {
                        updateSelectedItem(data.files[i]);
                    }
                }
            }

            if (error && error.message) {
                error.message = error.message.replace(METHOD, op._props[API_INTERFACE_METHOD]);
            }
        },

        _cleanup: function() {
            var op = this,
                props = op._props,
                channel = op._channel,
                resizeHandler = props.resizeHandler,
                keydownHandler = props.keydownHandler;

            op._clearTimeout();
            wl_event.unsubscribe(EVENT_AUTH_SESSIONCHANGE, op._authDelegate);

            if (props.lightBox) {
                props.frame.style.display = DOM_DISPLAY_NONE;
                props.lightBox.style.display = DOM_DISPLAY_NONE;

                document.body.removeChild(props.frame);
                document.body.removeChild(props.lightBox);

                delete props.lightBox;
                delete props.frame;
            }

            if (channel) {
                channel.dispose();
                delete op._channel;
            }

            if (resizeHandler) {
                detachDOMEvent(window, "resize", resizeHandler);
            }

            if (keydownHandler) {
                detachDOMEvent(window, "keydown", keydownHandler);
            }

            delete wl_app._pendingPickerOp;
        }
    };

    function showPicker(props) {
        var isOpenPicker = (props[FILEDIALOG_PARAM_MODE] === FILEDIALOG_PARAM_MODE_OPEN),
            width = isOpenPicker ? 800 : 500,
            height = isOpenPicker ? 600 : 450,
            lightbox = props[FILEDIALOG_PARAM_LIGHTBOX],
            isLightboxTransparent = (lightbox === FILEDIALOG_PARAM_LIGHTBOX_TRANSPARENT),
            opacity = isLightboxTransparent ? 0 : 60,
            backgroundColor = (lightbox === FILEDIALOG_PARAM_LIGHTBOX_WHITE) ? "white" : "rgb(0,0,0)",
            opaque = (opacity / 100),
            position = computePopoverPosition(width, height),
            url = buildPickerUrl(props);

        var container = document.createElement("div");
        container.style.top = "0px";
        container.style.left = "0px";
        container.style.position = "fixed";
        container.style.width = "100%";
        container.style.height = "100%";
        container.style.display = "block";
        container.style.backgroundColor = backgroundColor;
        container.style.opacity = opaque;
        container.style.MozOpacity = opaque;
        container.style.filter = 'alpha(opacity=' + opacity + ')';
        container.style.zIndex = "2600000";

        var iframe = document.createElement("iframe");
        iframe.id = "picker" + new Date().getTime();
        iframe.style.top = position.top;
        iframe.style.left = position.left;
        iframe.style.position = "fixed";
        iframe.style.width = width + "px";
        iframe.style.height = height + "px";
        iframe.style.display = "block";
        iframe.style.zIndex = "2600001";
        iframe.style.borderWidth = "1px";
        iframe.style.borderColor = "gray";
        iframe.style.maxHeight = "100%";
        iframe.style.maxWidth = "100%";
        iframe.src = url;
        iframe.setAttribute("sutra", "picker");

        document.body.appendChild(iframe);
        document.body.appendChild(container);

        props.resizeHandler = function () {
            var position = computePopoverPosition(width, height);
            iframe.style.top = position.top;
            iframe.style.left = position.left;
        };

        attachDOMEvent(window, "resize", props.resizeHandler);

        props.lightBox = container;
        props.frame = iframe;
        props.url = url;
    }

    function buildPickerUrl(props) {
        var isOpenPicker = (props[FILEDIALOG_PARAM_MODE] === FILEDIALOG_PARAM_MODE_OPEN),
            pickerview = isOpenPicker ? FILEDIALOG_PARAM_VIEWTYPE_FILEPICKER : FILEDIALOG_PARAM_VIEWTYPE_FOLDERPICKER,
            params = {},
            pickerScript = getSDKRootPath() + wl_app[FILEDIALOG_PARAM_PICKER_SCRIPT];

        if (pickerScript.charAt(0) === '/') {
            // SkyDrive does not accept url begining with "//", so we need to ensure the Url starts with "https".
            // Assuming the url is a not a relative path, we only check '/'.
            pickerScript = SCHEME_HTTPS + pickerScript;
        }

        params[FILEDIALOG_PARAM_VIEWTYPE] = pickerview;
        params[FILEDIALOG_PARAM_AUTH] = wl_app._rpsAuth ? FILEDIALOG_PARAM_AUTH_RPS : FILEDIALOG_PARAM_AUTH_OAUTH;
        params[FILEDIALOG_PARAM_DOMAIN] = window.location.hostname.toLowerCase();
        params[FILEDIALOG_PARAM_LIVESDK] = pickerScript;
        params[AK_CLIENT_ID] = wl_app._appId;
        params[AK_REQUEST_TS] = new Date().getTime();
        params[FILEDIALOG_PARAM_MKT] = wl_app._locale;

        if (!wl_app._rpsAuth) {
            params[AK_ACCESS_TOKEN] = props[AK_ACCESS_TOKEN];
        }

        if (isOpenPicker) {
            params[FILEDIALOG_PARAM_SELECT] = props[FILEDIALOG_PARAM_SELECT];
        }

        return appendUrlParameters(wl_app[WL_SKYDRIVE_URI], params);
    }

    function computePopoverPosition(width, height) {
        var left, top;
        if (wl_app._browser.ie) {
            var docElement = document.documentElement;

            left = (docElement.clientWidth - width) / 2;
            top = (docElement.clientHeight - height) / 2;
        }
        else {
            left = (window.innerWidth - width) / 2;
            top = (window.innerHeight - height) / 2;
        }

        left = (left < 10) ? 10 : left;
        top = (top < 10) ? 10 : top;

        return {
            left: left + "px",
            top: top + "px"
        };
    }
})();

// wl.app.upload.web.js
// WL.upload methods specific to the Live SDK for Web JavaScript.

var UPLOAD_TIMEOUT = 60 * 1000;  // milliseconds

UploadOperation.prototype._getStrategy = function (properties) {
    var self = this,
        interfaceMethod = properties[API_INTERFACE_METHOD],
        element = properties[API_PARAM_ELEMENT],
        fileName = properties[API_PARAM_FILENAME];

    validateProperties(
       properties,
       [{ name: API_PARAM_ELEMENT, type: TYPE_DOM, optional: false }],
       interfaceMethod);

    if (typeof(element) === TYPE_STRING) {
        element = document.getElementById(element);
    }

    // The element must be <input type="file" />
    if (!(element instanceof HTMLInputElement) ||
        element.type !== DOM_FILE) {
        throw createInvalidParamValue(
                API_PARAM_ELEMENT,
                interfaceMethod,
                "It must be an HTMLInputElement with its type attribute set to\"file\" (i.e., <input type=\"file\" />).");
    }

    // The element must have a file selected (i.e., value must not be empty).
    if (element.value === "") {
        throw createInvalidParamValue(
                API_PARAM_ELEMENT,
                interfaceMethod,
                "It did not contain a selected file.");
    }

    // if the input element has a files property and there is FileReader type available, then the browser supports
    // the file API and we will use that.
    if (element.files && window.FileReader) {
        if (element.files.length !== 1) {
            throw createInvalidParamValue(
                    API_PARAM_ELEMENT,
                    interfaceMethod,
                    "It must contain one selected file.");
        }

        var fileInput = element.files[0];

        if (fileInput.size > FORM_UPLOAD_SIZE_LIMIT) {
            throw createInvalidParamValue(
                    API_PARAM_ELEMENT,
                    interfaceMethod,
                    "Max supported file size for form uploads is 100MB.");
        }

        // If the caller supplied a file name, use that, otherwise get the file name from the input element.
        self.setFileName(fileName || fileInput.name);

        return new XhrUploadStrategy(self, fileInput);
    }

    // if they did not specify a name on the input element, change it to
    // the proper name of "file".
    if (element.name === "") {
        element.name = DOM_FILE;
    }

    // otherwise, if the browser does not support the File API, then we have
    // to upload using a multipart form post. That means the HTMLInputElement must
    // have a parent form.
    // TODO(skrueger): There is also an issue with the api service where sending over more than
    // 1 form element can cause issues. For now, only allow 1 form control element.
    var errorMessage = null;
    if (element.form === null) {
        errorMessage = "It must be a child of an HTMLFormElement.";
    } else if (element.form.length !== 1) {
        // the api service can only handle one input element
        errorMessage = "It must be the only HTMLInputElement in its parent HTMLFormElement.";
    } else if (element.name !== DOM_FILE) {
        // the input element must be named file
        errorMessage = "Its name attribute must be set to \"file\" (i.e., <input name=\"file\" />).";
    }

    if (errorMessage !== null) {
        throw createInvalidParamValue(API_PARAM_ELEMENT, interfaceMethod, errorMessage);
    }

    return new MultiPartFormUploadStrategy(self, element, interfaceMethod);
};

/**
 * Logic for performing a Multipart form upload.
 */
function MultiPartFormUploadStrategy(operation, element, interfaceMethod) {
    var self = this;
    self._op = operation;
    self._iMethod = interfaceMethod;
    self._element = element;
    self._uploadIFrame = null;
    self._uploadTimeoutId = null;

    // for a multipartform upload, we can NOT change the file name. The file name
    // will be sent in the body of the request, and there is no way to change it in the
    // body. So just set the FileName to undefined.
    operation.setFileName(undefined);
}

MultiPartFormUploadStrategy.prototype = {
    getId: function() {
        var self = this;
        if (self._uploadIFrame === null) {
            self._createUploadIFrame();
        }

        return self._uploadIFrame.id;
    },

    setUploadTimeout: function() {
        var self = this;
        self._uploadTimeoutId =
            window.setTimeout(function() { self.onTimeout(); }, UPLOAD_TIMEOUT);
    },

    clearUploadTimeout: function() {
        var self = this;
        if (self._uploadTimeoutId === null) {
            return;
        }

        window.clearTimeout(self._uploadTimeoutId);
        self._uploadTimeoutId = null;
    },

    /**
     * Callback used when the upload timeout is reached.
     */
    onTimeout: function() {
        var self = this;
        self._uploadTimeoutId = null;

        var errorDescription = self._iMethod + ": did not receive a response in " +
                               UPLOAD_TIMEOUT + " milliseconds.";
        var errorResponse = createErrorResponse(ERROR_TIMEDOUT, errorDescription);
        self._op.uploadComplete(false, errorResponse);
    },

    /**
     * Callback used when the upload completes.
     */
    onUploadComplete: function(result) {
        var self = this;
        self.clearUploadTimeout();

        self._removeUploadIFrame();

        result = deserializeJSON(result);

        var error = result.error;
        var successful = typeof(error) === TYPE_UNDEFINED;
        self._op.uploadComplete(successful, result);
    },

    /**
     * Public call to perform the upload.
     */
    upload: function(uploadPath) {
        this._multiPartFormUpload(uploadPath);
    },

    /**
     * Builds the request url from the path. Adds all necessary
     * query parameters.
     */
    _getRequestUrl: function(path) {
        var params = {};
        params[AK_REDIRECT_URI] = wl_app._redirect_uri;

        // The API service will be an HTML response that redirects to the redirect_uri and its
        // url will pass back whatever state query parameter we send it.
        // When this library loads up, it will know it is an upload redirect because the state
        // parameter will say that it is an upload redirect.
        var state = {};
        state[REDIRECT_TYPE] = REDIRECT_TYPE_UPLOAD;
        state[UPLOAD_STATE_ID] = this.getId();
        params[AK_STATE] = serializeParameters(state);

        return appendUrlParameters(path, params);
    },

    /**
     * Creates an iframe for the target of the multipart form upload
     * if one has NOT been created yet.
     */
    _createUploadIFrame: function() {
        var self = this;
        if (self._uploadIFrame !== null) {
            return;
        }

        self._uploadIFrame = createHiddenElement('iframe');
        self._uploadIFrame.name = (self._uploadIFrame.id = createUniqueElementId());
        document.body.appendChild(self._uploadIFrame);
    },

    /**
     * Performs the multipart form upload.
     */
    _multiPartFormUpload: function(uploadPath) {
        var self = this;
        self._createUploadIFrame();
        var requestUrl = self._getRequestUrl(uploadPath);
        self._submitForm(requestUrl);

        self.setUploadTimeout();
        pendingFormUploads.add(self);

        pollUploadResponseCookie();
    },

    /**
     * Removes the upload iframe from the DOM, and sets
     * the member variable to null.
     */
    _removeUploadIFrame: function() {
        var self = this;
        if (self._uploadIFrame === null) {
            return;
        }

        self._uploadIFrame.parentNode.removeChild(self._uploadIFrame);
        self._uploadIFrame = null;
    },

    /**
     * Sets up the form and submits it to the server.
     */
    _submitForm: function(requestUrl) {
        var self = this;
        var form = self._element.form;
        form.target = self._uploadIFrame.name;
        form.method = HTTP_METHOD_POST;
        form.enctype = 'multipart/form-data';  // must be set for input type="file"
        form.action = requestUrl;
        form.submit();
    }
};

/**
 * Keeps track of all the pending form uploads.
 */
function PendingMultipartFormUploadManager() {
    this._pendingUploads = {};
}

PendingMultipartFormUploadManager.prototype = {
    /**
     * Adds a pending MultipartFormUploadStrategy.
     */
    add: function(uploadStrategy) {
        var id = uploadStrategy.getId();
        this._pendingUploads[id] = uploadStrategy;
    },

    /**
     * returns true if there are any pending uploads.
     */
    hasPendingUploads: function() {
        for (var id in this._pendingUploads) {
            return true;
        }

        return false;
    },

    /**
     * returns true if the given MultipartFormUploadStrategy's
     * id is pending.
     */
    isPending: function(uploadStrategyId) {
        return uploadStrategyId in this._pendingUploads;
    },

    /**
     * returns the multipart form upload strategy with the given id.
     */
    get: function(uploadStrategyId) {
        return this._pendingUploads[uploadStrategyId];
    },

    /**
     * removes the multipart form upload strategy with the given id.
     */
    remove: function(uploadStrategyId) {
        delete this._pendingUploads[uploadStrategyId];
    }
};

/**
 * internal global for managing pending multipart form uploads.
 */
var pendingFormUploads = new PendingMultipartFormUploadManager();

/**
 * Polls the upload response by polling the wl_upload cookie.
 * Create closure to capture local variables.
 */
var pollUploadResponseCookie = (function() {
    var hasAddedObserver = false;
    var uploadCookie = new CookieState(COOKIE_UPLOAD);

    // onChange is called when the wl_upload cookie is changed.
    var onChange = function() {
        var states = uploadCookie.getStates();
        var cookieChanged = false;

        for (var uploadId in states) {
            // check to see if there is any pending request.
            // if there are none, then just move to the next request.
            if (!pendingFormUploads.isPending(uploadId)) {
                continue;
            }

            var result = states[uploadId];
            var multipartFormUploadStrategy = pendingFormUploads.get(uploadId);

            pendingFormUploads.remove(uploadId);
            uploadCookie.remove(uploadId);
            cookieChanged = true;

            multipartFormUploadStrategy.onUploadComplete(result);
        }

        // stop monitoring the cookie if there are no more pending uploads.
        if (!pendingFormUploads.hasPendingUploads()) {
            uploadCookie.stopMonitor();
        }

        if (cookieChanged) {
            uploadCookie.save();
            cookieChanged = false;
        }
    };

    return (function() {
        // if we are already monitoring the cookie then leave.
        if (uploadCookie.isBeingMonitored()) {
            return;
        }

        if (hasAddedObserver) {
            uploadCookie.startMonitor();
        } else {
            uploadCookie.addCookieChanged(onChange);
            hasAddedObserver = true;
        }
    });
})();


function XhrUploadStrategy(operation, uploadSource) {
    /// <summary>
    /// Performs an upload via an XMLHttpRequest.
    /// </summary>

    this.upload = function (requestUrl) {
        var reader = null;

        if (window.File && uploadSource instanceof window.File) {
            reader = new FileReader();
        }

        reader.onerror = function(evt) {
            error = evt.target.error;
            operation.onErr(error.name); // TODO: name?
        };

        reader.onload = function(evt) {
            var data = evt.target.result;
            var xhr = new XMLHttpRequest();
            xhr.open(HTTP_METHOD_PUT, requestUrl, true);

            xhr.onload = function(e) {
                operation.onResp(e.currentTarget.responseText);
            };

            xhr.onerror = function(e) {
                operation.onErr(e.currentTarget.statusText);
            };

            if (xhr.upload) {
                xhr.upload.onprogress = function(e) {
                    if (e.lengthComputable) {
                        var uploadProgress = {
                            bytesTransferred: e.loaded,
                            totalBytes: e.total,
                            progressPercentage: (e.total === 0) ? 0 : (e.loaded / e.total) * 100
                        };

                        operation.uploadProgress(uploadProgress);
                    }
                };
            }

            operation._cancel = createDelegate(xhr, xhr.abort);
            xhr.send(data);
        };

        reader.readAsArrayBuffer(uploadSource);
    };
};

function detectSecureConnection() {
    wl_app._isHttps = document.location.protocol.toLowerCase() == SCHEME_HTTPS;
}

function detectFlash() {
    if (wl_app._browser.flash !== undefined)
        return;

    var version = 0;
    try {
        if (wl_app._browser.ie) {
            var axo = new ActiveXObject("ShockwaveFlash.ShockwaveFlash.7");
            var fullVersion = axo.GetVariable("$version");
            var tempArray = fullVersion.split(" ");    // ["WIN", "2,0,0,11"]
            var tempString = tempArray[1];             // "2,0,0,11"
            var versionArray = tempString.split(",");  // ['2', '0', '0', '11']
            version = versionArray[0];
        }
        else if (navigator.plugins && navigator.plugins.length > 0) {
            if (navigator.plugins["Shockwave Flash 2.0"] || navigator.plugins["Shockwave Flash"]) {
                var swVer2 = navigator.plugins["Shockwave Flash 2.0"] ? " 2.0" : "";
                var description = navigator.plugins["Shockwave Flash" + swVer2].description;
                var descArray = description.split(" ");
                var tempArrayMajor = descArray[2].split(".");
                version = tempArrayMajor[0];
            }
        }
    }
    catch (e) {
    }

    wl_app._browser.flashVersion = version;
    wl_app._browser.flash = (version >= 8);
}

function onDocumentReady() {
    if (wl_app._documentReady === undefined) {
        wl_app._documentReady = new Date().getTime();
    }
}

function onCreateIframeReady(createIframe) {
    checkDocumentReady(function () {
        if (wl_app._browser.firefox &&
            (!(wl_app._documentReady) || ((new Date().getTime() - wl_app._documentReady) < 1000))) {
            // In Firefox, iframe loading will not going out and return cached content if loading too early.
            // For a simple page, 1 millisecond may be adequate. For a complex page, it takes a bit longer.
            window.setTimeout(createIframe, 1000);
        }
        else {
            createIframe();
        }
    });
}

/**
 * The Web version of checkDocumentReady() method.
 */
function checkDocumentReady(onDocReady) {
    if (document.body) {
        switch (document.readyState) {
            case "complete":                 // All
                onDocReady();
                return;
            case "loaded":                   // WebKit < 534.10
                if (wl_app._browser.webkit) {
                    onDocReady();
                    return;
                }
                break;
            case "interactive":              // Firefox >= 3.6 and WebKit >= 534.10
            case undefined:                  // Firefox <  3.6
                if (wl_app._browser.firefox || wl_app._browser.webkit) {
                    onDocReady();
                    return;
                }
                break;
        }
    }

    if (document.addEventListener) {
        document.addEventListener("DOMContentLoaded", onDocReady, false);
        document.addEventListener("load", onDocReady, false);
    }
    else if (window.attachEvent) {
        window.attachEvent("onload", onDocReady);
    }

    if (wl_app._browser.ie && document.attachEvent) {
        document.attachEvent("onreadystatechange", function () {
            if (document.readyState === "complete") {
                document.detachEvent("onreadystatechange", arguments.callee);
                onDocReady();
            }
        });
    }
}

/**
 * The Web version of setInnerHtml() method.
 */
function setInnerHtml(element, content) {
    element.innerHTML = content;
}

function attachDOMEvent(dom, eventName, handler) {
    if (dom.addEventListener) {
        dom.addEventListener(eventName, handler, false);
    }
    else if (dom.attachEvent) {
        dom.attachEvent("on" + eventName, handler);
    }
}

function detachDOMEvent(dom, eventName, handler) {
    if (dom.removeEventListener) {
        dom.removeEventListener(eventName, handler, false);
    }
    else if (dom.detachEvent) {
        dom.detachEvent("on" + eventName, handler);
    }
}

function getClientIdFromDOM() {
    var element = getElementById(DOM_ID_SDK);
    if (element) {
        var id = element.getAttribute(DOM_ATTR_CLIENTID);
        !id && logError(stringFormat("Could not find attribute '{0}' on element with id '{1}'.", DOM_ATTR_CLIENTID, DOM_ID_SDK), ONEDRIVE_PREFIX);

        return id;
    }
    else {
        logError(stringFormat("Could not find element with id '{0}'.", DOM_ID_SDK), ONEDRIVE_PREFIX);
        return null;
    }
}

var ChannelManager = {
        registerOuterChannel: function (name, allowedDomain, targetWindow, targetUrl, onMessageReceived) {
        // name: We use a channel name to identify each channel if more than one are used.
        // allowedDomain: We verify the domain when we receive messages
        // targetWindow: The window instance we are trying to connect.
        // targetUrl: The target Url of the remote channel.
        // onMessageReceived: The message handling function.
        return PostMessageChannelManager.registerChannel(name, allowedDomain, targetWindow, targetUrl, onMessageReceived);
    },

    registerInnerChannel: function (name, allowedDomain, onMessageReceived) {
        // name: We use a channel name to identify each channel if more than one are used.
        // allowedDomain: We verify the domain when we receive messages
        // onMessageReceived: The message handling function.
        return PostMessageChannelManager.registerChannel(name, allowedDomain, null, null, onMessageReceived);
    },

    isSupported: function () {
        return PostMessageChannelManager.isSupported();
    }
};

var PostMessageChannelManager = {
    _channels: [],

    isSupported: function () {
        return (window.postMessage != null);
    },

    registerChannel: function (name, allowedDomain, targetWindow, targetUrl, onMessageReceived) {
        var manager = PostMessageChannelManager,
            channels = manager._channels,
            channel = null;

        if (manager.isSupported()) {
            // PostMessage channel is supported, then create one.
            channel = new PostMessageChannel(name, allowedDomain, targetWindow, targetUrl, onMessageReceived);

            // ensure we listen to the message event.
            if (channels.length === 0) {
                attachDOMEvent(window, "message", manager._onMessage);
            }

            // Add the channel to the list
            channels.push(channel);
        }

        return channel;
    },

    unregisterChannel: function (channel) {
        var manager = PostMessageChannelManager,
            channels = manager._channels;

        for (var i = 0; i < channels.length; i++) {
            if (channels[i] == channel) {
                channels.splice(i, 1);
                break;
            }
        }

        if (channels.length === 0) {
            detachDOMEvent(window, "message", manager._onMessage);
        }
    },

    _onMessage: function (e) {
        var manager = PostMessageChannelManager,
            e = e || window.event, // Let's read e first, then window.event to cover browser differences.
            msg = readPostMessage(e);

        if (msg != null) {
            var channel = manager._findChannel(e, msg);
            if (channel != null) {
                switch (msg.text) {
                    case CONNECT_REQ:
                        channel._onConnectReq(e.source, e.origin);
                        break;

                    case CONNECT_ACK:
                        channel._onConnectAck(e.source);
                        break;

                    default:
                        channel._onMessage(msg.text);
                        break;
                }
            }
        }
    },

    _findChannel: function (e, msg) {
        var manager = PostMessageChannelManager,
            channels = manager._channels,
            domain = getDomainName(e.origin);

        for (var i = 0; i < channels.length; i++) {
            var channel = channels[i];

            if (stringsAreEqualIgnoreCase(channel._name, msg.name) &&
                stringsAreEqualIgnoreCase(channel._allowedDomain, domain)) {
                return channel;
            }
        }

        return null;
    }
};

var CONNECT_REQ = "@ConnectReq",
    CONNECT_ACK = "@ConnectAck";

function PostMessageChannel(name, allowedDomain, targetWindow, targetUrl, onMessageReceived) {
    var self = this;
    self._name = name;
    self._allowedDomain = allowedDomain;
    self._msgHandler = onMessageReceived;

    if (targetWindow) {
        self._targetWindow = targetWindow;
        self._targetUrl = getTargetOrigin(targetUrl);
        self._connect();
    }
}

PostMessageChannel.prototype = {
    _disposed: false,
    _connected: false,
    _allowedDomain: null,
    _targetWindow: null,
    _targetUrl: null,
    _msgHandler: null,
    _connectSchedule: -1,
    _pendingQueue: [],
    _recvQueue: [],

    dispose: function () {
        var self = this;
        if (!self._disposed) {

            self._disposed = true;
            self._cancelConnect();

            PostMessageChannelManager.unregisterChannel(self);
        }
    },

    send: function (command, args) {
        var self = this;
        if (self._disposed) {
            return;
        }

        var text = encodeChannelMessage(command, args);

        if (self._connected) {
            self._post(text);
        }
        else {
            self._pendingQueue.push(text);
        }
    },

    _connect: function () {
        var self = this;
        if (self._disposed || self._connected) {
            return;
        }

        var tryConnect = function () {
            self._post(CONNECT_REQ);
        };

        if (self._connectSchedule === -1) {
            self._connectSchedule = window.setInterval(tryConnect, 500);
            tryConnect();
        }
    },

    _post: function (text) {
        var self = this,
            msg = encodePostMessage(self._name, text);

        self._targetWindow.postMessage(msg, self._targetUrl);
    },

    _onConnectReq: function (source, targetUrl) {
        var self = this;
        if (!self._connected) {
            self._connected = true;
            self._targetWindow = source;
            self._targetUrl = targetUrl;

            self._post(CONNECT_ACK);
            self._onConnected();
        }
    },

    _onConnectAck: function (source) {
        var self = this;
        if (!self._connected) {

            self._targetWindow = source;

            self._onConnected();
        }
    },

    _onConnected: function () {

        var self = this,
            pendingMsgs = self._pendingQueue;

        // mark as connected
        self._connected = true;

        // send pending messages
        for (var i = 0; i < pendingMsgs.length; i++) {
            self._post(pendingMsgs[i]);
        }

        self._pendingQueue = [];
        self._cancelConnect();
    },

    _cancelConnect: function () {
        // clear the connect-scheduler
        var self = this;
        if (self._connectSchedule != -1) {
            window.clearInterval(self._connectSchedule);
            self._connectSchedule = -1;
        }
    },

    _onMessage: function (text) {
        if (this._msgHandler) {
            var msg = decodeChannelMessage(text);
            this._msgHandler(msg.cmd, msg.args);
        }
    }
};

/**
 * Gets the target origin by stripping the path, query string and hash fragments.
 */
function getTargetOrigin(url) {

    var idx = url.indexOf("://");
    if (idx >= 0) {
        idx = url.indexOf("/", idx + 3);
        if (idx >= 0) {
            url = url.substring(0, idx);
        }
    }

    return url;
}

function readPostMessage(e) {

    var msg = null;

    if (!stringIsNullOrEmpty(e.origin) &&
        !stringIsNullOrEmpty(e.data) &&
        e.source != null) {

        msg = decodePostMessage(e.data);
    }

    return msg;
}

function encodePostMessage(name, text) {
    return "<" + name + ">" + text;
}

function decodePostMessage(data) {
    var msg = null;
    if (data.charAt(0) == '<') {
        var index = data.indexOf('>');
        if (index > 0) {
            var name = data.substring(1, index),
                text = data.substr(index + 1);

            msg = {
                name: name,
                text: text
            };
        }
    }

    return msg;
}

function encodeChannelMessage(command, args) {
    var message = {
        cmd: command,
        args: args
    };

    return JSON.stringify(message);
}

function decodeChannelMessage(text) {
    // JSON is available in IE8 above and other modern browsers.
    // So, we can rely on browser native JSON support.
    var message = JSON.parse(text);
    return message;
}

if (window.WL) {
    // We intend to hide the channel interface from public.
    // If WL.JS is loaded in full, channel service is attached to wl_app instance.
    wl_app.channel = ChannelManager
}
else {
    // If only wl.channel.js is loaded, channel service is attach to WL object.
    window.WL = {
        channel: ChannelManager
    };
}

        var
//! Copyright (c) Microsoft Corporation. All rights reserved.


////////////////////////////////////////////////////////////////////////////////
// WLText

WLText = {
    connect: 'Connect',
    signIn: 'Sign in',
    signOut: 'Sign out',
    login: 'Log in',
    logout: 'Log out',
    skyDriveOpenPickerButtonText: 'Open from OneDrive',
    skyDriveOpenPickerButtonTooltip: 'Open from OneDrive',
    skyDriveSavePickerButtonText: 'Save to OneDrive',
    skyDriveSavePickerButtonTooltip: 'Save to OneDrive'
};



// ---- Do not remove this footer ----
// This script was generated using Script# v0.6.0.0 (http://projects.nikhilk.net/ScriptSharp)
// -----------------------------------

/**
 * Locale of this script.
 */
wl_app._locale = "en";

        wl_app[API_X_HTTP_LIVE_LIBRARY] = "Web/DEVICE_" + trimVersionBuildNumber("5.5.8816.3000");

        wl_app.testInit = function(properties) {

            if (properties.env) {
                wl_app._settings.init(properties.env);
            }

            if (properties.skydrive_uri) {
                wl_app._settings[WL_SKYDRIVE_URI] = properties.skydrive_uri;
            }

            if (properties[ANALYTICS_LISTENER]) {
                wl_app[ANALYTICS_LISTENER] = properties[ANALYTICS_LISTENER];
            }
        };

        var prodSettings = {};
        prodSettings[WL_AUTH_SERVER] = "login.live.com";
        prodSettings[WL_APISERVICE_URI] = "https://apis.live.net/v5.0/";
        prodSettings[WL_SKYDRIVE_URI] = "https://onedrive.live.com/";
        prodSettings[WL_SDK_ROOT] = "//js.live.net/v5.0/";
        prodSettings[WL_ONEDRIVE_API] = "https://api.onedrive.com/v1.0/";

        var dfSettings = {};
        dfSettings[WL_AUTH_SERVER] = "login.live.com";
        dfSettings[WL_APISERVICE_URI] = "https://apis.live.net/v5.0/";
        dfSettings[WL_SKYDRIVE_URI] = "https://onedrive.live.com/";
        dfSettings[WL_SDK_ROOT] = "//df-js.live.net/v5.0/";
        dfSettings[WL_ONEDRIVE_API] = "https://df.api.onedrive.com/v1.0/";

        var intSettings = {};
        intSettings[WL_AUTH_SERVER] = "login.live-int.com";
        intSettings[WL_APISERVICE_URI] = "https://apis.live-int.net/v5.0/";
        intSettings[WL_SKYDRIVE_URI] = "https://onedrive.live-int.com/";
        intSettings[WL_SDK_ROOT] = "//js.live-int.net/v5.0/";
        intSettings[WL_ONEDRIVE_API] = "https://newapi.storage.live-int.com/v1.0/";

        wl_app._settings =
        {
            PROD: prodSettings,
            DF: dfSettings,
            INT: intSettings,

            init: function(env) {
                env = env.toUpperCase();
                var envSettings = this[env];
                if (envSettings) {
                    cloneObject(envSettings, wl_app);
                }
            }
        };

        wl_app._settings.init("PROD");

        wl_app[FILEDIALOG_PARAM_PICKER_SCRIPT] = "wl.skydrivepicker.debug.js";
        wl_app.onloadInit();
        OneDriveApp.onloadInit();
    }
})();
