//! Copyright (c) Microsoft Corporation. All rights reserved.
var __extends = (this && this.__extends) || function (d, b) {for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];function __() { this.constructor = d; }__.prototype = b.prototype;d.prototype = new __();};
!function(e){if("object"==typeof exports)module.exports=e();else if("function"==typeof define&&define.amd)define(e);else{var f;"undefined"!=typeof window?f=window:"undefined"!=typeof global?f=global:"undefined"!=typeof self&&(f=self),f.OneDrive=e()}}(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);throw new Error("Cannot find module '"+o+"'")}var f=n[o]={exports:{}};t[o][0].call(f.exports,function(e){var n=t[o][1][e];return s(n?n:e)},f,f.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(_dereq_,module,exports){
var ErrorType = _dereq_('./models/ErrorType');
var Constants = function () {
        function Constants() {
        }
        Constants.ERROR_ACCESS_DENIED = 'access_denied';
        Constants.ERROR_POPUP_OPEN = {
            errorCode: ErrorType.popupOpen,
            message: 'popup window is already open'
        };
        Constants.ERROR_WEB_REQUEST = {
            errorCode: ErrorType.webRequestFailure,
            message: 'web request failed, see console logs for details'
        };
        Constants.HTTP_GET = 'GET';
        Constants.HTTP_POST = 'POST';
        Constants.HTTP_PUT = 'PUT';
        Constants.LINKTYPE_WEB = 'webLink';
        Constants.LINKTYPE_DOWNLOAD = 'downloadLink';
        Constants.PARAM_ACCESS_TOKEN = 'access_token';
        Constants.PARAM_ERROR = 'error';
        Constants.PARAM_STATE = 'state';
        Constants.PARAM_SDK_STATE = 'sdk_state';
        Constants.PARAM_ID_TOKEN = 'id_token';
        Constants.PARAM_SPREDIRECT = 'sp';
        Constants.SDK_VERSION = 'js-v2.1';
        Constants.STATE_AAD_LOGIN = 'aad_login';
        Constants.STATE_AAD_PICKER = 'aad_picker';
        Constants.STATE_MSA_PICKER = 'msa_picker';
        Constants.STATE_OPEN_POPUP = 'open_popup';
        Constants.STATE_GRAPH = 'graph';
        Constants.TYPE_BOOLEAN = 'boolean';
        Constants.TYPE_FUNCTION = 'function';
        Constants.TYPE_OBJECT = 'object';
        Constants.TYPE_STRING = 'string';
        Constants.TYPE_NUMBER = 'number';
        Constants.VROOM_URL = 'https://api.onedrive.com/v1.0/';
        Constants.VROOM_INT_URL = 'https://newapi.storage.live-int.com/v1.0/';
        Constants.GRAPH_URL = 'https://graph.microsoft.com/v1.0/';
        Constants.NONCE_LENGTH = 5;
        Constants.CUSTOMER_TID = '9188040d-6c67-4c5b-b112-36a304b66dad';
        Constants.DEFAULT_QUERY_ITEM_PARAMETER = 'expand=thumbnails&select=id,name,size,webUrl,folder';
        Constants.GLOBAL_FUNCTION_PREFIX = 'oneDriveFilePicker';
        return Constants;
    }();
module.exports = Constants;
},{"./models/ErrorType":6}],2:[function(_dereq_,module,exports){
var Constants = _dereq_('./Constants'), OneDriveApp = _dereq_('./OneDriveApp');
var OneDrive = function () {
        function OneDrive() {
        }
        OneDrive.open = function (options) {
            OneDriveApp.open(options);
        };
        OneDrive.save = function (options) {
            OneDriveApp.save(options);
        };
        OneDrive.webLink = Constants.LINKTYPE_WEB;
        OneDrive.downloadLink = Constants.LINKTYPE_DOWNLOAD;
        return OneDrive;
    }();
OneDriveApp.onloadInit();
module.exports = OneDrive;
},{"./Constants":1,"./OneDriveApp":3}],3:[function(_dereq_,module,exports){
var DomUtilities = _dereq_('./utilities/DomUtilities'), ErrorHandler = _dereq_('./utilities/ErrorHandler'), Logging = _dereq_('./utilities/Logging'), OneDriveState = _dereq_('./OneDriveState'), Picker = _dereq_('./utilities/Picker'), PickerMode = _dereq_('./models/PickerMode'), RedirectUtilities = _dereq_('./utilities/RedirectUtilities'), ResponseParser = _dereq_('./utilities/ResponseParser'), Saver = _dereq_('./utilities/Saver');
var OneDriveApp = function () {
        function OneDriveApp() {
        }
        OneDriveApp.onloadInit = function () {
            ErrorHandler.registerErrorObserver(OneDriveState.clearState);
            DomUtilities.getScriptInput();
            Logging.logMessage('initialized');
            var redirectResponse = RedirectUtilities.handleRedirect();
            if (!redirectResponse) {
                return;
            }
            var pickerResponse = ResponseParser.parsePickerResponse(redirectResponse);
            var options = redirectResponse.windowState.options;
            var optionsMode = redirectResponse.windowState.optionsMode;
            if (options.clientId) {
                OneDriveState.clientId = options.clientId;
            } else {
                ErrorHandler.throwError('client id is missing in options');
            }
            switch (optionsMode) {
            case PickerMode[PickerMode.open]:
                var picker = new Picker(options);
                if (pickerResponse.error) {
                    picker.handlePickerError(pickerResponse);
                } else {
                    picker.handlePickerSuccess(pickerResponse);
                }
                break;
            case PickerMode[PickerMode.save]:
                var saver = new Saver(options);
                if (pickerResponse.error) {
                    saver.handleSaverError(pickerResponse);
                } else {
                    saver.handleSaverSuccess(pickerResponse);
                }
                break;
            default:
                ErrorHandler.throwError('invalid value for options.mode: ' + optionsMode);
            }
        };
        OneDriveApp.open = function (options) {
            if (!OneDriveState.readyCheck()) {
                return;
            }
            if (!options) {
                ErrorHandler.throwError('missing picker options');
            }
            Logging.logMessage('open started');
            var picker = new Picker(options);
            picker.launchPicker();
        };
        OneDriveApp.save = function (options) {
            if (!OneDriveState.readyCheck()) {
                return;
            }
            if (!options) {
                ErrorHandler.throwError('missing saver options');
            }
            Logging.logMessage('save started');
            var saver = new Saver(options);
            saver.launchSaver();
        };
        return OneDriveApp;
    }();
module.exports = OneDriveApp;
},{"./OneDriveState":4,"./models/PickerMode":9,"./utilities/DomUtilities":15,"./utilities/ErrorHandler":16,"./utilities/Logging":18,"./utilities/Picker":20,"./utilities/RedirectUtilities":22,"./utilities/ResponseParser":23,"./utilities/Saver":24}],4:[function(_dereq_,module,exports){
var OneDriveState = function () {
        function OneDriveState() {
        }
        OneDriveState.clearState = function () {
            window.name = '';
            OneDriveState._isSdkReady = true;
        };
        OneDriveState.readyCheck = function () {
            if (!OneDriveState._isSdkReady) {
                return false;
            }
            OneDriveState._isSdkReady = false;
            return true;
        };
        OneDriveState.getODCHost = function () {
            return (OneDriveState.debug ? 'live-int' : 'live') + '.com';
        };
        OneDriveState.debug = false;
        OneDriveState._isSdkReady = true;
        return OneDriveState;
    }();
module.exports = OneDriveState;
},{}],5:[function(_dereq_,module,exports){
var ApiEndpoint;
(function (ApiEndpoint) {
    ApiEndpoint[ApiEndpoint['filesV2'] = 0] = 'filesV2';
    ApiEndpoint[ApiEndpoint['graph_odc'] = 1] = 'graph_odc';
    ApiEndpoint[ApiEndpoint['graph_odb'] = 2] = 'graph_odb';
    ApiEndpoint[ApiEndpoint['other'] = 3] = 'other';
}(ApiEndpoint || (ApiEndpoint = {})));
module.exports = ApiEndpoint;
},{}],6:[function(_dereq_,module,exports){
var ErrorType;
(function (ErrorType) {
    ErrorType[ErrorType['badResponse'] = 0] = 'badResponse';
    ErrorType[ErrorType['fileReaderFailure'] = 1] = 'fileReaderFailure';
    ErrorType[ErrorType['popupOpen'] = 2] = 'popupOpen';
    ErrorType[ErrorType['unknown'] = 3] = 'unknown';
    ErrorType[ErrorType['unsupportedFeature'] = 4] = 'unsupportedFeature';
    ErrorType[ErrorType['webRequestFailure'] = 5] = 'webRequestFailure';
}(ErrorType || (ErrorType = {})));
module.exports = ErrorType;
},{}],7:[function(_dereq_,module,exports){
var CallbackInvoker = _dereq_('../utilities/CallbackInvoker'), Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('../utilities/ErrorHandler'), Logging = _dereq_('../utilities/Logging'), OneDriveState = _dereq_('../OneDriveState'), StringUtilities = _dereq_('../utilities/StringUtilities'), TypeValidators = _dereq_('../utilities/TypeValidators'), UrlUtilities = _dereq_('../utilities/UrlUtilities');
var AAD_APPID_PATTERN = new RegExp('^[a-fA-F\\d]{8}-([a-fA-F\\d]{4}-){3}[a-fA-F\\d]{12}$');
var InvokerOptions = function () {
        function InvokerOptions(options) {
            this.openInNewWindow = TypeValidators.validateType(options.openInNewWindow, Constants.TYPE_BOOLEAN, true, true);
            this.expectGlobalFunction = !this.openInNewWindow;
            if (this.expectGlobalFunction) {
                this.cancelName = options.cancel;
                this.errorName = options.error;
            }
            var cancelCallback = TypeValidators.validateCallback(options.cancel, true, this.expectGlobalFunction);
            this.cancel = function () {
                Logging.logMessage('user cancelled operation');
                CallbackInvoker.invokeAppCallback(cancelCallback, true);
            };
            var errorCallback = TypeValidators.validateCallback(options.error, true, this.expectGlobalFunction);
            this.error = function (error) {
                Logging.logError(StringUtilities.format('error occured - code: \'{0}\', message: \'{1}\'', error.errorCode, error.message));
                CallbackInvoker.invokeAppCallback(errorCallback, true, error);
            };
            this.advanced = TypeValidators.validateType(options.advanced, Constants.TYPE_OBJECT, true, { redirectUri: UrlUtilities.trimUrlQuery(window.location.href) });
            if (!this.advanced.redirectUri) {
                this.advanced.redirectUri = UrlUtilities.trimUrlQuery(window.location.href);
            }
            this.clientId = TypeValidators.validateType(options.clientId, Constants.TYPE_STRING);
            this.isSharePointRedirect = !!this.advanced.sharePointTenantPersonalUrl && !!this.advanced.accessToken;
            InvokerOptions.checkClientId(this.clientId);
        }
        InvokerOptions.checkClientId = function (clientId) {
            if (clientId) {
                if (AAD_APPID_PATTERN.test(clientId)) {
                    Logging.logMessage('parsed AAD client id: ' + clientId);
                } else {
                    ErrorHandler.throwError(StringUtilities.format('invalid format for client id \'{0}\' - AAD: 32 characters (HEX) GUID', clientId));
                }
                OneDriveState.clientId = clientId;
            } else {
                ErrorHandler.throwError('client id is missing in options');
            }
        };
        return InvokerOptions;
    }();
module.exports = InvokerOptions;
},{"../Constants":1,"../OneDriveState":4,"../utilities/CallbackInvoker":14,"../utilities/ErrorHandler":16,"../utilities/Logging":18,"../utilities/StringUtilities":25,"../utilities/TypeValidators":26,"../utilities/UrlUtilities":27}],8:[function(_dereq_,module,exports){
var PickerActionType;
(function (PickerActionType) {
    PickerActionType[PickerActionType['download'] = 0] = 'download';
    PickerActionType[PickerActionType['query'] = 1] = 'query';
    PickerActionType[PickerActionType['share'] = 2] = 'share';
}(PickerActionType || (PickerActionType = {})));
module.exports = PickerActionType;
},{}],9:[function(_dereq_,module,exports){
var PickerMode;
(function (PickerMode) {
    PickerMode[PickerMode['open'] = 0] = 'open';
    PickerMode[PickerMode['save'] = 1] = 'save';
}(PickerMode || (PickerMode = {})));
module.exports = PickerMode;
},{}],10:[function(_dereq_,module,exports){
var PickerActionType = _dereq_('./PickerActionType'), CallbackInvoker = _dereq_('../utilities/CallbackInvoker'), Constants = _dereq_('../Constants'), InvokerOptions = _dereq_('./InvokerOptions'), Logging = _dereq_('../utilities/Logging'), PickerMode = _dereq_('./PickerMode'), TypeValidators = _dereq_('../utilities/TypeValidators');
var PickerOptions = function (_super) {
        __extends(PickerOptions, _super);
        function PickerOptions(options) {
            _super.call(this, options);
            if (this.expectGlobalFunction) {
                this.successName = options.success;
            }
            var successCallback = TypeValidators.validateCallback(options.success, false, this.expectGlobalFunction);
            this.success = function (files) {
                Logging.logMessage('picker operation succeeded');
                CallbackInvoker.invokeAppCallback(successCallback, true, files);
            };
            this.multiSelect = TypeValidators.validateType(options.multiSelect, Constants.TYPE_BOOLEAN, true, false);
            var actionName = TypeValidators.validateType(options.action, Constants.TYPE_STRING);
            this.action = PickerActionType[actionName];
        }
        PickerOptions.prototype.isSharing = function () {
            return this.action === PickerActionType.share;
        };
        PickerOptions.prototype.serializeState = function () {
            return {
                optionsMode: PickerMode[PickerMode.open],
                options: {
                    action: PickerActionType[this.action],
                    advanced: this.advanced,
                    clientId: this.clientId,
                    success: this.successName,
                    cancel: this.cancelName,
                    error: this.errorName,
                    multiSelect: this.multiSelect,
                    openInNewWindow: this.openInNewWindow
                }
            };
        };
        return PickerOptions;
    }(InvokerOptions);
module.exports = PickerOptions;
},{"../Constants":1,"../utilities/CallbackInvoker":14,"../utilities/Logging":18,"../utilities/TypeValidators":26,"./InvokerOptions":7,"./PickerActionType":8,"./PickerMode":9}],11:[function(_dereq_,module,exports){
var SaverActionType;
(function (SaverActionType) {
    SaverActionType[SaverActionType['save'] = 0] = 'save';
    SaverActionType[SaverActionType['query'] = 1] = 'query';
}(SaverActionType || (SaverActionType = {})));
module.exports = SaverActionType;
},{}],12:[function(_dereq_,module,exports){
var CallbackInvoker = _dereq_('../utilities/CallbackInvoker'), Constants = _dereq_('../Constants'), DomUtilities = _dereq_('../utilities/DomUtilities'), ErrorHandler = _dereq_('../utilities/ErrorHandler'), InvokerOptions = _dereq_('./InvokerOptions'), Logging = _dereq_('../utilities/Logging'), PickerMode = _dereq_('./PickerMode'), SaverActionType = _dereq_('./SaverActionType'), StringUtilities = _dereq_('../utilities/StringUtilities'), TypeValidators = _dereq_('../utilities/TypeValidators'), UploadType = _dereq_('./UploadType'), UrlUtilities = _dereq_('../utilities/UrlUtilities');
var FORM_UPLOAD_SIZE_LIMIT = 104857600;
var FORM_UPLOAD_SIZE_LIMIT_STRING = '100 MB';
var SaverOptions = function (_super) {
        __extends(SaverOptions, _super);
        function SaverOptions(options) {
            _super.call(this, options);
            this.invalidFile = false;
            if (this.expectGlobalFunction) {
                this.successName = options.success;
                this.progressName = options.progress;
            }
            var successCallback = TypeValidators.validateCallback(options.success, false, this.expectGlobalFunction);
            this.success = function (folder) {
                Logging.logMessage('saver operation succeeded');
                CallbackInvoker.invokeAppCallback(successCallback, true, folder);
            };
            var progressCallback = TypeValidators.validateCallback(options.progress, true, this.expectGlobalFunction);
            this.progress = function (percentage) {
                Logging.logMessage(StringUtilities.format('upload progress: {0}%', percentage));
                CallbackInvoker.invokeAppCallback(progressCallback, false, percentage);
            };
            var actionName = TypeValidators.validateType(options.action, Constants.TYPE_STRING, true, 'query');
            this.action = SaverActionType[actionName];
            if (this.action === SaverActionType.save) {
                this._setFileInfo(options);
            }
        }
        SaverOptions.prototype.serializeState = function () {
            return {
                optionsMode: PickerMode[PickerMode.save],
                options: {
                    action: SaverActionType[this.action],
                    advanced: this.advanced,
                    success: this.successName,
                    progress: this.progressName,
                    cancel: this.cancelName,
                    error: this.errorName,
                    fileName: this.fileName,
                    openInNewWindow: this.openInNewWindow,
                    clientId: this.clientId
                }
            };
        };
        SaverOptions.prototype._setFileInfo = function (options) {
            if (options.sourceInputElementId && options.sourceUri) {
                ErrorHandler.throwError('Only one type of file to save.');
            }
            this.sourceInputElementId = options.sourceInputElementId;
            this.sourceUri = options.sourceUri;
            var fileName = TypeValidators.validateType(options.fileName, Constants.TYPE_STRING, true, null);
            if (this.sourceUri) {
                if (UrlUtilities.isPathFullUrl(this.sourceUri)) {
                    this.uploadType = UploadType.url;
                    this.fileName = fileName || UrlUtilities.getFileNameFromUrl(this.sourceUri);
                    if (!this.fileName) {
                        ErrorHandler.throwError('must supply a file name or a URL that ends with a file name');
                    }
                } else if (UrlUtilities.isPathDataUrl(this.sourceUri)) {
                    this.uploadType = UploadType.dataUrl;
                    this.fileName = fileName;
                    if (!this.fileName) {
                        ErrorHandler.throwError('must supply a file name for data URL uploads');
                    }
                }
            } else if (this.sourceInputElementId) {
                this.uploadType = UploadType.form;
                var fileInputElement = DomUtilities.getElementById(this.sourceInputElementId);
                if (fileInputElement instanceof HTMLInputElement) {
                    if (fileInputElement.type !== 'file') {
                        ErrorHandler.throwError('input elemenet must be of type \'file\'');
                    }
                    if (!fileInputElement.value) {
                        this.error({
                            errorCode: 0,
                            message: 'user has not supplied a file to upload'
                        });
                        this.invalidFile = true;
                        return;
                    }
                    var files = fileInputElement.files;
                    if (!files || !window['FileReader']) {
                        ErrorHandler.throwError('browser does not support Files API');
                    }
                    if (files.length !== 1) {
                        ErrorHandler.throwError('can not upload more than one file at a time');
                    }
                    var fileInput = files[0];
                    if (!fileInput) {
                        ErrorHandler.throwError('missing file input');
                    }
                    if (fileInput.size > FORM_UPLOAD_SIZE_LIMIT) {
                        this.error({
                            errorCode: 1,
                            message: 'the user has selected a file larger than ' + FORM_UPLOAD_SIZE_LIMIT_STRING
                        });
                        this.invalidFile = true;
                        return;
                    }
                    this.fileName = fileName || fileInput.name;
                    this.fileInput = fileInput;
                } else {
                    ErrorHandler.throwError('element was not an instance of an HTMLInputElement');
                }
            } else {
                ErrorHandler.throwError('please specified one type of resource to save');
            }
        };
        return SaverOptions;
    }(InvokerOptions);
module.exports = SaverOptions;
},{"../Constants":1,"../utilities/CallbackInvoker":14,"../utilities/DomUtilities":15,"../utilities/ErrorHandler":16,"../utilities/Logging":18,"../utilities/StringUtilities":25,"../utilities/TypeValidators":26,"../utilities/UrlUtilities":27,"./InvokerOptions":7,"./PickerMode":9,"./SaverActionType":11,"./UploadType":13}],13:[function(_dereq_,module,exports){
var UploadType;
(function (UploadType) {
    UploadType[UploadType['dataUrl'] = 0] = 'dataUrl';
    UploadType[UploadType['form'] = 1] = 'form';
    UploadType[UploadType['url'] = 2] = 'url';
}(UploadType || (UploadType = {})));
module.exports = UploadType;
},{}],14:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), OneDriveState = _dereq_('../OneDriveState');
var CallbackInvoker = function () {
        function CallbackInvoker() {
        }
        CallbackInvoker.invokeAppCallback = function (callback, isFinalCallback) {
            var args = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                args[_i - 2] = arguments[_i];
            }
            if (isFinalCallback) {
                OneDriveState.clearState();
            }
            if (typeof callback === Constants.TYPE_FUNCTION) {
                callback.apply(null, args);
            }
        };
        CallbackInvoker.invokeCallbackAsynchronous = function (callback) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            window.setTimeout(function () {
                callback.apply(null, args);
            }, 0);
        };
        return CallbackInvoker;
    }();
module.exports = CallbackInvoker;
},{"../Constants":1,"../OneDriveState":4}],15:[function(_dereq_,module,exports){
var Logging = _dereq_('./Logging'), OneDriveState = _dereq_('../OneDriveState');
var DOM_DEBUG = 'debug';
var DOM_LOGGING_ID = 'enable-logging';
var DOM_SDK_ID = 'onedrive-js';
var DomUtilities = function () {
        function DomUtilities() {
        }
        DomUtilities.getScriptInput = function () {
            var element = DomUtilities.getElementById(DOM_SDK_ID);
            if (element) {
                var enableLogging = element.getAttribute(DOM_LOGGING_ID);
                if (enableLogging === 'true') {
                    Logging.loggingEnabled = true;
                }
                var debugMode = element.getAttribute(DOM_DEBUG);
                if (debugMode === 'true') {
                    OneDriveState.debug = true;
                }
            }
        };
        DomUtilities.getElementById = function (id) {
            return document.getElementById(id);
        };
        DomUtilities.onDocumentReady = function (callback) {
            if (document.readyState === 'interactive' || document.readyState === 'complete') {
                callback();
            } else {
                document.addEventListener('DOMContentLoaded', callback, false);
            }
        };
        return DomUtilities;
    }();
module.exports = DomUtilities;
},{"../OneDriveState":4,"./Logging":18}],16:[function(_dereq_,module,exports){
var Logging = _dereq_('./Logging');
var ERROR_PREFIX = '[OneDriveSDK Error] ';
var ErrorHandler = function () {
        function ErrorHandler() {
        }
        ErrorHandler.registerErrorObserver = function (callback) {
            ErrorHandler._errorObservers.push(callback);
        };
        ErrorHandler.throwError = function (message) {
            var callbacks = ErrorHandler._errorObservers;
            for (var index in callbacks) {
                try {
                    callbacks[index]();
                } catch (error) {
                    Logging.logError('exception thrown invoking error observer', error);
                }
            }
            throw new Error(ERROR_PREFIX + message);
        };
        ErrorHandler._errorObservers = [];
        return ErrorHandler;
    }();
module.exports = ErrorHandler;
},{"./Logging":18}],17:[function(_dereq_,module,exports){
var ApiEndpoint = _dereq_('../models/ApiEndpoint'), Constants = _dereq_('../Constants'), Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities'), OneDriveState = _dereq_('../OneDriveState'), UrlUtilities = _dereq_('./UrlUtilities'), XHR = _dereq_('./XHR'), StringUtilities = _dereq_('./StringUtilities');
var BATCH_SIZE = 10;
var GraphWrapper = function () {
        function GraphWrapper() {
        }
        GraphWrapper.callGraphShareBatch = function (response, items, createLinkParameters, finished) {
            var failedItemIds = [];
            var totalResponses = 0;
            var invokeCallbacks;
            var sharedItems = [];
            var runBatch = function (batchStart, batchEnd) {
                for (var i = batchStart; i < batchEnd; i++) {
                    var handleShareSuccess = function (sharedItem, permissionFacet) {
                        sharedItem.permissions = permissionFacet;
                        sharedItems.push(sharedItem);
                        invokeCallbacks();
                    };
                    var handleSShareFailed = function (item) {
                        failedItemIds.push(item.id);
                        invokeCallbacks();
                    };
                    GraphWrapper._callGraphShare(response, items[i], createLinkParameters, handleShareSuccess, handleSShareFailed);
                }
            };
            invokeCallbacks = function () {
                if (++totalResponses === items.length) {
                    if (failedItemIds.length) {
                        Logging.logMessage(StringUtilities.format('Create sharing link failed for {0} items', failedItemIds.length));
                    }
                    finished(sharedItems, failedItemIds);
                } else if (totalResponses % BATCH_SIZE === 0) {
                    runBatch(totalResponses, Math.min(items.length, totalResponses + BATCH_SIZE));
                }
            };
            runBatch(0, Math.min(items.length, BATCH_SIZE));
        };
        GraphWrapper.callGraphGetODC = function (response, itemQueryParameters, success, error, createLinkParameters) {
            var apiEndpoint = response.apiEndpoint;
            var accessToken = response.accessToken;
            var itemId = response.itemId;
            var headers = { 'Authorization': 'bearer ' + accessToken };
            switch (apiEndpoint) {
            case ApiEndpoint.graph_odc:
                break;
            case ApiEndpoint.graph_odb:
                break;
            }
            var apiEndpointUrl = UrlUtilities.appendToPath(response.apiEndpointUrl, 'drive/items/' + itemId);
            var queryParameters = {
                    'select': 'id,webUrl',
                    'expand': 'children(' + itemQueryParameters.replace('&', ';') + ')'
                };
            var xhr = new XHR({
                    url: UrlUtilities.appendQueryStrings(apiEndpointUrl, queryParameters),
                    clientId: OneDriveState.clientId,
                    method: Constants.HTTP_GET,
                    apiEndpoint: apiEndpoint,
                    headers: headers
                });
            Logging.logMessage('performing GET on sharing bundle with id: ' + itemId);
            xhr.start(function (xhr, statusCode) {
                var item = ObjectUtilities.deserializeJSON(xhr.responseText);
                if (response.action === 'share') {
                    GraphWrapper._callGraphShare(response, item, createLinkParameters, function (permissionFacet) {
                        item.webUrl = permissionFacet.link.webUrl;
                        success(item);
                    }, function () {
                        Logging.logError(StringUtilities.format('Create link failed for bundle with id {0}:', item.id));
                        error();
                    });
                } else {
                    success(item);
                }
            }, function (xhr, statusCode, timeout) {
                error();
            });
        };
        GraphWrapper.callGraphGetODB = function (response, itemQueryParameters, success, error) {
            var apiEndpointUrl = response.apiEndpointUrl;
            var apiEndpoint = response.apiEndpoint;
            var accessToken = response.accessToken;
            var itemIds = response.itemIds;
            var headers = {
                    'Authorization': 'bearer ' + accessToken,
                    'Cache-Control': 'no-cache, no-store, must-revalidate'
                };
            var successObjects = [];
            var errorCount = 0;
            var totalResponses = 0;
            var numItems = itemIds.length;
            var invokeCallbacks;
            var runBatch = function (batchStart, batchEnd) {
                Logging.logMessage(StringUtilities.format('running batch for items {0} - {1}', batchStart + 1, batchEnd + 1));
                for (var i = batchStart; i < batchEnd; i++) {
                    var itemId = itemIds[i];
                    var url = UrlUtilities.appendToPath(apiEndpointUrl, 'drive/items/' + itemId + '/?' + itemQueryParameters);
                    var xhr = new XHR({
                            url: url,
                            clientId: OneDriveState.clientId,
                            method: Constants.HTTP_GET,
                            apiEndpoint: apiEndpoint,
                            headers: headers
                        });
                    Logging.logMessage('performing GET on item with id: ' + itemId);
                    xhr.start(function (xhr, statusCode, url) {
                        var successResponse = ObjectUtilities.deserializeJSON(xhr.responseText);
                        successObjects.push(successResponse);
                        invokeCallbacks();
                    }, function (xhr, statusCode, timeout) {
                        errorCount++;
                        invokeCallbacks();
                    });
                }
            };
            invokeCallbacks = function () {
                if (++totalResponses === numItems) {
                    if (successObjects.length) {
                        Logging.logMessage(StringUtilities.format('GET metadata succeeded for \'{0}\' items', successObjects.length));
                        success(successObjects);
                    }
                    if (errorCount) {
                        Logging.logMessage(StringUtilities.format('GET metadata failed for \'{0}\' items', errorCount));
                        error(errorCount);
                    }
                } else if (totalResponses % BATCH_SIZE === 0) {
                    runBatch(totalResponses, Math.min(numItems, totalResponses + BATCH_SIZE));
                }
            };
            runBatch(0, Math.min(numItems, BATCH_SIZE));
        };
        GraphWrapper._callGraphShare = function (response, item, createLinkParameters, success, error) {
            var shareRequestUrl = response.apiEndpointUrl + 'drive/items/' + item.id + '/' + response.apiActionNamingSpace + '.createLink';
            var postXhr = new XHR({
                    url: shareRequestUrl,
                    clientId: OneDriveState.clientId,
                    method: Constants.HTTP_POST,
                    apiEndpoint: response.apiEndpoint,
                    headers: { 'Authorization': 'bearer ' + response.accessToken },
                    json: JSON.stringify(createLinkParameters)
                });
            postXhr.start(function (postXhr, statusCode, url) {
                Logging.logMessage(StringUtilities.format('POST createLink succeeded via path {0}', shareRequestUrl));
                success(item, ObjectUtilities.deserializeJSON(postXhr.responseText));
            }, function (xhr, statusCode, timeout) {
                Logging.logMessage(StringUtilities.format('POST createLink failed via path {0}', shareRequestUrl));
                error(item);
            });
        };
        return GraphWrapper;
    }();
module.exports = GraphWrapper;
},{"../Constants":1,"../OneDriveState":4,"../models/ApiEndpoint":5,"./Logging":18,"./ObjectUtilities":19,"./StringUtilities":25,"./UrlUtilities":27,"./XHR":29}],18:[function(_dereq_,module,exports){
var ENABLE_LOGGING = 'onedrive_enable_logging';
var LOG_PREFIX = '[OneDriveSDK] ';
var Logging = function () {
        function Logging() {
        }
        Logging.logError = function (message) {
            var objects = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                objects[_i - 1] = arguments[_i];
            }
            Logging._log(message, true, objects);
        };
        Logging.logMessage = function (message) {
            Logging._log(message, false);
        };
        Logging._log = function (message, isError) {
            var objects = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                objects[_i - 2] = arguments[_i];
            }
            if (isError || Logging.loggingEnabled || window[ENABLE_LOGGING]) {
                console.log(LOG_PREFIX + message, objects);
            }
        };
        Logging.loggingEnabled = false;
        return Logging;
    }();
module.exports = Logging;
},{}],19:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), Logging = _dereq_('./Logging');
var ObjectUtilities = function () {
        function ObjectUtilities() {
        }
        ObjectUtilities.shallowClone = function (object) {
            if (typeof object !== Constants.TYPE_OBJECT || !object) {
                return null;
            }
            var clonedObject = {};
            for (var key in object) {
                if (object.hasOwnProperty(key)) {
                    clonedObject[key] = object[key];
                }
            }
            return clonedObject;
        };
        ObjectUtilities.deserializeJSON = function (text) {
            var deserializedObject = null;
            try {
                deserializedObject = JSON.parse(text);
            } catch (error) {
                Logging.logError('deserialization error' + error);
            }
            if (typeof deserializedObject !== Constants.TYPE_OBJECT || deserializedObject === null) {
                deserializedObject = {};
            }
            return deserializedObject;
        };
        ObjectUtilities.serializeJSON = function (value) {
            return JSON.stringify(value);
        };
        return ObjectUtilities;
    }();
module.exports = ObjectUtilities;
},{"../Constants":1,"./Logging":18}],20:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('./ErrorHandler'), ErrorType = _dereq_('../models/ErrorType'), Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities'), Popup = _dereq_('./Popup'), PickerOptions = _dereq_('../models/PickerOptions'), RedirectUtilities = _dereq_('./RedirectUtilities'), StringUtilities = _dereq_('./StringUtilities'), UrlUtilities = _dereq_('./UrlUtilities'), GraphWrapper = _dereq_('./GraphWrapper'), PickerActionType = _dereq_('../models/PickerActionType');
var VROOM_THUMBNAIL_SIZES = [
        'large',
        'medium',
        'small'
    ];
var Picker = function () {
        function Picker(options) {
            var clonedOptions = ObjectUtilities.shallowClone(options);
            this._pickerOptions = new PickerOptions(clonedOptions);
        }
        Picker.prototype.launchPicker = function () {
            var _this = this;
            var pickerOptions = this._pickerOptions;
            var windowState = pickerOptions.serializeState();
            if (pickerOptions.openInNewWindow) {
                var popupUrl = UrlUtilities.appendQueryStrings(pickerOptions.advanced.redirectUri, {
                        'sdk_state': JSON.stringify(windowState),
                        'state': Constants.STATE_OPEN_POPUP
                    });
                if (pickerOptions.isSharePointRedirect) {
                    popupUrl = UrlUtilities.appendQueryString(popupUrl, 'sp', '1');
                }
                var popup = new Popup(popupUrl, function (response) {
                        _this.handlePickerSuccess(response);
                    }, function (response) {
                        _this.handlePickerError(response);
                    });
                if (!popup.openPopup()) {
                    pickerOptions.error(Constants.ERROR_POPUP_OPEN);
                }
            } else {
                var url = UrlUtilities.trimUrlQuery(window.location.href);
                if (pickerOptions.isSharePointRedirect) {
                    url = UrlUtilities.appendQueryString(url, 'sp', '1');
                }
                RedirectUtilities.redirectToAADLogin(pickerOptions, windowState);
            }
        };
        Picker.prototype.handlePickerSuccess = function (pickerResponse) {
            var pickerType = pickerResponse.pickerType;
            var options = this._pickerOptions;
            var queryParameter = Constants.DEFAULT_QUERY_ITEM_PARAMETER;
            if (options.action === PickerActionType.query && options.advanced.queryParameters) {
                queryParameter = options.advanced.queryParameters;
            } else if (options.action === PickerActionType.download) {
                queryParameter += ',@content.downloadUrl';
            }
            switch (pickerType) {
            case Constants.STATE_MSA_PICKER:
                this._handleMSAOpenResponse(pickerResponse, queryParameter);
                break;
            case Constants.STATE_AAD_PICKER:
                this._handleAADOpenResponse(pickerResponse, queryParameter);
                break;
            default:
                ErrorHandler.throwError('invalid value for picker type: ' + pickerType);
            }
        };
        Picker.prototype.handlePickerError = function (errorResponse) {
            if (errorResponse.error === Constants.ERROR_ACCESS_DENIED) {
                this._pickerOptions.cancel();
            } else {
                this._pickerOptions.error({
                    errorCode: ErrorType.unknown,
                    message: 'something went wrong: ' + errorResponse.error
                });
            }
        };
        Picker.prototype._handleMSAOpenResponse = function (pickerResponse, queryParameters) {
            var _this = this;
            var options = this._pickerOptions;
            var createLinkParameters;
            var handleGetSuccess;
            if (options.action === PickerActionType.share) {
                createLinkParameters = options.advanced.createLinkParameters || { 'type': 'view' };
                handleGetSuccess = function (getReponse) {
                    GraphWrapper.callGraphShareBatch(pickerResponse, getReponse.children, createLinkParameters, function (sharedItems, sharingFailedItems) {
                        _this._handleSuccessResponse({
                            webUrl: getReponse.webUrl,
                            files: sharedItems
                        });
                    });
                };
            } else {
                handleGetSuccess = function (getReponse) {
                    _this._handleSuccessResponse({
                        webUrl: getReponse.webUrl,
                        files: getReponse.children
                    }, true);
                };
            }
            GraphWrapper.callGraphGetODC(pickerResponse, queryParameters, handleGetSuccess, function () {
                options.error(Constants.ERROR_WEB_REQUEST);
            }, createLinkParameters);
        };
        Picker.prototype._handleAADOpenResponse = function (pickerResponse, queryParameters) {
            var _this = this;
            var options = this._pickerOptions;
            var createLinkParameters;
            var handleGetSuccess;
            if (options.action === PickerActionType.share) {
                createLinkParameters = options.advanced.createLinkParameters || {
                    'type': 'view',
                    'scope': 'organization'
                };
                handleGetSuccess = function (getReponse) {
                    GraphWrapper.callGraphShareBatch(pickerResponse, getReponse, createLinkParameters, function (sharedItems, sharingFailedItems) {
                        _this._handleSuccessResponse({
                            webUrl: null,
                            files: sharedItems
                        });
                    });
                };
            } else {
                handleGetSuccess = function (getReponse) {
                    _this._handleSuccessResponse({
                        webUrl: null,
                        files: getReponse
                    }, true);
                };
            }
            GraphWrapper.callGraphGetODB(pickerResponse, queryParameters, handleGetSuccess, function () {
                options.error(Constants.ERROR_WEB_REQUEST);
            });
        };
        Picker.prototype._handleSuccessResponse = function (response, isMSA) {
            var options = this._pickerOptions;
            var files = {
                    link: options.getWebLinks ? response.webUrl : null,
                    values: []
                };
            var pickerFiles = response.files;
            if (!pickerFiles || !pickerFiles.length) {
                options.error({
                    errorCode: ErrorType.badResponse,
                    message: 'no files returned'
                });
            }
            Logging.logMessage(StringUtilities.format('returning \'{0}\' files picked', pickerFiles.length));
            for (var i = 0; i < pickerFiles.length; i++) {
                var pickerFile = pickerFiles[i];
                if (isMSA) {
                    var thumbnails = [];
                    var fileThumbnails = pickerFile.thumbnails && pickerFile.thumbnails[0];
                    if (fileThumbnails) {
                        for (var j = 0; j < VROOM_THUMBNAIL_SIZES.length; j++) {
                            thumbnails.push(fileThumbnails[VROOM_THUMBNAIL_SIZES[j]].url);
                        }
                        pickerFile.thumbnails = thumbnails;
                    }
                }
                files.values.push(pickerFile);
            }
            options.success(files);
        };
        return Picker;
    }();
module.exports = Picker;
},{"../Constants":1,"../models/ErrorType":6,"../models/PickerActionType":8,"../models/PickerOptions":10,"./ErrorHandler":16,"./GraphWrapper":17,"./Logging":18,"./ObjectUtilities":19,"./Popup":21,"./RedirectUtilities":22,"./StringUtilities":25,"./UrlUtilities":27}],21:[function(_dereq_,module,exports){
var CallbackInvoker = _dereq_('./CallbackInvoker'), Constants = _dereq_('../Constants'), Logging = _dereq_('./Logging'), ResponseParser = _dereq_('./ResponseParser');
var POPUP_WIDTH = 800;
var POPUP_HEIGHT = 650;
var POPUP_PINGER_INTERVAL = 500;
var Popup = function () {
        function Popup(url, successCallback, errorCallabck) {
            this._messageCallbackInvoked = false;
            this._url = url;
            this._successCallback = successCallback;
            this._failureCallback = errorCallabck;
        }
        Popup.canReceiveMessage = function (event) {
            return event.origin === window.location.origin;
        };
        Popup._createPopupFeatures = function () {
            var left = window.screenX + Math.max(window.outerWidth - POPUP_WIDTH, 0) / 2;
            var top = window.screenY + Math.max(window.outerHeight - POPUP_HEIGHT, 0) / 2;
            var features = [
                    'width=' + POPUP_WIDTH,
                    'height=' + POPUP_HEIGHT,
                    'top=' + top,
                    'left=' + left,
                    'status=no',
                    'resizable=yes',
                    'toolbar=no',
                    'menubar=no',
                    'scrollbars=yes'
                ];
            return features.join(',');
        };
        Popup.prototype.openPopup = function () {
            if (Popup._currentPopup && Popup._currentPopup._isPopupOpen()) {
                return false;
            }
            if (!window.onedriveReceiveMessage) {
                window.onedriveReceiveMessage = function (data) {
                    var currentPopup = Popup._currentPopup;
                    if (currentPopup && currentPopup._isPopupOpen()) {
                        Popup._currentPopup = null;
                        var response = ResponseParser.parsePickerResponse(data);
                        currentPopup._messageCallbackInvoked = true;
                        if (response.error === undefined) {
                            CallbackInvoker.invokeCallbackAsynchronous(currentPopup._successCallback, response);
                        } else {
                            CallbackInvoker.invokeCallbackAsynchronous(currentPopup._failureCallback, response);
                        }
                    }
                };
            }
            this._popup = window.open(this._url, '_blank', Popup._createPopupFeatures());
            this._popup.focus();
            this._createPopupPinger();
            Popup._currentPopup = this;
            return true;
        };
        Popup.prototype._createPopupPinger = function () {
            var _this = this;
            var interval = window.setInterval(function () {
                    if (_this._isPopupOpen()) {
                        _this._popup.postMessage('ping', '*');
                    } else {
                        window.clearInterval(interval);
                        Popup._currentPopup = null;
                        if (!_this._messageCallbackInvoked) {
                            Logging.logMessage('closed callback');
                            _this._failureCallback({ error: Constants.ERROR_ACCESS_DENIED });
                        }
                    }
                }, POPUP_PINGER_INTERVAL);
        };
        Popup.prototype._isPopupOpen = function () {
            return this._popup !== null && !this._popup.closed;
        };
        return Popup;
    }();
module.exports = Popup;
},{"../Constants":1,"./CallbackInvoker":14,"./Logging":18,"./ResponseParser":23}],22:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), DomUtilities = _dereq_('./DomUtilities'), ErrorHandler = _dereq_('./ErrorHandler'), Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities'), OneDriveState = _dereq_('../OneDriveState'), PickerMode = _dereq_('../models/PickerMode'), StringUtilities = _dereq_('./StringUtilities'), TypeValidators = _dereq_('./TypeValidators'), UrlUtilities = _dereq_('./UrlUtilities'), WindowState = _dereq_('./WindowState'), XHR = _dereq_('./XHR');
var AAD_LOGIN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
var RedirectUtilities = function () {
        function RedirectUtilities() {
        }
        RedirectUtilities.redirect = function (url, values, windowState) {
            if (values === void 0) {
                values = null;
            }
            if (windowState === void 0) {
                windowState = null;
            }
            if (values) {
                WindowState.setWindowState(values, windowState);
            }
            window.location.replace(url);
        };
        RedirectUtilities.handleRedirect = function () {
            var queryParameters = UrlUtilities.readCurrentUrlParameters();
            var serializedState = WindowState.getWindowState();
            var state = queryParameters[Constants.PARAM_STATE] || serializedState[Constants.PARAM_STATE];
            if (!state && queryParameters[Constants.PARAM_ERROR] === Constants.ERROR_ACCESS_DENIED) {
                queryParameters[Constants.PARAM_STATE] = Constants.STATE_MSA_PICKER;
            } else if (state === Constants.STATE_AAD_PICKER) {
                queryParameters[Constants.PARAM_STATE] = Constants.STATE_AAD_PICKER;
            }
            var redirectState = queryParameters[Constants.PARAM_STATE];
            if (!redirectState) {
                return null;
            }
            Logging.logMessage('current state: ' + redirectState);
            if (redirectState === Constants.STATE_OPEN_POPUP) {
                serializedState = JSON.parse(queryParameters[Constants.PARAM_SDK_STATE]);
            }
            var options = serializedState['options'];
            var optionsModeName = serializedState['optionsMode'];
            var optionsMode = PickerMode[optionsModeName];
            if (!options) {
                ErrorHandler.throwError('missing options from serialized state');
            }
            var inPopupFlow = TypeValidators.validateType(options.openInNewWindow, Constants.TYPE_BOOLEAN);
            if (inPopupFlow) {
                RedirectUtilities._displayOverlay();
            }
            if (queryParameters[Constants.PARAM_SPREDIRECT]) {
                RedirectUtilities._redirectToTenant(options, optionsMode, options.openInNewWindow);
                return null;
            }
            switch (redirectState) {
            case Constants.STATE_OPEN_POPUP:
                RedirectUtilities.redirectToAADLogin(options, serializedState);
                break;
            case Constants.STATE_AAD_LOGIN:
                RedirectUtilities._handleAADLogin(queryParameters, options, optionsMode);
                break;
            case Constants.STATE_MSA_PICKER:
            case Constants.STATE_AAD_PICKER:
                var pickerResponse = {
                        windowState: serializedState,
                        queryParameters: queryParameters
                    };
                Logging.logMessage('sending invoker response');
                if (inPopupFlow) {
                    RedirectUtilities._sendResponse(pickerResponse);
                } else {
                    return pickerResponse;
                }
                break;
            default:
                ErrorHandler.throwError('invalid value for redirect state: ' + redirectState);
            }
            return null;
        };
        RedirectUtilities.redirectToAADLogin = function (options, stateValues) {
            if (!options.openInNewWindow) {
                RedirectUtilities._displayOverlay();
            }
            if (!options.clientId) {
                ErrorHandler.throwError('clientId is missing in options');
            }
            var url = UrlUtilities.appendQueryStrings(AAD_LOGIN_URL, {
                    'redirect_uri': options.advanced.redirectUri,
                    'client_id': options.clientId,
                    'scope': 'openid https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/User.Read',
                    'response_mode': 'fragment',
                    'state': Constants.STATE_AAD_LOGIN,
                    'nonce': UrlUtilities.generateNonce()
                });
            url += '&response_type=id_token+token';
            RedirectUtilities.redirect(url, stateValues);
        };
        RedirectUtilities._handleAADLogin = function (queryParameters, options, optionsMode) {
            if (!options.openInNewWindow) {
                RedirectUtilities._displayOverlay();
            }
            var id_token = queryParameters[Constants.PARAM_ID_TOKEN];
            if (!id_token) {
                ErrorHandler.throwError('id_toekn is missing in returned parameters');
            }
            options.advanced.accessToken = queryParameters[Constants.PARAM_ACCESS_TOKEN];
            var open_id = JSON.parse(atob(id_token.split('.')[1]));
            if (open_id.tid === Constants.CUSTOMER_TID) {
                RedirectUtilities._redirectToODCPicker(options, optionsMode);
            } else {
                RedirectUtilities._redirectToTenant(options, optionsMode, options.openInNewWindow);
            }
        };
        RedirectUtilities._redirectToODCPicker = function (options, optionsMode) {
            var access, viewType, selectionMode;
            switch (optionsMode) {
            case PickerMode.open:
                access = 'read';
                viewType = 'file';
                selectionMode = options.multiSelect ? 'multi' : 'single';
                break;
            case PickerMode.save:
                access = 'readwrite';
                viewType = 'folder';
                selectionMode = 'single';
                break;
            }
            var baseUrl = 'https://login.' + OneDriveState.getODCHost() + '/oauth20_authorize.srf';
            var url = UrlUtilities.appendQueryStrings(baseUrl, {
                    'client_id': options.clientId,
                    'redirect_uri': options.advanced.redirectUri,
                    'response_type': 'token',
                    'scope': 'onedrive_onetime.access:' + access + viewType + '|' + selectionMode + '|downloadLink',
                    'state': Constants.STATE_MSA_PICKER
                });
            RedirectUtilities.redirect(url, { 'options': options });
        };
        RedirectUtilities._redirectToTenant = function (options, optionsMode, inPopupFlow) {
            var redirectToTenant = function (tenantUrl) {
                var access, viewType, selectionMode;
                switch (optionsMode) {
                case PickerMode.open:
                    access = 'read';
                    viewType = 'file';
                    selectionMode = options.multiSelect ? 'multi' : 'single';
                    break;
                case PickerMode.save:
                    access = 'readwrite';
                    viewType = 'folder';
                    selectionMode = 'single';
                    break;
                }
                RedirectUtilities.redirect(tenantUrl + '_layouts/onedrive.aspx', {
                    'ODBParams': {
                        'p': '2',
                        'ru': options.advanced.redirectUri,
                        'selection_mode': selectionMode,
                        'access': access,
                        'view_type': viewType
                    },
                    'options': options
                });
            };
            if (options.advanced.sharePointTenantPersonalUrl) {
                redirectToTenant(options.advanced.sharePointTenantPersonalUrl);
            } else {
                var xhr = new XHR({
                        url: UrlUtilities.appendQueryString(Constants.GRAPH_URL + 'me', '$select', 'mySite'),
                        method: Constants.HTTP_GET,
                        headers: {
                            'Authorization': 'bearer ' + options.advanced.accessToken,
                            'Accept': 'application/json'
                        }
                    });
                xhr.start(function (xhr, statusCode) {
                    var response = ObjectUtilities.deserializeJSON(xhr.responseText);
                    if (response.mySite) {
                        redirectToTenant(response.mySite);
                    } else {
                        ErrorHandler.throwError(StringUtilities.format('Cannot find the personal tenant url, response text: {0}', xhr.responseText));
                    }
                }, function (xhr, statusCode, timeout) {
                    RedirectUtilities._handleError(StringUtilities.format('graph/me request failed, status code: \'{0}\', response text: \'{1}\'', XHR.statusCodeToString(statusCode), xhr.responseText), inPopupFlow);
                });
            }
        };
        RedirectUtilities._sendResponse = function (response) {
            var parentWindow = window.opener;
            if (parentWindow.onedriveReceiveMessage) {
                parentWindow.onedriveReceiveMessage(response);
            } else {
                Logging.logError('error in window\'s opener, pop up will close.');
                RedirectUtilities._handleError('SDK message receiver is undefined.', true);
            }
            window.close();
        };
        RedirectUtilities._handleError = function (error, inPopupFlow) {
            var errorQueryParametere = {};
            errorQueryParametere[Constants.PARAM_ERROR] = error;
            if (inPopupFlow) {
                RedirectUtilities._sendResponse({ queryParameters: errorQueryParametere });
            } else {
                Logging.logMessage('error in picker flow, redirecting back to app');
                errorQueryParametere[Constants.PARAM_STATE] = Constants.STATE_AAD_PICKER;
                var redirectUrl = UrlUtilities.trimUrlQuery(window.location.href);
                RedirectUtilities.redirect(UrlUtilities.appendQueryStrings(redirectUrl, errorQueryParametere));
            }
        };
        RedirectUtilities._displayOverlay = function () {
            var overlay = document.createElement('div');
            var overlayStyle = [
                    'position: fixed',
                    'width: 100%',
                    'height: 100%',
                    'top: 0px',
                    'left: 0px',
                    'background-color: white',
                    'opacity: 1',
                    'z-index: 10000'
                ];
            overlay.id = 'od-overlay';
            overlay.style.cssText = overlayStyle.join(';');
            var spinner = document.createElement('img');
            var spinnerStyle = [
                    'position: absolute',
                    'top: calc(50% - 40px)',
                    'left: calc(50% - 40px)'
                ];
            spinner.id = 'od-spinner';
            spinner.src = 'https://p.sfx.ms/common/spinner_grey_40_transparent.gif';
            spinner.style.cssText = spinnerStyle.join(';');
            overlay.appendChild(spinner);
            var hiddenStyle = document.createElement('style');
            hiddenStyle.type = 'text/css';
            hiddenStyle.innerHTML = 'body { visibility: hidden !important; }';
            document.head.appendChild(hiddenStyle);
            DomUtilities.onDocumentReady(function () {
                var documentBody = document.body;
                if (documentBody !== null) {
                    documentBody.insertBefore(overlay, documentBody.firstChild);
                } else {
                    document.createElement('body').appendChild(overlay);
                }
                document.head.removeChild(hiddenStyle);
            });
        };
        return RedirectUtilities;
    }();
module.exports = RedirectUtilities;
},{"../Constants":1,"../OneDriveState":4,"../models/PickerMode":9,"./DomUtilities":15,"./ErrorHandler":16,"./Logging":18,"./ObjectUtilities":19,"./StringUtilities":25,"./TypeValidators":26,"./UrlUtilities":27,"./WindowState":28,"./XHR":29}],23:[function(_dereq_,module,exports){
var ApiEndpoint = _dereq_('../models/ApiEndpoint'), Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('./ErrorHandler'), Logging = _dereq_('./Logging');
var CID_PADDING = '0000000000000000';
var CID_PADDING_LENGTH = CID_PADDING.length;
var MSA_SCOPE_RESPONSE_PATTERN = new RegExp('^\\w+\\.\\w+:\\w+[\\|\\w+]+:([\\w]+\\!\\d+)(?:\\!(.+))*$');
var ResponseParser = function () {
        function ResponseParser() {
        }
        ResponseParser.parsePickerResponse = function (response) {
            Logging.logMessage('parsing picker response');
            var serializedState = response.windowState;
            if (!serializedState) {
                ErrorHandler.throwError('missing windowState from picker response');
            }
            var queryParameters = response.queryParameters;
            if (!queryParameters) {
                ErrorHandler.throwError('missing queryParameters from picker response');
            }
            var responseError = queryParameters[Constants.PARAM_ERROR];
            if (responseError) {
                return { error: responseError };
            }
            var pickerType = queryParameters[Constants.PARAM_STATE];
            var result = {
                    pickerType: pickerType,
                    accessToken: serializedState.options.advanced.accessToken
                };
            result.action = serializedState.options.action;
            switch (pickerType) {
            case Constants.STATE_MSA_PICKER:
                ResponseParser._parseMSAResponse(result, queryParameters);
                break;
            case Constants.STATE_AAD_PICKER:
                ResponseParser._parseAADResponse(result, queryParameters, serializedState);
                break;
            default:
                ErrorHandler.throwError('invalid value for picker type: ' + pickerType);
            }
            if (!result.accessToken) {
                ErrorHandler.throwError('missing access token');
            }
            if (!result.apiEndpointUrl) {
                ErrorHandler.throwError('missing API endpoint URL');
            }
            return result;
        };
        ResponseParser._parseMSAResponse = function (result, queryParameters) {
            result.apiEndpoint = ApiEndpoint.graph_odc;
            result.apiEndpointUrl = Constants.GRAPH_URL;
            result.apiActionNamingSpace = 'microsoft.graph';
            var responseScope = queryParameters['scope'];
            if (!responseScope) {
                ErrorHandler.throwError('missing \'scope\' paramter from MSA picker response');
            }
            var scopes = responseScope.split(' ');
            var scopeResult;
            for (var i = 0; i < scopes.length && !scopeResult; i++) {
                scopeResult = MSA_SCOPE_RESPONSE_PATTERN.exec(scopes[i]);
            }
            if (!scopeResult) {
                ErrorHandler.throwError('scope was not formatted correctly');
            }
            var rawResult = scopeResult[1].split('_');
            var rawItemId = rawResult[1];
            var splitIndex = rawItemId.indexOf('!');
            var rawItemIdPart1 = rawItemId.substring(0, splitIndex);
            var rawItemIdPart2 = rawItemId.substring(splitIndex);
            var ownerCid = ResponseParser._leftPadCid(rawItemIdPart1);
            var itemId = ownerCid + rawItemIdPart2;
            result.ownerCid = ownerCid;
            result.itemId = itemId;
            result.authKey = scopeResult[2];
        };
        ResponseParser._parseAADResponse = function (result, queryParameters, state) {
            if (state.options.advanced.sharePointTenantPersonalUrl) {
                result.apiEndpointUrl = state.options.advanced.sharePointTenantPersonalUrl + '_api/v2.0/';
                result.apiEndpoint = ApiEndpoint.filesV2;
                result.apiActionNamingSpace = 'action';
            } else {
                result.apiEndpoint = ApiEndpoint.graph_odb;
                result.apiEndpointUrl = Constants.GRAPH_URL + 'me/';
                result.apiActionNamingSpace = 'microsoft.graph';
            }
            var itemIds = queryParameters['item-id'].split(',');
            if (!itemIds.length) {
                ErrorHandler.throwError('missing item id(s)');
            }
            result.itemIds = itemIds;
        };
        ResponseParser._leftPadCid = function (cid) {
            if (cid.length === CID_PADDING_LENGTH) {
                return cid;
            }
            return CID_PADDING.substring(0, CID_PADDING_LENGTH - cid.length) + cid;
        };
        return ResponseParser;
    }();
module.exports = ResponseParser;
},{"../Constants":1,"../models/ApiEndpoint":5,"./ErrorHandler":16,"./Logging":18}],24:[function(_dereq_,module,exports){
var CallbackInvoker = _dereq_('./CallbackInvoker'), Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('./ErrorHandler'), ErrorType = _dereq_('../models/ErrorType'), GraphWrapper = _dereq_('./GraphWrapper'), Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities'), OneDriveState = _dereq_('../OneDriveState'), Popup = _dereq_('./Popup'), RedirectUtilities = _dereq_('./RedirectUtilities'), SaverActionType = _dereq_('../models/SaverActionType'), SaverOptions = _dereq_('../models/SaverOptions'), StringUtilities = _dereq_('./StringUtilities'), UploadType = _dereq_('../models/UploadType'), UrlUtilities = _dereq_('./UrlUtilities'), XHR = _dereq_('./XHR');
var POLLING_INTERVAL = 1000;
var POLLING_COUNTER = 5;
var Saver = function () {
        function Saver(options) {
            var clonedOptions = ObjectUtilities.shallowClone(options);
            this._saverOptions = new SaverOptions(clonedOptions);
        }
        Saver.prototype.launchSaver = function () {
            var _this = this;
            var saverOptions = this._saverOptions;
            if (saverOptions.invalidFile) {
                return;
            }
            var windowState = saverOptions.serializeState();
            if (saverOptions.openInNewWindow) {
                var popupUrl = UrlUtilities.appendQueryStrings(saverOptions.advanced.redirectUri, {
                        'sdk_state': JSON.stringify(windowState),
                        'state': Constants.STATE_OPEN_POPUP
                    });
                if (saverOptions.isSharePointRedirect) {
                    popupUrl = UrlUtilities.appendQueryString(popupUrl, 'sp', '1');
                }
                var popup = new Popup(popupUrl, function (response) {
                        _this.handleSaverSuccess(response);
                    }, function (response) {
                        _this.handleSaverError(response);
                    });
                if (!popup.openPopup()) {
                    saverOptions.error(Constants.ERROR_POPUP_OPEN);
                }
            } else {
                var url = UrlUtilities.trimUrlQuery(window.location.href);
                if (saverOptions.isSharePointRedirect) {
                    url = UrlUtilities.appendQueryString(url, 'sp', '1');
                }
                RedirectUtilities.redirectToAADLogin(saverOptions, windowState);
            }
        };
        Saver.prototype.handleSaverSuccess = function (saverResponse) {
            var _this = this;
            var pickerType = saverResponse.pickerType;
            var queryParameters = Constants.DEFAULT_QUERY_ITEM_PARAMETER;
            var options = this._saverOptions;
            if (options.action === SaverActionType.query && options.advanced.queryParameters) {
                queryParameters = options.advanced.queryParameters;
            }
            switch (pickerType) {
            case Constants.STATE_MSA_PICKER:
                GraphWrapper.callGraphGetODC(saverResponse, queryParameters, function (apiResponse) {
                    var apiResponseValue = apiResponse.children;
                    if (!apiResponseValue) {
                        ErrorHandler.throwError('empty API response');
                    }
                    var folder = apiResponseValue[0];
                    if (!folder || apiResponseValue.length !== 1) {
                        ErrorHandler.throwError('incorrect number of folders returned');
                    }
                    if (options.action === SaverActionType.query) {
                        options.success({
                            link: null,
                            values: [folder]
                        });
                    } else if (options.action === SaverActionType.save) {
                        _this._executeUpload(saverResponse, folder);
                    }
                }, function () {
                    options.error(Constants.ERROR_WEB_REQUEST);
                });
                break;
            case Constants.STATE_AAD_PICKER:
                var folderIds = saverResponse.itemIds;
                if (folderIds.length !== 1) {
                    ErrorHandler.throwError('incorrect number of folders returned');
                }
                var folderId = folderIds[0];
                if (!folderId) {
                    folderId = 'root';
                }
                if (OneDriveState.debug && OneDriveDebug) {
                    OneDriveDebug.accessToken = saverResponse.accessToken;
                }
                if (options.action === SaverActionType.query) {
                    GraphWrapper.callGraphGetODB(saverResponse, queryParameters, function (apiResponse) {
                        options.success({
                            link: null,
                            values: [apiResponse]
                        });
                    }, function () {
                        options.error(Constants.ERROR_WEB_REQUEST);
                    });
                } else {
                    this._executeUpload(saverResponse, { id: folderId });
                }
                break;
            default:
                ErrorHandler.throwError('invalid value for picker type: ' + pickerType);
            }
        };
        Saver.prototype.handleSaverError = function (errorResponse) {
            if (errorResponse.error === Constants.ERROR_ACCESS_DENIED) {
                this._saverOptions.cancel();
            } else {
                this._saverOptions.error({
                    errorCode: ErrorType.unknown,
                    message: 'something went wrong: ' + errorResponse.error
                });
            }
        };
        Saver.prototype._executeUpload = function (saverResponse, folder) {
            var uploadType = this._saverOptions.uploadType;
            Logging.logMessage(StringUtilities.format('beginning \'{0}\' upload', UploadType[uploadType]));
            var accessToken = saverResponse.accessToken;
            switch (uploadType) {
            case UploadType.dataUrl:
            case UploadType.url:
                this._executeUrlUpload(saverResponse, folder, accessToken, uploadType);
                break;
            case UploadType.form:
                this._executeFormUpload(saverResponse, folder, accessToken);
                break;
            default:
                ErrorHandler.throwError('invalid value for upload type: ' + uploadType);
            }
        };
        Saver.prototype._executeUrlUpload = function (saverResponse, folder, accessToken, uploadType) {
            var _this = this;
            var options = this._saverOptions;
            if (uploadType === UploadType.url && saverResponse.pickerType === Constants.STATE_AAD_PICKER) {
                options.error({
                    errorCode: ErrorType.unsupportedFeature,
                    message: 'URL upload not supported for AAD'
                });
                return;
            }
            var uploadUrl = UrlUtilities.appendToPath(saverResponse.apiEndpointUrl, 'drive/items/' + folder.id + '/children');
            var requestHeaders = {};
            requestHeaders['Prefer'] = 'respond-async';
            requestHeaders['Authorization'] = 'bearer ' + accessToken;
            var body = {
                    '@microsoft.graph.sourceUrl': options.sourceUri,
                    'name': options.fileName,
                    'file': {}
                };
            var xhr = new XHR({
                    url: uploadUrl,
                    clientId: OneDriveState.clientId,
                    method: Constants.HTTP_POST,
                    headers: requestHeaders,
                    json: ObjectUtilities.serializeJSON(body),
                    apiEndpoint: saverResponse.apiEndpoint
                });
            xhr.start(function (xhr, statusCode) {
                if (uploadType === UploadType.dataUrl && (statusCode === 200 || statusCode === 201)) {
                    if (OneDriveState.debug && OneDriveDebug) {
                        OneDriveDebug.itemUrl = ObjectUtilities.deserializeJSON(xhr.responseText)['@odata.id'];
                    }
                    options.success(ObjectUtilities.deserializeJSON(xhr.responseText));
                } else if (uploadType === UploadType.url && statusCode === 202) {
                    var location_1 = xhr.getResponseHeader('Location');
                    if (!location_1) {
                        options.error({
                            errorCode: ErrorType.badResponse,
                            message: 'missing \'Location\' header on response'
                        });
                    }
                    _this._beginPolling(location_1, accessToken);
                } else {
                    options.error(Constants.ERROR_WEB_REQUEST);
                }
            }, function (xhr, statusCode, timeout) {
                options.error(Constants.ERROR_WEB_REQUEST);
            });
        };
        Saver.prototype._executeFormUpload = function (saverResponse, folder, accessToken) {
            var options = this._saverOptions;
            var uploadSource = options.fileInput;
            var reader = null;
            if (window['File'] && uploadSource instanceof window['File']) {
                reader = new FileReader();
            } else {
                ErrorHandler.throwError('file reader not supported');
            }
            reader.onerror = function (event) {
                Logging.logError('failed to read or upload the file', event);
                options.error({
                    errorCode: ErrorType.fileReaderFailure,
                    message: 'failed to read or upload the file, see console log for details'
                });
            };
            reader.onload = function (event) {
                var uploadUrl = UrlUtilities.appendToPath(saverResponse.apiEndpointUrl, 'drive/items/' + folder.id + '/children/' + options.fileName + '/content');
                var queryParameters = {};
                queryParameters['@name.conflictBehavior'] = saverResponse.pickerType === Constants.STATE_AAD_PICKER ? 'fail' : 'rename';
                var requestHeaders = {};
                requestHeaders['Authorization'] = 'bearer ' + accessToken;
                requestHeaders['Content-Type'] = 'multipart/form-data';
                var xhr = new XHR({
                        url: UrlUtilities.appendQueryStrings(uploadUrl, queryParameters),
                        clientId: OneDriveState.clientId,
                        headers: requestHeaders,
                        apiEndpoint: saverResponse.apiEndpoint
                    });
                var data = event.target.result;
                xhr.upload(data, function (xhr, statusCode) {
                    options.success({
                        link: null,
                        values: [JSON.parse(xhr.responseText)]
                    });
                }, function (xhr, statusCode, timeout) {
                    options.error(Constants.ERROR_WEB_REQUEST);
                }, function (xhr, uploadProgress) {
                    options.progress(uploadProgress.progressPercentage);
                });
            };
            reader.readAsArrayBuffer(uploadSource);
        };
        Saver.prototype._beginPolling = function (location, accessToken) {
            Logging.logMessage('polling for URL upload completion');
            var pollingInterval = POLLING_INTERVAL;
            var pollCount = POLLING_COUNTER;
            var xhrOptions = {
                    url: UrlUtilities.appendQueryString(location, Constants.PARAM_ACCESS_TOKEN, accessToken),
                    method: Constants.HTTP_GET
                };
            var options = this._saverOptions;
            var pollForProgress = function () {
                var xhr = new XHR(xhrOptions);
                xhr.start(function (xhr, statusCode) {
                    switch (statusCode) {
                    case 202:
                        var apiResponse = ObjectUtilities.deserializeJSON(xhr.responseText);
                        options.progress(apiResponse['percentageComplete']);
                        if (!pollCount--) {
                            pollingInterval *= 2;
                            pollCount = POLLING_COUNTER;
                        }
                        CallbackInvoker.invokeCallbackAsynchronous(pollForProgress, pollingInterval);
                        break;
                    case 200:
                        options.progress(100);
                        options.success({
                            link: null,
                            values: []
                        });
                        break;
                    default:
                        options.error(Constants.ERROR_WEB_REQUEST);
                    }
                }, function (xhr, statusCode, timeout) {
                    options.error(Constants.ERROR_WEB_REQUEST);
                });
            };
            CallbackInvoker.invokeCallbackAsynchronous(pollForProgress, pollingInterval);
        };
        return Saver;
    }();
module.exports = Saver;
},{"../Constants":1,"../OneDriveState":4,"../models/ErrorType":6,"../models/SaverActionType":11,"../models/SaverOptions":12,"../models/UploadType":13,"./CallbackInvoker":14,"./ErrorHandler":16,"./GraphWrapper":17,"./Logging":18,"./ObjectUtilities":19,"./Popup":21,"./RedirectUtilities":22,"./StringUtilities":25,"./UrlUtilities":27,"./XHR":29}],25:[function(_dereq_,module,exports){
var FORMAT_ARGS_REGEX = /[\{\}]/g;
var FORMAT_REGEX = /\{\d+\}/g;
var StringUtilities = function () {
        function StringUtilities() {
        }
        StringUtilities.format = function (str) {
            var values = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                values[_i - 1] = arguments[_i];
            }
            var replacer = function (match) {
                var replacement = values[match.replace(FORMAT_ARGS_REGEX, '')];
                if (replacement === null) {
                    replacement = '';
                }
                return replacement;
            };
            return str.replace(FORMAT_REGEX, replacer);
        };
        return StringUtilities;
    }();
module.exports = StringUtilities;
},{}],26:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('./ErrorHandler'), Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities'), StringUtilities = _dereq_('./StringUtilities');
var TypeValidators = function () {
        function TypeValidators() {
        }
        TypeValidators.validateType = function (object, expectedType, optional, defaultValue, validValues) {
            if (optional === void 0) {
                optional = false;
            }
            if (object === undefined) {
                if (optional) {
                    if (defaultValue === undefined) {
                        ErrorHandler.throwError('default value missing');
                    }
                    Logging.logMessage('applying default value: ' + defaultValue);
                    return defaultValue;
                } else {
                    ErrorHandler.throwError('object was missing and not optional');
                }
            }
            var objectType = typeof object;
            if (objectType !== expectedType) {
                ErrorHandler.throwError(StringUtilities.format('expected object type: \'{0}\', actual object type: \'{1}\'', expectedType, objectType));
            }
            if (!TypeValidators._isValidValue(object, validValues)) {
                ErrorHandler.throwError(StringUtilities.format('object does not match any valid values: \'{0}\'', ObjectUtilities.serializeJSON(validValues)));
            }
            return object;
        };
        TypeValidators.validateCallback = function (functionOption, optional, expectGlobalFunction) {
            if (optional === void 0) {
                optional = false;
            }
            if (expectGlobalFunction === void 0) {
                expectGlobalFunction = false;
            }
            if (functionOption === undefined) {
                if (optional) {
                    return null;
                } else {
                    ErrorHandler.throwError('function was missing and not optional');
                }
            }
            var functionType = typeof functionOption;
            if (functionType !== Constants.TYPE_STRING && functionType !== Constants.TYPE_FUNCTION) {
                ErrorHandler.throwError(StringUtilities.format('expected function type: \'function | string\', actual type: \'{0}\'', functionType));
            }
            var returnFunction = null;
            if (functionType === Constants.TYPE_STRING) {
                var globalFunction = window[functionOption];
                if (typeof globalFunction === Constants.TYPE_FUNCTION) {
                    returnFunction = globalFunction;
                } else {
                    ErrorHandler.throwError(StringUtilities.format('could not find a function with name \'{0}\' on the window object', functionOption));
                }
            } else if (expectGlobalFunction) {
                ErrorHandler.throwError('expected a global function');
            } else {
                returnFunction = functionOption;
            }
            return returnFunction;
        };
        TypeValidators._isValidValue = function (object, validValues) {
            if (!Array.isArray(validValues)) {
                return true;
            }
            for (var index in validValues) {
                if (object === validValues[index]) {
                    return true;
                }
            }
            return false;
        };
        return TypeValidators;
    }();
module.exports = TypeValidators;
},{"../Constants":1,"./ErrorHandler":16,"./Logging":18,"./ObjectUtilities":19,"./StringUtilities":25}],27:[function(_dereq_,module,exports){
var Constants = _dereq_('../Constants'), StringUtilities = _dereq_('./StringUtilities');
var UrlUtilities = function () {
        function UrlUtilities() {
        }
        UrlUtilities.appendToPath = function (baseUrl, path) {
            return baseUrl + (baseUrl.charAt(baseUrl.length - 1) !== '/' ? '/' : '') + path;
        };
        UrlUtilities.appendQueryString = function (baseUrl, queryKey, queryValue) {
            return UrlUtilities.appendQueryStrings(baseUrl, (_a = {}, _a[queryKey] = queryValue, _a));
            var _a;
        };
        UrlUtilities.appendQueryStrings = function (baseUrl, queryParameters, isAspx) {
            if (!queryParameters || Object.keys(queryParameters).length === 0) {
                return baseUrl;
            }
            if (isAspx) {
                baseUrl += '#';
            } else if (baseUrl.indexOf('?') === -1) {
                baseUrl += '?';
            } else if (baseUrl.charAt(baseUrl.length - 1) !== '&') {
                baseUrl += '&';
            }
            var queryString = '';
            for (var key in queryParameters) {
                queryString += (queryString.length ? '&' : '') + StringUtilities.format('{0}={1}', encodeURIComponent(key), encodeURIComponent(queryParameters[key]));
            }
            return baseUrl + queryString;
        };
        UrlUtilities.readCurrentUrlParameters = function () {
            return UrlUtilities.readUrlParameters(window.location.href);
        };
        UrlUtilities.readUrlParameters = function (url) {
            var queryParamters = {};
            var queryStart = url.indexOf('?') + 1;
            var hashStart = url.indexOf('#') + 1;
            if (queryStart > 0) {
                var queryEnd = hashStart > queryStart ? hashStart - 1 : url.length;
                UrlUtilities._deserializeParameters(url.substring(queryStart, queryEnd), queryParamters);
            }
            if (hashStart > 0) {
                UrlUtilities._deserializeParameters(url.substring(hashStart), queryParamters);
            }
            return queryParamters;
        };
        UrlUtilities.trimUrlQuery = function (url) {
            var separators = [
                    '?',
                    '#'
                ];
            for (var index in separators) {
                var charIndex = url.indexOf(separators[index]);
                if (charIndex > 0) {
                    url = url.substring(0, charIndex);
                }
            }
            return url;
        };
        UrlUtilities.getFileNameFromUrl = function (url) {
            var trimmedUrl = UrlUtilities.trimUrlQuery(url);
            return trimmedUrl.substr(trimmedUrl.lastIndexOf('/') + 1);
        };
        UrlUtilities.isPathFullUrl = function (path) {
            return path.indexOf('https://') === 0 || path.indexOf('http://') === 0;
        };
        UrlUtilities.isPathDataUrl = function (path) {
            return path.indexOf('data:') === 0;
        };
        UrlUtilities.generateNonce = function () {
            var text = '';
            var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
            for (var i = 0; i < Constants.NONCE_LENGTH; i++) {
                text += possible.charAt(Math.floor(Math.random() * possible.length));
            }
            return text;
        };
        UrlUtilities._deserializeParameters = function (query, queryParameters) {
            var properties = query.split('&');
            for (var i = 0; i < properties.length; i++) {
                var property = properties[i].split('=');
                if (property.length === 2) {
                    queryParameters[decodeURIComponent(property[0])] = decodeURIComponent(property[1]);
                }
            }
        };
        return UrlUtilities;
    }();
module.exports = UrlUtilities;
},{"../Constants":1,"./StringUtilities":25}],28:[function(_dereq_,module,exports){
var Logging = _dereq_('./Logging'), ObjectUtilities = _dereq_('./ObjectUtilities');
var WindowState = function () {
        function WindowState() {
        }
        WindowState.getWindowState = function () {
            return ObjectUtilities.deserializeJSON(window.name || '{}');
        };
        WindowState.setWindowState = function (values, windowState) {
            if (windowState === void 0) {
                windowState = null;
            }
            if (windowState === null) {
                windowState = WindowState.getWindowState();
            }
            for (var property in values) {
                windowState[property] = values[property];
            }
            var serializedWindowState = ObjectUtilities.serializeJSON(windowState);
            Logging.logMessage('window.name = ' + serializedWindowState);
            window.name = serializedWindowState;
        };
        return WindowState;
    }();
module.exports = WindowState;
},{"./Logging":18,"./ObjectUtilities":19}],29:[function(_dereq_,module,exports){
var ApiEndpoint = _dereq_('../models/ApiEndpoint'), Constants = _dereq_('../Constants'), ErrorHandler = _dereq_('./ErrorHandler'), Logging = _dereq_('./Logging'), StringUtilities = _dereq_('./StringUtilities');
var REQUEST_TIMEOUT = 30000;
var EXCEPTION_STATUS = -1;
var TIMEOUT_STATUS = -2;
var ABORT_STATUS = -3;
var XHR = function () {
        function XHR(options) {
            this._url = options.url;
            this._json = options.json;
            this._headers = options.headers || {};
            this._method = options.method;
            this._clientId = options.clientId;
            this._apiEndpoint = options.apiEndpoint || ApiEndpoint.other;
            ErrorHandler.registerErrorObserver(this._abortRequest);
        }
        XHR.statusCodeToString = function (statusCode) {
            switch (statusCode) {
            case -1:
                return 'EXCEPTION';
            case -2:
                return 'TIMEOUT';
            case -3:
                return 'REQUEST ABORTED';
            default:
                return statusCode.toString();
            }
        };
        XHR.prototype.start = function (successCallback, failureCallback) {
            var _this = this;
            try {
                this._successCallback = successCallback;
                this._failureCallback = failureCallback;
                this._request = new XMLHttpRequest();
                this._request.ontimeout = this._onTimeout;
                this._request.onreadystatechange = function () {
                    if (!_this._completed && _this._request.readyState === 4) {
                        _this._completed = true;
                        var status_1 = _this._request.status;
                        if (status_1 < 400 && status_1 > 0) {
                            _this._callSuccessCallback(status_1);
                        } else {
                            _this._callFailureCallback(status_1);
                        }
                    }
                };
                if (!this._method) {
                    this._method = this._json ? Constants.HTTP_POST : Constants.HTTP_GET;
                }
                this._request.open(this._method, this._url, true);
                this._request.timeout = REQUEST_TIMEOUT;
                this._setHeaders();
                Logging.logMessage('starting request to: ' + this._url);
                this._request.send(this._json);
            } catch (error) {
                this._callFailureCallback(EXCEPTION_STATUS, error);
            }
        };
        XHR.prototype.upload = function (data, successCallback, failureCallback, progressCallback) {
            var _this = this;
            try {
                this._successCallback = successCallback;
                this._progressCallback = progressCallback;
                this._failureCallback = failureCallback;
                this._request = new XMLHttpRequest();
                this._request.ontimeout = this._onTimeout;
                this._method = Constants.HTTP_PUT;
                this._request.open(this._method, this._url, true);
                this._setHeaders();
                this._request.onload = function (event) {
                    _this._completed = true;
                    var status = _this._request.status;
                    if (status === 200 || status === 201) {
                        _this._callSuccessCallback(status);
                    } else {
                        _this._callFailureCallback(status, event);
                    }
                };
                this._request.onerror = function (event) {
                    _this._completed = true;
                    _this._callFailureCallback(_this._request.status, event);
                };
                this._request.upload.onprogress = function (event) {
                    if (event.lengthComputable) {
                        var uploadProgress = {
                                bytesTransferred: event.loaded,
                                totalBytes: event.total,
                                progressPercentage: event.total === 0 ? 0 : event.loaded / event.total * 100
                            };
                        _this._callProgressCallback(uploadProgress);
                    }
                };
                Logging.logMessage('starting upload to: ' + this._url);
                this._request.send(data);
            } catch (error) {
                this._callFailureCallback(EXCEPTION_STATUS, error);
            }
        };
        XHR.prototype._callSuccessCallback = function (status) {
            Logging.logMessage('calling xhr success callback, status: ' + XHR.statusCodeToString(status));
            this._successCallback(this._request, status, this._url);
        };
        XHR.prototype._callFailureCallback = function (status, error) {
            Logging.logError('calling xhr failure callback, status: ' + XHR.statusCodeToString(status), this._request, error);
            this._failureCallback(this._request, status, status === TIMEOUT_STATUS);
        };
        XHR.prototype._callProgressCallback = function (uploadProgress) {
            Logging.logMessage('calling xhr upload progress callback');
            this._progressCallback(this._request, uploadProgress);
        };
        XHR.prototype._abortRequest = function () {
            if (!this._completed) {
                this._completed = true;
                if (this._request) {
                    try {
                        this._request.abort();
                    } catch (error) {
                    }
                }
                this._callFailureCallback(ABORT_STATUS);
            }
        };
        XHR.prototype._onTimeout = function () {
            if (!this._completed) {
                this._completed = true;
                this._callFailureCallback(TIMEOUT_STATUS);
            }
        };
        XHR.prototype._setHeaders = function () {
            for (var x in this._headers) {
                this._request.setRequestHeader(x, this._headers[x]);
            }
            if (this._clientId && this._apiEndpoint !== ApiEndpoint.other) {
                this._request.setRequestHeader('Application', '0x' + this._clientId);
            }
            var sdkVersion = StringUtilities.format('{0}={1}', 'SDK-Version', Constants.SDK_VERSION);
            switch (this._apiEndpoint) {
            case ApiEndpoint.graph_odb:
                this._request.setRequestHeader('X-ClientService-ClientTag', sdkVersion);
                break;
            case ApiEndpoint.graph_odc:
                this._request.setRequestHeader('X-RequestStats', sdkVersion);
                break;
            case ApiEndpoint.other:
                break;
            default:
                ErrorHandler.throwError('invalid API endpoint: ' + this._apiEndpoint);
            }
            if (this._method === Constants.HTTP_POST) {
                this._request.setRequestHeader('Content-Type', this._json ? 'application/json' : 'text/plain');
            }
        };
        return XHR;
    }();
module.exports = XHR;
},{"../Constants":1,"../models/ApiEndpoint":5,"./ErrorHandler":16,"./Logging":18,"./StringUtilities":25}]},{},[2])
(2)
});