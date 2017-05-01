//! Copyright (c) Microsoft Corporation. All rights reserved.
var __extends = (this && this.__extends) || function (d, b) {for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];function __() { this.constructor = d; }__.prototype = b.prototype;d.prototype = new __();};
(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.OneDrive = f()}})(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);var f=new Error("Cannot find module '"+o+"'");throw f.code="MODULE_NOT_FOUND",f}var l=n[o]={exports:{}};t[o][0].call(l.exports,function(e){var n=t[o][1][e];return s(n?n:e)},l,l.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var Constants = function () {
        function Constants() {
        }
        Constants.SDK_VERSION_NUMBER = '7.2';
        Constants.SDK_VERSION = 'js-v' + Constants.SDK_VERSION_NUMBER;
        Constants.TYPE_BOOLEAN = 'boolean';
        Constants.TYPE_FUNCTION = 'function';
        Constants.TYPE_OBJECT = 'object';
        Constants.TYPE_STRING = 'string';
        Constants.TYPE_NUMBER = 'number';
        Constants.VROOM_URL = 'https://api.onedrive.com/v1.0/';
        Constants.VROOM_ENDPOINT_HINT = 'api.onedrive.com';
        Constants.GRAPH_URL = 'https://graph.microsoft.com/v1.0/';
        Constants.NONCE_LENGTH = 5;
        Constants.CUSTOMER_TID = '9188040d-6c67-4c5b-b112-36a304b66dad';
        Constants.DEFAULT_QUERY_ITEM_PARAMETER = 'select=id';
        return Constants;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Constants;
}(require, exports));
},{}],2:[function(require,module,exports){
module.exports = function (require, exports, OneDriveApp_1, Oauth_1) {
    'use strict';
    var OneDrive = function () {
        function OneDrive() {
        }
        OneDrive.open = function (options) {
            OneDriveApp_1.default.open(options);
        };
        OneDrive.save = function (options) {
            OneDriveApp_1.default.save(options);
        };
        return OneDrive;
    }();
    Oauth_1.onAuth();
    return OneDrive;
}(require, exports, require('./OneDriveApp'), require('./controllers/Oauth'));
},{"./OneDriveApp":3,"./controllers/Oauth":7}],3:[function(require,module,exports){
(function (require, exports, ErrorHandler_1, ErrorType_1, Logging_1, OneDriveSdkError_1, Picker_1, Saver_1) {
    'use strict';
    var OneDriveApp = function () {
        function OneDriveApp() {
        }
        OneDriveApp.open = function (options) {
            if (!OneDriveApp.isReady()) {
                return;
            }
            if (!options) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unknown, 'missing picker options')).exposeToPublic();
            }
            Logging_1.default.logMessage('open started');
            var picker = new Picker_1.default(options);
            picker.launchPicker().then(function () {
                OneDriveApp.reset();
            });
        };
        OneDriveApp.save = function (options) {
            if (!OneDriveApp.isReady()) {
                return;
            }
            if (!options) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unknown, 'missing saver options'));
            }
            Logging_1.default.logMessage('save started');
            var saver = new Saver_1.default(options);
            saver.launchSaver().then(function () {
                OneDriveApp.reset();
            });
        };
        OneDriveApp.reset = function () {
            OneDriveApp.checked = false;
        };
        OneDriveApp.isReady = function () {
            if (OneDriveApp.checked) {
                return false;
            }
            OneDriveApp.checked = true;
            return true;
        };
        OneDriveApp.checked = false;
        return OneDriveApp;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = OneDriveApp;
}(require, exports, require('./utilities/ErrorHandler'), require('./models/ErrorType'), require('./utilities/Logging'), require('./models/OneDriveSdkError'), require('./controllers/Picker'), require('./controllers/Saver')));
},{"./controllers/Picker":8,"./controllers/Saver":10,"./models/ErrorType":13,"./models/OneDriveSdkError":16,"./utilities/ErrorHandler":26,"./utilities/Logging":27}],4:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, ErrorType_1, Logging_1, ObjectUtilities_1, OneDriveSdkError_1, StringUtilities_1, UrlUtilities_1, XHR_1, es6_promise_1) {
    'use strict';
    var POLLING_INTERVAL = 100;
    var POLLING_COUNTER = 5;
    var MAXIMUM_POLLING_INTERVAL = 30 * 60 * 1000;
    var ROOT_ID = 'root';
    function getItem(item, apiRequestConfig, queryParameters) {
        var getRequestUrl = buildPathToItem(item, apiRequestConfig.apiEndpointUrl);
        if (queryParameters) {
            getRequestUrl = UrlUtilities_1.appendToPath(getRequestUrl, '?' + queryParameters);
        }
        var xhr = new XHR_1.default({
            url: getRequestUrl,
            clientId: apiRequestConfig.clientId,
            method: XHR_1.default.HTTP_GET,
            apiEndpoint: apiRequestConfig.apiEndpoint,
            headers: { 'Authorization': 'bearer ' + apiRequestConfig.accessToken }
        });
        Logging_1.default.logMessage('performing GET on item with id: ' + item.id);
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode) {
                var itemInDetail = JSON.parse(xhr.responseText);
                resolve(itemInDetail);
            }, function (xhr, statusCode, timeout) {
                reject({
                    errorCode: ErrorType_1.default[ErrorType_1.default.webRequestFailure],
                    message: 'HTTP error status: ' + statusCode
                });
            });
        });
    }
    exports.getItem = getItem;
    function getItems(items, apiRequestConfig, queryParameters) {
        var filesPromises = [];
        var processedFiles = {
            webUrl: null,
            value: []
        };
        for (var _i = 0, _a = items.value; _i < _a.length; _i++) {
            var item = _a[_i];
            filesPromises.push(getItem(item, apiRequestConfig, queryParameters));
        }
        return es6_promise_1.Promise.all(filesPromises).then(function (successFiles) {
            processedFiles.value = successFiles;
            return processedFiles;
        }, function (error) {
            Logging_1.default.logError('Received ajax error.', error);
            return error;
        });
    }
    exports.getItems = getItems;
    function shareItem(item, apiRequestConfig, createLinkParameters) {
        var shareRequestUrl = UrlUtilities_1.appendToPath(buildPathToItem(item, apiRequestConfig.apiEndpointUrl), StringUtilities_1.format('{0}.createLink', apiRequestConfig.apiActionNamingSpace));
        var xhr = new XHR_1.default({
            url: shareRequestUrl,
            clientId: apiRequestConfig.clientId,
            method: XHR_1.default.HTTP_POST,
            apiEndpoint: apiRequestConfig.apiEndpoint,
            headers: { 'Authorization': 'bearer ' + apiRequestConfig.accessToken },
            json: JSON.stringify(createLinkParameters)
        });
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode, url) {
                Logging_1.default.logMessage(StringUtilities_1.format('POST createLink succeeded via path {0}', shareRequestUrl));
                item.permissions = [JSON.parse(xhr.responseText)];
                resolve(item);
            }, function (xhr, statusCode, timeout) {
                Logging_1.default.logMessage(StringUtilities_1.format('POST createLink failed via path {0}', shareRequestUrl));
                reject({
                    errorCode: ErrorType_1.default[ErrorType_1.default.webRequestFailure],
                    message: statusCode
                });
            });
        });
    }
    exports.shareItem = shareItem;
    function shareItems(items, apiRequestConfig, createLinkParameters) {
        var sharePromises = [];
        var processedFiles = {
            webUrl: null,
            value: []
        };
        for (var _i = 0, _a = items.value; _i < _a.length; _i++) {
            var item = _a[_i];
            sharePromises.push(shareItem(item, apiRequestConfig, createLinkParameters));
        }
        return es6_promise_1.Promise.all(sharePromises).then(function (successFiles) {
            processedFiles.value = successFiles;
            return processedFiles;
        }, function (error) {
            Logging_1.default.logError('Received sharing error.', error);
            return error;
        });
    }
    exports.shareItems = shareItems;
    function saveItemByFormUpload(folder, itemDescription, file, apiRequestConfig, progressCallback) {
        var reader = null;
        return new es6_promise_1.Promise(function (resolve, reject) {
            if (window['File'] && file instanceof window['File']) {
                reader = new FileReader();
            } else {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'FileReader is not supported in this browser'));
            }
            reader.onerror = function (event) {
                Logging_1.default.logError('failed to read or upload the file', event);
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.fileReaderFailure, 'failed to read or upload the file, see console log for details'));
            };
            reader.onload = function (event) {
                var uploadUrl = UrlUtilities_1.appendToPath(buildPathToItem(folder, apiRequestConfig.apiEndpointUrl), 'children(\'' + itemDescription.name + '\')/content');
                var queryParameters = {};
                queryParameters['@name.conflictBehavior'] = itemDescription['@name.conflictBehavior'];
                var requestHeaders = {};
                requestHeaders['Authorization'] = 'bearer ' + apiRequestConfig.accessToken;
                requestHeaders['Content-Type'] = 'multipart/form-data';
                var xhr = new XHR_1.default({
                    url: UrlUtilities_1.appendQueryStrings(uploadUrl, queryParameters),
                    clientId: apiRequestConfig.clientId,
                    headers: requestHeaders,
                    apiEndpoint: apiRequestConfig.apiEndpoint
                });
                var data = event.target.result;
                xhr.upload(data, function (xhr, statusCode) {
                    resolve({
                        webUrl: null,
                        value: [JSON.parse(xhr.responseText)]
                    });
                }, function (xhr, statusCode, timeout) {
                    reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('file uploading failed by form uplaod with HTTP status: {0}', statusCode)));
                }, function (xhr, uploadProgress) {
                    progressCallback(uploadProgress.progressPercentage);
                });
            };
            reader.readAsArrayBuffer(file);
        });
    }
    exports.saveItemByFormUpload = saveItemByFormUpload;
    function saveItemByUriUpload(folder, itemDescription, sourceUri, apiRequestConfig) {
        if (apiRequestConfig.apiEndpoint === ApiEndpoint_1.default.filesV2 || apiRequestConfig.apiEndpoint === ApiEndpoint_1.default.graph_odb) {
            return new es6_promise_1.Promise(function (resolve, reject) {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'URL upload not supported for OneDrive business'));
            });
        }
        var uploadUrl = UrlUtilities_1.appendToPath(buildPathToItem(folder, apiRequestConfig.apiEndpointUrl), 'children');
        var requestHeaders = {};
        requestHeaders['Prefer'] = 'respond-async';
        requestHeaders['Authorization'] = 'bearer ' + apiRequestConfig.accessToken;
        itemDescription[getContentSourceUrl(apiRequestConfig.apiEndpoint)] = sourceUri;
        itemDescription['file'] = {};
        var xhr = new XHR_1.default({
            url: uploadUrl,
            clientId: apiRequestConfig.clientId,
            method: XHR_1.default.HTTP_POST,
            headers: requestHeaders,
            json: JSON.stringify(itemDescription),
            apiEndpoint: apiRequestConfig.apiEndpoint
        });
        if (UrlUtilities_1.isPathDataUrl(sourceUri)) {
            return saveItemByDateUriUpload(xhr);
        } else if (UrlUtilities_1.isPathFullUrl(sourceUri)) {
            return saveItemByHttpUrlUpload(xhr).then(function (location) {
                return beginPolling(location).then(function (resourceId) {
                    var file = { id: resourceId };
                    return getItem(file, apiRequestConfig).then(function (file) {
                        var response = {
                            webUrl: null,
                            value: [file]
                        };
                        return es6_promise_1.Promise.resolve(response);
                    });
                });
            });
        }
    }
    exports.saveItemByUriUpload = saveItemByUriUpload;
    function getUserTenantUrl(apiRequestConfig) {
        var queryUrl = UrlUtilities_1.appendQueryString(apiRequestConfig.apiEndpointUrl, '$select', 'mySite');
        var requestHeaders = {};
        requestHeaders['Authorization'] = 'bearer ' + apiRequestConfig.accessToken;
        requestHeaders['Accept'] = 'application/json';
        var xhr = new XHR_1.default({
            url: queryUrl,
            clientId: apiRequestConfig.clientId,
            method: XHR_1.default.HTTP_GET,
            headers: requestHeaders,
            apiEndpoint: apiRequestConfig.apiEndpoint
        });
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode) {
                var response = ObjectUtilities_1.deserializeJSON(xhr.responseText);
                if (response.mySite) {
                    resolve(response.mySite);
                } else {
                    reject(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, StringUtilities_1.format('Cannot find the personal tenant url, response text: {0}', xhr.responseText)));
                }
            }, function (xhr, statusCode, timeout) {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('graph/me request failed, status code: \'{0}\', response text: \'{1}\'', XHR_1.default.statusCodeToString(statusCode), xhr.responseText)));
            });
        });
    }
    exports.getUserTenantUrl = getUserTenantUrl;
    function saveItemByDateUriUpload(xhr) {
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode) {
                if (statusCode === 200 || statusCode === 201) {
                    var uploadedFile = {
                        webUrl: null,
                        value: [ObjectUtilities_1.deserializeJSON(xhr.responseText)]
                    };
                    resolve(uploadedFile);
                } else {
                    reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('file uploading failed by data uri with HTTP status: {0}', statusCode)));
                }
            }, function (xhr, statusCode, timeout) {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('file uploading failed with HTTP status: {0}', statusCode)));
            });
        });
    }
    function saveItemByHttpUrlUpload(xhr) {
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode) {
                if (statusCode === 202) {
                    var location_1 = xhr.getResponseHeader('Location');
                    if (!location_1) {
                        reject({
                            errorCode: ErrorType_1.default.badResponse,
                            message: 'missing \'Location\' header on response'
                        });
                    }
                    resolve(location_1);
                } else {
                    reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('create upload by url job failed with HTTP status: {0}', statusCode)));
                }
            }, function (xhr, statusCode, timeout) {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('create upload by url job failed with HTTP status: {0}', statusCode)));
            });
        });
    }
    function getContentSourceUrl(apiEndpoint) {
        if (apiEndpoint === ApiEndpoint_1.default.graph_odb || apiEndpoint === ApiEndpoint_1.default.graph_odc) {
            return '@microsoft.graph.sourceUrl';
        } else {
            return '@content.sourceUrl';
        }
    }
    function beginPolling(location) {
        return new es6_promise_1.Promise(function (resolve, reject) {
            (function ping(retry, pollingInterval) {
                if (!retry--) {
                    pollingInterval *= 2;
                    retry = POLLING_COUNTER;
                }
                pollForProgress(location).then(function (monitorResponse) {
                    if (monitorResponse.resourceId) {
                        resolve(monitorResponse.resourceId);
                    } else if (pollingInterval <= MAXIMUM_POLLING_INTERVAL) {
                        setTimeout(ping(retry, pollingInterval), pollingInterval);
                    } else {
                        reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, 'polling the uploading job takes too much time'));
                    }
                });
            }(POLLING_COUNTER, POLLING_INTERVAL));
        });
    }
    function pollForProgress(location) {
        var xhr = new XHR_1.default({
            url: location,
            method: XHR_1.default.HTTP_GET
        });
        return new es6_promise_1.Promise(function (resolve, reject) {
            xhr.start(function (xhr, statusCode) {
                switch (statusCode) {
                case 202:
                case 200:
                    var successResponse = ObjectUtilities_1.deserializeJSON(xhr.responseText);
                    resolve(successResponse);
                    break;
                default:
                    reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('polling upload job failed with HTTP status: {0}', statusCode)));
                }
            }, function (xhr, statusCode, timeout) {
                reject(new OneDriveSdkError_1.default(ErrorType_1.default.webRequestFailure, StringUtilities_1.format('polling upload job failed with HTTP status: {0}', statusCode)));
            });
        });
    }
    function buildPathToItem(item, apiEndpointUrl) {
        var subPath;
        if (item.parentReference && item.parentReference.driveId) {
            subPath = UrlUtilities_1.appendToPath('drives', item.parentReference.driveId);
        } else {
            subPath = 'drive';
        }
        subPath = UrlUtilities_1.appendToPath(subPath, item.id === ROOT_ID ? 'root' : 'items/' + item.id);
        return UrlUtilities_1.appendToPath(apiEndpointUrl, subPath);
    }
}(require, exports, require('../models/ApiEndpoint'), require('../models/ErrorType'), require('../utilities/Logging'), require('../utilities/ObjectUtilities'), require('../models/OneDriveSdkError'), require('../utilities/StringUtilities'), require('../utilities/UrlUtilities'), require('../utilities/XHR'), require('es6-promise')));
},{"../models/ApiEndpoint":11,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"../utilities/Logging":27,"../utilities/ObjectUtilities":28,"../utilities/StringUtilities":30,"../utilities/UrlUtilities":32,"../utilities/XHR":33,"es6-promise":34}],5:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, ApiRequest_1, Constants_1, DomainHint_1, ErrorHandler_1, ErrorType_1, Logging_1, LoginCache_1, Oauth_1, OneDriveSdkError_1, PickerUX_1, Popup_1, es6_promise_1, StringUtilities_1, UrlUtilities_1) {
    'use strict';
    var Invoker = function () {
        function Invoker(invokerOptions) {
            this.invokerOptions = invokerOptions;
            this.popup = new Popup_1.default();
        }
        Invoker.prototype.launchInvoker = function () {
            var _this = this;
            var invokerOptions = this.invokerOptions;
            return this.launch().catch(function (error) {
                Logging_1.default.logError('Failed due to unknown error: ', error);
                invokerOptions.error(error);
            }).then(function () {
                _this.cleanPopupAndIFrame();
            });
        };
        Invoker.prototype.launch = function (switchAccount) {
            var _this = this;
            return this.buildOauthPromise(switchAccount).then(function (oauthResponse) {
                if (oauthResponse && oauthResponse.type === 'cancel') {
                    return oauthResponse;
                } else {
                    return _this.buildPickerUI(oauthResponse);
                }
            }).then(function (response) {
                var invokerOptions = _this.invokerOptions;
                var type = response.type;
                if (!type) {
                    ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, StringUtilities_1.format('received bad response from picker UX: {0}', JSON.stringify(response)))).exposeToPublic();
                }
                if (response.type === 'switch') {
                    return _this.launch(true);
                } else if (response.type === 'success') {
                    var items = [];
                    var isMicrosoftApp = false;
                    for (var _i = 0, _a = response.items; _i < _a.length; _i++) {
                        var item = _a[_i];
                        if (item.driveItem && !isMicrosoftApp) {
                            isMicrosoftApp = true;
                        }
                        items.push(item);
                    }
                    var successResponse = {
                        webUrl: null,
                        value: items
                    };
                    var successPromise = void 0;
                    if (!invokerOptions.needAPICall() || isMicrosoftApp && invokerOptions.accessToken.toLowerCase() === 'rps') {
                        successPromise = es6_promise_1.Promise.resolve(successResponse);
                    } else {
                        _this.apiRequestConfig = _this.buildApiConfig();
                        successPromise = _this.makeApiRequest(successResponse);
                    }
                    return successPromise.then(function (files) {
                        if (_this.oauthResponse) {
                            files.accessToken = _this.oauthResponse.accessToken;
                        }
                        if (_this.apiRequestConfig) {
                            files.apiEndpoint = _this.apiRequestConfig.apiEndpointUrl;
                        } else if (_this.loginHint && _this.loginHint.endpointHint === DomainHint_1.default.aad) {
                            files.apiEndpoint = UrlUtilities_1.appendToPath(Constants_1.default.GRAPH_URL, 'me');
                        }
                        invokerOptions.success(files);
                        return files;
                    });
                } else if (response.type === 'cancel') {
                    invokerOptions.cancel();
                    return es6_promise_1.Promise.resolve({
                        webUrl: null,
                        value: null
                    });
                }
            });
        };
        Invoker.prototype.buildOauthPromise = function (switchAccount) {
            var oauthPromise;
            if (switchAccount || this.invokerOptions.needOauth()) {
                oauthPromise = Oauth_1.auth(Oauth_1.buildOauthConfig(this.invokerOptions, switchAccount), this.popup);
            } else {
                oauthPromise = es6_promise_1.Promise.resolve(null);
            }
            return oauthPromise;
        };
        Invoker.prototype.buildPickerUI = function (oauthResponse) {
            var _this = this;
            var loginHint;
            if (oauthResponse) {
                this.oauthResponse = oauthResponse;
                loginHint = LoginCache_1.updateLoginHint(this.invokerOptions.clientId, this.oauthResponse.idToken, this.invokerOptions);
            } else {
                if (this.invokerOptions.endpointHint !== DomainHint_1.default.msa && this.invokerOptions.endpointHint !== DomainHint_1.default.tenant) {
                    ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'must specify the endpointHint in advanced options as \'api.onedrive.com\' for customer picker or the url for business picker/tenant')).exposeToPublic();
                }
                loginHint = {
                    loginHint: null,
                    domainHint: null,
                    timeStamp: null,
                    apiEndpoint: this.invokerOptions.endpointHint === DomainHint_1.default.msa ? ApiEndpoint_1.default.msa : ApiEndpoint_1.default.filesV2,
                    endpointHint: this.invokerOptions.endpointHint === DomainHint_1.default.msa ? DomainHint_1.default.msa : DomainHint_1.default.tenant
                };
            }
            this.loginHint = loginHint;
            var tenantPromise;
            if (loginHint.apiEndpoint === ApiEndpoint_1.default.graph_odb) {
                tenantPromise = ApiRequest_1.getUserTenantUrl(this.buildApiConfig());
            } else {
                tenantPromise = es6_promise_1.Promise.resolve(undefined);
            }
            return tenantPromise.then(function (tenantUrl) {
                _this.pickerUX = PickerUX_1.generatePickerUX(loginHint.apiEndpoint, loginHint.endpointHint === DomainHint_1.default.tenant ? _this.invokerOptions.siteUrl : tenantUrl);
                var pickerUXConfig = _this.buildPickerUXConfig(_this.invokerOptions);
                if (_this.invokerOptions.navEntryLocation) {
                    pickerUXConfig.entryLocation = _this.invokerOptions.navEntryLocation;
                }
                if (_this.invokerOptions.navSourceTypes) {
                    pickerUXConfig.sourceTypes = _this.invokerOptions.navSourceTypes;
                }
                if (_this.invokerOptions.linkType) {
                    pickerUXConfig.linkType = _this.invokerOptions.linkType;
                }
                return _this.pickerUX.invokePickerUX(pickerUXConfig, _this.popup);
            });
        };
        Invoker.prototype.getApiRequestConfig = function () {
            return this.apiRequestConfig;
        };
        Invoker.prototype.buildApiConfig = function () {
            if (!this.loginHint) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.internalError, 'missing loginHint when trying to build API request')).exposeToPublic();
            }
            if (!this.oauthResponse && !this.invokerOptions.accessToken) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.internalError, 'missing access token when trying to build API request')).exposeToPublic();
            }
            var apiEndpointUrl = null;
            var apiActionNamingSpace = null;
            switch (this.loginHint.apiEndpoint) {
            case ApiEndpoint_1.default.graph_odb:
            case ApiEndpoint_1.default.graph_odc:
                apiEndpointUrl = UrlUtilities_1.appendToPath(Constants_1.default.GRAPH_URL, 'me');
                apiActionNamingSpace = 'microsoft.graph';
                break;
            case ApiEndpoint_1.default.msa:
                apiEndpointUrl = Constants_1.default.VROOM_URL;
                apiActionNamingSpace = 'action';
                break;
            case ApiEndpoint_1.default.filesV2:
                apiEndpointUrl = UrlUtilities_1.appendToPath(this.invokerOptions.siteUrl, '_api/v2.0/');
                apiActionNamingSpace = 'action';
                break;
            default:
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.internalError, 'apiEndpoint in loginHint is not correct'));
            }
            var apiRequestConfig = {
                accessToken: this.oauthResponse ? this.oauthResponse.accessToken : this.invokerOptions.accessToken,
                apiEndpoint: this.loginHint.apiEndpoint,
                apiEndpointUrl: apiEndpointUrl,
                apiActionNamingSpace: apiActionNamingSpace,
                clientId: this.invokerOptions.clientId
            };
            return apiRequestConfig;
        };
        Invoker.prototype.cleanPopupAndIFrame = function () {
            if (this.popup) {
                this.popup.close();
            }
            if (this.pickerUX) {
                this.pickerUX.removeIFrame();
            }
        };
        return Invoker;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Invoker;
}(require, exports, require('../models/ApiEndpoint'), require('./ApiRequest'), require('../Constants'), require('../models/DomainHint'), require('../utilities/ErrorHandler'), require('../models/ErrorType'), require('../utilities/Logging'), require('./LoginCache'), require('./Oauth'), require('../models/OneDriveSdkError'), require('./PickerUX'), require('../utilities/Popup'), require('es6-promise'), require('../utilities/StringUtilities'), require('../utilities/UrlUtilities')));
},{"../Constants":1,"../models/ApiEndpoint":11,"../models/DomainHint":12,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"../utilities/ErrorHandler":26,"../utilities/Logging":27,"../utilities/Popup":29,"../utilities/StringUtilities":30,"../utilities/UrlUtilities":32,"./ApiRequest":4,"./LoginCache":6,"./Oauth":7,"./PickerUX":9,"es6-promise":34}],6:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, Cache_1, Constants_1, DomainHint_1, ErrorHandler_1, ErrorType_1, OneDriveSdkError_1, ObjectUtilities_1) {
    'use strict';
    var ACCESS_TOKEN_LIFESPAN = 3600000;
    var ID_TOKEN_TID = 'tid';
    var ID_TOKEN_PREFERRED_USERNAME = 'preferred_username';
    var LOGINHINT_KEY = 'odsdkLoginHint';
    var LOGINHINT_ITEM_PREFIX = 'od7-';
    function getLoginHint(invokerOptions) {
        if (!invokerOptions) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'invoker option is not defined!'));
        }
        var loginCache = Cache_1.getCacheItem(LOGINHINT_KEY) || {};
        var loginHint = loginCache[LOGINHINT_ITEM_PREFIX + invokerOptions.clientId];
        if (!loginHint) {
            return null;
        }
        if (invokerOptions.endpointHint !== loginHint.endpointHint) {
            return null;
        }
        if (invokerOptions.loginHint && invokerOptions.loginHint !== loginHint.loginHint) {
            return null;
        }
        return loginHint;
    }
    exports.getLoginHint = getLoginHint;
    function loginHintExpired() {
        var loginHint = Cache_1.getCacheItem(LOGINHINT_KEY);
        return new Date().getTime() - loginHint.timeStamp > ACCESS_TOKEN_LIFESPAN;
    }
    exports.loginHintExpired = loginHintExpired;
    function updateLoginHint(clientId, idToken, invokerOptions) {
        var loginHint;
        var domainHint;
        var endpointHint;
        var apiEndpoint;
        switch (invokerOptions.endpointHint) {
        case DomainHint_1.default.aad:
            var idTokenObj = this.parseIdToken(idToken);
            loginHint = idTokenObj.preferredUserName;
            if (idTokenObj.tid === Constants_1.default.CUSTOMER_TID) {
                apiEndpoint = ApiEndpoint_1.default.graph_odc;
                domainHint = "consumers";
            } else {
                apiEndpoint = ApiEndpoint_1.default.graph_odb;
                domainHint = "organizations";
            }
            endpointHint = DomainHint_1.default.aad;
            break;
        case DomainHint_1.default.msa:
            apiEndpoint = ApiEndpoint_1.default.msa;
            endpointHint = DomainHint_1.default.msa;
            loginHint = invokerOptions.loginHint;
            domainHint = "consumers";
            break;
        case DomainHint_1.default.tenant:
            apiEndpoint = ApiEndpoint_1.default.filesV2;
            endpointHint = DomainHint_1.default.tenant;
            loginHint = invokerOptions.loginHint;
            domainHint = "organizations";
            break;
        }
        var newLoginHint = {
            apiEndpoint: apiEndpoint,
            loginHint: loginHint,
            domainHint: domainHint,
            endpointHint: endpointHint,
            timeStamp: new Date().getTime()
        };
        var loginCache = Cache_1.getCacheItem(LOGINHINT_KEY) || {};
        loginCache[LOGINHINT_ITEM_PREFIX + clientId] = newLoginHint;
        Cache_1.setCacheItem(LOGINHINT_KEY, loginCache);
        return newLoginHint;
    }
    exports.updateLoginHint = updateLoginHint;
    function parseIdToken(idToken) {
        if (!idToken) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'id_token is missing in oauth response'));
        }
        var userInfoPart = idToken.split('.')[1];
        var urlFriendlyValue = userInfoPart.replace('-', '+').replace('_', '/');
        var userInfo = ObjectUtilities_1.deserializeJSON(atob(urlFriendlyValue));
        if (!userInfo[ID_TOKEN_TID]) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'tid is missing in id_token response'));
        }
        if (!userInfo[ID_TOKEN_PREFERRED_USERNAME]) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'preferred_username is missing in id_token response'));
        }
        return {
            tid: userInfo[ID_TOKEN_TID],
            preferredUserName: userInfo[ID_TOKEN_PREFERRED_USERNAME]
        };
    }
    exports.parseIdToken = parseIdToken;
}(require, exports, require('../models/ApiEndpoint'), require('../utilities/Cache'), require('../Constants'), require('../models/DomainHint'), require('../utilities/ErrorHandler'), require('../models/ErrorType'), require('../models/OneDriveSdkError'), require('../utilities/ObjectUtilities')));
},{"../Constants":1,"../models/ApiEndpoint":11,"../models/DomainHint":12,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"../utilities/Cache":22,"../utilities/ErrorHandler":26,"../utilities/ObjectUtilities":28}],7:[function(require,module,exports){
(function (require, exports, Channel_1, DomainHint_1, DomUtilities_1, ErrorHandler_1, ErrorType_1, LoginCache_1, OauthEndpoint_1, OneDriveSdkError_1, es6_promise_1, UrlUtilities_1) {
    'use strict';
    var PARAM_ACCESS_TOKEN = 'access_token';
    var PARAM_ERROR = 'error';
    var PARAM_ERROR_DESCRIPTION = 'error_description';
    var PARAM_ID_TOKEN = 'id_token';
    var PARAM_OAUTH_CONFIG = 'oauth';
    var PARAM_STATE = 'state';
    var AAD_OAUTH_ENDPOINT = 'https://login.microsoftonline.com/common/oauth2/authorize';
    var AADV2_OAUTH_ENDPOINT = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
    var MSA_OAUTH_ENDPONT = 'https://login.live.com/oauth20_authorize.srf';
    var OAUTH_RESPONSE_HEADER = '[OneDriveSDK-OauthResponse]';
    function onAuth() {
        DomUtilities_1.onDocumentReady(function () {
            var redirResults = UrlUtilities_1.readCurrentUrlParameters();
            var isOauth = redirResults[PARAM_OAUTH_CONFIG] || redirResults[PARAM_ERROR] || redirResults[PARAM_ACCESS_TOKEN];
            if (isOauth && window.opener) {
                handleOauth(redirResults, new Channel_1.default(window.opener));
            }
        });
    }
    exports.onAuth = onAuth;
    function handleOauth(redirResults, channel) {
        DomUtilities_1.displayOverlay();
        if (redirResults[PARAM_OAUTH_CONFIG]) {
            redirectToOauthPage(JSON.parse(redirResults[PARAM_OAUTH_CONFIG]));
        } else if (redirResults[PARAM_ERROR]) {
            sendResponseToParent(generateErrorResponse(redirResults), channel);
        } else if (redirResults[PARAM_ACCESS_TOKEN]) {
            sendResponseToParent(generateSuccessResponse(redirResults), channel);
        }
    }
    exports.handleOauth = handleOauth;
    function generateErrorResponse(redirResults) {
        var error = new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, redirResults[PARAM_ERROR_DESCRIPTION]);
        return {
            type: 'error',
            error: error,
            state: redirResults[PARAM_STATE]
        };
    }
    function generateSuccessResponse(redirResults) {
        return {
            type: 'success',
            accessToken: redirResults[PARAM_ACCESS_TOKEN],
            idToken: redirResults[PARAM_ID_TOKEN],
            state: redirResults[PARAM_STATE]
        };
    }
    function sendResponseToParent(response, channel) {
        if (response.state) {
            var responseFragment = response.state.split('_');
            if (responseFragment.length !== 2) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'received bad state parameter from Oauth endpoint, state received: ' + response.state)).exposeToPublic();
            }
            var origin = responseFragment[0];
            if (channel) {
                channel.send(OAUTH_RESPONSE_HEADER + JSON.stringify(response), origin);
            } else {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.popupOpen, 'opener is not defined')).exposeToPublic();
            }
        } else {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'missing state parameter from Oauth redirect')).exposeToPublic();
        }
    }
    function redirectToOauthPage(oauthConfig) {
        var url;
        switch (oauthConfig.endpoint) {
        case OauthEndpoint_1.default.AAD:
            url = buildAADOauthUrl(oauthConfig);
            break;
        case OauthEndpoint_1.default.AADv2:
            url = buildAADOauthV2Url(oauthConfig);
            break;
        case OauthEndpoint_1.default.MSA:
            url = buildMSAOauthUrl(oauthConfig);
            break;
        default:
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'received bad oauth endpoint, endpoint value is: ' + oauthConfig.endpoint));
            break;
        }
        if (oauthConfig.switchAccount) {
            url = UrlUtilities_1.appendQueryString(url, 'prompt', 'select_account');
        } else if (oauthConfig.loginHint) {
            url = UrlUtilities_1.appendQueryString(url, 'login_hint', oauthConfig.loginHint);
            if (oauthConfig.domainHint) {
               url = UrlUtilities_1.appendQueryString(url, 'domain_hint', oauthConfig.domainHint);
            }
        }
        
        UrlUtilities_1.redirect(url);
    }
    function buildAADOauthUrl(config) {
        return UrlUtilities_1.appendQueryStrings(AAD_OAUTH_ENDPOINT, {
            redirect_uri: config.redirectUri,
            client_id: config.clientId,
            response_type: 'token',
            state: config.state,
            resource: config.origin
        });
    }
    function buildAADOauthV2Url(config) {
        var scope = 'profile openid https://graph.microsoft.com/User.Read ' + config.scopes.map(function (s) {
            return 'https://graph.microsoft.com/' + s;
        }).join(' ');
        var url = UrlUtilities_1.appendQueryStrings(AADV2_OAUTH_ENDPOINT, {
            redirect_uri: config.redirectUri,
            client_id: config.clientId,
            scope: scope,
            response_mode: 'fragment',
            state: config.state,
            nonce: UrlUtilities_1.generateNonce()
        });
        url += '&response_type=id_token+token';
        return url;
    }
    function buildMSAOauthUrl(config) {
        var needWritePermission = false;
        for (var _i = 0, _a = config.scopes; _i < _a.length; _i++) {
            var scope = _a[_i];
            needWritePermission = needWritePermission || scope.toLowerCase().indexOf('readwrite') > 1;
        }
        return UrlUtilities_1.appendQueryStrings(MSA_OAUTH_ENDPONT, {
            redirect_uri: config.redirectUri,
            client_id: config.clientId,
            response_type: 'token',
            state: config.state,
            scope: 'onedrive.' + (needWritePermission ? 'readwrite' : 'readonly')
        });
    }
    function auth(config, popupView) {
        var state = document.location.origin + '_' + UrlUtilities_1.generateNonce();
        config.state = state;
        return new es6_promise_1.Promise(function (resolve, reject) {
            var listenerId = DomUtilities_1.onMessage(function (event) {
                if (event.data && event.data.indexOf(OAUTH_RESPONSE_HEADER) === 0) {
                    var responseData = JSON.parse(event.data.substring(OAUTH_RESPONSE_HEADER.length));
                    if (responseData.state === state && event.source === popupView.getPopupWindow()) {
                        DomUtilities_1.removeMessageListener(listenerId);
                        if (responseData.type === 'error' || responseData.error) {
                            var errorCode = ErrorType_1.default[responseData.error.errorCode];
                            reject(new OneDriveSdkError_1.default(errorCode, responseData.error.message));
                        } else {
                            resolve(responseData);
                        }
                    } else {
                        reject(new OneDriveSdkError_1.default(ErrorType_1.default.popupOpen, 'Another popup is already opened.'));
                    }
                }
            });
            return popupView.openPopup(config.redirectUri + '?' + PARAM_OAUTH_CONFIG + '=' + JSON.stringify(config)).then(function () {
                resolve({
                    type: 'cancel',
                    state: state
                });
            });
        });
    }
    exports.auth = auth;
    function buildOauthConfig(invokerOptions, switchAccount) {
        var endpoint;
        switch (invokerOptions.endpointHint) {
        case DomainHint_1.default.aad:
            endpoint = OauthEndpoint_1.default.AADv2;
            break;
        case DomainHint_1.default.msa:
            endpoint = OauthEndpoint_1.default.MSA;
            break;
        case DomainHint_1.default.tenant:
            endpoint = OauthEndpoint_1.default.AAD;
            break;
        }
        var loginHint = LoginCache_1.getLoginHint(invokerOptions);
        var scopes = invokerOptions.scopes.map(function (s) {
            return s + (s.indexOf('Files.') > -1 && invokerOptions.needSharePointPermission ? '.All' : '');
        });
        return {
            clientId: invokerOptions.clientId,
            endpoint: endpoint,
            scopes: scopes,
            loginHint: invokerOptions.loginHint || (loginHint ? loginHint.loginHint : null),
            domainHint: invokerOptions.domainHint || (loginHint ? loginHint.domainHint : null),
            origin: window.location.origin,
            redirectUri: invokerOptions.redirectUri,
            switchAccount: switchAccount
        };
    }
    exports.buildOauthConfig = buildOauthConfig;
}(require, exports, require('../utilities/Channel'), require('../models/DomainHint'), require('../utilities/DomUtilities'), require('../utilities/ErrorHandler'), require('../models/ErrorType'), require('./LoginCache'), require('../models/OauthEndpoint'), require('../models/OneDriveSdkError'), require('es6-promise'), require('../utilities/UrlUtilities')));
},{"../models/DomainHint":12,"../models/ErrorType":13,"../models/OauthEndpoint":15,"../models/OneDriveSdkError":16,"../utilities/Channel":24,"../utilities/DomUtilities":25,"../utilities/ErrorHandler":26,"../utilities/UrlUtilities":32,"./LoginCache":6,"es6-promise":34}],8:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, ApiRequest_1, Constants_1, Invoker_1, ObjectUtilities_1, PickerOptions_1, PickerActionType_1, StringUtilities_1, UrlUtilities_1) {
    'use strict';
    var Picker = function (_super) {
        __extends(Picker, _super);
        function Picker(options) {
            var clonedOptions = ObjectUtilities_1.shallowClone(options);
            var pickerOptions = new PickerOptions_1.default(clonedOptions);
            _super.call(this, pickerOptions);
        }
        Picker.prototype.launchPicker = function () {
            return _super.prototype.launchInvoker.call(this);
        };
        Picker.prototype.buildPickerUXConfig = function (pickerOptions) {
            var pickerUXConfig = {
                applicationId: pickerOptions.clientId,
                accessLevel: Picker.ACCESS_LEVEL,
                filter: pickerOptions.filter,
                id: UrlUtilities_1.generateNonce(),
                navEnabled: pickerOptions.navEnabled,
                origin: window.location.origin,
                parentDiv: pickerOptions.parentDiv,
                redirectUri: pickerOptions.redirectUri,
                selectionMode: pickerOptions.multiSelect ? 'multiple' : 'single',
                viewType: Picker.VIEW_TYPE
            };
            return pickerUXConfig;
        };
        Picker.prototype.makeApiRequest = function (files) {
            if (this.invokerOptions.action === PickerActionType_1.default.share) {
                return this.shareItems(files);
            } else {
                var isDownload = this.invokerOptions.action === PickerActionType_1.default.download;
                return this.queryItems(files, isDownload);
            }
        };
        Picker.prototype.queryItems = function (files, isDownload) {
            var itemQuery = this.invokerOptions.queryParameters || Constants_1.default.DEFAULT_QUERY_ITEM_PARAMETER;
            if (isDownload) {
                itemQuery = StringUtilities_1.format('{0}{1}{2}', itemQuery, itemQuery.indexOf('select') === -1 ? '&select=' : ',', 'name,size,@content.downloadUrl');
            }
            return ApiRequest_1.getItems(files, this.getApiRequestConfig(), itemQuery);
        };
        Picker.prototype.shareItems = function (files) {
            var _this = this;
            var pickerOptions = this.invokerOptions;
            var createLinkParameters = pickerOptions.createLinkParameters || this.getDefaultSharingConfig();
            return ApiRequest_1.getItems(files, this.getApiRequestConfig()).then(function (files) {
                return ApiRequest_1.shareItems(files, _this.getApiRequestConfig(), createLinkParameters);
            });
        };
        Picker.prototype.getDefaultSharingConfig = function () {
            var createLinkParameters = { 'type': 'view' };
            if (this.getApiRequestConfig().apiEndpoint === ApiEndpoint_1.default.graph_odc || this.getApiRequestConfig().apiEndpoint === ApiEndpoint_1.default.msa) {
                return createLinkParameters;
            }
            createLinkParameters['scope'] = 'organization';
            return createLinkParameters;
        };
        Picker.ACCESS_LEVEL = 'read';
        Picker.VIEW_TYPE = 'files';
        return Picker;
    }(Invoker_1.default);
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Picker;
}(require, exports, require('../models/ApiEndpoint'), require('./ApiRequest'), require('../Constants'), require('./Invoker'), require('../utilities/ObjectUtilities'), require('../models/PickerOptions'), require('../models/PickerActionType'), require('../utilities/StringUtilities'), require('../utilities/UrlUtilities')));
},{"../Constants":1,"../models/ApiEndpoint":11,"../models/PickerActionType":17,"../models/PickerOptions":18,"../utilities/ObjectUtilities":28,"../utilities/StringUtilities":30,"../utilities/UrlUtilities":32,"./ApiRequest":4,"./Invoker":5}],9:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, Channel_1, DomUtilities_1, ErrorHandler_1, ErrorType_1, Logging_1, OneDriveSdkError_1, UrlUtilities_1, es6_promise_1, Constants_1) {
    'use strict';
    var CUSTOMER_PICKER_BASE_URL = 'https://onedrive.live.com/';
    var RESPONSE_PREFIX = '[OneDrive-FromPicker]';
    var MESSAGE_PREFIX = '[OneDrive-ToPicker]';
    var INITIALIZE_RESPONSE = 'initialize';
    function generatePickerUX(apiEndpoint, tenantUrl) {
        return new PickerUX(apiEndpoint, tenantUrl);
    }
    exports.generatePickerUX = generatePickerUX;
    var PickerUX = function () {
        function PickerUX(apiEndpoint, tenantUrl) {
            if (apiEndpoint === ApiEndpoint_1.default.graph_odc || apiEndpoint === ApiEndpoint_1.default.msa) {
                this.url = UrlUtilities_1.appendQueryStrings(CUSTOMER_PICKER_BASE_URL, { 'v': '2' });
            } else if (apiEndpoint === ApiEndpoint_1.default.graph_odb || apiEndpoint === ApiEndpoint_1.default.filesV2) {
                if (!tenantUrl) {
                    ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'the site url must be specified')).exposeToPublic();
                }
                UrlUtilities_1.validateUrlProtocol(tenantUrl, ['HTTPS']);
                if (apiEndpoint === ApiEndpoint_1.default.graph_odb) {
                    tenantUrl = UrlUtilities_1.appendToPath(tenantUrl, '_layouts/onedrive.aspx');
                }
                this.url = UrlUtilities_1.appendQueryString(tenantUrl, 'p', '2');
            }
        }
        PickerUX.prototype.invokePickerUX = function (uxConfig, popupView) {
            var _this = this;
            var receivedResponse = false;
            var pickerUXPromise = new es6_promise_1.Promise(function (resolve, reject) {
                var channelId = DomUtilities_1.onMessage(function (event) {
                    var urlParts = _this.url.split('/');
                    var channel = new Channel_1.default(_this.iframe ? _this.iframe.contentWindow : popupView.getPopupWindow());
                    if (event.origin === urlParts[0] + '//' + urlParts[2]) {
                        var message = '' + (event.data || '');
                        if (message.indexOf(RESPONSE_PREFIX) === 0 && event.source === channel.getReceiver()) {
                            var response = JSON.parse(message.substring(RESPONSE_PREFIX.length));
                            var pickerId = response.pickerId, conversationId = response.conversationId, type = response.type;
                            if (pickerId === uxConfig.id) {
                                if (type === INITIALIZE_RESPONSE) {
                                    channel.send(MESSAGE_PREFIX + JSON.stringify({
                                        pickerId: pickerId,
                                        conversationId: conversationId,
                                        type: 'activate'
                                    }), event.origin);
                                } else {
                                    receivedResponse = true;
                                    resolve(response);
                                    DomUtilities_1.removeMessageListener(channelId);
                                }
                            }
                        } else {
                            reject(new OneDriveSdkError_1.default(ErrorType_1.default.badResponse, 'received invalid response from picker UI'));
                        }
                    }
                });
                var pickerOption = {
                    aid: uxConfig.applicationId,
                    a: uxConfig.accessLevel,
                    id: uxConfig.id,
                    l: uxConfig.linkType,
                    ln: uxConfig.navEnabled,
                    s: uxConfig.selectionMode,
                    f: uxConfig.filter,
                    v: uxConfig.viewType,
                    ru: uxConfig.redirectUri,
                    o: uxConfig.origin,
                    sdk: Constants_1.default.SDK_VERSION_NUMBER,
                    e: uxConfig.entryLocation,
                    st: uxConfig.sourceTypes,
                    sn: !uxConfig.parentDiv,
                    ss: !uxConfig.parentDiv
                };
                var pickerUrl = UrlUtilities_1.appendQueryString(_this.url, 'picker', JSON.stringify(pickerOption));
                Logging_1.default.logMessage('invoke picker with url: ' + pickerUrl);
                if (uxConfig.parentDiv) {
                    popupView.close();
                    var iframe = document.createElement('iframe');
                    iframe.id = 'odpicker' + new Date().getTime();
                    iframe.style.position = 'relative';
                    iframe.style.width = '100%';
                    iframe.style.height = '100%';
                    iframe.src = pickerUrl;
                    uxConfig.parentDiv.appendChild(iframe);
                    _this.iframe = iframe;
                } else {
                    return popupView.openPopup(pickerUrl).then(function () {
                        resolve({ type: 'cancel' });
                    });
                }
            });
            return pickerUXPromise;
        };
        PickerUX.prototype.removeIFrame = function () {
            if (this.iframe) {
                this.iframe.parentNode.removeChild(this.iframe);
                this.iframe = null;
            }
        };
        return PickerUX;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = PickerUX;
}(require, exports, require('../models/ApiEndpoint'), require('../utilities/Channel'), require('../utilities/DomUtilities'), require('../utilities/ErrorHandler'), require('../models/ErrorType'), require('../utilities/Logging'), require('../models/OneDriveSdkError'), require('../utilities/UrlUtilities'), require('es6-promise'), require('../Constants')));
},{"../Constants":1,"../models/ApiEndpoint":11,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"../utilities/Channel":24,"../utilities/DomUtilities":25,"../utilities/ErrorHandler":26,"../utilities/Logging":27,"../utilities/UrlUtilities":32,"es6-promise":34}],10:[function(require,module,exports){
(function (require, exports, ApiRequest_1, Constants_1, Invoker_1, ObjectUtilities_1, SaverActionType_1, SaverOptions_1, UploadType_1, UrlUtilities_1) {
    'use strict';
    var ACCESS_LEVEL = 'readwrite';
    var VIEW_TYPE = 'folders';
    var SELECTION_MODE = 'single';
    var Saver = function (_super) {
        __extends(Saver, _super);
        function Saver(options) {
            var clonedOptions = ObjectUtilities_1.shallowClone(options);
            var saverOptions = new SaverOptions_1.default(clonedOptions);
            _super.call(this, saverOptions);
        }
        Saver.prototype.launchSaver = function () {
            return _super.prototype.launchInvoker.call(this);
        };
        Saver.prototype.buildPickerUXConfig = function (saverOptions) {
            return {
                applicationId: saverOptions.clientId,
                accessLevel: ACCESS_LEVEL,
                id: UrlUtilities_1.generateNonce(),
                navEnabled: saverOptions.navEnabled,
                filter: saverOptions.filter,
                origin: window.location.origin,
                parentDiv: saverOptions.parentDiv,
                redirectUri: saverOptions.redirectUri,
                selectionMode: SELECTION_MODE,
                viewType: VIEW_TYPE
            };
        };
        Saver.prototype.makeApiRequest = function (files) {
            var saverOptions = this.invokerOptions;
            if (this.invokerOptions.action === SaverActionType_1.default.query) {
                var itemQuery = this.invokerOptions.queryParameters || Constants_1.default.DEFAULT_QUERY_ITEM_PARAMETER;
                return ApiRequest_1.getItems(files, this.apiRequestConfig, itemQuery);
            } else if (saverOptions.uploadType === UploadType_1.default.dataUrl || saverOptions.uploadType === UploadType_1.default.url) {
                var uploadItemDescription = { name: saverOptions.fileName };
                return ApiRequest_1.saveItemByUriUpload(files.value[0], uploadItemDescription, saverOptions.sourceUri, this.apiRequestConfig);
            } else if (saverOptions.uploadType === UploadType_1.default.form) {
                var uploadItemDescription = {
                    name: saverOptions.fileName,
                    '@name.conflictBehavior': saverOptions.nameConflictBehavior
                };
                return ApiRequest_1.saveItemByFormUpload(files.value[0], uploadItemDescription, saverOptions.fileInput, this.apiRequestConfig, saverOptions.progress);
            }
        };
        return Saver;
    }(Invoker_1.default);
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Saver;
}(require, exports, require('./ApiRequest'), require('../Constants'), require('./Invoker'), require('../utilities/ObjectUtilities'), require('../models/SaverActionType'), require('../models/SaverOptions'), require('../models/UploadType'), require('../utilities/UrlUtilities')));
},{"../Constants":1,"../models/SaverActionType":19,"../models/SaverOptions":20,"../models/UploadType":21,"../utilities/ObjectUtilities":28,"../utilities/UrlUtilities":32,"./ApiRequest":4,"./Invoker":5}],11:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var ApiEndpoint;
    (function (ApiEndpoint) {
        ApiEndpoint[ApiEndpoint['filesV2'] = 0] = 'filesV2';
        ApiEndpoint[ApiEndpoint['graph_odc'] = 1] = 'graph_odc';
        ApiEndpoint[ApiEndpoint['graph_odb'] = 2] = 'graph_odb';
        ApiEndpoint[ApiEndpoint['msa'] = 3] = 'msa';
    }(ApiEndpoint || (ApiEndpoint = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = ApiEndpoint;
}(require, exports));
},{}],12:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var DomainHint;
    (function (DomainHint) {
        DomainHint[DomainHint['aad'] = 0] = 'aad';
        DomainHint[DomainHint['msa'] = 1] = 'msa';
        DomainHint[DomainHint['tenant'] = 2] = 'tenant';
    }(DomainHint || (DomainHint = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = DomainHint;
}(require, exports));
},{}],13:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var ErrorType;
    (function (ErrorType) {
        ErrorType[ErrorType['badResponse'] = 0] = 'badResponse';
        ErrorType[ErrorType['fileReaderFailure'] = 1] = 'fileReaderFailure';
        ErrorType[ErrorType['popupOpen'] = 2] = 'popupOpen';
        ErrorType[ErrorType['unknown'] = 3] = 'unknown';
        ErrorType[ErrorType['unsupportedFeature'] = 4] = 'unsupportedFeature';
        ErrorType[ErrorType['webRequestFailure'] = 5] = 'webRequestFailure';
        ErrorType[ErrorType['internalError'] = 6] = 'internalError';
        ErrorType[ErrorType['optionsError'] = 7] = 'optionsError';
        ErrorType[ErrorType['typeError'] = 8] = 'typeError';
        ErrorType[ErrorType['popupClosed'] = 9] = 'popupClosed';
    }(ErrorType || (ErrorType = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = ErrorType;
}(require, exports));
},{}],14:[function(require,module,exports){
(function (require, exports, CallbackInvoker_1, Constants_1, DomainHint_1, ErrorHandler_1, ErrorType_1, Logging_1, OneDriveSdkError_1, StringUtilities_1, TypeValidators_1, UrlUtilities_1) {
    'use strict';
    var AAD_APPID_PATTERN = new RegExp('^[a-fA-F\\d]{8}-([a-fA-F\\d]{4}-){3}[a-fA-F\\d]{12}$');
    var InvokerOptions = function () {
        function InvokerOptions(options) {
            this.navEnabled = true;
            this.needSharePointPermission = true;
            this.clientId = TypeValidators_1.validateType(options.clientId, Constants_1.default.TYPE_STRING);
            var cancelCallback = TypeValidators_1.validateCallback(options.cancel, true);
            this.cancel = function () {
                Logging_1.default.logMessage('user cancelled operation');
                if (cancelCallback) {
                    CallbackInvoker_1.invokeAppCallback(cancelCallback, true);
                }
            };
            var errorCallback = TypeValidators_1.validateCallback(options.error, true);
            this.error = function (error) {
                if (errorCallback) {
                    CallbackInvoker_1.invokeAppCallback(errorCallback, true, error);
                } else {
                    throw error;
                }
            };
            this.parseAdvancedOptions(options);
            this.redirectUri = this.redirectUri || UrlUtilities_1.trimUrlQuery(window.location.href);
            this.endpointHint = this.endpointHint || DomainHint_1.default.aad;
            InvokerOptions.checkClientId(this.clientId);
        }
        InvokerOptions.checkClientId = function (clientId) {
            if (clientId) {
                if (AAD_APPID_PATTERN.test(clientId)) {
                    Logging_1.default.logMessage('parsed client id: ' + clientId);
                } else {
                    ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unknown, StringUtilities_1.format('invalid format for client id \'{0}\' - AAD: 32 characters (HEX) GUID', clientId)));
                }
            } else {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unknown, 'client id is missing in options'));
            }
        };
        InvokerOptions.prototype.needOauth = function () {
            return this.needAPICall() && !this.accessToken || this.endpointHint === DomainHint_1.default.aad;
        };
        InvokerOptions.prototype.parseAdvancedOptions = function (options) {
            if (options.advanced) {
                if (!!options.advanced.redirectUri) {
                    UrlUtilities_1.validateRedirectUrlHost(options.advanced.redirectUri);
                    this.redirectUri = options.advanced.redirectUri;
                }
                if (!!options.advanced.queryParameters) {
                    var itemQueries = UrlUtilities_1.readUrlParameters('?' + options.advanced.queryParameters);
                    for (var key in itemQueries) {
                        if (key.toLowerCase() !== 'select' && key.toLowerCase() !== 'expand') {
                            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, StringUtilities_1.format('unexpected query key: {0} is found in advanced.queryParameters', key)));
                        }
                    }
                    var selectQuery = itemQueries['select'];
                    var expandQuery = itemQueries['expand'];
                    if (selectQuery && expandQuery) {
                        this.queryParameters = StringUtilities_1.format('expand={0}&select={1}', expandQuery, selectQuery);
                    } else if (expandQuery) {
                        this.queryParameters = StringUtilities_1.format('expand={0}', expandQuery);
                    } else if (selectQuery) {
                        if ('select=' + selectQuery.split(',').sort().join(',') !== Constants_1.default.DEFAULT_QUERY_ITEM_PARAMETER) {
                            this.queryParameters = StringUtilities_1.format('select={0}', selectQuery);
                        }
                    }
                }
                if (!!options.advanced.endpointHint) {
                    if (options.advanced.endpointHint.toLowerCase() === Constants_1.default.VROOM_ENDPOINT_HINT) {
                        this.endpointHint = DomainHint_1.default.msa;
                    } else {
                        var domainHint = TypeValidators_1.validateType(options.advanced.endpointHint, 'string', false);
                        UrlUtilities_1.validateUrlProtocol(domainHint);
                        this.endpointHint = DomainHint_1.default.tenant;
                        this.siteUrl = domainHint;
                    }
                    if (!!options.advanced.accessToken) {
                        this.accessToken = options.advanced.accessToken;
                    }
                }
                if (!!options.advanced.iframeParentDiv) {
                    if (!options.advanced.iframeParentDiv.nodeName || options.advanced.iframeParentDiv.nodeName.toLowerCase() !== 'div') {
                        ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'the iframe\'s parent div element is not a DOM object')).exposeToPublic();
                    }
                    this.parentDiv = options.advanced.iframeParentDiv;
                }
                if (!!options.advanced.scopes) {
                    if (typeof options.advanced.scopes === 'string') {
                        this.scopes = [options.advanced.scopes];
                    } else if (options.advanced.scopes instanceof Array) {
                        this.scopes = options.advanced.scopes;
                    }
                }
                this.linkType = options.advanced.linkType;
                this.parseNavigationOptions(options.advanced.navigation);
                this.loginHint = options.advanced.loginHint;
                this.filter = options.advanced.filter;
            }
        };
        InvokerOptions.prototype.parseNavigationOptions = function (navigation) {
            if (navigation) {
                var entryLocation = navigation.entryLocation;
                if (entryLocation) {
                    var _a = entryLocation.sharePoint, sitePath = _a.sitePath, listPath = _a.listPath;
                    if (sitePath) {
                        UrlUtilities_1.validateUrlProtocol(sitePath, ['HTTPS']);
                    }
                    if (listPath) {
                        UrlUtilities_1.validateUrlProtocol(listPath, ['HTTPS']);
                    }
                    this.navEntryLocation = entryLocation;
                }
                var sourceTypes = navigation.sourceTypes instanceof Array ? navigation.sourceTypes : navigation.sourceTypes ? [navigation.sourceTypes] : null;
                if (!!sourceTypes) {
                    this.needSharePointPermission = !(sourceTypes.length === 1 && sourceTypes[0].toLowerCase() === 'onedrive');
                    this.navSourceTypes = sourceTypes;
                }
                this.navEnabled = !navigation.disable;
            }
        };
        return InvokerOptions;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = InvokerOptions;
}(require, exports, require('../utilities/CallbackInvoker'), require('../Constants'), require('./DomainHint'), require('../utilities/ErrorHandler'), require('./ErrorType'), require('../utilities/Logging'), require('./OneDriveSdkError'), require('../utilities/StringUtilities'), require('../utilities/TypeValidators'), require('../utilities/UrlUtilities')));
},{"../Constants":1,"../utilities/CallbackInvoker":23,"../utilities/ErrorHandler":26,"../utilities/Logging":27,"../utilities/StringUtilities":30,"../utilities/TypeValidators":31,"../utilities/UrlUtilities":32,"./DomainHint":12,"./ErrorType":13,"./OneDriveSdkError":16}],15:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var OauthEndpoint;
    (function (OauthEndpoint) {
        OauthEndpoint[OauthEndpoint['AAD'] = 0] = 'AAD';
        OauthEndpoint[OauthEndpoint['AADv2'] = 1] = 'AADv2';
        OauthEndpoint[OauthEndpoint['MSA'] = 2] = 'MSA';
    }(OauthEndpoint || (OauthEndpoint = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = OauthEndpoint;
}(require, exports));
},{}],16:[function(require,module,exports){
(function (require, exports, ErrorType_1, StringUtilities_1) {
    'use strict';
    var OneDriveSdkError = function (_super) {
        __extends(OneDriveSdkError, _super);
        function OneDriveSdkError(errorCode, message) {
            _super.call(this, message);
            this.errorCode = ErrorType_1.default[errorCode];
            this.message = message;
        }
        OneDriveSdkError.prototype.toString = function () {
            return StringUtilities_1.format('[OneDriveSDK Error] errorType: {0}, message: {1}', this.errorCode, this.message);
        };
        return OneDriveSdkError;
    }(Error);
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = OneDriveSdkError;
}(require, exports, require('../models/ErrorType'), require('../utilities/StringUtilities')));
},{"../models/ErrorType":13,"../utilities/StringUtilities":30}],17:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var PickerActionType;
    (function (PickerActionType) {
        PickerActionType[PickerActionType['download'] = 0] = 'download';
        PickerActionType[PickerActionType['query'] = 1] = 'query';
        PickerActionType[PickerActionType['share'] = 2] = 'share';
    }(PickerActionType || (PickerActionType = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = PickerActionType;
}(require, exports));
},{}],18:[function(require,module,exports){
(function (require, exports, CallbackInvoker_1, Constants_1, InvokerOptions_1, Logging_1, PickerActionType_1, TypeValidators_1) {
    'use strict';
    var PickerOptions = function (_super) {
        __extends(PickerOptions, _super);
        function PickerOptions(options) {
            _super.call(this, options);
            var successCallback = TypeValidators_1.validateCallback(options.success, false);
            this.success = function (files) {
                Logging_1.default.logMessage('picker operation succeeded');
                CallbackInvoker_1.invokeAppCallback(successCallback, true, files);
            };
            this.multiSelect = TypeValidators_1.validateType(options.multiSelect, Constants_1.default.TYPE_BOOLEAN, true, false);
            var actionName = TypeValidators_1.validateType(options.action, Constants_1.default.TYPE_STRING, true, PickerActionType_1.default[PickerActionType_1.default.query]);
            this.action = PickerActionType_1.default[actionName];
            if (options.advanced) {
                this.createLinkParameters = options.advanced.createLinkParameters;
            }
            if (!this.scopes) {
                this.scopes = [this.action === PickerActionType_1.default.share ? 'Files.ReadWrite' : 'Files.Read'];
            }
        }
        PickerOptions.prototype.needAPICall = function () {
            return !!this.queryParameters || this.action !== PickerActionType_1.default.query;
        };
        return PickerOptions;
    }(InvokerOptions_1.default);
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = PickerOptions;
}(require, exports, require('../utilities/CallbackInvoker'), require('../Constants'), require('./InvokerOptions'), require('../utilities/Logging'), require('./PickerActionType'), require('../utilities/TypeValidators')));
},{"../Constants":1,"../utilities/CallbackInvoker":23,"../utilities/Logging":27,"../utilities/TypeValidators":31,"./InvokerOptions":14,"./PickerActionType":17}],19:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var SaverActionType;
    (function (SaverActionType) {
        SaverActionType[SaverActionType['save'] = 0] = 'save';
        SaverActionType[SaverActionType['query'] = 1] = 'query';
    }(SaverActionType || (SaverActionType = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = SaverActionType;
}(require, exports));
},{}],20:[function(require,module,exports){
(function (require, exports, CallbackInvoker_1, Constants_1, DomUtilities_1, ErrorHandler_1, ErrorType_1, InvokerOptions_1, Logging_1, OneDriveSdkError_1, SaverActionType_1, StringUtilities_1, TypeValidators_1, UploadType_1, UrlUtilities_1) {
    'use strict';
    var SaverOptions = function (_super) {
        __extends(SaverOptions, _super);
        function SaverOptions(options) {
            _super.call(this, options);
            var successCallback = TypeValidators_1.validateCallback(options.success, false);
            this.success = function (folder) {
                Logging_1.default.logMessage('saver operation succeeded');
                CallbackInvoker_1.invokeAppCallback(successCallback, true, folder);
            };
            var progressCallback = TypeValidators_1.validateCallback(options.progress, true);
            this.progress = function (percentage) {
                Logging_1.default.logMessage(StringUtilities_1.format('upload progress: {0}%', percentage));
                if (progressCallback) {
                    CallbackInvoker_1.invokeAppCallback(progressCallback, false, percentage);
                }
            };
            var actionName = TypeValidators_1.validateType(options.action, Constants_1.default.TYPE_STRING, true, SaverActionType_1.default[SaverActionType_1.default.query]);
            this.action = SaverActionType_1.default[actionName];
            if (this.action === SaverActionType_1.default.save) {
                this._setFileInfo(options);
            }
            this.nameConflictBehavior = TypeValidators_1.validateType(options.nameConflictBehavior, Constants_1.default.TYPE_STRING, true, 'rename');
            if (!this.scopes) {
                this.scopes = ['Files.ReadWrite'];
            }
        }
        SaverOptions.prototype.needAPICall = function () {
            return !!this.queryParameters || this.action === SaverActionType_1.default.save;
        };
        SaverOptions.prototype._setFileInfo = function (options) {
            if (options.sourceInputElementId && options.sourceUri) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'sourceUri and sourceInputElementId' + ' are provided, only one is required.'));
            }
            this.sourceInputElementId = options.sourceInputElementId;
            this.sourceUri = options.sourceUri;
            var fileName = TypeValidators_1.validateType(options.fileName, Constants_1.default.TYPE_STRING, true, null);
            if (this.sourceUri) {
                if (UrlUtilities_1.isPathFullUrl(this.sourceUri)) {
                    this.uploadType = UploadType_1.default.url;
                    this.fileName = fileName || UrlUtilities_1.getFileNameFromUrl(this.sourceUri);
                    if (!this.fileName) {
                        ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'must supply a file name or a URL that ends with a file name'));
                    }
                } else if (UrlUtilities_1.isPathDataUrl(this.sourceUri)) {
                    this.uploadType = UploadType_1.default.dataUrl;
                    this.fileName = fileName;
                    if (!this.fileName) {
                        ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'must supply a file name for data URL uploads'));
                    }
                }
            } else if (this.sourceInputElementId) {
                this.uploadType = UploadType_1.default.form;
                this.fileInput = DomUtilities_1.getFileInput(this.sourceInputElementId);
                this.fileName = fileName || this.fileInput.name;
            } else {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'please specified one type of resource to save'));
            }
        };
        return SaverOptions;
    }(InvokerOptions_1.default);
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = SaverOptions;
}(require, exports, require('../utilities/CallbackInvoker'), require('../Constants'), require('../utilities/DomUtilities'), require('../utilities/ErrorHandler'), require('./ErrorType'), require('./InvokerOptions'), require('../utilities/Logging'), require('./OneDriveSdkError'), require('./SaverActionType'), require('../utilities/StringUtilities'), require('../utilities/TypeValidators'), require('./UploadType'), require('../utilities/UrlUtilities')));
},{"../Constants":1,"../utilities/CallbackInvoker":23,"../utilities/DomUtilities":25,"../utilities/ErrorHandler":26,"../utilities/Logging":27,"../utilities/StringUtilities":30,"../utilities/TypeValidators":31,"../utilities/UrlUtilities":32,"./ErrorType":13,"./InvokerOptions":14,"./OneDriveSdkError":16,"./SaverActionType":19,"./UploadType":21}],21:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var UploadType;
    (function (UploadType) {
        UploadType[UploadType['dataUrl'] = 0] = 'dataUrl';
        UploadType[UploadType['form'] = 1] = 'form';
        UploadType[UploadType['url'] = 2] = 'url';
    }(UploadType || (UploadType = {})));
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = UploadType;
}(require, exports));
},{}],22:[function(require,module,exports){
(function (require, exports, ErrorType_1, ErrorHandler_1, OneDriveSdkError_1) {
    'use strict';
    var CACHE_NAME = 'odpickerv7cache';
    function getCacheItem(key) {
        var cache = getCache();
        return cache[key];
    }
    exports.getCacheItem = getCacheItem;
    function setCacheItem(key, object) {
        var cache = getCache();
        cache[key] = object;
        setCache(cache);
    }
    exports.setCacheItem = setCacheItem;
    function removeCacheItem(key) {
        var cache = getCache();
        var deletedItem = cache[key];
        delete cache[key];
        setCache(cache);
        return deletedItem;
    }
    exports.removeCacheItem = removeCacheItem;
    function getCache() {
        if (Storage) {
            var cacheString = localStorage.getItem(CACHE_NAME);
            var cache = JSON.parse(cacheString || '{}');
            return cache;
        } else {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'cache based on Storage is not supported in this browser'));
        }
    }
    function setCache(cache) {
        if (Storage) {
            localStorage.setItem(CACHE_NAME, JSON.stringify(cache));
        } else {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'cache based on Storage is not supported in this browser'));
        }
    }
}(require, exports, require('../models/ErrorType'), require('./ErrorHandler'), require('../models/OneDriveSdkError')));
},{"../models/ErrorType":13,"../models/OneDriveSdkError":16,"./ErrorHandler":26}],23:[function(require,module,exports){
(function (require, exports, Constants_1) {
    'use strict';
    function invokeAppCallback(callback, isFinalCallback) {
        var args = [];
        for (var _i = 2; _i < arguments.length; _i++) {
            args[_i - 2] = arguments[_i];
        }
        if (typeof callback === Constants_1.default.TYPE_FUNCTION) {
            callback.apply(null, args);
        }
    }
    exports.invokeAppCallback = invokeAppCallback;
    function invokeCallbackAsynchronous(callback) {
        var args = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            args[_i - 1] = arguments[_i];
        }
        window.setTimeout(function () {
            callback.apply(null, args);
        }, 0);
    }
    exports.invokeCallbackAsynchronous = invokeCallbackAsynchronous;
}(require, exports, require('../Constants')));
},{"../Constants":1}],24:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var Channel = function () {
        function Channel(receiver) {
            this.receiver = receiver;
        }
        Channel.prototype.send = function (data, origin) {
            if (this.receiver) {
                this.receiver.postMessage(data, origin);
            }
        };
        Channel.prototype.getReceiver = function () {
            return this.receiver;
        };
        return Channel;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Channel;
}(require, exports));
},{}],25:[function(require,module,exports){
(function (require, exports, ErrorHandler_1, ErrorType_1, OneDriveSdkError_1, UrlUtilities_1) {
    'use strict';
    var FORM_UPLOAD_SIZE_LIMIT = 104857600;
    var FORM_UPLOAD_SIZE_LIMIT_STRING = '100 MB';
    var MESSAGE_LISTENERS = {};
    function getElementById(id) {
        return document.getElementById(id);
    }
    exports.getElementById = getElementById;
    function onDocumentReady(callback) {
        if (document.readyState === 'interactive' || document.readyState === 'complete') {
            callback();
        } else {
            document.addEventListener('DOMContentLoaded', callback, false);
        }
    }
    exports.onDocumentReady = onDocumentReady;
    function onMessage(callback) {
        var id = UrlUtilities_1.generateNonce() + '_' + new Date().getMilliseconds();
        window.addEventListener('message', callback);
        MESSAGE_LISTENERS[id] = callback;
        return id;
    }
    exports.onMessage = onMessage;
    function removeMessageListener(id) {
        var callback = MESSAGE_LISTENERS[id];
        if (callback) {
            window.removeEventListener('message', callback);
        }
    }
    exports.removeMessageListener = removeMessageListener;
    function getFileInput(elementId) {
        var fileInputElement = getElementById(elementId);
        if (fileInputElement instanceof HTMLInputElement) {
            if (fileInputElement.type !== 'file') {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'input elemenet must be of type \'file\''));
            }
            if (!fileInputElement.value) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'please a file to upload'));
                return null;
            }
            var files = fileInputElement.files;
            if (!files || !window['FileReader']) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'browser does not support Files API'));
            }
            if (files.length !== 1) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unsupportedFeature, 'can not upload more than one file at a time'));
            }
            var fileInput = files[0];
            if (!fileInput) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'missing file input'));
            }
            if (fileInput.size > FORM_UPLOAD_SIZE_LIMIT) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'the user has selected a file larger than ' + FORM_UPLOAD_SIZE_LIMIT_STRING));
                return null;
            }
            return fileInput;
        } else {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.unknown, 'element was not an instance of an HTMLInputElement'));
        }
    }
    exports.getFileInput = getFileInput;
    function displayOverlay() {
        var overlay = document.createElement('div');
        var overlayStyle = [
            'position: fixed',
            'width: 100%',
            'height: 100%',
            'top: 0px',
            'left: 0px',
            'background-color: white',
            'opacity: 1',
            'z-index: 10000',
            'min-width: 40px',
            'min-height: 40px'
        ];
        overlay.id = 'od-overlay';
        overlay.style.cssText = overlayStyle.join(';');
        var spinner = document.createElement('img');
        var spinnerStyle = [
            'position: absolute',
            'margin: auto',
            'top: 0',
            'left: 0',
            'right: 0',
            'bottom: 0'
        ];
        spinner.id = 'od-spinner';
        spinner.src = 'https://p.sfx.ms/common/spinner_grey_40_transparent.gif';
        spinner.style.cssText = spinnerStyle.join(';');
        overlay.appendChild(spinner);
        var hiddenStyle = document.createElement('style');
        hiddenStyle.type = 'text/css';
        hiddenStyle.innerHTML = 'body { visibility: hidden !important; }';
        document.head.appendChild(hiddenStyle);
        onDocumentReady(function () {
            var documentBody = document.body;
            if (documentBody !== null) {
                documentBody.insertBefore(overlay, documentBody.firstChild);
            } else {
                document.createElement('body').appendChild(overlay);
            }
            document.head.removeChild(hiddenStyle);
        });
    }
    exports.displayOverlay = displayOverlay;
}(require, exports, require('./ErrorHandler'), require('../models/ErrorType'), require('../models/OneDriveSdkError'), require('./UrlUtilities')));
},{"../models/ErrorType":13,"../models/OneDriveSdkError":16,"./ErrorHandler":26,"./UrlUtilities":32}],26:[function(require,module,exports){
(function (require, exports, ErrorType_1) {
    'use strict';
    var ERROR_HANDLERS = [];
    function registerErrorObserver(callback) {
        ERROR_HANDLERS.push(callback);
    }
    exports.registerErrorObserver = registerErrorObserver;
    function throwError(sdkError) {
        if (sdkError.errorCode !== ErrorType_1.default[ErrorType_1.default.unknown]) {
            for (var _i = 0, ERROR_HANDLERS_1 = ERROR_HANDLERS; _i < ERROR_HANDLERS_1.length; _i++) {
                var callback = ERROR_HANDLERS_1[_i];
                callback(sdkError);
            }
            return {
                exposeToPublic: function () {
                    throw sdkError;
                }
            };
        } else {
            throw sdkError;
        }
    }
    exports.throwError = throwError;
}(require, exports, require('../models/ErrorType')));
},{"../models/ErrorType":13}],27:[function(require,module,exports){
(function (require, exports) {
    'use strict';
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
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Logging;
}(require, exports));
},{}],28:[function(require,module,exports){
(function (require, exports, Constants_1, Logging_1) {
    'use strict';
    function shallowClone(object) {
        if (typeof object !== Constants_1.default.TYPE_OBJECT || !object) {
            return null;
        }
        var clonedObject = {};
        for (var key in object) {
            if (object.hasOwnProperty(key)) {
                clonedObject[key] = object[key];
            }
        }
        return clonedObject;
    }
    exports.shallowClone = shallowClone;
    function deserializeJSON(text) {
        var deserializedObject = null;
        try {
            deserializedObject = JSON.parse(text);
        } catch (error) {
            Logging_1.default.logError('deserialization error' + error);
        }
        if (typeof deserializedObject !== Constants_1.default.TYPE_OBJECT || deserializedObject === null) {
            deserializedObject = {};
        }
        return deserializedObject;
    }
    exports.deserializeJSON = deserializeJSON;
    function serializeJSON(value) {
        return JSON.stringify(value);
    }
    exports.serializeJSON = serializeJSON;
}(require, exports, require('../Constants'), require('./Logging')));
},{"../Constants":1,"./Logging":27}],29:[function(require,module,exports){
(function (require, exports, ErrorHandler_1, ErrorType_1, Logging_1, OneDriveSdkError_1, UrlUtilities_1) {
    'use strict';
    var POPUP_WIDTH = 1024;
    var POPUP_HEIGHT = 650;
    exports.POPUP_PINGER_INTERVAL = 500;
    var Popup = function () {
        function Popup() {
        }
        Popup.getCurrentPopup = function () {
            return Popup._currentPopup || new Popup();
        };
        Popup.setCurrentPopup = function (popup) {
            Popup._currentPopup = popup;
        };
        Popup.createPopupFeatures = function () {
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
        Popup.prototype.close = function () {
            if (Popup.getCurrentPopup().isPopupOpen()) {
                Popup.getCurrentPopup()._popup.close();
                Popup.setCurrentPopup(null);
            }
        };
        Popup.prototype.openPopup = function (url) {
            var _this = this;
            UrlUtilities_1.validateUrlProtocol(url);
            if (Popup.getCurrentPopup().isPopupOpen()) {
                Logging_1.default.logMessage('leaving current url: ' + this._url);
                this._url = url;
                Popup.getCurrentPopup().getPopupWindow().location.href = url;
            } else {
                this._url = url;
                this._popup = window.open(url, '_blank', Popup.createPopupFeatures());
                if (!this._popup) {
                    ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.popupOpen, 'popup window is disconnected')).exposeToPublic();
                } else {
                    this._popup.focus();
                }
                Popup.setCurrentPopup(this);
            }
            return new Promise(function (resolve, reject) {
                var interval = setInterval(function () {
                    if (!_this.isPopupOpen()) {
                        window.clearInterval(interval);
                        resolve();
                    }
                }, exports.POPUP_PINGER_INTERVAL);
            });
        };
        Popup.prototype.getPopupWindow = function () {
            return this._popup;
        };
        Popup.prototype.getCurrentUrl = function () {
            return this._url;
        };
        Popup.prototype.isPopupOpen = function () {
            return !!this._popup && !this._popup.closed;
        };
        return Popup;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = Popup;
}(require, exports, require('./ErrorHandler'), require('../models/ErrorType'), require('./Logging'), require('../models/OneDriveSdkError'), require('../utilities/UrlUtilities')));
},{"../models/ErrorType":13,"../models/OneDriveSdkError":16,"../utilities/UrlUtilities":32,"./ErrorHandler":26,"./Logging":27}],30:[function(require,module,exports){
(function (require, exports) {
    'use strict';
    var FORMAT_ARGS_REGEX = /[\{\}]/g;
    var FORMAT_REGEX = /\{\d+\}/g;
    function format(str) {
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
    }
    exports.format = format;
}(require, exports));
},{}],31:[function(require,module,exports){
(function (require, exports, Constants_1, ErrorHandler_1, ErrorType_1, Logging_1, ObjectUtilities_1, OneDriveSdkError_1, StringUtilities_1) {
    'use strict';
    function validateType(object, expectedType, optional, defaultValue, validValues) {
        if (optional === void 0) {
            optional = false;
        }
        if (object === undefined) {
            if (optional) {
                if (defaultValue !== undefined) {
                    Logging_1.default.logMessage('applying default value: ' + defaultValue);
                    return defaultValue;
                }
            } else {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.typeError, 'object was missing and not optional'));
            }
            return null;
        }
        var objectType = typeof object;
        if (objectType !== expectedType) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.typeError, StringUtilities_1.format('expected object type: \'{0}\', actual object type: \'{1}\'', expectedType, objectType)));
            return null;
        }
        if (!isValidValue(object, validValues)) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.typeError, StringUtilities_1.format('object does not match any valid values {0}', ObjectUtilities_1.serializeJSON(validValues))));
            return null;
        }
        return object;
    }
    exports.validateType = validateType;
    function validateCallback(functionOption, optional) {
        if (optional === void 0) {
            optional = false;
        }
        if (functionOption === undefined) {
            if (!optional) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.typeError, 'function was missing and not optional'));
            }
            return null;
        }
        var functionType = typeof functionOption;
        if (functionType !== Constants_1.default.TYPE_FUNCTION) {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.typeError, StringUtilities_1.format('expected function type: function | string, actual type: {0}', functionType)));
        }
        return functionOption;
    }
    exports.validateCallback = validateCallback;
    function isValidValue(object, validValues) {
        if (!Array.isArray(validValues)) {
            return true;
        }
        for (var index in validValues) {
            if (object === validValues[index]) {
                return true;
            }
        }
        return false;
    }
}(require, exports, require('../Constants'), require('./ErrorHandler'), require('../models/ErrorType'), require('./Logging'), require('./ObjectUtilities'), require('../models/OneDriveSdkError'), require('./StringUtilities')));
},{"../Constants":1,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"./ErrorHandler":26,"./Logging":27,"./ObjectUtilities":28,"./StringUtilities":30}],32:[function(require,module,exports){
(function (require, exports, Constants_1, ErrorHandler_1, ErrorType_1, OneDriveSdkError_1, StringUtilities_1) {
    'use strict';
    var PROTOCOL_HTTP = 'HTTP';
    var PROTOCOL_HTTPS = 'HTTPS';
    function appendToPath(baseUrl, path) {
        return baseUrl + (baseUrl.charAt(baseUrl.length - 1) !== '/' ? '/' : '') + path;
    }
    exports.appendToPath = appendToPath;
    function appendQueryString(baseUrl, queryKey, queryValue) {
        return appendQueryStrings(baseUrl, (_a = {}, _a[queryKey] = queryValue, _a));
        var _a;
    }
    exports.appendQueryString = appendQueryString;
    function appendQueryStrings(baseUrl, queryParameters, isAspx) {
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
            queryString += (queryString.length ? '&' : '') + StringUtilities_1.format('{0}={1}', encodeURIComponent(key), encodeURIComponent(queryParameters[key]));
        }
        return baseUrl + queryString;
    }
    exports.appendQueryStrings = appendQueryStrings;
    function readCurrentUrlParameters() {
        return readUrlParameters(window.location.href);
    }
    exports.readCurrentUrlParameters = readCurrentUrlParameters;
    function readUrlParameters(url) {
        var queryParamters = {};
        var queryStart = url.indexOf('?') + 1;
        var hashStart = url.indexOf('#') + 1;
        if (queryStart > 0) {
            var queryEnd = hashStart > queryStart ? hashStart - 1 : url.length;
            deserializeParameters(url.substring(queryStart, queryEnd), queryParamters);
        }
        if (hashStart > 0) {
            deserializeParameters(url.substring(hashStart), queryParamters);
        }
        return queryParamters;
    }
    exports.readUrlParameters = readUrlParameters;
    function redirect(url) {
        validateUrlProtocol(url);
        window.location.replace(url);
    }
    exports.redirect = redirect;
    function trimUrlQuery(url) {
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
    }
    exports.trimUrlQuery = trimUrlQuery;
    function getFileNameFromUrl(url) {
        var trimmedUrl = trimUrlQuery(url);
        return trimmedUrl.substr(trimmedUrl.lastIndexOf('/') + 1);
    }
    exports.getFileNameFromUrl = getFileNameFromUrl;
    function getOrigin(url) {
        return appendToPath(url.replace(/^((\w+:)?\/\/[^\/]+\/?).*$/, '$1'), '');
    }
    exports.getOrigin = getOrigin;
    function isPathFullUrl(path) {
        return path.indexOf('https://') === 0 || path.indexOf('http://') === 0;
    }
    exports.isPathFullUrl = isPathFullUrl;
    function isPathDataUrl(path) {
        return path.indexOf('data:') === 0;
    }
    exports.isPathDataUrl = isPathDataUrl;
    function generateNonce() {
        var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
        var text = '';
        for (var i = 0; i < Constants_1.default.NONCE_LENGTH; i++) {
            text += possible.charAt(Math.floor(Math.random() * possible.length));
        }
        return text;
    }
    exports.generateNonce = generateNonce;
    function validateUrlProtocol(url, protocols) {
        protocols = protocols ? protocols : [
            PROTOCOL_HTTP,
            PROTOCOL_HTTPS
        ];
        for (var _i = 0, protocols_1 = protocols; _i < protocols_1.length; _i++) {
            var protocol = protocols_1[_i];
            if (url.toUpperCase().indexOf(protocol) === 0) {
                return;
            }
        }
        ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, StringUtilities_1.format('uri {0} does not match protocol(s): ' + protocols, url))).exposeToPublic();
    }
    exports.validateUrlProtocol = validateUrlProtocol;
    function validateRedirectUrlHost(url) {
        validateUrlProtocol(url);
        if (url.indexOf('://') > -1) {
            var domain = url.split('/')[2];
            if (domain !== window.location.host) {
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'redirect uri is not in the same domain as picker sdk')).exposeToPublic();
            }
        } else {
            ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.optionsError, 'redirect uri is not an absolute url')).exposeToPublic();
        }
    }
    exports.validateRedirectUrlHost = validateRedirectUrlHost;
    function deserializeParameters(query, queryParameters) {
        var properties = query.split('&');
        for (var i = 0; i < properties.length; i++) {
            var property = properties[i].split('=');
            if (property.length === 2) {
                queryParameters[decodeURIComponent(property[0])] = decodeURIComponent(property[1]);
            }
        }
    }
}(require, exports, require('../Constants'), require('./ErrorHandler'), require('../models/ErrorType'), require('../models/OneDriveSdkError'), require('./StringUtilities')));
},{"../Constants":1,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"./ErrorHandler":26,"./StringUtilities":30}],33:[function(require,module,exports){
(function (require, exports, ApiEndpoint_1, Constants_1, ErrorHandler_1, ErrorType_1, Logging_1, OneDriveSdkError_1, StringUtilities_1) {
    'use strict';
    var REQUEST_TIMEOUT = 30000;
    var EXCEPTION_STATUS = -1;
    var TIMEOUT_STATUS = -2;
    var ABORT_STATUS = -3;
    var LEGACY_MSA_APPID_PATTERN = new RegExp('^([a-fA-F0-9]){16}$');
    var XHR = function () {
        function XHR(options) {
            this._url = options.url;
            this._json = options.json;
            this._headers = options.headers || {};
            this._method = options.method;
            this._clientId = options.clientId;
            this._apiEndpoint = options.apiEndpoint || ApiEndpoint_1.default.msa;
            ErrorHandler_1.registerErrorObserver(this._abortRequest);
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
                this._request.onload = function () {
                    var status = _this._request.status;
                    if (status < 400 && status > 0) {
                        _this._callSuccessCallback(status);
                    } else {
                        _this._callFailureCallback(status);
                    }
                };
                if (!this._method) {
                    this._method = this._json ? XHR.HTTP_POST : XHR.HTTP_GET;
                }
                this._request.open(this._method, this._url, true);
                this._request.timeout = REQUEST_TIMEOUT;
                this._setHeaders();
                Logging_1.default.logMessage('starting request to: ' + this._url);
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
                this._method = XHR.HTTP_PUT;
                this._request.open(this._method, this._url, true);
                this._setHeaders();
                this._request.onload = function (event) {
                    _this._completed = true;
                    var status = _this._request.status;
                    if (status === 200 || status === 201 || status === 202) {
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
                Logging_1.default.logMessage('starting upload to: ' + this._url);
                this._request.send(data);
            } catch (error) {
                this._callFailureCallback(EXCEPTION_STATUS, error);
            }
        };
        XHR.prototype._callSuccessCallback = function (status) {
            Logging_1.default.logMessage('calling xhr success callback, status: ' + XHR.statusCodeToString(status));
            this._successCallback(this._request, status, this._url);
        };
        XHR.prototype._callFailureCallback = function (status, error) {
            Logging_1.default.logError('calling xhr failure callback, status: ' + XHR.statusCodeToString(status), this._request, error);
            this._failureCallback(this._request, status, status === TIMEOUT_STATUS);
        };
        XHR.prototype._callProgressCallback = function (uploadProgress) {
            Logging_1.default.logMessage('calling xhr upload progress callback');
            this._progressCallback(this._request, uploadProgress);
        };
        XHR.prototype._abortRequest = function () {
            if (this && !this._completed) {
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
            if (this._clientId) {
                var idToRecord = this._clientId;
                if (LEGACY_MSA_APPID_PATTERN.test(this._clientId)) {
                    idToRecord = '0x' + this._clientId;
                }
                this._request.setRequestHeader('Application', '0x' + idToRecord);
            }
            var sdkVersion = StringUtilities_1.format('{0}={1}', 'SDK-Version', Constants_1.default.SDK_VERSION);
            switch (this._apiEndpoint) {
            case ApiEndpoint_1.default.graph_odb:
            case ApiEndpoint_1.default.filesV2:
                this._request.setRequestHeader('X-ClientService-ClientTag', sdkVersion);
                break;
            case ApiEndpoint_1.default.graph_odc:
            case ApiEndpoint_1.default.msa:
                this._request.setRequestHeader('X-RequestStats', sdkVersion);
                break;
            default:
                ErrorHandler_1.throwError(new OneDriveSdkError_1.default(ErrorType_1.default.internalError, 'invalid API endpoint: ' + this._apiEndpoint));
            }
            if (this._method === XHR.HTTP_POST) {
                this._request.setRequestHeader('Content-Type', this._json ? 'application/json' : 'text/plain');
            }
        };
        XHR.HTTP_GET = 'GET';
        XHR.HTTP_POST = 'POST';
        XHR.HTTP_PUT = 'PUT';
        return XHR;
    }();
    Object.defineProperty(exports, '__esModule', { value: true });
    exports.default = XHR;
}(require, exports, require('../models/ApiEndpoint'), require('../Constants'), require('./ErrorHandler'), require('../models/ErrorType'), require('./Logging'), require('../models/OneDriveSdkError'), require('./StringUtilities')));
},{"../Constants":1,"../models/ApiEndpoint":11,"../models/ErrorType":13,"../models/OneDriveSdkError":16,"./ErrorHandler":26,"./Logging":27,"./StringUtilities":30}],34:[function(require,module,exports){
/*!
 * @overview es6-promise - a tiny implementation of Promises/A+.
 * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
 * @license   Licensed under MIT license
 *            See https://raw.githubusercontent.com/stefanpenner/es6-promise/master/LICENSE
 * @version   3.3.1
 */

(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) :
    (global.ES6Promise = factory());
}(this, (function () { 'use strict';

function objectOrFunction(x) {
  return typeof x === 'function' || typeof x === 'object' && x !== null;
}

function isFunction(x) {
  return typeof x === 'function';
}

var _isArray = undefined;
if (!Array.isArray) {
  _isArray = function (x) {
    return Object.prototype.toString.call(x) === '[object Array]';
  };
} else {
  _isArray = Array.isArray;
}

var isArray = _isArray;

var len = 0;
var vertxNext = undefined;
var customSchedulerFn = undefined;

var asap = function asap(callback, arg) {
  queue[len] = callback;
  queue[len + 1] = arg;
  len += 2;
  if (len === 2) {
    // If len is 2, that means that we need to schedule an async flush.
    // If additional callbacks are queued before the queue is flushed, they
    // will be processed by this flush that we are scheduling.
    if (customSchedulerFn) {
      customSchedulerFn(flush);
    } else {
      scheduleFlush();
    }
  }
};

function setScheduler(scheduleFn) {
  customSchedulerFn = scheduleFn;
}

function setAsap(asapFn) {
  asap = asapFn;
}

var browserWindow = typeof window !== 'undefined' ? window : undefined;
var browserGlobal = browserWindow || {};
var BrowserMutationObserver = browserGlobal.MutationObserver || browserGlobal.WebKitMutationObserver;
var isNode = typeof self === 'undefined' && typeof process !== 'undefined' && ({}).toString.call(process) === '[object process]';

// test for web worker but not in IE10
var isWorker = typeof Uint8ClampedArray !== 'undefined' && typeof importScripts !== 'undefined' && typeof MessageChannel !== 'undefined';

// node
function useNextTick() {
  // node version 0.10.x displays a deprecation warning when nextTick is used recursively
  // see https://github.com/cujojs/when/issues/410 for details
  return function () {
    return process.nextTick(flush);
  };
}

// vertx
function useVertxTimer() {
  return function () {
    vertxNext(flush);
  };
}

function useMutationObserver() {
  var iterations = 0;
  var observer = new BrowserMutationObserver(flush);
  var node = document.createTextNode('');
  observer.observe(node, { characterData: true });

  return function () {
    node.data = iterations = ++iterations % 2;
  };
}

// web worker
function useMessageChannel() {
  var channel = new MessageChannel();
  channel.port1.onmessage = flush;
  return function () {
    return channel.port2.postMessage(0);
  };
}

function useSetTimeout() {
  // Store setTimeout reference so es6-promise will be unaffected by
  // other code modifying setTimeout (like sinon.useFakeTimers())
  var globalSetTimeout = setTimeout;
  return function () {
    return globalSetTimeout(flush, 1);
  };
}

var queue = new Array(1000);
function flush() {
  for (var i = 0; i < len; i += 2) {
    var callback = queue[i];
    var arg = queue[i + 1];

    callback(arg);

    queue[i] = undefined;
    queue[i + 1] = undefined;
  }

  len = 0;
}

function attemptVertx() {
  try {
    var r = require;
    var vertx = r('vertx');
    vertxNext = vertx.runOnLoop || vertx.runOnContext;
    return useVertxTimer();
  } catch (e) {
    return useSetTimeout();
  }
}

var scheduleFlush = undefined;
// Decide what async method to use to triggering processing of queued callbacks:
if (isNode) {
  scheduleFlush = useNextTick();
} else if (BrowserMutationObserver) {
  scheduleFlush = useMutationObserver();
} else if (isWorker) {
  scheduleFlush = useMessageChannel();
} else if (browserWindow === undefined && typeof require === 'function') {
  scheduleFlush = attemptVertx();
} else {
  scheduleFlush = useSetTimeout();
}

function then(onFulfillment, onRejection) {
  var _arguments = arguments;

  var parent = this;

  var child = new this.constructor(noop);

  if (child[PROMISE_ID] === undefined) {
    makePromise(child);
  }

  var _state = parent._state;

  if (_state) {
    (function () {
      var callback = _arguments[_state - 1];
      asap(function () {
        return invokeCallback(_state, child, callback, parent._result);
      });
    })();
  } else {
    subscribe(parent, child, onFulfillment, onRejection);
  }

  return child;
}

/**
  `Promise.resolve` returns a promise that will become resolved with the
  passed `value`. It is shorthand for the following:

  ```javascript
  let promise = new Promise(function(resolve, reject){
    resolve(1);
  });

  promise.then(function(value){
    // value === 1
  });
  ```

  Instead of writing the above, your code now simply becomes the following:

  ```javascript
  let promise = Promise.resolve(1);

  promise.then(function(value){
    // value === 1
  });
  ```

  @method resolve
  @static
  @param {Any} value value that the returned promise will be resolved with
  Useful for tooling.
  @return {Promise} a promise that will become fulfilled with the given
  `value`
*/
function resolve(object) {
  /*jshint validthis:true */
  var Constructor = this;

  if (object && typeof object === 'object' && object.constructor === Constructor) {
    return object;
  }

  var promise = new Constructor(noop);
  _resolve(promise, object);
  return promise;
}

var PROMISE_ID = Math.random().toString(36).substring(16);

function noop() {}

var PENDING = void 0;
var FULFILLED = 1;
var REJECTED = 2;

var GET_THEN_ERROR = new ErrorObject();

function selfFulfillment() {
  return new TypeError("You cannot resolve a promise with itself");
}

function cannotReturnOwn() {
  return new TypeError('A promises callback cannot return that same promise.');
}

function getThen(promise) {
  try {
    return promise.then;
  } catch (error) {
    GET_THEN_ERROR.error = error;
    return GET_THEN_ERROR;
  }
}

function tryThen(then, value, fulfillmentHandler, rejectionHandler) {
  try {
    then.call(value, fulfillmentHandler, rejectionHandler);
  } catch (e) {
    return e;
  }
}

function handleForeignThenable(promise, thenable, then) {
  asap(function (promise) {
    var sealed = false;
    var error = tryThen(then, thenable, function (value) {
      if (sealed) {
        return;
      }
      sealed = true;
      if (thenable !== value) {
        _resolve(promise, value);
      } else {
        fulfill(promise, value);
      }
    }, function (reason) {
      if (sealed) {
        return;
      }
      sealed = true;

      _reject(promise, reason);
    }, 'Settle: ' + (promise._label || ' unknown promise'));

    if (!sealed && error) {
      sealed = true;
      _reject(promise, error);
    }
  }, promise);
}

function handleOwnThenable(promise, thenable) {
  if (thenable._state === FULFILLED) {
    fulfill(promise, thenable._result);
  } else if (thenable._state === REJECTED) {
    _reject(promise, thenable._result);
  } else {
    subscribe(thenable, undefined, function (value) {
      return _resolve(promise, value);
    }, function (reason) {
      return _reject(promise, reason);
    });
  }
}

function handleMaybeThenable(promise, maybeThenable, then$$) {
  if (maybeThenable.constructor === promise.constructor && then$$ === then && maybeThenable.constructor.resolve === resolve) {
    handleOwnThenable(promise, maybeThenable);
  } else {
    if (then$$ === GET_THEN_ERROR) {
      _reject(promise, GET_THEN_ERROR.error);
    } else if (then$$ === undefined) {
      fulfill(promise, maybeThenable);
    } else if (isFunction(then$$)) {
      handleForeignThenable(promise, maybeThenable, then$$);
    } else {
      fulfill(promise, maybeThenable);
    }
  }
}

function _resolve(promise, value) {
  if (promise === value) {
    _reject(promise, selfFulfillment());
  } else if (objectOrFunction(value)) {
    handleMaybeThenable(promise, value, getThen(value));
  } else {
    fulfill(promise, value);
  }
}

function publishRejection(promise) {
  if (promise._onerror) {
    promise._onerror(promise._result);
  }

  publish(promise);
}

function fulfill(promise, value) {
  if (promise._state !== PENDING) {
    return;
  }

  promise._result = value;
  promise._state = FULFILLED;

  if (promise._subscribers.length !== 0) {
    asap(publish, promise);
  }
}

function _reject(promise, reason) {
  if (promise._state !== PENDING) {
    return;
  }
  promise._state = REJECTED;
  promise._result = reason;

  asap(publishRejection, promise);
}

function subscribe(parent, child, onFulfillment, onRejection) {
  var _subscribers = parent._subscribers;
  var length = _subscribers.length;

  parent._onerror = null;

  _subscribers[length] = child;
  _subscribers[length + FULFILLED] = onFulfillment;
  _subscribers[length + REJECTED] = onRejection;

  if (length === 0 && parent._state) {
    asap(publish, parent);
  }
}

function publish(promise) {
  var subscribers = promise._subscribers;
  var settled = promise._state;

  if (subscribers.length === 0) {
    return;
  }

  var child = undefined,
      callback = undefined,
      detail = promise._result;

  for (var i = 0; i < subscribers.length; i += 3) {
    child = subscribers[i];
    callback = subscribers[i + settled];

    if (child) {
      invokeCallback(settled, child, callback, detail);
    } else {
      callback(detail);
    }
  }

  promise._subscribers.length = 0;
}

function ErrorObject() {
  this.error = null;
}

var TRY_CATCH_ERROR = new ErrorObject();

function tryCatch(callback, detail) {
  try {
    return callback(detail);
  } catch (e) {
    TRY_CATCH_ERROR.error = e;
    return TRY_CATCH_ERROR;
  }
}

function invokeCallback(settled, promise, callback, detail) {
  var hasCallback = isFunction(callback),
      value = undefined,
      error = undefined,
      succeeded = undefined,
      failed = undefined;

  if (hasCallback) {
    value = tryCatch(callback, detail);

    if (value === TRY_CATCH_ERROR) {
      failed = true;
      error = value.error;
      value = null;
    } else {
      succeeded = true;
    }

    if (promise === value) {
      _reject(promise, cannotReturnOwn());
      return;
    }
  } else {
    value = detail;
    succeeded = true;
  }

  if (promise._state !== PENDING) {
    // noop
  } else if (hasCallback && succeeded) {
      _resolve(promise, value);
    } else if (failed) {
      _reject(promise, error);
    } else if (settled === FULFILLED) {
      fulfill(promise, value);
    } else if (settled === REJECTED) {
      _reject(promise, value);
    }
}

function initializePromise(promise, resolver) {
  try {
    resolver(function resolvePromise(value) {
      _resolve(promise, value);
    }, function rejectPromise(reason) {
      _reject(promise, reason);
    });
  } catch (e) {
    _reject(promise, e);
  }
}

var id = 0;
function nextId() {
  return id++;
}

function makePromise(promise) {
  promise[PROMISE_ID] = id++;
  promise._state = undefined;
  promise._result = undefined;
  promise._subscribers = [];
}

function Enumerator(Constructor, input) {
  this._instanceConstructor = Constructor;
  this.promise = new Constructor(noop);

  if (!this.promise[PROMISE_ID]) {
    makePromise(this.promise);
  }

  if (isArray(input)) {
    this._input = input;
    this.length = input.length;
    this._remaining = input.length;

    this._result = new Array(this.length);

    if (this.length === 0) {
      fulfill(this.promise, this._result);
    } else {
      this.length = this.length || 0;
      this._enumerate();
      if (this._remaining === 0) {
        fulfill(this.promise, this._result);
      }
    }
  } else {
    _reject(this.promise, validationError());
  }
}

function validationError() {
  return new Error('Array Methods must be provided an Array');
};

Enumerator.prototype._enumerate = function () {
  var length = this.length;
  var _input = this._input;

  for (var i = 0; this._state === PENDING && i < length; i++) {
    this._eachEntry(_input[i], i);
  }
};

Enumerator.prototype._eachEntry = function (entry, i) {
  var c = this._instanceConstructor;
  var resolve$$ = c.resolve;

  if (resolve$$ === resolve) {
    var _then = getThen(entry);

    if (_then === then && entry._state !== PENDING) {
      this._settledAt(entry._state, i, entry._result);
    } else if (typeof _then !== 'function') {
      this._remaining--;
      this._result[i] = entry;
    } else if (c === Promise) {
      var promise = new c(noop);
      handleMaybeThenable(promise, entry, _then);
      this._willSettleAt(promise, i);
    } else {
      this._willSettleAt(new c(function (resolve$$) {
        return resolve$$(entry);
      }), i);
    }
  } else {
    this._willSettleAt(resolve$$(entry), i);
  }
};

Enumerator.prototype._settledAt = function (state, i, value) {
  var promise = this.promise;

  if (promise._state === PENDING) {
    this._remaining--;

    if (state === REJECTED) {
      _reject(promise, value);
    } else {
      this._result[i] = value;
    }
  }

  if (this._remaining === 0) {
    fulfill(promise, this._result);
  }
};

Enumerator.prototype._willSettleAt = function (promise, i) {
  var enumerator = this;

  subscribe(promise, undefined, function (value) {
    return enumerator._settledAt(FULFILLED, i, value);
  }, function (reason) {
    return enumerator._settledAt(REJECTED, i, reason);
  });
};

/**
  `Promise.all` accepts an array of promises, and returns a new promise which
  is fulfilled with an array of fulfillment values for the passed promises, or
  rejected with the reason of the first passed promise to be rejected. It casts all
  elements of the passed iterable to promises as it runs this algorithm.

  Example:

  ```javascript
  let promise1 = resolve(1);
  let promise2 = resolve(2);
  let promise3 = resolve(3);
  let promises = [ promise1, promise2, promise3 ];

  Promise.all(promises).then(function(array){
    // The array here would be [ 1, 2, 3 ];
  });
  ```

  If any of the `promises` given to `all` are rejected, the first promise
  that is rejected will be given as an argument to the returned promises's
  rejection handler. For example:

  Example:

  ```javascript
  let promise1 = resolve(1);
  let promise2 = reject(new Error("2"));
  let promise3 = reject(new Error("3"));
  let promises = [ promise1, promise2, promise3 ];

  Promise.all(promises).then(function(array){
    // Code here never runs because there are rejected promises!
  }, function(error) {
    // error.message === "2"
  });
  ```

  @method all
  @static
  @param {Array} entries array of promises
  @param {String} label optional string for labeling the promise.
  Useful for tooling.
  @return {Promise} promise that is fulfilled when all `promises` have been
  fulfilled, or rejected if any of them become rejected.
  @static
*/
function all(entries) {
  return new Enumerator(this, entries).promise;
}

/**
  `Promise.race` returns a new promise which is settled in the same way as the
  first passed promise to settle.

  Example:

  ```javascript
  let promise1 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 1');
    }, 200);
  });

  let promise2 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 2');
    }, 100);
  });

  Promise.race([promise1, promise2]).then(function(result){
    // result === 'promise 2' because it was resolved before promise1
    // was resolved.
  });
  ```

  `Promise.race` is deterministic in that only the state of the first
  settled promise matters. For example, even if other promises given to the
  `promises` array argument are resolved, but the first settled promise has
  become rejected before the other promises became fulfilled, the returned
  promise will become rejected:

  ```javascript
  let promise1 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 1');
    }, 200);
  });

  let promise2 = new Promise(function(resolve, reject){
    setTimeout(function(){
      reject(new Error('promise 2'));
    }, 100);
  });

  Promise.race([promise1, promise2]).then(function(result){
    // Code here never runs
  }, function(reason){
    // reason.message === 'promise 2' because promise 2 became rejected before
    // promise 1 became fulfilled
  });
  ```

  An example real-world use case is implementing timeouts:

  ```javascript
  Promise.race([ajax('foo.json'), timeout(5000)])
  ```

  @method race
  @static
  @param {Array} promises array of promises to observe
  Useful for tooling.
  @return {Promise} a promise which settles in the same way as the first passed
  promise to settle.
*/
function race(entries) {
  /*jshint validthis:true */
  var Constructor = this;

  if (!isArray(entries)) {
    return new Constructor(function (_, reject) {
      return reject(new TypeError('You must pass an array to race.'));
    });
  } else {
    return new Constructor(function (resolve, reject) {
      var length = entries.length;
      for (var i = 0; i < length; i++) {
        Constructor.resolve(entries[i]).then(resolve, reject);
      }
    });
  }
}

/**
  `Promise.reject` returns a promise rejected with the passed `reason`.
  It is shorthand for the following:

  ```javascript
  let promise = new Promise(function(resolve, reject){
    reject(new Error('WHOOPS'));
  });

  promise.then(function(value){
    // Code here doesn't run because the promise is rejected!
  }, function(reason){
    // reason.message === 'WHOOPS'
  });
  ```

  Instead of writing the above, your code now simply becomes the following:

  ```javascript
  let promise = Promise.reject(new Error('WHOOPS'));

  promise.then(function(value){
    // Code here doesn't run because the promise is rejected!
  }, function(reason){
    // reason.message === 'WHOOPS'
  });
  ```

  @method reject
  @static
  @param {Any} reason value that the returned promise will be rejected with.
  Useful for tooling.
  @return {Promise} a promise rejected with the given `reason`.
*/
function reject(reason) {
  /*jshint validthis:true */
  var Constructor = this;
  var promise = new Constructor(noop);
  _reject(promise, reason);
  return promise;
}

function needsResolver() {
  throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
}

function needsNew() {
  throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
}

/**
  Promise objects represent the eventual result of an asynchronous operation. The
  primary way of interacting with a promise is through its `then` method, which
  registers callbacks to receive either a promise's eventual value or the reason
  why the promise cannot be fulfilled.

  Terminology
  -----------

  - `promise` is an object or function with a `then` method whose behavior conforms to this specification.
  - `thenable` is an object or function that defines a `then` method.
  - `value` is any legal JavaScript value (including undefined, a thenable, or a promise).
  - `exception` is a value that is thrown using the throw statement.
  - `reason` is a value that indicates why a promise was rejected.
  - `settled` the final resting state of a promise, fulfilled or rejected.

  A promise can be in one of three states: pending, fulfilled, or rejected.

  Promises that are fulfilled have a fulfillment value and are in the fulfilled
  state.  Promises that are rejected have a rejection reason and are in the
  rejected state.  A fulfillment value is never a thenable.

  Promises can also be said to *resolve* a value.  If this value is also a
  promise, then the original promise's settled state will match the value's
  settled state.  So a promise that *resolves* a promise that rejects will
  itself reject, and a promise that *resolves* a promise that fulfills will
  itself fulfill.


  Basic Usage:
  ------------

  ```js
  let promise = new Promise(function(resolve, reject) {
    // on success
    resolve(value);

    // on failure
    reject(reason);
  });

  promise.then(function(value) {
    // on fulfillment
  }, function(reason) {
    // on rejection
  });
  ```

  Advanced Usage:
  ---------------

  Promises shine when abstracting away asynchronous interactions such as
  `XMLHttpRequest`s.

  ```js
  function getJSON(url) {
    return new Promise(function(resolve, reject){
      let xhr = new XMLHttpRequest();

      xhr.open('GET', url);
      xhr.onreadystatechange = handler;
      xhr.responseType = 'json';
      xhr.setRequestHeader('Accept', 'application/json');
      xhr.send();

      function handler() {
        if (this.readyState === this.DONE) {
          if (this.status === 200) {
            resolve(this.response);
          } else {
            reject(new Error('getJSON: `' + url + '` failed with status: [' + this.status + ']'));
          }
        }
      };
    });
  }

  getJSON('/posts.json').then(function(json) {
    // on fulfillment
  }, function(reason) {
    // on rejection
  });
  ```

  Unlike callbacks, promises are great composable primitives.

  ```js
  Promise.all([
    getJSON('/posts'),
    getJSON('/comments')
  ]).then(function(values){
    values[0] // => postsJSON
    values[1] // => commentsJSON

    return values;
  });
  ```

  @class Promise
  @param {function} resolver
  Useful for tooling.
  @constructor
*/
function Promise(resolver) {
  this[PROMISE_ID] = nextId();
  this._result = this._state = undefined;
  this._subscribers = [];

  if (noop !== resolver) {
    typeof resolver !== 'function' && needsResolver();
    this instanceof Promise ? initializePromise(this, resolver) : needsNew();
  }
}

Promise.all = all;
Promise.race = race;
Promise.resolve = resolve;
Promise.reject = reject;
Promise._setScheduler = setScheduler;
Promise._setAsap = setAsap;
Promise._asap = asap;

Promise.prototype = {
  constructor: Promise,

  /**
    The primary way of interacting with a promise is through its `then` method,
    which registers callbacks to receive either a promise's eventual value or the
    reason why the promise cannot be fulfilled.
  
    ```js
    findUser().then(function(user){
      // user is available
    }, function(reason){
      // user is unavailable, and you are given the reason why
    });
    ```
  
    Chaining
    --------
  
    The return value of `then` is itself a promise.  This second, 'downstream'
    promise is resolved with the return value of the first promise's fulfillment
    or rejection handler, or rejected if the handler throws an exception.
  
    ```js
    findUser().then(function (user) {
      return user.name;
    }, function (reason) {
      return 'default name';
    }).then(function (userName) {
      // If `findUser` fulfilled, `userName` will be the user's name, otherwise it
      // will be `'default name'`
    });
  
    findUser().then(function (user) {
      throw new Error('Found user, but still unhappy');
    }, function (reason) {
      throw new Error('`findUser` rejected and we're unhappy');
    }).then(function (value) {
      // never reached
    }, function (reason) {
      // if `findUser` fulfilled, `reason` will be 'Found user, but still unhappy'.
      // If `findUser` rejected, `reason` will be '`findUser` rejected and we're unhappy'.
    });
    ```
    If the downstream promise does not specify a rejection handler, rejection reasons will be propagated further downstream.
  
    ```js
    findUser().then(function (user) {
      throw new PedagogicalException('Upstream error');
    }).then(function (value) {
      // never reached
    }).then(function (value) {
      // never reached
    }, function (reason) {
      // The `PedgagocialException` is propagated all the way down to here
    });
    ```
  
    Assimilation
    ------------
  
    Sometimes the value you want to propagate to a downstream promise can only be
    retrieved asynchronously. This can be achieved by returning a promise in the
    fulfillment or rejection handler. The downstream promise will then be pending
    until the returned promise is settled. This is called *assimilation*.
  
    ```js
    findUser().then(function (user) {
      return findCommentsByAuthor(user);
    }).then(function (comments) {
      // The user's comments are now available
    });
    ```
  
    If the assimliated promise rejects, then the downstream promise will also reject.
  
    ```js
    findUser().then(function (user) {
      return findCommentsByAuthor(user);
    }).then(function (comments) {
      // If `findCommentsByAuthor` fulfills, we'll have the value here
    }, function (reason) {
      // If `findCommentsByAuthor` rejects, we'll have the reason here
    });
    ```
  
    Simple Example
    --------------
  
    Synchronous Example
  
    ```javascript
    let result;
  
    try {
      result = findResult();
      // success
    } catch(reason) {
      // failure
    }
    ```
  
    Errback Example
  
    ```js
    findResult(function(result, err){
      if (err) {
        // failure
      } else {
        // success
      }
    });
    ```
  
    Promise Example;
  
    ```javascript
    findResult().then(function(result){
      // success
    }, function(reason){
      // failure
    });
    ```
  
    Advanced Example
    --------------
  
    Synchronous Example
  
    ```javascript
    let author, books;
  
    try {
      author = findAuthor();
      books  = findBooksByAuthor(author);
      // success
    } catch(reason) {
      // failure
    }
    ```
  
    Errback Example
  
    ```js
  
    function foundBooks(books) {
  
    }
  
    function failure(reason) {
  
    }
  
    findAuthor(function(author, err){
      if (err) {
        failure(err);
        // failure
      } else {
        try {
          findBoooksByAuthor(author, function(books, err) {
            if (err) {
              failure(err);
            } else {
              try {
                foundBooks(books);
              } catch(reason) {
                failure(reason);
              }
            }
          });
        } catch(error) {
          failure(err);
        }
        // success
      }
    });
    ```
  
    Promise Example;
  
    ```javascript
    findAuthor().
      then(findBooksByAuthor).
      then(function(books){
        // found books
    }).catch(function(reason){
      // something went wrong
    });
    ```
  
    @method then
    @param {Function} onFulfilled
    @param {Function} onRejected
    Useful for tooling.
    @return {Promise}
  */
  then: then,

  /**
    `catch` is simply sugar for `then(undefined, onRejection)` which makes it the same
    as the catch block of a try/catch statement.
  
    ```js
    function findAuthor(){
      throw new Error('couldn't find that author');
    }
  
    // synchronous
    try {
      findAuthor();
    } catch(reason) {
      // something went wrong
    }
  
    // async with promises
    findAuthor().catch(function(reason){
      // something went wrong
    });
    ```
  
    @method catch
    @param {Function} onRejection
    Useful for tooling.
    @return {Promise}
  */
  'catch': function _catch(onRejection) {
    return this.then(null, onRejection);
  }
};

function polyfill() {
    var local = undefined;

    if (typeof global !== 'undefined') {
        local = global;
    } else if (typeof self !== 'undefined') {
        local = self;
    } else {
        try {
            local = Function('return this')();
        } catch (e) {
            throw new Error('polyfill failed because global object is unavailable in this environment');
        }
    }

    var P = local.Promise;

    if (P) {
        var promiseToString = null;
        try {
            promiseToString = Object.prototype.toString.call(P.resolve());
        } catch (e) {
            // silently ignored
        }

        if (promiseToString === '[object Promise]' && !P.cast) {
            return;
        }
    }

    local.Promise = Promise;
}

polyfill();
// Strange compat..
Promise.polyfill = polyfill;
Promise.Promise = Promise;

return Promise;

})));

},{}]},{},[2])(2)
});