/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/


// Sources:
// osfweb: 16.0\16.0.15928.10000
// runtime: 16.0\16.0.15928.10000
// core: 16.0\16.0.15928.10000
// host: 16.0\16.0.15928.10000



var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b)
                if (b.hasOwnProperty(p))
                    d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OfficeExt;
(function (OfficeExt) {
    var MicrosoftAjaxFactory = (function () {
        function MicrosoftAjaxFactory() {
        }
        MicrosoftAjaxFactory.prototype.isMsAjaxLoaded = function () {
            if (typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function" &&
                typeof (Function._validateParams) === "function" &&
                Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize) === "function") {
                return true;
            }
            else {
                return false;
            }
        };
        MicrosoftAjaxFactory.prototype.loadMsAjaxFull = function (callback) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            OSF.OUtil.loadScript(msAjaxCDNPath, callback);
        };
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
            get: function () {
                if (this._msAjaxError == null && this.isMsAjaxLoaded()) {
                    this._msAjaxError = Error;
                }
                return this._msAjaxError;
            },
            set: function (errorClass) {
                this._msAjaxError = errorClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
            get: function () {
                if (this._msAjaxString == null && this.isMsAjaxLoaded()) {
                    this._msAjaxString = String;
                }
                return this._msAjaxString;
            },
            set: function (stringClass) {
                this._msAjaxString = stringClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
            get: function () {
                if (this._msAjaxDebug == null && this.isMsAjaxLoaded()) {
                    this._msAjaxDebug = Sys.Debug;
                }
                return this._msAjaxDebug;
            },
            set: function (debugClass) {
                this._msAjaxDebug = debugClass;
            },
            enumerable: true,
            configurable: true
        });
        return MicrosoftAjaxFactory;
    }());
    OfficeExt.MicrosoftAjaxFactory = MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory();
var OSF = OSF || {};
(function (OfficeExt) {
    var SafeStorage = (function () {
        function SafeStorage(_internalStorage) {
            this._internalStorage = _internalStorage;
        }
        SafeStorage.prototype.getItem = function (key) {
            try {
                return this._internalStorage && this._internalStorage.getItem(key);
            }
            catch (e) {
                return null;
            }
        };
        SafeStorage.prototype.setItem = function (key, data) {
            try {
                this._internalStorage && this._internalStorage.setItem(key, data);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.clear = function () {
            try {
                this._internalStorage && this._internalStorage.clear();
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.removeItem = function (key) {
            try {
                this._internalStorage && this._internalStorage.removeItem(key);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.getKeysWithPrefix = function (keyPrefix) {
            var keyList = [];
            try {
                var len = this._internalStorage && this._internalStorage.length || 0;
                for (var i = 0; i < len; i++) {
                    var key = this._internalStorage.key(i);
                    if (key.indexOf(keyPrefix) === 0) {
                        keyList.push(key);
                    }
                }
            }
            catch (e) {
            }
            return keyList;
        };
        SafeStorage.prototype.isLocalStorageAvailable = function () {
            return (this._internalStorage != null);
        };
        return SafeStorage;
    }());
    OfficeExt.SafeStorage = SafeStorage;
})(OfficeExt || (OfficeExt = {}));
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.TestFlightStart = 1000;
OSF.TestFlightEnd = 1009;
OSF.FlightNames = {
    UseOriginNotUrl: 0,
    AddinEnforceHttps: 2,
    FirstPartyAnonymousProxyReadyCheckTimeout: 6,
    AddinRibbonIdAllowUnknown: 9,
    ManifestParserDevConsoleLog: 15,
    AddinActionDefinitionHybridMode: 18,
    UseActionIdForUILessCommand: 20,
    RequirementSetRibbonApiOnePointTwo: 21,
    SetFocusToTaskpaneIsEnabled: 22,
    ShortcutInfoArrayInUserPreferenceData: 23,
    OSFTestFlight1000: OSF.TestFlightStart,
    OSFTestFlight1001: OSF.TestFlightStart + 1,
    OSFTestFlight1002: OSF.TestFlightStart + 2,
    OSFTestFlight1003: OSF.TestFlightStart + 3,
    OSFTestFlight1004: OSF.TestFlightStart + 4,
    OSFTestFlight1005: OSF.TestFlightStart + 5,
    OSFTestFlight1006: OSF.TestFlightStart + 6,
    OSFTestFlight1007: OSF.TestFlightStart + 7,
    OSFTestFlight1008: OSF.TestFlightStart + 8,
    OSFTestFlight1009: OSF.TestFlightEnd
};
OSF.TrustUXFlightValues = {
    TrustUXControlA: 0,
    TrustUXExperimentB: 1,
    TrustUXExperimentC: 2
};
OSF.FlightTreatmentNames = {
    AddinTrustUXImprovement: "Microsoft.Office.SharedOnline.AddinTrustUXImprovement",
    BlockAutoOpenAddInIfStoreDisabled: "Microsoft.Office.SharedOnline.BlockAutoOpenAddInIfStoreDisabled",
    Bug7083046KillSwitch: "Microsoft.Office.SharedOnline.Bug7083046KillSwitch",
    CheckProxyIsReadyRetry: "Microsoft.Office.SharedOnline.OEP.CheckProxyIsReadyRetry",
    InsertionDialogFixesEnabled: "Microsoft.Office.SharedOnline.InsertionDialogFixesEnabled",
    TeachingUIForPrivateCatelogEnabled: "Microsoft.Office.SharedOnline.TeachingUIForPrivateCatelogEnabled",
    WopiPreinstalledAddInsEnabled: "Microsoft.Office.SharedOnline.WopiPreinstalledAddInsEnabled",
    WopiUseNewActivate: "Microsoft.Office.SharedOnline.WopiUseNewActivate",
    MosManifestEnabled: "Microsoft.Office.SharedOnline.OEP.MosManifest",
    RemoveGetTrustNoPrompt: "Microsoft.Office.SharedOnline.removeGetTrustNoPrompt",
    HostTrustDialog: "Microsoft.Office.SharedOnline.HostTrustDialog",
    WordEditorAddinThemeSupportEnabled: "Microsoft.Office.WordOnline.WordEditorAddinThemeSupportEnabled"
};
OSF.Flights = [];
OSF.IntFlights = {};
OSF.Settings = {};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    SerializerVersion: "serializerVersion",
    AppContext: "appContext",
    Flights: "flights"
};
OSF.OUtil = (function () {
    var _uniqueId = -1;
    var _xdmInfoKey = '&_xdm_Info=';
    var _serializerVersionKey = '&_serializer_version=';
    var _flightsKey = '&_flights=';
    var _xdmSessionKeyPrefix = '_xdm_';
    var _serializerVersionKeyPrefix = '_serializer_version=';
    var _flightsKeyPrefix = '_flights=';
    var _fragmentSeparator = '#';
    var _fragmentInfoDelimiter = '&';
    var _classN = "class";
    var _loadedScripts = {};
    var _defaultScriptLoadingTimeout = 30000;
    var _safeSessionStorage = null;
    var _safeLocalStorage = null;
    var _rndentropy = new Date().getTime();
    function _random() {
        var nextrand = 0x7fffffff * (Math.random());
        nextrand ^= _rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
        return nextrand.toString(16);
    }
    ;
    function _getSessionStorage() {
        if (!_safeSessionStorage) {
            try {
                var sessionStorage = window.sessionStorage;
            }
            catch (ex) {
                sessionStorage = null;
            }
            _safeSessionStorage = new OfficeExt.SafeStorage(sessionStorage);
        }
        return _safeSessionStorage;
    }
    ;
    function _reOrderTabbableElements(elements) {
        var bucket0 = [];
        var bucketPositive = [];
        var i;
        var len = elements.length;
        var ele;
        for (i = 0; i < len; i++) {
            ele = elements[i];
            if (ele.tabIndex) {
                if (ele.tabIndex > 0) {
                    bucketPositive.push(ele);
                }
                else if (ele.tabIndex === 0) {
                    bucket0.push(ele);
                }
            }
            else {
                bucket0.push(ele);
            }
        }
        bucketPositive = bucketPositive.sort(function (left, right) {
            var diff = left.tabIndex - right.tabIndex;
            if (diff === 0) {
                diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
            }
            return diff;
        });
        return [].concat(bucketPositive, bucket0);
    }
    ;
    return {
        set_entropy: function OSF_OUtil$set_entropy(entropy) {
            if (typeof entropy == "string") {
                for (var i = 0; i < entropy.length; i += 4) {
                    var temp = 0;
                    for (var j = 0; j < 4 && i + j < entropy.length; j++) {
                        temp = (temp << 8) + entropy.charCodeAt(i + j);
                    }
                    _rndentropy ^= temp;
                }
            }
            else if (typeof entropy == "number") {
                _rndentropy ^= entropy;
            }
            else {
                _rndentropy ^= 0x7fffffff * Math.random();
            }
            _rndentropy &= 0x7fffffff;
        },
        extend: function OSF_OUtil$extend(child, parent) {
            var F = function () { };
            F.prototype = parent.prototype;
            child.prototype = new F();
            child.prototype.constructor = child;
            child.uber = parent.prototype;
            if (parent.prototype.constructor === Object.prototype.constructor) {
                parent.prototype.constructor = parent;
            }
        },
        setNamespace: function OSF_OUtil$setNamespace(name, parent) {
            if (parent && name && !parent[name]) {
                parent[name] = {};
            }
        },
        unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
            if (parent && name && parent[name]) {
                delete parent[name];
            }
        },
        serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection) {
            var ret = {};
            for (var key in settingsCollection) {
                var value = settingsCollection[key];
                try {
                    if (JSON) {
                        value = JSON.stringify(value, function dateReplacer(k, v) {
                            return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings) {
            var ret = {};
            serializedSettings = serializedSettings || {};
            for (var key in serializedSettings) {
                var value = serializedSettings[key];
                try {
                    if (JSON) {
                        value = JSON.parse(value, function dateReviver(k, v) {
                            var d;
                            if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                                d = new Date(parseInt(v.slice(5, -1)));
                                if (d) {
                                    return d;
                                }
                            }
                            return v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
            if (url && callback) {
                var doc = window.document;
                var _loadedScriptEntry = _loadedScripts[url];
                if (!_loadedScriptEntry) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    _loadedScriptEntry = { loaded: false, pendingCallbacks: [callback], timer: null };
                    _loadedScripts[url] = _loadedScriptEntry;
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        _loadedScriptEntry.loaded = true;
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    var onLoadTimeOut = function OSF_OUtil_loadScript$onLoadTimeOut() {
                        if (window.navigator.userAgent.indexOf("Trident") > 0) {
                            onLoadError(null);
                        }
                        else {
                            onLoadError(new Event("Script load timed out"));
                        }
                    };
                    var onLoadError = function OSF_OUtil_loadScript$onLoadError(errorEvent) {
                        delete _loadedScripts[url];
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback(errorEvent);
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                    _loadedScriptEntry.timer = setTimeout(onLoadTimeOut, timeoutInMs);
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (_loadedScriptEntry.loaded) {
                    callback();
                }
                else {
                    _loadedScriptEntry.pendingCallbacks.push(callback);
                }
            }
        },
        loadCSS: function OSF_OUtil$loadCSS(url) {
            if (url) {
                var doc = window.document;
                var link = doc.createElement("link");
                link.type = "text/css";
                link.rel = "stylesheet";
                link.href = url;
                doc.getElementsByTagName("head")[0].appendChild(link);
            }
        },
        parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
            var parsed = enumObject[str.trim()];
            if (typeof (parsed) == 'undefined') {
                OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + str);
                throw OsfMsAjaxFactory.msAjaxError.argument("str");
            }
            return parsed;
        },
        delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
            var obj = { calc: arguments[0] };
            return function () {
                if (obj.calc) {
                    obj.val = obj.calc.apply(this, arguments);
                    delete obj.calc;
                }
                return obj.val;
            };
        },
        getUniqueId: function OSF_OUtil$getUniqueId() {
            _uniqueId = _uniqueId + 1;
            return _uniqueId.toString();
        },
        formatString: function OSF_OUtil$formatString() {
            var args = arguments;
            var source = args[0];
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10) + 1;
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        },
        generateConversationId: function OSF_OUtil$generateConversationId() {
            return [_random(), _random(), (new Date()).getTime().toString()].join('_');
        },
        getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
            return _xdmSessionKeyPrefix + cacheKey + this.generateConversationId();
        },
        addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
            return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
        },
        addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
            return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
        },
        addFlightsAsHash: function OSF_OUtil$addFlightsAsHash(url, flights) {
            return OSF.OUtil.addInfoAsHash(url, _flightsKey, flights, true);
        },
        addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
            url = url.trim() || '';
            var urlParts = url.split(_fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(_fragmentSeparator);
            var newFragment;
            if (encodeInfo) {
                newFragment = [keyName, encodeURIComponent(infoValue), fragment].join('');
            }
            else {
                newFragment = [fragment, keyName, infoValue].join('');
            }
            return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
        },
        parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
        },
        parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
            var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
            if (!xdmInfoValue) {
                xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
            }
            return xdmInfoValue;
        },
        parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
        },
        parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
        },
        parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
            var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
            if (isNaN(serializerVersion)) {
                serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
            }
            return serializerVersion;
        },
        parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
            return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
        },
        parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseFlights: function OSF_OUtil$parseFlights(skipSessionStorage) {
            var flights = OSF.OUtil.parseFlightsWithGivenFragment(skipSessionStorage, window.location.hash);
            if (flights.length == 0) {
                flights = OSF.OUtil.parseFlightsFromWindowName(skipSessionStorage, window.name);
            }
            return flights;
        },
        checkFlight: function OSF_OUtil$checkFlightEnabled(flight) {
            return OSF.Flights && OSF.Flights.indexOf(flight) >= 0;
        },
        pushFlight: function OSF_OUtil$pushFlight(flight) {
            if (OSF.Flights.indexOf(flight) < 0) {
                OSF.Flights.push(flight);
                return true;
            }
            return false;
        },
        getBooleanSetting: function OSF_OUtil$getSetting(settingName) {
            return OSF.OUtil.getBooleanFromDictionary(OSF.Settings, settingName);
        },
        getBooleanFromDictionary: function OSF_OUtil$getBooleanFromDictionary(settings, settingName) {
            var result = (settings && settingName && settings[settingName] !== undefined && settings[settingName] &&
                ((typeof (settings[settingName]) === "string" && settings[settingName].toUpperCase() === 'TRUE') ||
                    (typeof (settings[settingName]) === "boolean" && settings[settingName])));
            return result !== undefined ? result : false;
        },
        getIntFromDictionary: function OSF_OUtil$getIntFromDictionary(settings, settingName) {
            if (settings && settingName && settings[settingName] !== undefined && (typeof settings[settingName] === "string")) {
                return parseInt(settings[settingName]);
            }
            else {
                return NaN;
            }
        },
        pushIntFlight: function OSF_OUtil$pushIntFlight(flight, flightValue) {
            if (!(flight in OSF.IntFlights)) {
                OSF.IntFlights[flight] = flightValue;
                return true;
            }
            return false;
        },
        getIntFlight: function OSF_OUtil$getIntFlight(flight) {
            if (OSF.IntFlights && (flight in OSF.IntFlights)) {
                return OSF.IntFlights[flight];
            }
            else {
                return NaN;
            }
        },
        parseFlightsFromWindowName: function OSF_OUtil$parseFlightsFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.Flights));
        },
        parseFlightsWithGivenFragment: function OSF_OUtil$parseFlightsWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment(_flightsKey, _flightsKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseArrayWithDefault: function OSF_OUtil$parseArrayWithDefault(jsonString) {
            var array = [];
            try {
                array = JSON.parse(jsonString);
            }
            catch (ex) { }
            if (!Array.isArray(array)) {
                array = [];
            }
            return array;
        },
        parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
            try {
                var windowNameObj = JSON.parse(windowName);
                var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                var osfSessionStorage = _getSessionStorage();
                if (!skipSessionStorage && osfSessionStorage && windowNameObj != null) {
                    var sessionKey = windowNameObj[OSF.WindowNameItemKeys.BaseFrameName] + infoKey;
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
                return infoValue;
            }
            catch (Exception) {
                return null;
            }
        },
        parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            if (decodeInfo && infoValue != null) {
                if (infoValue.indexOf(_fragmentInfoDelimiter) >= 0) {
                    infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                }
                infoValue = decodeURIComponent(infoValue);
            }
            var osfSessionStorage = _getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return infoValue;
        },
        getConversationId: function OSF_OUtil$getConversationId() {
            var searchString = window.location.search;
            var conversationId = null;
            if (searchString) {
                var index = searchString.indexOf("&");
                conversationId = index > 0 ? searchString.substring(1, index) : searchString.substr(1);
                if (conversationId && conversationId.charAt(conversationId.length - 1) === '=') {
                    conversationId = conversationId.substring(0, conversationId.length - 1);
                    if (conversationId) {
                        conversationId = decodeURIComponent(conversationId);
                    }
                }
            }
            return conversationId;
        },
        getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
            var items = strInfo.split("$");
            if (typeof items[1] == "undefined") {
                items = strInfo.split("|");
            }
            if (typeof items[1] == "undefined") {
                items = strInfo.split("%7C");
            }
            return items;
        },
        getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
            var fieldValue = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    switch (xdmFieldName) {
                        case OSF.XdmFieldName.ConversationUrl:
                            fieldValue = items[2];
                            break;
                        case OSF.XdmFieldName.AppId:
                            fieldValue = items[1];
                            break;
                    }
                }
            }
            return fieldValue;
        },
        validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "expectedProperties", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: true }
            ]);
            if (e)
                throw e;
            for (var p in expectedProperties) {
                e = Function._validateParameter(params[p], expectedProperties[p], p);
                if (e)
                    throw e;
            }
        },
        writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(text);
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        outputDebug: function OSF_OUtil$outputDebug(text) {
            if (typeof (OsfMsAjaxFactory) !== 'undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
            descriptor = descriptor || {};
            for (var nd in attributes) {
                var attribute = attributes[nd];
                if (descriptor[attribute] == undefined) {
                    descriptor[attribute] = true;
                }
            }
            Object.defineProperty(obj, prop, descriptor);
            return obj;
        },
        defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
            descriptors = descriptors || {};
            for (var prop in descriptors) {
                OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
            }
            return obj;
        },
        defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
        },
        defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
        },
        defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
        },
        defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
        },
        finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
            descriptor = descriptor || {};
            var props = Object.getOwnPropertyNames(obj);
            var propsLength = props.length;
            for (var i = 0; i < propsLength; i++) {
                var prop = props[i];
                var desc = Object.getOwnPropertyDescriptor(obj, prop);
                if (!desc.get && !desc.set) {
                    desc.writable = descriptor.writable || false;
                }
                desc.configurable = descriptor.configurable || false;
                desc.enumerable = descriptor.enumerable || true;
                Object.defineProperty(obj, prop, desc);
            }
            return obj;
        },
        mapList: function OSF_OUtil$MapList(list, mapFunction) {
            var ret = [];
            if (list) {
                for (var item in list) {
                    ret.push(mapFunction(list[item]));
                }
            }
            return ret;
        },
        listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
            for (var item in list) {
                if (key == item) {
                    return true;
                }
            }
            return false;
        },
        listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
            for (var item in list) {
                if (value == list[item]) {
                    return true;
                }
            }
            return false;
        },
        augmentList: function OSF_OUtil$augmentList(list, addenda) {
            var add = list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key] = value; };
            for (var key in addenda) {
                add(key, addenda[key]);
            }
        },
        redefineList: function OSF_Outil$redefineList(oldList, newList) {
            for (var key1 in oldList) {
                delete oldList[key1];
            }
            for (var key2 in newList) {
                oldList[key2] = newList[key2];
            }
        },
        isArray: function OSF_OUtil$isArray(obj) {
            return Object.prototype.toString.apply(obj) === "[object Array]";
        },
        isFunction: function OSF_OUtil$isFunction(obj) {
            return Object.prototype.toString.apply(obj) === "[object Function]";
        },
        isDate: function OSF_OUtil$isDate(obj) {
            return Object.prototype.toString.apply(obj) === "[object Date]";
        },
        addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
            if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = listener;
            }
        },
        removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
            if (element.removeEventListener) {
                element.removeEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.detachEvent) {
                element.detachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = null;
            }
        },
        xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
            var xmlhttp;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp.responseText);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        encodeBase64: function OSF_Outil$encodeBase64(input) {
            if (!input)
                return input;
            var codex = "ABCDEFGHIJKLMNOP" + "QRSTUVWXYZabcdef" + "ghijklmnopqrstuv" + "wxyz0123456789+/=";
            var output = [];
            var temp = [];
            var index = 0;
            var c1, c2, c3, a, b, c;
            var i;
            var length = input.length;
            do {
                c1 = input.charCodeAt(index++);
                c2 = input.charCodeAt(index++);
                c3 = input.charCodeAt(index++);
                i = 0;
                a = c1 & 255;
                b = c1 >> 8;
                c = c2 & 255;
                temp[i++] = a >> 2;
                temp[i++] = ((a & 3) << 4) | (b >> 4);
                temp[i++] = ((b & 15) << 2) | (c >> 6);
                temp[i++] = c & 63;
                if (!isNaN(c2)) {
                    a = c2 >> 8;
                    b = c3 & 255;
                    c = c3 >> 8;
                    temp[i++] = a >> 2;
                    temp[i++] = ((a & 3) << 4) | (b >> 4);
                    temp[i++] = ((b & 15) << 2) | (c >> 6);
                    temp[i++] = c & 63;
                }
                if (isNaN(c2)) {
                    temp[i - 1] = 64;
                }
                else if (isNaN(c3)) {
                    temp[i - 2] = 64;
                    temp[i - 1] = 64;
                }
                for (var t = 0; t < i; t++) {
                    output.push(codex.charAt(temp[t]));
                }
            } while (index < length);
            return output.join("");
        },
        getSessionStorage: function OSF_Outil$getSessionStorage() {
            return _getSessionStorage();
        },
        getLocalStorage: function OSF_Outil$getLocalStorage() {
            if (!_safeLocalStorage) {
                try {
                    var localStorage = window.localStorage;
                }
                catch (ex) {
                    localStorage = null;
                }
                _safeLocalStorage = new OfficeExt.SafeStorage(localStorage);
            }
            return _safeLocalStorage;
        },
        convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
            var hex = "#" + (Number(val) + 0x1000000).toString(16).slice(-6);
            return hex;
        },
        attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
            element.onclick = function (e) {
                handler();
            };
            element.ontouchend = function (e) {
                handler();
                e.preventDefault();
            };
        },
        getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
            var e = Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
                { name: "paramName", type: String, mayBeNull: false }
            ]);
            if (e) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                return "";
            }
            var queryExp = new RegExp("[\\?&]" + paramName + "=([^&#]*)", "i");
            if (!queryExp.test(queryString)) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                return "";
            }
            return queryExp.exec(queryString)[1];
        },
        getHostnamePortionForLogging: function OSF_Outil$getHostnamePortionForLogging(hostname) {
            var e = Function._validateParams(arguments, [{ name: "hostname", type: String, mayBeNull: false }
            ]);
            if (e) {
                return "";
            }
            var hostnameSubstrings = hostname.split('.');
            var len = hostnameSubstrings.length;
            if (len >= 2) {
                return hostnameSubstrings[len - 2] + "." + hostnameSubstrings[len - 1];
            }
            else if (len == 1) {
                return hostnameSubstrings[0];
            }
        },
        isiOS: function OSF_Outil$isiOS() {
            return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
        },
        isChrome: function OSF_Outil$isChrome() {
            return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
        },
        isEdge: function OSF_Outil$isEdge() {
            return window.navigator.userAgent.indexOf("Edge") > 0;
        },
        isIE: function OSF_Outil$isIE() {
            return window.navigator.userAgent.indexOf("Trident") > 0;
        },
        isFirefox: function OSF_Outil$isFirefox() {
            return window.navigator.userAgent.indexOf("Firefox") > 0;
        },
        startsWith: function OSF_Outil$startsWith(originalString, patternToCheck, browserIsIE) {
            if (browserIsIE) {
                return originalString.substr(0, patternToCheck.length) === patternToCheck;
            }
            else {
                return originalString.startsWith(patternToCheck);
            }
        },
        containsPort: function OSF_Outil$containsPort(url, protocol, hostname, portNumber) {
            return this.startsWith(url, protocol + "//" + hostname + ":" + portNumber, true) || this.startsWith(url, hostname + ":" + portNumber, true);
        },
        getRedundandPortString: function OSF_Outil$getRedundandPortString(url, parser) {
            if (!url || !parser)
                return "";
            if (parser.protocol == "https:" && this.containsPort(url, "https:", parser.hostname, "443"))
                return ":443";
            else if (parser.protocol == "http:" && this.containsPort(url, "http:", parser.hostname, "80"))
                return ":80";
            return "";
        },
        removeChar: function OSF_Outil$removeChar(url, indexOfCharToRemove) {
            if (indexOfCharToRemove < url.length - 1)
                return url.substring(0, indexOfCharToRemove) + url.substring(indexOfCharToRemove + 1);
            else if (indexOfCharToRemove == url.length - 1)
                return url.substring(0, url.length - 1);
            else
                return url;
        },
        cleanUrlOfChar: function OSF_Outil$cleanUrlOfChar(url, charToClean) {
            var i;
            for (i = 0; i < url.length; i++) {
                if (url.charAt(i) === charToClean) {
                    if (i + 1 >= url.length) {
                        return this.removeChar(url, i);
                    }
                    else if (charToClean === '/') {
                        if (url.charAt(i + 1) === '?' || url.charAt(i + 1) === '#') {
                            return this.removeChar(url, i);
                        }
                    }
                    else if (charToClean === '?') {
                        if (url.charAt(i + 1) === '#') {
                            return this.removeChar(url, i);
                        }
                    }
                }
            }
            return url;
        },
        cleanUrl: function OSF_Outil$cleanUrl(url) {
            url = this.cleanUrlOfChar(url, '/');
            url = this.cleanUrlOfChar(url, '?');
            url = this.cleanUrlOfChar(url, '#');
            if (url.substr(0, 8) == "https://") {
                var portIndex = url.indexOf(":443");
                if (portIndex != -1) {
                    if (portIndex == url.length - 4 || url.charAt(portIndex + 4) == "/" || url.charAt(portIndex + 4) == "?" || url.charAt(portIndex + 4) == "#") {
                        url = url.substring(0, portIndex) + url.substring(portIndex + 4);
                    }
                }
            }
            else if (url.substr(0, 7) == "http://") {
                var portIndex = url.indexOf(":80");
                if (portIndex != -1) {
                    if (portIndex == url.length - 3 || url.charAt(portIndex + 3) == "/" || url.charAt(portIndex + 3) == "?" || url.charAt(portIndex + 3) == "#") {
                        url = url.substring(0, portIndex) + url.substring(portIndex + 3);
                    }
                }
            }
            return url;
        },
        parseUrl: function OSF_Outil$parseUrl(url, enforceHttps) {
            if (enforceHttps === void 0) {
                enforceHttps = false;
            }
            if (typeof url === "undefined" || !url) {
                return undefined;
            }
            var notHttpsErrorMessage = "NotHttps";
            var invalidUrlErrorMessage = "InvalidUrl";
            var isIEBoolean = this.isIE();
            var parsedUrlObj = {
                protocol: undefined,
                hostname: undefined,
                host: undefined,
                port: undefined,
                pathname: undefined,
                search: undefined,
                hash: undefined,
                isPortPartOfUrl: undefined
            };
            try {
                if (isIEBoolean) {
                    var parser = document.createElement("a");
                    parser.href = url;
                    if (!parser || !parser.protocol || !parser.host || !parser.hostname || !parser.href
                        || this.cleanUrl(parser.href).toLowerCase() !== this.cleanUrl(url).toLowerCase()) {
                        throw invalidUrlErrorMessage;
                    }
                    if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                        if (enforceHttps && parser.protocol != "https:")
                            throw new Error(notHttpsErrorMessage);
                    }
                    var redundandPortString = this.getRedundandPortString(url, parser);
                    parsedUrlObj.protocol = parser.protocol;
                    parsedUrlObj.hostname = parser.hostname;
                    parsedUrlObj.port = (redundandPortString == "") ? parser.port : "";
                    parsedUrlObj.host = (redundandPortString != "") ? parser.hostname : parser.host;
                    parsedUrlObj.pathname = (isIEBoolean ? "/" : "") + parser.pathname;
                    parsedUrlObj.search = parser.search;
                    parsedUrlObj.hash = parser.hash;
                    parsedUrlObj.isPortPartOfUrl = this.containsPort(url, parser.protocol, parser.hostname, parser.port);
                }
                else {
                    var urlObj = new URL(url);
                    if (urlObj && urlObj.protocol && urlObj.host && urlObj.hostname) {
                        if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps)) {
                            if (enforceHttps && urlObj.protocol != "https:")
                                throw new Error(notHttpsErrorMessage);
                        }
                        parsedUrlObj.protocol = urlObj.protocol;
                        parsedUrlObj.hostname = urlObj.hostname;
                        parsedUrlObj.port = urlObj.port;
                        parsedUrlObj.host = urlObj.host;
                        parsedUrlObj.pathname = urlObj.pathname;
                        parsedUrlObj.search = urlObj.search;
                        parsedUrlObj.hash = urlObj.hash;
                        parsedUrlObj.isPortPartOfUrl = urlObj.host.lastIndexOf(":" + urlObj.port) == (urlObj.host.length - urlObj.port.length - 1);
                    }
                }
            }
            catch (err) {
                if (err.message === notHttpsErrorMessage)
                    throw err;
            }
            return parsedUrlObj;
        },
        shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
            if (sourceObj == null) {
                return null;
            }
            else if (!(sourceObj instanceof Object)) {
                return sourceObj;
            }
            else if (Array.isArray(sourceObj)) {
                var copyArr = [];
                for (var i = 0; i < sourceObj.length; i++) {
                    copyArr.push(sourceObj[i]);
                }
                return copyArr;
            }
            else {
                var copyObj = sourceObj.constructor();
                for (var property in sourceObj) {
                    if (sourceObj.hasOwnProperty(property)) {
                        copyObj[property] = sourceObj[property];
                    }
                }
                return copyObj;
            }
        },
        createObject: function OSF_Outil$createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        },
        addClass: function OSF_OUtil$addClass(elmt, val) {
            if (!OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                if (className) {
                    elmt.setAttribute(_classN, className + " " + val);
                }
                else {
                    elmt.setAttribute(_classN, val);
                }
            }
        },
        removeClass: function OSF_OUtil$removeClass(elmt, val) {
            if (OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                var reg = new RegExp('(\\s|^)' + val + '(\\s|$)');
                className = className.replace(reg, '');
                elmt.setAttribute(_classN, className);
            }
        },
        hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
            var className = elmt.getAttribute(_classN);
            return className && className.match(new RegExp('(\\s|^)' + clsName + '(\\s|$)'));
        },
        focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findNextPos = function (allLen, currPos, backward) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && backward) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !backward) {
                    return -1;
                }
                if (backward) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            next = backward ? all.length - 1 : 0;
            if (all.length === 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                window.focus();
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, backward);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
            var currPos;
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findCurrPos = function (all, curr) {
                var i = 0;
                for (; i < all.length; i++) {
                    if (all[i] === curr) {
                        return i;
                    }
                }
                return -1;
            };
            var findNextPos = function (allLen, currPos, shift) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && shift) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !shift) {
                    return -1;
                }
                if (shift) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            currPos = findCurrPos(all, curr);
            next = findNextPos(all.length, currPos, shift);
            if (next < 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, shift);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        isNullOrUndefined: function OSF_OUtil$isNullOrUndefined(value) {
            if (typeof (value) === "undefined") {
                return true;
            }
            if (value === null) {
                return true;
            }
            return false;
        },
        stringEndsWith: function OSF_OUtil$stringEndsWith(value, subString) {
            if (!OSF.OUtil.isNullOrUndefined(value) && !OSF.OUtil.isNullOrUndefined(subString)) {
                if (subString.length > value.length) {
                    return false;
                }
                if (value.substr(value.length - subString.length) === subString) {
                    return true;
                }
            }
            return false;
        },
        hashCode: function OSF_OUtil$hashCode(str) {
            var hash = 0;
            if (!OSF.OUtil.isNullOrUndefined(str)) {
                var i = 0;
                var len = str.length;
                while (i < len) {
                    hash = (hash << 5) - hash + str.charCodeAt(i++) | 0;
                }
            }
            return hash;
        },
        getValue: function OSF_OUtil$getValue(value, defaultValue) {
            if (OSF.OUtil.isNullOrUndefined(value)) {
                return defaultValue;
            }
            return value;
        },
        externalNativeFunctionExists: function OSF_OUtil$externalNativeFunctionExists(type) {
            return type === 'unknown' || type !== 'undefined';
        }
    };
})();
OSF.OUtil.Guid = (function () {
    var hexCode = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    return {
        generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
            var result = "";
            var tick = (new Date()).getTime();
            var index = 0;
            for (; index < 32 && tick > 0; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[tick % 16];
                tick = Math.floor(tick / 16);
            }
            for (; index < 32; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[Math.floor(Math.random() * 16)];
            }
            return result;
        }
    };
})();
try {
    (function () {
        OSF.Flights = OSF.OUtil.parseFlights(true);
    })();
}
catch (ex) { }
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs = {
    "FetchBundleUrl": 0,
    "LoadReactBundle": 1,
    "LoadBundleSuccess": 2,
    "LoadBundleError": 3
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144,
    ExcelWinRT: 524288,
    WordWinRT: 1048576,
    PowerpointWinRT: 2097152,
    OutlookAndroid: 4194304,
    OneNoteWinRT: 8388608,
    ExcelAndroid: 8388609,
    VisioWebApp: 8388610,
    OneNoteIOS: 8388611,
    WordAndroid: 8388613,
    PowerpointAndroid: 8388614,
    Visio: 8388615,
    OneNoteAndroid: 4194305
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    "Select": 0,
    "UnSelect": 1,
    "CancelDialog": 2,
    "InsertAgave": 3,
    "CtrlF6In": 4,
    "CtrlF6Exit": 5,
    "CtrlF6ExitShift": 6,
    "SelectWithError": 7,
    "NotifyHostError": 8,
    "RefreshAddinCommands": 9,
    "PageIsReady": 10,
    "TabIn": 11,
    "TabInShift": 12,
    "TabExit": 13,
    "TabExitShift": 14,
    "EscExit": 15,
    "F2Exit": 16,
    "ExitNoFocusable": 17,
    "ExitNoFocusableShift": 18,
    "MouseEnter": 19,
    "MouseLeave": 20,
    "UpdateTargetUrl": 21,
    "InstallCustomFunctions": 22,
    "SendTelemetryEvent": 23,
    "UninstallCustomFunctions": 24,
    "SendMessage": 25,
    "LaunchExtensionComponent": 26,
    "StopExtensionComponent": 27,
    "RestartExtensionComponent": 28,
    "EnableTaskPaneHeaderButton": 29,
    "DisableTaskPaneHeaderButton": 30,
    "TaskPaneHeaderButtonClicked": 31,
    "RemoveAppCommandsAddin": 32,
    "RefreshRibbonGallery": 33,
    "GetOriginalControlId": 34,
    "OfficeJsReady": 35,
    "InsertDevManifest": 36,
    "InsertDevManifestError": 37,
    "SendCustomerContent": 38,
    "KeyboardShortcuts": 39
};
OSF.SharedConstants = {
    "NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogParentMessageReceived: 1,
    DialogClosed: 12006
};
OSF.OfficeAppContext = function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix, featureGates, officeTheme, initialDisplayMode) {
    this._id = id;
    this._appName = appName;
    this._appVersion = appVersion;
    this._appUILocale = appUILocale;
    this._dataLocale = dataLocale;
    this._docUrl = docUrl;
    this._clientMode = clientMode;
    this._settings = settings;
    this._reason = reason;
    this._osfControlType = osfControlType;
    this._eToken = eToken;
    this._correlationId = correlationId;
    this._appInstanceId = appInstanceId;
    this._touchEnabled = touchEnabled;
    this._commerceAllowed = commerceAllowed;
    this._appMinorVersion = appMinorVersion;
    this._requirementMatrix = requirementMatrix;
    this._hostCustomMessage = hostCustomMessage;
    this._hostFullVersion = hostFullVersion;
    this._isDialog = false;
    this._clientWindowHeight = clientWindowHeight;
    this._clientWindowWidth = clientWindowWidth;
    this._addinName = addinName;
    this._appDomains = appDomains;
    this._dialogRequirementMatrix = dialogRequirementMatrix;
    this._featureGates = featureGates;
    this._officeTheme = officeTheme;
    this._initialDisplayMode = initialDisplayMode;
    this.get_id = function get_id() { return this._id; };
    this.get_appName = function get_appName() { return this._appName; };
    this.get_appVersion = function get_appVersion() { return this._appVersion; };
    this.get_appUILocale = function get_appUILocale() { return this._appUILocale; };
    this.get_dataLocale = function get_dataLocale() { return this._dataLocale; };
    this.get_docUrl = function get_docUrl() { return this._docUrl; };
    this.get_clientMode = function get_clientMode() { return this._clientMode; };
    this.get_bindings = function get_bindings() { return this._bindings; };
    this.get_settings = function get_settings() { return this._settings; };
    this.get_reason = function get_reason() { return this._reason; };
    this.get_osfControlType = function get_osfControlType() { return this._osfControlType; };
    this.get_eToken = function get_eToken() { return this._eToken; };
    this.get_correlationId = function get_correlationId() { return this._correlationId; };
    this.get_appInstanceId = function get_appInstanceId() { return this._appInstanceId; };
    this.get_touchEnabled = function get_touchEnabled() { return this._touchEnabled; };
    this.get_commerceAllowed = function get_commerceAllowed() { return this._commerceAllowed; };
    this.get_appMinorVersion = function get_appMinorVersion() { return this._appMinorVersion; };
    this.get_requirementMatrix = function get_requirementMatrix() { return this._requirementMatrix; };
    this.get_dialogRequirementMatrix = function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
    this.get_hostCustomMessage = function get_hostCustomMessage() { return this._hostCustomMessage; };
    this.get_hostFullVersion = function get_hostFullVersion() { return this._hostFullVersion; };
    this.get_isDialog = function get_isDialog() { return this._isDialog; };
    this.get_clientWindowHeight = function get_clientWindowHeight() { return this._clientWindowHeight; };
    this.get_clientWindowWidth = function get_clientWindowWidth() { return this._clientWindowWidth; };
    this.get_addinName = function get_addinName() { return this._addinName; };
    this.get_appDomains = function get_appDomains() { return this._appDomains; };
    this.get_featureGates = function get_featureGates() { return this._featureGates; };
    this.get_officeTheme = function get_officeTheme() { return this._officeTheme; };
    this.get_initialDisplayMode = function get_initialDisplayMode() { return this._initialDisplayMode ? this._initialDisplayMode : 0; };
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened",
    ControlActivation: "controlActivation"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    PlatformType: "platformType",
    HostType: "hostType",
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    AllowConsentPrompt: "allowConsentPrompt",
    ForMSGraphAccess: "forMSGraphAccess",
    AllowSignInPrompt: "allowSignInPrompt",
    JsonPayload: "jsonPayload",
    EnableNewHosts: "enableNewHosts",
    AccountTypeFilter: "accountTypeFilter",
    AddinTrustId: "addinTrustId",
    Reserved: "reserved",
    Tcid: "tcid",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    CustomFieldId: "customFieldId",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    DisplayInIframe: "displayInIframe",
    MessageContent: "messageContent",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain",
    UrlNoHostInfo: "urlNoHostInfo",
    TargetOrigin: "targetOrigin",
    AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
    Base64: "base64",
    FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function OSF_DDA$GetXdmEventName(id, eventType) {
    if (eventType == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
        return id + "_" + eventType;
    }
    else {
        return eventType;
    }
};
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidCloseContainerMethod: 97,
    dispidGetAccessTokenMethod: 98,
    dispidGetAuthContextMethod: 99,
    dispidOpenBrowserWindow: 102,
    dispidCreateDocumentMethod: 105,
    dispidInsertFormMethod: 106,
    dispidDisplayRibbonCalloutAsyncMethod: 109,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidCreateTaskMethod: 124,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMessageParentMethod: 144,
    dispidSendMessageMethod: 145,
    dispidExecuteFeature: 146,
    dispidQueryFeature: 147,
    dispidMethodMax: 147
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidDialogNotificationShownInAddinEvent: 11,
    dispidDialogParentMessageReceivedEvent: 12,
    dispidObjectDeletedEvent: 13,
    dispidObjectSelectionChangedEvent: 14,
    dispidObjectDataChangedEvent: 15,
    dispidContentControlAddedEvent: 16,
    dispidActivationStatusChangedEvent: 32,
    dispidRichApiMessageEvent: 33,
    dispidAppCommandInvokedEvent: 39,
    dispidOlkItemSelectedChangedEvent: 46,
    dispidOlkRecipientsChangedEvent: 47,
    dispidOlkAppointmentTimeChangedEvent: 48,
    dispidOlkRecurrenceChangedEvent: 49,
    dispidOlkAttachmentsChangedEvent: 50,
    dispidOlkEnhancedLocationsChangedEvent: 51,
    dispidOlkInfobarClickedEvent: 52,
    dispidOlkSelectedItemsChangedEvent: 53,
    dispidOlkSensitivityLabelChangedEvent: 54,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager = (function () {
    var _errorMappings = {};
    return {
        getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
            var errorArgs = _errorMappings[errorCode];
            if (!errorArgs) {
                errorArgs = _errorMappings[this.errorCodes.ooeInternalError];
            }
            else {
                if (!errorArgs.name) {
                    errorArgs.name = _errorMappings[this.errorCodes.ooeInternalError].name;
                }
                if (!errorArgs.message) {
                    errorArgs.message = _errorMappings[this.errorCodes.ooeInternalError].message;
                }
            }
            return errorArgs;
        },
        addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
            _errorMappings[errorCode] = errorNameMessage;
        },
        errorCodes: {
            ooeSuccess: 0,
            ooeChunkResult: 1,
            ooeCoercionTypeNotSupported: 1000,
            ooeGetSelectionNotMatchDataType: 1001,
            ooeCoercionTypeNotMatchBinding: 1002,
            ooeInvalidGetRowColumnCounts: 1003,
            ooeSelectionNotSupportCoercionType: 1004,
            ooeInvalidGetStartRowColumn: 1005,
            ooeNonUniformPartialGetNotSupported: 1006,
            ooeGetDataIsTooLarge: 1008,
            ooeFileTypeNotSupported: 1009,
            ooeGetDataParametersConflict: 1010,
            ooeInvalidGetColumns: 1011,
            ooeInvalidGetRows: 1012,
            ooeInvalidReadForBlankRow: 1013,
            ooeUnsupportedDataObject: 2000,
            ooeCannotWriteToSelection: 2001,
            ooeDataNotMatchSelection: 2002,
            ooeOverwriteWorksheetData: 2003,
            ooeDataNotMatchBindingSize: 2004,
            ooeInvalidSetStartRowColumn: 2005,
            ooeInvalidDataFormat: 2006,
            ooeDataNotMatchCoercionType: 2007,
            ooeDataNotMatchBindingType: 2008,
            ooeSetDataIsTooLarge: 2009,
            ooeNonUniformPartialSetNotSupported: 2010,
            ooeInvalidSetColumns: 2011,
            ooeInvalidSetRows: 2012,
            ooeSetDataParametersConflict: 2013,
            ooeCellDataAmountBeyondLimits: 2014,
            ooeSelectionCannotBound: 3000,
            ooeBindingNotExist: 3002,
            ooeBindingToMultipleSelection: 3003,
            ooeInvalidSelectionForBindingType: 3004,
            ooeOperationNotSupportedOnThisBindingType: 3005,
            ooeNamedItemNotFound: 3006,
            ooeMultipleNamedItemFound: 3007,
            ooeInvalidNamedItemForBindingType: 3008,
            ooeUnknownBindingType: 3009,
            ooeOperationNotSupportedOnMatrixData: 3010,
            ooeInvalidColumnsForBinding: 3011,
            ooeSettingNameNotExist: 4000,
            ooeSettingsCannotSave: 4001,
            ooeSettingsAreStale: 4002,
            ooeOperationNotSupported: 5000,
            ooeInternalError: 5001,
            ooeDocumentReadOnly: 5002,
            ooeEventHandlerNotExist: 5003,
            ooeInvalidApiCallInContext: 5004,
            ooeShuttingDown: 5005,
            ooeUnsupportedEnumeration: 5007,
            ooeIndexOutOfRange: 5008,
            ooeBrowserAPINotSupported: 5009,
            ooeInvalidParam: 5010,
            ooeRequestTimeout: 5011,
            ooeInvalidOrTimedOutSession: 5012,
            ooeInvalidApiArguments: 5013,
            ooeOperationCancelled: 5014,
            ooeWorkbookHidden: 5015,
            ooeWriteNotSupportedWhenModalDialogOpen: 5016,
            ooeTooManyIncompleteRequests: 5100,
            ooeRequestTokenUnavailable: 5101,
            ooeActivityLimitReached: 5102,
            ooeRequestPayloadSizeLimitExceeded: 5103,
            ooeResponsePayloadSizeLimitExceeded: 5104,
            ooeCustomXmlNodeNotFound: 6000,
            ooeCustomXmlError: 6100,
            ooeCustomXmlExceedQuota: 6101,
            ooeCustomXmlOutOfDate: 6102,
            ooeNoCapability: 7000,
            ooeCannotNavTo: 7001,
            ooeSpecifiedIdNotExist: 7002,
            ooeNavOutOfBound: 7004,
            ooeElementMissing: 8000,
            ooeProtectedError: 8001,
            ooeInvalidCellsValue: 8010,
            ooeInvalidTableOptionValue: 8011,
            ooeInvalidFormatValue: 8012,
            ooeRowIndexOutOfRange: 8020,
            ooeColIndexOutOfRange: 8021,
            ooeFormatValueOutOfRange: 8022,
            ooeCellFormatAmountBeyondLimits: 8023,
            ooeMemoryFileLimit: 11000,
            ooeNetworkProblemRetrieveFile: 11001,
            ooeInvalidSliceSize: 11002,
            ooeInvalidCallback: 11101,
            ooeInvalidWidth: 12000,
            ooeInvalidHeight: 12001,
            ooeNavigationError: 12002,
            ooeInvalidScheme: 12003,
            ooeAppDomains: 12004,
            ooeRequireHTTPS: 12005,
            ooeWebDialogClosed: 12006,
            ooeDialogAlreadyOpened: 12007,
            ooeEndUserAllow: 12008,
            ooeEndUserIgnore: 12009,
            ooeNotUILessDialog: 12010,
            ooeCrossZone: 12011,
            ooeModalDialogOpen: 12012,
            ooeDocumentIsInactive: 12013,
            ooeDialogParentIsMinimized: 12014,
            ooeNotSSOAgave: 13000,
            ooeSSOUserNotSignedIn: 13001,
            ooeSSOUserAborted: 13002,
            ooeSSOUnsupportedUserIdentity: 13003,
            ooeSSOInvalidResourceUrl: 13004,
            ooeSSOInvalidGrant: 13005,
            ooeSSOClientError: 13006,
            ooeSSOServerError: 13007,
            ooeAddinIsAlreadyRequestingToken: 13008,
            ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
            ooeSSOConnectionLost: 13010,
            ooeResourceNotAllowed: 13011,
            ooeSSOUnsupportedPlatform: 13012,
            ooeSSOCallThrottled: 13013,
            ooeAccessDenied: 13990,
            ooeGeneralException: 13991
        },
        initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = { name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = { name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = { name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = { name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = { name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = { name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = { name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = { name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = { name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = { name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = { name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = { name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = { name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = { name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = { name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = { name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = { name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = { name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = { name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = { name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = { name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = { name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = { name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden] = { name: stringNS.L_APICallFailed, message: stringNS.L_WorkbookHiddenMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWriteNotSupportedWhenModalDialogOpen] = { name: stringNS.L_APICallFailed, message: stringNS.L_WriteNotSupportedWhenModalDialogOpen };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = { name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = { name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestPayloadSizeLimitExceeded] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestPayloadSizeLimitExceededMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeResponsePayloadSizeLimitExceeded] = { name: stringNS.L_APICallFailed, message: stringNS.L_ResponsePayloadSizeLimitExceededMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = { name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = { name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = { name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = { name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = { name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = { name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = { name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = { name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = { name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = { name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = { name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = { name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = { name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = { name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeModalDialogOpen] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_ModalDialogOpen };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentIsInactive] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DocumentIsInactive };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogParentIsMinimized] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogParentIsMinimized };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = { name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = { name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = { name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = { name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = { name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = { name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = { name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = { name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = { name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = { name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = { name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform] = { name: stringNS.L_APINotSupported, message: stringNS.L_SSOUnsupportedPlatform };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOCallThrottled] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = { name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
        }
    };
})();
(function (OfficeExt) {
    var Requirement;
    (function (Requirement) {
        var RequirementVersion = (function () {
            function RequirementVersion() {
            }
            return RequirementVersion;
        }());
        Requirement.RequirementVersion = RequirementVersion;
        var RequirementMatrix = (function () {
            function RequirementMatrix(_setMap) {
                this.isSetSupported = function _isSetSupported(name, minVersion) {
                    if (name == undefined) {
                        return false;
                    }
                    if (minVersion == undefined) {
                        minVersion = 0;
                    }
                    var setSupportArray = this._setMap;
                    var sets = setSupportArray._sets;
                    if (sets.hasOwnProperty(name.toLowerCase())) {
                        var setMaxVersion = sets[name.toLowerCase()];
                        try {
                            var setMaxVersionNum = this._getVersion(setMaxVersion);
                            minVersion = minVersion + "";
                            var minVersionNum = this._getVersion(minVersion);
                            if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
                                return true;
                            }
                            if (setMaxVersionNum.major > 0 &&
                                setMaxVersionNum.minor >= 0 &&
                                setMaxVersionNum.major == minVersionNum.major &&
                                setMaxVersionNum.minor >= minVersionNum.minor) {
                                return true;
                            }
                        }
                        catch (e) {
                            return false;
                        }
                    }
                    return false;
                };
                this._getVersion = function (version) {
                    version = version + "";
                    var temp = version.split(".");
                    var major = 0;
                    var minor = 0;
                    if (temp.length < 2 && isNaN(Number(version))) {
                        throw "version format incorrect";
                    }
                    else {
                        major = Number(temp[0]);
                        if (temp.length >= 2) {
                            minor = Number(temp[1]);
                        }
                        if (isNaN(major) || isNaN(minor)) {
                            throw "version format incorrect";
                        }
                    }
                    var result = { "minor": minor, "major": major };
                    return result;
                };
                this._setMap = _setMap;
                this.isSetSupported = this.isSetSupported.bind(this);
            }
            return RequirementMatrix;
        }());
        Requirement.RequirementMatrix = RequirementMatrix;
        var DefaultSetRequirement = (function () {
            function DefaultSetRequirement(setMap) {
                this._addSetMap = function DefaultSetRequirement_addSetMap(addedSet) {
                    for (var name in addedSet) {
                        this._sets[name] = addedSet[name];
                    }
                };
                this._sets = setMap;
            }
            return DefaultSetRequirement;
        }());
        Requirement.DefaultSetRequirement = DefaultSetRequirement;
        var DefaultRequiredDialogSetRequirement = (function (_super) {
            __extends(DefaultRequiredDialogSetRequirement, _super);
            function DefaultRequiredDialogSetRequirement() {
                return _super.call(this, {
                    "dialogapi": 1.1
                }) || this;
            }
            return DefaultRequiredDialogSetRequirement;
        }(DefaultSetRequirement));
        Requirement.DefaultRequiredDialogSetRequirement = DefaultRequiredDialogSetRequirement;
        var DefaultOptionalDialogSetRequirement = (function (_super) {
            __extends(DefaultOptionalDialogSetRequirement, _super);
            function DefaultOptionalDialogSetRequirement() {
                return _super.call(this, {
                    "dialogorigin": 1.1
                }) || this;
            }
            return DefaultOptionalDialogSetRequirement;
        }(DefaultSetRequirement));
        Requirement.DefaultOptionalDialogSetRequirement = DefaultOptionalDialogSetRequirement;
        var ExcelClientDefaultSetRequirement = (function (_super) {
            __extends(ExcelClientDefaultSetRequirement, _super);
            function ExcelClientDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "excelapi": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ExcelClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelClientDefaultSetRequirement = ExcelClientDefaultSetRequirement;
        var ExcelClientV1DefaultSetRequirement = (function (_super) {
            __extends(ExcelClientV1DefaultSetRequirement, _super);
            function ExcelClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return ExcelClientV1DefaultSetRequirement;
        }(ExcelClientDefaultSetRequirement));
        Requirement.ExcelClientV1DefaultSetRequirement = ExcelClientV1DefaultSetRequirement;
        var OutlookClientDefaultSetRequirement = (function (_super) {
            __extends(OutlookClientDefaultSetRequirement, _super);
            function OutlookClientDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.3
                }) || this;
            }
            return OutlookClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookClientDefaultSetRequirement = OutlookClientDefaultSetRequirement;
        var WordClientDefaultSetRequirement = (function (_super) {
            __extends(WordClientDefaultSetRequirement, _super);
            function WordClientDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1,
                    "wordapi": 1.1
                }) || this;
            }
            return WordClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordClientDefaultSetRequirement = WordClientDefaultSetRequirement;
        var WordClientV1DefaultSetRequirement = (function (_super) {
            __extends(WordClientV1DefaultSetRequirement, _super);
            function WordClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2,
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return WordClientV1DefaultSetRequirement;
        }(WordClientDefaultSetRequirement));
        Requirement.WordClientV1DefaultSetRequirement = WordClientV1DefaultSetRequirement;
        var PowerpointClientDefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientDefaultSetRequirement, _super);
            function PowerpointClientDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return PowerpointClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointClientDefaultSetRequirement = PowerpointClientDefaultSetRequirement;
        var PowerpointClientV1DefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientV1DefaultSetRequirement, _super);
            function PowerpointClientV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "imagecoercion": 1.1
                });
                return _this;
            }
            return PowerpointClientV1DefaultSetRequirement;
        }(PowerpointClientDefaultSetRequirement));
        Requirement.PowerpointClientV1DefaultSetRequirement = PowerpointClientV1DefaultSetRequirement;
        var ProjectClientDefaultSetRequirement = (function (_super) {
            __extends(ProjectClientDefaultSetRequirement, _super);
            function ProjectClientDefaultSetRequirement() {
                return _super.call(this, {
                    "selection": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ProjectClientDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ProjectClientDefaultSetRequirement = ProjectClientDefaultSetRequirement;
        var ExcelWebDefaultSetRequirement = (function (_super) {
            __extends(ExcelWebDefaultSetRequirement, _super);
            function ExcelWebDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "file": 1.1
                }) || this;
            }
            return ExcelWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelWebDefaultSetRequirement = ExcelWebDefaultSetRequirement;
        var WordWebDefaultSetRequirement = (function (_super) {
            __extends(WordWebDefaultSetRequirement, _super);
            function WordWebDefaultSetRequirement() {
                return _super.call(this, {
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "imagecoercion": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablecoercion": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                }) || this;
            }
            return WordWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordWebDefaultSetRequirement = WordWebDefaultSetRequirement;
        var PowerpointWebDefaultSetRequirement = (function (_super) {
            __extends(PowerpointWebDefaultSetRequirement, _super);
            function PowerpointWebDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "settings": 1.1
                }) || this;
            }
            return PowerpointWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointWebDefaultSetRequirement = PowerpointWebDefaultSetRequirement;
        var OutlookWebDefaultSetRequirement = (function (_super) {
            __extends(OutlookWebDefaultSetRequirement, _super);
            function OutlookWebDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.3
                }) || this;
            }
            return OutlookWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookWebDefaultSetRequirement = OutlookWebDefaultSetRequirement;
        var SwayWebDefaultSetRequirement = (function (_super) {
            __extends(SwayWebDefaultSetRequirement, _super);
            function SwayWebDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "documentevents": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return SwayWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.SwayWebDefaultSetRequirement = SwayWebDefaultSetRequirement;
        var AccessWebDefaultSetRequirement = (function (_super) {
            __extends(AccessWebDefaultSetRequirement, _super);
            function AccessWebDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "partialtablebindings": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1
                }) || this;
            }
            return AccessWebDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.AccessWebDefaultSetRequirement = AccessWebDefaultSetRequirement;
        var ExcelIOSDefaultSetRequirement = (function (_super) {
            __extends(ExcelIOSDefaultSetRequirement, _super);
            function ExcelIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return ExcelIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.ExcelIOSDefaultSetRequirement = ExcelIOSDefaultSetRequirement;
        var WordIOSDefaultSetRequirement = (function (_super) {
            __extends(WordIOSDefaultSetRequirement, _super);
            function WordIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                }) || this;
            }
            return WordIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.WordIOSDefaultSetRequirement = WordIOSDefaultSetRequirement;
        var WordIOSV1DefaultSetRequirement = (function (_super) {
            __extends(WordIOSV1DefaultSetRequirement, _super);
            function WordIOSV1DefaultSetRequirement() {
                var _this = _super.call(this) || this;
                _this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2
                });
                return _this;
            }
            return WordIOSV1DefaultSetRequirement;
        }(WordIOSDefaultSetRequirement));
        Requirement.WordIOSV1DefaultSetRequirement = WordIOSV1DefaultSetRequirement;
        var PowerpointIOSDefaultSetRequirement = (function (_super) {
            __extends(PowerpointIOSDefaultSetRequirement, _super);
            function PowerpointIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                }) || this;
            }
            return PowerpointIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.PowerpointIOSDefaultSetRequirement = PowerpointIOSDefaultSetRequirement;
        var OutlookIOSDefaultSetRequirement = (function (_super) {
            __extends(OutlookIOSDefaultSetRequirement, _super);
            function OutlookIOSDefaultSetRequirement() {
                return _super.call(this, {
                    "mailbox": 1.1
                }) || this;
            }
            return OutlookIOSDefaultSetRequirement;
        }(DefaultSetRequirement));
        Requirement.OutlookIOSDefaultSetRequirement = OutlookIOSDefaultSetRequirement;
        var RequirementsMatrixFactory = (function () {
            function RequirementsMatrixFactory() {
            }
            RequirementsMatrixFactory.initializeOsfDda = function () {
                OSF.OUtil.setNamespace("Requirement", OSF.DDA);
            };
            RequirementsMatrixFactory.getDefaultRequirementMatrix = function (appContext) {
                this.initializeDefaultSetMatrix();
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_requirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    var appLocator = RequirementsMatrixFactory.getClientFullVersionString(appContext);
                    if (RequirementsMatrixFactory.DefaultSetArrayMatrix != undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] != undefined) {
                        defaultRequirementMatrix = new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
                    }
                    else {
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement({}));
                    }
                }
                return defaultRequirementMatrix;
            };
            RequirementsMatrixFactory.getDefaultDialogRequirementMatrix = function (appContext) {
                var setRequirements = undefined;
                var clientRequirement = appContext.get_dialogRequirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    setRequirements = new DefaultSetRequirement(matrixItem);
                }
                else {
                    setRequirements = new DefaultRequiredDialogSetRequirement();
                    var mainRequirement = appContext.get_requirementMatrix();
                    if (mainRequirement != undefined && mainRequirement.length > 0 && typeof (JSON) !== "undefined") {
                        var matrixItem = JSON.parse(mainRequirement.toLowerCase());
                        for (var name in setRequirements._sets) {
                            if (matrixItem.hasOwnProperty(name)) {
                                setRequirements._sets[name] = matrixItem[name];
                            }
                        }
                        var dialogOptionalSetRequirement = new DefaultOptionalDialogSetRequirement();
                        for (var name in dialogOptionalSetRequirement._sets) {
                            if (matrixItem.hasOwnProperty(name)) {
                                setRequirements._sets[name] = matrixItem[name];
                            }
                        }
                    }
                }
                return new RequirementMatrix(setRequirements);
            };
            RequirementsMatrixFactory.getClientFullVersionString = function (appContext) {
                var appMinorVersion = appContext.get_appMinorVersion();
                var appMinorVersionString = "";
                var appFullVersion = "";
                var appName = appContext.get_appName();
                var isIOSClient = appName == 1024 ||
                    appName == 4096 ||
                    appName == 8192 ||
                    appName == 65536;
                if (isIOSClient && appContext.get_appVersion() == 1) {
                    if (appName == 4096 && appMinorVersion >= 15) {
                        appFullVersion = "16.00.01";
                    }
                    else {
                        appFullVersion = "16.00";
                    }
                }
                else if (appContext.get_appName() == 64) {
                    appFullVersion = appContext.get_appVersion();
                }
                else {
                    if (appMinorVersion < 10) {
                        appMinorVersionString = "0" + appMinorVersion;
                    }
                    else {
                        appMinorVersionString = "" + appMinorVersion;
                    }
                    appFullVersion = appContext.get_appVersion() + "." + appMinorVersionString;
                }
                return appContext.get_appName() + "-" + appFullVersion;
            };
            RequirementsMatrixFactory.initializeDefaultSetMatrix = function () {
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600] = new ExcelClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600] = new WordClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600] = new PowerpointClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601] = new ExcelClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601] = new WordClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601] = new PowerpointClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600] = new OutlookClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600] = new ExcelWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600] = new WordWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600] = new ProjectClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600] = new AccessWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600] = new PowerpointWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600] = new ExcelIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600] = new SwayWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600] = new WordIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001] = new WordIOSV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600] = new PowerpointIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600] = new OutlookIOSDefaultSetRequirement();
            };
            RequirementsMatrixFactory.Excel_RCLIENT_1600 = "1-16.00";
            RequirementsMatrixFactory.Excel_RCLIENT_1601 = "1-16.01";
            RequirementsMatrixFactory.Word_RCLIENT_1600 = "2-16.00";
            RequirementsMatrixFactory.Word_RCLIENT_1601 = "2-16.01";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1600 = "4-16.00";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1601 = "4-16.01";
            RequirementsMatrixFactory.Outlook_RCLIENT_1600 = "8-16.00";
            RequirementsMatrixFactory.Excel_WAC_1600 = "16-16.00";
            RequirementsMatrixFactory.Word_WAC_1600 = "32-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1600 = "64-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1601 = "64-16.01";
            RequirementsMatrixFactory.Project_RCLIENT_1600 = "128-16.00";
            RequirementsMatrixFactory.Access_WAC_1600 = "256-16.00";
            RequirementsMatrixFactory.PowerPoint_WAC_1600 = "512-16.00";
            RequirementsMatrixFactory.Excel_IOS_1600 = "1024-16.00";
            RequirementsMatrixFactory.SWAY_WAC_1600 = "2048-16.00";
            RequirementsMatrixFactory.Word_IOS_1600 = "4096-16.00";
            RequirementsMatrixFactory.Word_IOS_16001 = "4096-16.00.01";
            RequirementsMatrixFactory.PowerPoint_IOS_1600 = "8192-16.00";
            RequirementsMatrixFactory.Outlook_IOS_1600 = "65536-16.00";
            RequirementsMatrixFactory.DefaultSetArrayMatrix = {};
            return RequirementsMatrixFactory;
        }());
        Requirement.RequirementsMatrixFactory = RequirementsMatrixFactory;
    })(Requirement = OfficeExt.Requirement || (OfficeExt.Requirement = {}));
})(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
OSF.NamespaceManager = (function OSF_NamespaceManager() {
    var _userOffice;
    var _useShortcut = false;
    return {
        enableShortcut: function OSF_NamespaceManager$enableShortcut() {
            if (!_useShortcut) {
                if (window.Office) {
                    _userOffice = window.Office;
                }
                else {
                    OSF.OUtil.setNamespace("Office", window);
                }
                window.Office = Microsoft.Office.WebExtension;
                _useShortcut = true;
            }
        },
        disableShortcut: function OSF_NamespaceManager$disableShortcut() {
            if (_useShortcut) {
                if (_userOffice) {
                    window.Office = _userOffice;
                }
                else {
                    OSF.OUtil.unsetNamespace("Office", window);
                }
                _useShortcut = false;
            }
        }
    };
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
    if (useShortcut) {
        OSF.NamespaceManager.enableShortcut();
    }
    else {
        OSF.NamespaceManager.disableShortcut();
    }
};
Microsoft.Office.WebExtension.select = function Microsoft_Office_WebExtension_select(str, errorCallback) {
    var promise;
    if (str && typeof str == "string") {
        var index = str.indexOf("#");
        if (index != -1) {
            var op = str.substring(0, index);
            var target = str.substring(index + 1);
            switch (op) {
                case "binding":
                case "bindings":
                    if (target) {
                        promise = new OSF.DDA.BindingPromise(target);
                    }
                    break;
            }
        }
    }
    if (!promise) {
        if (errorCallback) {
            var callbackType = typeof errorCallback;
            if (callbackType == "function") {
                var callArgs = {};
                callArgs[Microsoft.Office.WebExtension.Parameters.Callback] = errorCallback;
                OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
            }
            else {
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
            }
        }
    }
    else {
        promise.onFail = errorCallback;
        return promise;
    }
};
OSF.DDA.Context = function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "contentLanguage": {
            value: officeAppContext.get_dataLocale()
        },
        "displayLanguage": {
            value: officeAppContext.get_appUILocale()
        },
        "touchEnabled": {
            value: officeAppContext.get_touchEnabled()
        },
        "commerceAllowed": {
            value: officeAppContext.get_commerceAllowed()
        },
        "host": {
            value: OfficeExt.HostName.Host.getInstance().getHost()
        },
        "platform": {
            value: OfficeExt.HostName.Host.getInstance().getPlatform()
        },
        "isDialog": {
            value: OSF._OfficeAppFactory.getHostInfo().isDialog
        },
        "diagnostics": {
            value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
        }
    });
    if (license) {
        OSF.OUtil.defineEnumerableProperty(this, "license", {
            value: license
        });
    }
    if (officeAppContext.ui) {
        OSF.OUtil.defineEnumerableProperty(this, "ui", {
            value: officeAppContext.ui
        });
    }
    if (officeAppContext.auth) {
        OSF.OUtil.defineEnumerableProperty(this, "auth", {
            value: officeAppContext.auth
        });
    }
    if (officeAppContext.webAuth) {
        OSF.OUtil.defineEnumerableProperty(this, "webAuth", {
            value: officeAppContext.webAuth
        });
    }
    if (officeAppContext.application) {
        OSF.OUtil.defineEnumerableProperty(this, "application", {
            value: officeAppContext.application
        });
    }
    if (officeAppContext.extensionLifeCycle) {
        OSF.OUtil.defineEnumerableProperty(this, "extensionLifeCycle", {
            value: officeAppContext.extensionLifeCycle
        });
    }
    if (officeAppContext.messaging) {
        OSF.OUtil.defineEnumerableProperty(this, "messaging", {
            value: officeAppContext.messaging
        });
    }
    if (officeAppContext.ui && officeAppContext.ui.taskPaneAction) {
        OSF.OUtil.defineEnumerableProperty(this, "taskPaneAction", {
            value: officeAppContext.ui.taskPaneAction
        });
    }
    if (officeAppContext.ui && officeAppContext.ui.ribbonGallery) {
        OSF.OUtil.defineEnumerableProperty(this, "ribbonGallery", {
            value: officeAppContext.ui.ribbonGallery
        });
    }
    if (officeAppContext.get_isDialog()) {
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
    else {
        if (document) {
            OSF.OUtil.defineEnumerableProperty(this, "document", {
                value: document
            });
        }
        if (appOM) {
            var displayName = appOM.displayName || "appOM";
            delete appOM.displayName;
            OSF.OUtil.defineEnumerableProperty(this, displayName, {
                value: appOM
            });
        }
        if (officeAppContext.get_officeTheme()) {
            OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
                get: function () {
                    return officeAppContext.get_officeTheme();
                }
            });
        }
        else if (getOfficeTheme) {
            OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
                get: function () {
                    return getOfficeTheme();
                }
            });
        }
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
};
OSF.DDA.OutlookContext = function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
    OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
            value: settings
        });
    }
    if (appContext.sensitivityLabelsCatalog) {
        OSF.OUtil.defineEnumerableProperty(this, "sensitivityLabelsCatalog", {
            value: appContext.sensitivityLabelsCatalog()
        });
    }
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Application = function OSF_DDA_Application(officeAppContext) {
};
OSF.DDA.Document = function OSF_DDA_Document(officeAppContext, settings) {
    var mode;
    switch (officeAppContext.get_clientMode()) {
        case OSF.ClientMode.ReadOnly:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
            break;
        case OSF.ClientMode.ReadWrite:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
            break;
    }
    ;
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "settings", {
            value: settings
        });
    }
    ;
    OSF.OUtil.defineMutableProperties(this, {
        "mode": {
            value: mode
        },
        "url": {
            value: officeAppContext.get_docUrl()
        }
    });
};
OSF.DDA.JsomDocument = function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
    OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
    if (bindingFacade) {
        OSF.OUtil.defineEnumerableProperty(this, "bindings", {
            get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
        });
    }
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.GetSelectedDataAsync,
        am.SetSelectedDataAsync
    ]);
    OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
    get: function Microsoft_Office_WebExtension$GetContext() {
        var context;
        if (OSF && OSF._OfficeAppFactory) {
            context = OSF._OfficeAppFactory.getContext();
        }
        return context;
    }
});
OSF.DDA.License = function OSF_DDA_License(eToken) {
    OSF.OUtil.defineEnumerableProperty(this, "value", {
        value: eToken
    });
};
OSF.DDA.ApiMethodCall = function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var getInvalidParameterString = OSF.OUtil.delayExecutionAndCache(function () {
        return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
    });
    this.verifyArguments = function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
        for (var name in params) {
            var param = params[name];
            var arg = args[name];
            if (param["enum"]) {
                switch (typeof arg) {
                    case "string":
                        if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
                            break;
                        }
                    case "undefined":
                        throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                    default:
                        throw getInvalidParameterString();
                }
            }
            if (param["types"]) {
                if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
                    throw getInvalidParameterString();
                }
            }
        }
    };
    this.extractRequiredArguments = function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
        if (userArgs.length < requiredCount) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        }
        var requiredArgs = [];
        var index;
        for (index = 0; index < requiredCount; index++) {
            requiredArgs.push(userArgs[index]);
        }
        this.verifyArguments(requiredParameters, requiredArgs);
        var ret = {};
        for (index = 0; index < requiredCount; index++) {
            var param = requiredParameters[index];
            var arg = requiredArgs[index];
            if (param.verify) {
                var isValid = param.verify(arg, caller, stateInfo);
                if (!isValid) {
                    throw getInvalidParameterString();
                }
            }
            ret[param.name] = arg;
        }
        return ret;
    },
        this.fillOptions = function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
            options = options || {};
            for (var optionName in supportedOptions) {
                if (!OSF.OUtil.listContainsKey(options, optionName)) {
                    var value = undefined;
                    var option = supportedOptions[optionName];
                    if (option.calculate && requiredArgs) {
                        value = option.calculate(requiredArgs, caller, stateInfo);
                    }
                    if (!value && option.defaultValue !== undefined) {
                        value = option.defaultValue;
                    }
                    options[optionName] = value;
                }
            }
            return options;
        };
    this.constructCallArgs = function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
        var callArgs = {};
        for (var r in required) {
            callArgs[r] = required[r];
        }
        for (var o in options) {
            callArgs[o] = options[o];
        }
        for (var s in privateStateCallbacks) {
            callArgs[s] = privateStateCallbacks[s](caller, stateInfo);
        }
        if (checkCallArgs) {
            callArgs = checkCallArgs(callArgs, caller, stateInfo);
        }
        return callArgs;
    };
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.AsyncMethodNames[entry] = am;
    }
};
OSF.DDA.AsyncMethodCall = function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 2) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                case "function":
                    if (parameterCallback) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                    }
                    else {
                        parameterCallback = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        if (parameterCallback) {
            if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
                throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            }
            else {
                options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
            }
        }
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
    this.processResponse = function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
        var payload;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            if (onSucceeded) {
                payload = onSucceeded(response, caller, callArgs);
            }
            else {
                payload = response;
            }
        }
        else {
            if (onFailed) {
                payload = onFailed(status, response);
            }
            else {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
        }
        return payload;
    };
    this.getCallArgs = function (suppliedArgs) {
        var options, parameterCallback;
        for (var i = suppliedArgs.length - 1; i >= requiredCount; i--) {
            var argument = suppliedArgs[i];
            switch (typeof argument) {
                case "object":
                    options = argument;
                    break;
                case "function":
                    parameterCallback = argument;
                    break;
            }
        }
        options = options || {};
        if (parameterCallback) {
            options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
        }
        return options;
    };
};
OSF.DDA.AsyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            var privateStateCallbacks = params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
            return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.AsyncMethodCalls[callDefinition.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error = function OSF_DDA_Error(name, message, code) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "name": {
            value: name
        },
        "message": {
            value: message
        },
        "code": {
            value: code
        }
    });
};
OSF.DDA.AsyncResult = function OSF_DDA_AsyncResult(initArgs, errorArgs) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "value": {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
        },
        "status": {
            value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
        }
    });
    if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
        OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
        });
    }
    if (errorArgs) {
        OSF.OUtil.defineEnumerableProperty(this, "error", {
            value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
        });
    }
};
OSF.DDA.issueAsyncResult = function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
    var callback = callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
    if (callback) {
        var asyncInitArgs = {};
        asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context] = callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var errorArgs;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value] = payload;
        }
        else {
            errorArgs = {};
            payload = payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
        }
        callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
    }
};
OSF.DDA.SyncMethodNames = {};
OSF.DDA.SyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.SyncMethodNames[entry] = am;
    }
};
OSF.DDA.SyncMethodCall = function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 1) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
};
OSF.DDA.SyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.SyncMethodCalls = {};
OSF.DDA.SyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.SyncMethodCalls[callDefinition.method.id] = OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType = (function () {
    var listTypes = {};
    return {
        setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t] = prop; },
        isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
        getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
    };
})();
OSF.DDA.HostParameterMap = function (specialProcessor, mappings) {
    var toHostMap = "toHost";
    var fromHostMap = "fromHost";
    var sourceData = "sourceData";
    var self = "self";
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function (data) {
            if (data != null && data.rows !== undefined) {
                var tableData = {};
                tableData[OSF.DDA.TableDataProperties.TableRows] = data.rows;
                tableData[OSF.DDA.TableDataProperties.TableHeaders] = data.headers;
                data = tableData;
            }
            return data;
        },
        fromHost: function (args) {
            return args;
        }
    };
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData] = dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
    function mapValues(preimageSet, mapping) {
        var ret = preimageSet ? {} : undefined;
        for (var entry in preimageSet) {
            var preimage = preimageSet[entry];
            var image;
            if (OSF.DDA.ListType.isListType(entry)) {
                image = [];
                for (var subEntry in preimage) {
                    image.push(mapValues(preimage[subEntry], mapping));
                }
            }
            else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
                image = dynamicTypes[entry][mapping](preimage);
            }
            else if (mapping == fromHostMap && specialProcessor.preserveNesting(entry)) {
                image = mapValues(preimage, mapping);
            }
            else {
                var maps = mappings[entry];
                if (maps) {
                    var map = maps[mapping];
                    if (map) {
                        image = map[preimage];
                        if (image === undefined) {
                            image = preimage;
                        }
                    }
                }
                else {
                    image = preimage;
                }
            }
            ret[entry] = image;
        }
        return ret;
    }
    ;
    function generateArguments(imageSet, parameters) {
        var ret;
        for (var param in parameters) {
            var arg;
            if (specialProcessor.isComplexType(param)) {
                arg = generateArguments(imageSet, mappings[param][toHostMap]);
            }
            else {
                arg = imageSet[param];
            }
            if (arg != undefined) {
                if (!ret) {
                    ret = {};
                }
                var index = parameters[param];
                if (index == self) {
                    index = param;
                }
                ret[index] = specialProcessor.pack(param, arg);
            }
        }
        return ret;
    }
    ;
    function extractArguments(source, parameters, extracted) {
        if (!extracted) {
            extracted = {};
        }
        for (var param in parameters) {
            var index = parameters[param];
            var value;
            if (index == self) {
                value = source;
            }
            else if (index == sourceData) {
                extracted[param] = source.toArray();
                continue;
            }
            else {
                value = source[index];
            }
            if (value === null || value === undefined) {
                extracted[param] = undefined;
            }
            else {
                value = specialProcessor.unpack(param, value);
                var map;
                if (specialProcessor.isComplexType(param)) {
                    map = mappings[param][fromHostMap];
                    if (specialProcessor.preserveNesting(param)) {
                        extracted[param] = extractArguments(value, map);
                    }
                    else {
                        extractArguments(value, map, extracted);
                    }
                }
                else {
                    if (OSF.DDA.ListType.isListType(param)) {
                        map = {};
                        var entryDescriptor = OSF.DDA.ListType.getDescriptor(param);
                        map[entryDescriptor] = self;
                        var extractedValues = new Array(value.length);
                        for (var item in value) {
                            extractedValues[item] = extractArguments(value[item], map);
                        }
                        extracted[param] = extractedValues;
                    }
                    else {
                        extracted[param] = value;
                    }
                }
            }
        }
        return extracted;
    }
    ;
    function applyMap(mapName, preimage, mapping) {
        var parameters = mappings[mapName][mapping];
        var image;
        if (mapping == "toHost") {
            var imageSet = mapValues(preimage, mapping);
            image = generateArguments(imageSet, parameters);
        }
        else if (mapping == "fromHost") {
            var argumentSet = extractArguments(preimage, parameters);
            image = mapValues(argumentSet, mapping);
        }
        return image;
    }
    ;
    if (!mappings) {
        mappings = {};
    }
    this.addMapping = function (mapName, description) {
        var toHost, fromHost;
        if (description.map) {
            toHost = description.map;
            fromHost = {};
            for (var preimage in toHost) {
                var image = toHost[preimage];
                if (image == self) {
                    image = preimage;
                }
                fromHost[image] = preimage;
            }
        }
        else {
            toHost = description.toHost;
            fromHost = description.fromHost;
        }
        var pair = mappings[mapName];
        if (pair) {
            var currMap = pair[toHostMap];
            for (var th in currMap)
                toHost[th] = currMap[th];
            currMap = pair[fromHostMap];
            for (var fh in currMap)
                fromHost[fh] = currMap[fh];
        }
        else {
            pair = mappings[mapName] = {};
        }
        pair[toHostMap] = toHost;
        pair[fromHostMap] = fromHost;
    };
    this.toHost = function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
    this.fromHost = function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
    this.self = self;
    this.sourceData = sourceData;
    this.addComplexType = function (ct) { specialProcessor.addComplexType(ct); };
    this.getDynamicType = function (dt) { return specialProcessor.getDynamicType(dt); };
    this.setDynamicType = function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
    this.dynamicTypes = dynamicTypes;
    this.doMapValues = function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor = function (complexTypes, dynamicTypes) {
    this.addComplexType = function OSF_DDA_SpecialProcessor$addComplexType(ct) {
        complexTypes.push(ct);
    };
    this.getDynamicType = function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
        return dynamicTypes[dt];
    };
    this.setDynamicType = function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
        dynamicTypes[dt] = handler;
    };
    this.isComplexType = function OSF_DDA_SpecialProcessor$isComplexType(t) {
        return OSF.OUtil.listContainsValue(complexTypes, t);
    };
    this.isDynamicType = function OSF_DDA_SpecialProcessor$isDynamicType(p) {
        return OSF.OUtil.listContainsKey(dynamicTypes, p);
    };
    this.preserveNesting = function OSF_DDA_SpecialProcessor$preserveNesting(p) {
        var pn = [];
        if (OSF.DDA.PropertyDescriptors)
            pn.push(OSF.DDA.PropertyDescriptors.Subset);
        if (OSF.DDA.DataNodeEventProperties) {
            pn = pn.concat([
                OSF.DDA.DataNodeEventProperties.OldNode,
                OSF.DDA.DataNodeEventProperties.NewNode,
                OSF.DDA.DataNodeEventProperties.NextSiblingNode
            ]);
        }
        return OSF.OUtil.listContainsValue(pn, p);
    };
    this.pack = function OSF_DDA_SpecialProcessor$pack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].toHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
    this.unpack = function OSF_DDA_SpecialProcessor$unpack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].fromHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
};
OSF.DDA.getDecoratedParameterMap = function (specialProcessor, initialDefs) {
    var parameterMap = new OSF.DDA.HostParameterMap(specialProcessor);
    var self = parameterMap.self;
    function createObject(properties) {
        var obj = null;
        if (properties) {
            obj = {};
            var len = properties.length;
            for (var i = 0; i < len; i++) {
                obj[properties[i].name] = properties[i].value;
            }
        }
        return obj;
    }
    parameterMap.define = function define(definition) {
        var args = {};
        var toHost = createObject(definition.toHost);
        if (definition.invertible) {
            args.map = toHost;
        }
        else if (definition.canonical) {
            args.toHost = args.fromHost = toHost;
        }
        else {
            args.toHost = toHost;
            args.fromHost = createObject(definition.fromHost);
        }
        parameterMap.addMapping(definition.type, args);
        if (definition.isComplexType)
            parameterMap.addComplexType(definition.type);
    };
    for (var id in initialDefs)
        parameterMap.define(initialDefs[id]);
    return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade = function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
    var dispIdMap = {};
    var jsom = OSF.DDA.AsyncMethodNames;
    var did = OSF.DDA.MethodDispId;
    var methodMap = {
        "GoToByIdAsync": did.dispidNavigateToMethod,
        "GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
        "SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
        "GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
        "ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
        "GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
        "AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
        "AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
        "AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
        "GetAllAsync": did.dispidGetAllBindingsMethod,
        "GetByIdAsync": did.dispidGetBindingMethod,
        "ReleaseByIdAsync": did.dispidReleaseBindingMethod,
        "GetDataAsync": did.dispidGetBindingDataMethod,
        "SetDataAsync": did.dispidSetBindingDataMethod,
        "AddRowsAsync": did.dispidAddRowsMethod,
        "AddColumnsAsync": did.dispidAddColumnsMethod,
        "DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
        "RefreshAsync": did.dispidLoadSettingsMethod,
        "SaveAsync": did.dispidSaveSettingsMethod,
        "GetActiveViewAsync": did.dispidGetActiveViewMethod,
        "GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
        "GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
        "GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
        "ClearFormatsAsync": did.dispidClearFormatsMethod,
        "SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
        "SetFormatsAsync": did.dispidSetFormatsMethod,
        "GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
        "GetAuthContextAsync": did.dispidGetAuthContextMethod,
        "ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
        "AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
        "CloseContainerAsync": did.dispidCloseContainerMethod,
        "OpenBrowserWindow": did.dispidOpenBrowserWindow,
        "CreateDocumentAsync": did.dispidCreateDocumentMethod,
        "InsertFormAsync": did.dispidInsertFormMethod,
        "ExecuteFeature": did.dispidExecuteFeature,
        "QueryFeature": did.dispidQueryFeature,
        "AddDataPartAsync": did.dispidAddDataPartMethod,
        "GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
        "GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
        "GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
        "GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
        "DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
        "GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
        "GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
        "GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
        "SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
        "SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
        "AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
        "GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
        "GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
        "GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
        "SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
        "GetSelectedTask": did.dispidGetSelectedTaskMethod,
        "GetTask": did.dispidGetTaskMethod,
        "GetWSSUrl": did.dispidGetWSSUrlMethod,
        "GetTaskField": did.dispidGetTaskFieldMethod,
        "GetSelectedResource": did.dispidGetSelectedResourceMethod,
        "GetResourceField": did.dispidGetResourceFieldMethod,
        "GetProjectField": did.dispidGetProjectFieldMethod,
        "GetSelectedView": did.dispidGetSelectedViewMethod,
        "GetTaskByIndex": did.dispidGetTaskByIndexMethod,
        "GetResourceByIndex": did.dispidGetResourceByIndexMethod,
        "SetTaskField": did.dispidSetTaskFieldMethod,
        "SetResourceField": did.dispidSetResourceFieldMethod,
        "GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
        "GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
        "CreateTask": did.dispidCreateTaskMethod
    };
    for (var method in methodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = methodMap[method];
        }
    }
    jsom = OSF.DDA.SyncMethodNames;
    did = OSF.DDA.MethodDispId;
    var syncMethodMap = {
        "MessageParent": did.dispidMessageParentMethod,
        "SendMessage": did.dispidSendMessageMethod
    };
    for (var method in syncMethodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = syncMethodMap[method];
        }
    }
    jsom = Microsoft.Office.WebExtension.EventType;
    did = OSF.DDA.EventDispId;
    var eventMap = {
        "SettingsChanged": did.dispidSettingsChangedEvent,
        "DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
        "BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
        "BindingDataChanged": did.dispidBindingDataChangedEvent,
        "ActiveViewChanged": did.dispidActiveViewChangedEvent,
        "OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
        "DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
        "AppCommandInvoked": did.dispidAppCommandInvokedEvent,
        "DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
        "DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
        "ObjectDeleted": did.dispidObjectDeletedEvent,
        "ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
        "ObjectDataChanged": did.dispidObjectDataChangedEvent,
        "ContentControlAdded": did.dispidContentControlAddedEvent,
        "RichApiMessage": did.dispidRichApiMessageEvent,
        "ItemChanged": did.dispidOlkItemSelectedChangedEvent,
        "RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
        "AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
        "RecurrenceChanged": did.dispidOlkRecurrenceChangedEvent,
        "AttachmentsChanged": did.dispidOlkAttachmentsChangedEvent,
        "EnhancedLocationsChanged": did.dispidOlkEnhancedLocationsChangedEvent,
        "InfobarClicked": did.dispidOlkInfobarClickedEvent,
        "SelectedItemsChanged": did.dispidOlkSelectedItemsChangedEvent,
        "SensitivityLabelChanged": did.dispidOlkSensitivityLabelChangedEvent,
        "TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
        "ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
        "ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
        "DataNodeInserted": did.dispidDataNodeAddedEvent,
        "DataNodeReplaced": did.dispidDataNodeReplacedEvent,
        "DataNodeDeleted": did.dispidDataNodeDeletedEvent
    };
    for (var event in eventMap) {
        if (jsom[event]) {
            dispIdMap[jsom[event]] = eventMap[event];
        }
    }
    function IsObjectEvent(dispId) {
        return (dispId == OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidContentControlAddedEvent);
    }
    function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
        if (typeof ex == "number") {
            if (!callArgs) {
                callArgs = asyncMethodCall.getCallArgs(suppliedArgs);
            }
            OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
        }
        else {
            throw ex;
        }
    }
    ;
    this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
        var callArgs;
        try {
            var methodName = method.id;
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[methodName];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
            var dispId = dispIdMap[methodName];
            var delegate = getDelegateMethods(methodName);
            var richApiInExcelMethodSubstitution = null;
            if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
                window.Excel._RedirectV1APIs = true;
            }
            if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution = window.Excel._V1APIMap[methodName])) {
                var preprocessedCallArgs = OSF.OUtil.shallowCopy(callArgs);
                delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
                if (richApiInExcelMethodSubstitution.preprocess) {
                    preprocessedCallArgs = richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
                }
                var ctx = new window.Excel.RequestContext();
                var result = richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
                ctx.sync()
                    .then(function () {
                    var response = result.value;
                    var status = response.status;
                    delete response["status"];
                    delete response["@odata.type"];
                    if (richApiInExcelMethodSubstitution.postprocess) {
                        response = richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
                    }
                    if (status != 0) {
                        response = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                    }
                    OSF.DDA.issueAsyncResult(callArgs, status, response);
                })["catch"](function (error) {
                    OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
                });
            }
            else {
                var hostCallArgs;
                if (parameterMap.toHost) {
                    hostCallArgs = parameterMap.toHost(dispId, callArgs);
                }
                else {
                    hostCallArgs = callArgs;
                }
                var startTime = (new Date()).getTime();
                delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                    "dispId": dispId,
                    "hostCallArgs": hostCallArgs,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                    "onComplete": function (status, hostResponseArgs) {
                        var responseArgs;
                        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                            if (parameterMap.fromHost) {
                                responseArgs = parameterMap.fromHost(dispId, hostResponseArgs);
                            }
                            else {
                                responseArgs = hostResponseArgs;
                            }
                        }
                        else {
                            responseArgs = hostResponseArgs;
                        }
                        var payload = asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
                        OSF.DDA.issueAsyncResult(callArgs, status, payload);
                        if (OSF.AppTelemetry && !(OSF.ConstantNames && OSF.ConstantNames.IsCustomFunctionsRuntime)) {
                            OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
                        }
                    }
                });
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                var added = !isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
                    eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
                if (!added) {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
                }
            }
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            if (isPopupWindow) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                return;
            }
            var dispId_1 = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId_1);
            var targetId_1 = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId_1) : eventDispatch.getEventHandlerCount(eventType);
            if (count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId_1,
                    "targetId": targetId_1,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration,
                    "onEvent": function handleEvent(hostArgs) {
                        var args = null;
                        if (OSF.OUtil.checkFlight(OSF.FlightTreatmentNames.WordEditorAddinThemeSupportEnabled)) {
                            var hostInfo = OSF._OfficeAppFactory.getHostInfo();
                            if (hostInfo && hostInfo.hostPlatform.toLowerCase() == "web" && dispId_1 == OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent) {
                                args = hostArgs;
                            }
                            else {
                                args = parameterMap.fromHost(dispId_1, hostArgs);
                            }
                        }
                        else {
                            args = parameterMap.fromHost(dispId_1, hostArgs);
                        }
                        if (!isObjectEvent)
                            eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
                        else
                            eventDispatch.fireObjectEvent(targetId_1, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId_1, args));
                    }
                });
            }
            else {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var status, removeSuccess;
            if (handler === null) {
                removeSuccess = isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
            }
            else {
                removeSuccess = isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
                status = removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
            }
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
            if (removeSuccess && count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration
                });
            }
            else {
                onEnsureRegistration(status);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.OpenDialog] = function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller, isModal) {
        var callArgs;
        var targetId;
        var asyncMethodCall = null;
        var dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
        var dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
        function onEnsureRegistration(status) {
            var payload;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            else {
                var onSucceedArgs = {};
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id] = targetId;
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data] = eventDispatch;
                var payload = asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
                OSF.DialogShownStatus.hasDialogShown = true;
                eventDispatch.clearEventHandlers(dialogMessageEvent);
                eventDispatch.clearEventHandlers(dialogOtherEvent);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, payload);
        }
        try {
            if (dialogMessageEvent == undefined || dialogOtherEvent == undefined) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
            }
            if (!isModal) {
                if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync == null) {
                    onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                    return;
                }
                asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
            }
            else {
                if (OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync == null) {
                    onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                    return;
                }
                asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync.id];
            }
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
            callArgs["isModal"] = isModal;
            targetId = JSON.stringify(callArgs);
            if (!OSF.DialogShownStatus.hasDialogShown) {
                eventDispatch.clearQueuedEvent(dialogMessageEvent);
                eventDispatch.clearQueuedEvent(dialogOtherEvent);
                eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
            }
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": onEnsureRegistration,
                "onEvent": function handleEvent(hostArgs) {
                    var args = parameterMap.fromHost(dispId, hostArgs);
                    var event = OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
                    if (event.type == dialogOtherEvent) {
                        var payload = OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
                        var errorArgs = {};
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
                        event.error = new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
                    }
                    eventDispatch.fireOrQueueEvent(event);
                    if (args[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogClosed) {
                        eventDispatch.clearEventHandlers(dialogMessageEvent);
                        eventDispatch.clearEventHandlers(dialogOtherEvent);
                        eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
                        OSF.DialogShownStatus.hasDialogShown = false;
                    }
                }
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.CloseDialog] = function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
        var callArgs;
        var dialogMessageEvent, dialogOtherEvent;
        var closeStatus = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
        function closeCallback(status) {
            closeStatus = status;
            OSF.DialogShownStatus.hasDialogShown = false;
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
            dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
            eventDispatch.clearEventHandlers(dialogMessageEvent);
            eventDispatch.clearEventHandlers(dialogOtherEvent);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": closeCallback
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
        if (closeStatus != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.MessageParent] = function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
    this[OSF.DDA.DispIdHost.Methods.SendMessage] = function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
};
OSF.DDA.DispIdHost.addAsyncMethods = function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
    for (var entry in asyncMethodNames) {
        var method = asyncMethodNames[entry];
        var name = method.displayName;
        if (!target[name]) {
            OSF.OUtil.defineEnumerableProperty(target, name, {
                value: (function (asyncMethod) {
                    return function () {
                        var invokeMethod = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                        invokeMethod(asyncMethod, arguments, target, privateState);
                    };
                })(method)
            });
        }
    }
};
OSF.DDA.DispIdHost.addEventSupport = function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
    var add = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
    var remove = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    if (!target[add]) {
        OSF.OUtil.defineEnumerableProperty(target, add, {
            value: function () {
                var addEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
                addEventHandler(arguments, eventDispatch, target, isPopupWindow);
            }
        });
    }
    if (!target[remove]) {
        OSF.OUtil.defineEnumerableProperty(target, remove, {
            value: function () {
                var removeEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
                removeEventHandler(arguments, eventDispatch, target);
            }
        });
    }
};
(function (OfficeExt) {
    var MsAjaxTypeHelper = (function () {
        function MsAjaxTypeHelper() {
        }
        MsAjaxTypeHelper.isInstanceOfType = function (type, instance) {
            if (typeof (instance) === "undefined" || instance === null)
                return false;
            if (instance instanceof type)
                return true;
            var instanceType = instance.constructor;
            if (!instanceType || (typeof (instanceType) !== "function") || !instanceType.__typeName || instanceType.__typeName === 'Object') {
                instanceType = Object;
            }
            return !!(instanceType === type) ||
                (instanceType.__typeName && type.__typeName && instanceType.__typeName === type.__typeName);
        };
        return MsAjaxTypeHelper;
    }());
    OfficeExt.MsAjaxTypeHelper = MsAjaxTypeHelper;
    var MsAjaxError = (function () {
        function MsAjaxError() {
        }
        MsAjaxError.create = function (message, errorInfo) {
            var err = new Error(message);
            err.message = message;
            if (errorInfo) {
                for (var v in errorInfo) {
                    err[v] = errorInfo[v];
                }
            }
            err.popStackFrame();
            return err;
        };
        MsAjaxError.parameterCount = function (message) {
            var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argument = function (paramName, message) {
            var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentNull = function (paramName, message) {
            var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentOutOfRange = function (paramName, actualValue, message) {
            var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            if (typeof (actualValue) !== "undefined" && actualValue !== null) {
                displayMessage += "\n" + MsAjaxString.format("Actual value was {0}.", actualValue);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentOutOfRangeException",
                paramName: paramName,
                actualValue: actualValue
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentType = function (paramName, actualType, expectedType, message) {
            var displayMessage = "Sys.ArgumentTypeException: ";
            if (message) {
                displayMessage += message;
            }
            else if (actualType && expectedType) {
                displayMessage += MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
            }
            else {
                displayMessage += "Object cannot be converted to the required type.";
            }
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentTypeException",
                paramName: paramName,
                actualType: actualType,
                expectedType: expectedType
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentUndefined = function (paramName, message) {
            var displayMessage = "Sys.ArgumentUndefinedException: " + (message ? message : "Value cannot be undefined.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.invalidOperation = function (message) {
            var displayMessage = "Sys.InvalidOperationException: " + (message ? message : "Operation is not valid due to the current state of the object.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
            err.popStackFrame();
            return err;
        };
        return MsAjaxError;
    }());
    OfficeExt.MsAjaxError = MsAjaxError;
    var MsAjaxString = (function () {
        function MsAjaxString() {
        }
        MsAjaxString.format = function (format) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            var source = format;
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10);
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        };
        MsAjaxString.startsWith = function (str, prefix) {
            return (str.substr(0, prefix.length) === prefix);
        };
        return MsAjaxString;
    }());
    OfficeExt.MsAjaxString = MsAjaxString;
    var MsAjaxDebug = (function () {
        function MsAjaxDebug() {
        }
        MsAjaxDebug.trace = function (text) {
            if (typeof Debug !== "undefined" && Debug.writeln)
                Debug.writeln(text);
            if (window.console && window.console.log)
                window.console.log(text);
            if (window.opera && window.opera.postError)
                window.opera.postError(text);
            if (window.debugService && window.debugService.trace)
                window.debugService.trace(text);
            var a = document.getElementById("TraceConsole");
            if (a && a.tagName.toUpperCase() === "TEXTAREA") {
                a.appendChild(document.createTextNode(text + "\n"));
            }
        };
        return MsAjaxDebug;
    }());
    OfficeExt.MsAjaxDebug = MsAjaxDebug;
    if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
        var registerTypeInternal = function registerTypeInternal(type, name, isClass) {
            if (type.__typeName === undefined || type.__typeName === null) {
                type.__typeName = name;
            }
            if (type.__class === undefined || type.__class === null) {
                type.__class = isClass;
            }
        };
        registerTypeInternal(Function, "Function", true);
        registerTypeInternal(Error, "Error", true);
        registerTypeInternal(Object, "Object", true);
        registerTypeInternal(String, "String", true);
        registerTypeInternal(Boolean, "Boolean", true);
        registerTypeInternal(Date, "Date", true);
        registerTypeInternal(Number, "Number", true);
        registerTypeInternal(RegExp, "RegExp", true);
        registerTypeInternal(Array, "Array", true);
        if (!Function.createCallback) {
            Function.createCallback = function Function$createCallback(method, context) {
                var e = Function._validateParams(arguments, [
                    { name: "method", type: Function },
                    { name: "context", mayBeNull: true }
                ]);
                if (e)
                    throw e;
                return function () {
                    var l = arguments.length;
                    if (l > 0) {
                        var args = [];
                        for (var i = 0; i < l; i++) {
                            args[i] = arguments[i];
                        }
                        args[l] = context;
                        return method.apply(this, args);
                    }
                    return method.call(this, context);
                };
            };
        }
        if (!Function.createDelegate) {
            Function.createDelegate = function Function$createDelegate(instance, method) {
                var e = Function._validateParams(arguments, [
                    { name: "instance", mayBeNull: true },
                    { name: "method", type: Function }
                ]);
                if (e)
                    throw e;
                return function () {
                    return method.apply(instance, arguments);
                };
            };
        }
        if (!Function._validateParams) {
            Function._validateParams = function (params, expectedParams, validateParameterCount) {
                var e, expectedLength = expectedParams.length;
                validateParameterCount = validateParameterCount || (typeof (validateParameterCount) === "undefined");
                e = Function._validateParameterCount(params, expectedParams, validateParameterCount);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                for (var i = 0, l = params.length; i < l; i++) {
                    var expectedParam = expectedParams[Math.min(i, expectedLength - 1)], paramName = expectedParam.name;
                    if (expectedParam.parameterArray) {
                        paramName += "[" + (i - expectedLength + 1) + "]";
                    }
                    else if (!validateParameterCount && (i >= expectedLength)) {
                        break;
                    }
                    e = Function._validateParameter(params[i], expectedParam, paramName);
                    if (e) {
                        e.popStackFrame();
                        return e;
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterCount) {
            Function._validateParameterCount = function (params, expectedParams, validateParameterCount) {
                var i, error, expectedLen = expectedParams.length, actualLen = params.length;
                if (actualLen < expectedLen) {
                    var minParams = expectedLen;
                    for (i = 0; i < expectedLen; i++) {
                        var param = expectedParams[i];
                        if (param.optional || param.parameterArray) {
                            minParams--;
                        }
                    }
                    if (actualLen < minParams) {
                        error = true;
                    }
                }
                else if (validateParameterCount && (actualLen > expectedLen)) {
                    error = true;
                    for (i = 0; i < expectedLen; i++) {
                        if (expectedParams[i].parameterArray) {
                            error = false;
                            break;
                        }
                    }
                }
                if (error) {
                    var e = MsAjaxError.parameterCount();
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!Function._validateParameter) {
            Function._validateParameter = function (param, expectedParam, paramName) {
                var e, expectedType = expectedParam.type, expectedInteger = !!expectedParam.integer, expectedDomElement = !!expectedParam.domElement, mayBeNull = !!expectedParam.mayBeNull;
                e = Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                var expectedElementType = expectedParam.elementType, elementMayBeNull = !!expectedParam.elementMayBeNull;
                if (expectedType === Array && typeof (param) !== "undefined" && param !== null &&
                    (expectedElementType || !elementMayBeNull)) {
                    var expectedElementInteger = !!expectedParam.elementInteger, expectedElementDomElement = !!expectedParam.elementDomElement;
                    for (var i = 0; i < param.length; i++) {
                        var elem = param[i];
                        e = Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName + "[" + i + "]");
                        if (e) {
                            e.popStackFrame();
                            return e;
                        }
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterType) {
            Function._validateParameterType = function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
                var e, i;
                if (typeof (param) === "undefined") {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentUndefined(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (param === null) {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentNull(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
                    e = OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!window.Type) {
            window.Type = Function;
        }
        if (!Type.registerNamespace) {
            Type.registerNamespace = function (ns) {
                var namespaceParts = ns.split('.');
                var currentNamespace = window;
                for (var i = 0; i < namespaceParts.length; i++) {
                    currentNamespace[namespaceParts[i]] = currentNamespace[namespaceParts[i]] || {};
                    currentNamespace = currentNamespace[namespaceParts[i]];
                }
            };
        }
        if (!Type.prototype.registerClass) {
            Type.prototype.registerClass = function (cls) { cls = {}; };
        }
        if (typeof (Sys) === "undefined") {
            Type.registerNamespace('Sys');
        }
        if (!Error.prototype.popStackFrame) {
            Error.prototype.popStackFrame = function () {
                if (arguments.length !== 0)
                    throw MsAjaxError.parameterCount();
                if (typeof (this.stack) === "undefined" || this.stack === null ||
                    typeof (this.fileName) === "undefined" || this.fileName === null ||
                    typeof (this.lineNumber) === "undefined" || this.lineNumber === null) {
                    return;
                }
                var stackFrames = this.stack.split("\n");
                var currentFrame = stackFrames[0];
                var pattern = this.fileName + ":" + this.lineNumber;
                while (typeof (currentFrame) !== "undefined" &&
                    currentFrame !== null &&
                    currentFrame.indexOf(pattern) === -1) {
                    stackFrames.shift();
                    currentFrame = stackFrames[0];
                }
                var nextFrame = stackFrames[1];
                if (typeof (nextFrame) === "undefined" || nextFrame === null) {
                    return;
                }
                var nextFrameParts = nextFrame.match(/@(.*):(\d+)$/);
                if (typeof (nextFrameParts) === "undefined" || nextFrameParts === null) {
                    return;
                }
                this.fileName = nextFrameParts[1];
                this.lineNumber = parseInt(nextFrameParts[2]);
                stackFrames.shift();
                this.stack = stackFrames.join("\n");
            };
        }
        OsfMsAjaxFactory.msAjaxError = MsAjaxError;
        OsfMsAjaxFactory.msAjaxString = MsAjaxString;
        OsfMsAjaxFactory.msAjaxDebug = MsAjaxDebug;
    }
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response = {
    Status: 0,
    Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments = {
    Offset: "offset",
    Run: "run",
    BindingSpecificData: "bindingSpecificData",
    MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException = function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
    var status;
    var statusNumber = ex.number;
    if (statusNumber) {
        switch (statusNumber) {
            case -2146828218:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                break;
            case -2147467259:
                if (args.dispId == OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent) {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
                }
                else {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                }
                break;
            case -2146828283:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
                break;
            case -2147209089:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
                break;
            case -2147208704:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests;
                break;
            case -2146827850:
            default:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                break;
        }
    }
    if (args.onComplete) {
        args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
    }
};
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod = function OSF_DDA_SafeArray_Delegate$OnExceptionSyncMethod(ex, args) {
    var status;
    var number = ex.number;
    if (number) {
        switch (number) {
            case -2146828218:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                break;
            case -2146827850:
            default:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                break;
        }
    }
    return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor = function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
    function _2DVBArrayToJaggedArray(vbArr) {
        var ret;
        try {
            var rows = vbArr.ubound(1);
            var cols = vbArr.ubound(2);
            vbArr = vbArr.toArray();
            if (rows == 1 && cols == 1) {
                ret = [vbArr];
            }
            else {
                ret = [];
                for (var row = 0; row < rows; row++) {
                    var rowArr = [];
                    for (var col = 0; col < cols; col++) {
                        var datum = vbArr[row * cols + col];
                        if (datum != OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
                            rowArr.push(datum);
                        }
                    }
                    if (rowArr.length > 0) {
                        ret.push(rowArr);
                    }
                }
            }
        }
        catch (ex) {
        }
        return ret;
    }
    var complexTypes = [];
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = (function () {
        var tableRows = 0;
        var tableHeaders = 1;
        return {
            toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
                if (OSF.DDA.TableDataProperties && typeof data != "string" && data[OSF.DDA.TableDataProperties.TableRows] !== undefined) {
                    var tableData = [];
                    tableData[tableRows] = data[OSF.DDA.TableDataProperties.TableRows];
                    tableData[tableHeaders] = data[OSF.DDA.TableDataProperties.TableHeaders];
                    data = tableData;
                }
                return data;
            },
            fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
                var ret;
                if (hostArgs.toArray) {
                    var dimensions = hostArgs.dimensions();
                    if (dimensions === 2) {
                        ret = _2DVBArrayToJaggedArray(hostArgs);
                    }
                    else {
                        var array = hostArgs.toArray();
                        if (array.length === 2 && ((array[0] != null && array[0].toArray) || (array[1] != null && array[1].toArray))) {
                            ret = {};
                            ret[OSF.DDA.TableDataProperties.TableRows] = _2DVBArrayToJaggedArray(array[tableRows]);
                            ret[OSF.DDA.TableDataProperties.TableHeaders] = _2DVBArrayToJaggedArray(array[tableHeaders]);
                        }
                        else {
                            ret = array;
                        }
                    }
                }
                else {
                    ret = hostArgs;
                }
                return ret;
            }
        };
    })();
    OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
    this.unpack = function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
        var value;
        if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
            var toArraySupported = arg !== undefined && arg.toArray !== undefined;
            value = toArraySupported ? arg.toArray() : arg || {};
        }
        else if (this.isDynamicType(param)) {
            value = dynamicTypes[param].fromHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor(), [
    {
        type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
        toHost: [
            { name: Microsoft.Office.WebExtension.ValueFormat.Unformatted, value: 0 },
            { name: Microsoft.Office.WebExtension.ValueFormat.Formatted, value: 1 }
        ]
    },
    {
        type: Microsoft.Office.WebExtension.Parameters.FilterType,
        toHost: [
            { name: Microsoft.Office.WebExtension.FilterType.All, value: 0 }
        ]
    }
]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
    fromHost: [
        { name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded, value: 0 },
        { name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed, value: 1 }
    ]
});
OSF.DDA.SafeArray.Delegate.executeAsync = function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
    function toArray(args) {
        var arrArgs = args;
        if (OSF.OUtil.isArray(args)) {
            var len = arrArgs.length;
            for (var i = 0; i < len; i++) {
                arrArgs[i] = toArray(arrArgs[i]);
            }
        }
        else if (OSF.OUtil.isDate(args)) {
            arrArgs = args.getVarDate();
        }
        else if (typeof args === "object" && !OSF.OUtil.isArray(args)) {
            arrArgs = [];
            for (var index in args) {
                if (!OSF.OUtil.isFunction(args[index])) {
                    arrArgs[index] = toArray(args[index]);
                }
            }
        }
        return arrArgs;
    }
    function fromSafeArray(value) {
        var ret = value;
        if (value != null && value.toArray) {
            var arrayResult = value.toArray();
            ret = new Array(arrayResult.length);
            for (var i = 0; i < arrayResult.length; i++) {
                ret[i] = fromSafeArray(arrayResult[i]);
            }
        }
        return ret;
    }
    try {
        if (args.onCalling) {
            args.onCalling();
        }
        OSF.ClientHostController.execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
            var result;
            var status;
            if (typeof hostResponseArgs === "number") {
                result = [];
                status = hostResponseArgs;
            }
            else {
                result = hostResponseArgs.toArray();
                status = result[OSF.DDA.SafeArray.Response.Status];
            }
            if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
                var payload = result[OSF.DDA.SafeArray.Response.Payload];
                payload = fromSafeArray(payload);
                if (payload != null) {
                    if (!args._chunkResultData) {
                        args._chunkResultData = new Array();
                    }
                    args._chunkResultData[payload[0]] = payload[1];
                }
                return false;
            }
            if (args.onReceiving) {
                args.onReceiving();
            }
            if (args.onComplete) {
                var payload;
                if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                    if (result.length > 2) {
                        payload = [];
                        for (var i = 1; i < result.length; i++)
                            payload[i - 1] = result[i];
                    }
                    else {
                        payload = result[OSF.DDA.SafeArray.Response.Payload];
                    }
                    if (args._chunkResultData) {
                        payload = fromSafeArray(payload);
                        if (payload != null) {
                            var expectedChunkCount = payload[payload.length - 1];
                            if (args._chunkResultData.length == expectedChunkCount) {
                                payload[payload.length - 1] = args._chunkResultData;
                            }
                            else {
                                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                            }
                        }
                    }
                }
                else {
                    payload = result[OSF.DDA.SafeArray.Response.Payload];
                }
                args.onComplete(status, payload);
            }
            return true;
        });
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent = function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
    var startTime = (new Date()).getTime();
    return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
        if (args.onReceiving) {
            args.onReceiving();
        }
        var status = hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
        if (args.onComplete) {
            args.onComplete(status);
        }
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
        }
    };
};
OSF.DDA.SafeArray.Delegate.registerEventAsync = function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
    try {
        OSF.ClientHostController.registerEvent(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
            if (args.onEvent) {
                args.onEvent(payload);
            }
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.onEventDone(args.dispId);
            }
        }, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync = function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
    try {
        OSF.ClientHostController.unregisterEvent(args.dispId, args.targetId, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.ClientMode = {
    ReadWrite: 0,
    ReadOnly: 1
};
OSF.DDA.RichInitializationReason = {
    1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
    2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
    this._initializeSettings = this.initializeSettings;
};
OSF.InitializationHelper.prototype.deserializeSettings = function OSF_InitializationHelper$deserializeSettings(serializedSettings, refreshSupported) {
    var settings;
    var osfSessionStorage = OSF.OUtil.getSessionStorage();
    if (osfSessionStorage) {
        var storageSettings = osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
        if (storageSettings) {
            serializedSettings = JSON.parse(storageSettings);
        }
        else {
            storageSettings = JSON.stringify(serializedSettings);
            osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
        }
    }
    var deserializedSettings = OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
    if (refreshSupported) {
        settings = new OSF.DDA.RefreshableSettings(deserializedSettings);
    }
    else {
        settings = new OSF.DDA.Settings(deserializedSettings);
    }
    return settings;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
    this.prepareApiSurface(appContext);
    Microsoft.Office.WebExtension.initialize(this.getInitializationReason(appContext));
};
OSF.InitializationHelper.prototype.prepareApiSurface = function OSF_InitializationHelper$prepareApiSurfaceAndInitialize(appContext) {
    var license = new OSF.DDA.License(appContext.get_eToken());
    var getOfficeThemeHandler = (OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme) ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
    if (appContext.get_isDialog()) {
        if (OSF.DDA.UI.ChildUI) {
            appContext.ui = new OSF.DDA.UI.ChildUI();
        }
    }
    else {
        if (OSF.DDA.UI.ParentUI) {
            appContext.ui = new OSF.DDA.UI.ParentUI();
            if (OfficeExt.Container) {
                OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
            }
        }
    }
    if (OSF.DDA.OpenBrowser) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);
    }
    if (OSF.DDA.ExecuteFeature) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.ExecuteFeature]);
    }
    if (OSF.DDA.QueryFeature) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.QueryFeature]);
    }
    if (OSF.DDA.Auth) {
        appContext.auth = new OSF.DDA.Auth();
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
    }
    OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(appContext, appContext.doc, license, null, getOfficeThemeHandler));
    var getDelegateMethods, parameterMap;
    getDelegateMethods = OSF.DDA.DispIdHost.getClientDelegateMethods;
    parameterMap = OSF.DDA.SafeArray.Delegate.ParameterMap;
    OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(getDelegateMethods, parameterMap));
};
OSF.InitializationHelper.prototype.getInitializationReason = function (appContext) { return OSF.DDA.RichInitializationReason[appContext.get_reason()]; };
OSF.DDA.DispIdHost.getClientDelegateMethods = function (actionId) {
    var delegateMethods = {};
    delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.SafeArray.Delegate.executeAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.SafeArray.Delegate.registerEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] = OSF.DDA.SafeArray.Delegate.openDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] = OSF.DDA.SafeArray.Delegate.closeDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent] = OSF.DDA.SafeArray.Delegate.messageParent;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage] = OSF.DDA.SafeArray.Delegate.sendMessage;
    if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId == OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
        var readSerializedSettings = function (hostCallArgs, onCalling, onReceiving) {
            if (typeof (OSF.DDA.ClientSettingsManager.refresh) === "function") {
                return OSF.DDA.ClientSettingsManager.refresh(onCalling, onReceiving);
            }
            else {
                return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
            }
        };
        delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
    }
    if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId == OSF.DDA.AsyncMethodNames.SaveAsync.id) {
        var writeSerializedSettings = function (hostCallArgs, onCalling, onReceiving) {
            return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
        };
        delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
    }
    return delegateMethods;
};
(function (OfficeExt) {
    var RichClientHostController = (function () {
        function RichClientHostController() {
        }
        RichClientHostController.prototype.execute = function (id, params, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.Execute(id, params, callback, OsfOMToken);
            }
            else {
                window.external.Execute(id, params, callback);
            }
        };
        RichClientHostController.prototype.registerEvent = function (id, targetId, handler, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.RegisterEvent(id, targetId, handler, callback, OsfOMToken);
            }
            else {
                window.external.RegisterEvent(id, targetId, handler, callback);
            }
        };
        RichClientHostController.prototype.unregisterEvent = function (id, targetId, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.UnregisterEvent(id, targetId, callback, OsfOMToken);
            }
            else {
                window.external.UnregisterEvent(id, targetId, callback);
            }
        };
        RichClientHostController.prototype.closeSdxDialog = function (context) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof window.external.closeSdxDialog)) {
                window.external.closeSdxDialog(context);
            }
        };
        return RichClientHostController;
    }());
    OfficeExt.RichClientHostController = RichClientHostController;
})(OfficeExt || (OfficeExt = {}));
(function (OfficeExt) {
    var Win32RichClientHostController = (function (_super) {
        __extends(Win32RichClientHostController, _super);
        function Win32RichClientHostController() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Win32RichClientHostController.prototype.messageParent = function (params) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof window.external.MessageParent2)) {
                if (params) {
                    var messageToParent = params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                    if (typeof messageToParent === "boolean") {
                        if (messageToParent === true) {
                            params[Microsoft.Office.WebExtension.Parameters.MessageToParent] = "true";
                        }
                        else if (messageToParent === false) {
                            params[Microsoft.Office.WebExtension.Parameters.MessageToParent] = "";
                        }
                    }
                }
                if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                    window.external.MessageParent2(JSON.stringify(params), OsfOMToken);
                }
                else {
                    window.external.MessageParent2(JSON.stringify(params));
                }
            }
            else {
                var message = params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                window.external.MessageParent(message);
            }
        };
        Win32RichClientHostController.prototype.openDialog = function (id, targetId, handler, callback) {
            this.registerEvent(id, targetId, handler, callback);
        };
        Win32RichClientHostController.prototype.closeDialog = function (id, targetId, callback) {
            this.unregisterEvent(id, targetId, callback);
        };
        Win32RichClientHostController.prototype.sendMessage = function (params) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof window.external.MessageChild2)) {
                if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                    window.external.MessageChild2(JSON.stringify(params), OsfOMToken);
                }
                else {
                    window.external.MessageChild2(JSON.stringify(params));
                }
            }
            else {
                var message = params[Microsoft.Office.WebExtension.Parameters.MessageContent];
                window.external.MessageChild(message);
            }
        };
        return Win32RichClientHostController;
    }(OfficeExt.RichClientHostController));
    OfficeExt.Win32RichClientHostController = Win32RichClientHostController;
})(OfficeExt || (OfficeExt = {}));
(function (OfficeExt) {
    var OfficeTheme;
    (function (OfficeTheme) {
        var OfficeThemeManager = (function () {
            function OfficeThemeManager() {
                this._osfOfficeTheme = null;
                this._osfOfficeThemeTimeStamp = null;
            }
            OfficeThemeManager.prototype.getOfficeTheme = function () {
                if (OSF.DDA._OsfControlContext) {
                    if (this._osfOfficeTheme && this._osfOfficeThemeTimeStamp && ((new Date()).getTime() - this._osfOfficeThemeTimeStamp < OfficeThemeManager._osfOfficeThemeCacheValidPeriod)) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", 0);
                        }
                    }
                    else {
                        var startTime = (new Date()).getTime();
                        var osfOfficeTheme = OSF.DDA._OsfControlContext.GetOfficeThemeInfo();
                        var endTime = (new Date()).getTime();
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", Math.abs(endTime - startTime));
                        }
                        this._osfOfficeTheme = JSON.parse(osfOfficeTheme);
                        for (var color in this._osfOfficeTheme) {
                            this._osfOfficeTheme[color] = OSF.OUtil.convertIntToCssHexColor(this._osfOfficeTheme[color]);
                        }
                        this._osfOfficeThemeTimeStamp = endTime;
                    }
                    return this._osfOfficeTheme;
                }
            };
            OfficeThemeManager.instance = function () {
                if (OfficeThemeManager._instance == null) {
                    OfficeThemeManager._instance = new OfficeThemeManager();
                }
                return OfficeThemeManager._instance;
            };
            OfficeThemeManager._osfOfficeThemeCacheValidPeriod = 5000;
            OfficeThemeManager._instance = null;
            return OfficeThemeManager;
        }());
        OfficeTheme.OfficeThemeManager = OfficeThemeManager;
        OSF.OUtil.setNamespace("OfficeTheme", OSF.DDA);
        OSF.DDA.OfficeTheme.getOfficeTheme = OfficeExt.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme;
    })(OfficeTheme = OfficeExt.OfficeTheme || (OfficeExt.OfficeTheme = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.initializeRichCommon = function OSF_initializeRichCommon() {
    OSF.DDA.ClientSettingsManager = {
        getSettingsExecuteMethod: function OSF_DDA_ClientSettingsManager$getSettingsExecuteMethod(hostDelegateMethod) {
            return function (args) {
                var onComplete = function onComplete(status, response) {
                    if (args.onReceiving) {
                        args.onReceiving();
                    }
                    if (args.onComplete) {
                        args.onComplete(status, response);
                    }
                };
                var response;
                try {
                    response = hostDelegateMethod(args.hostCallArgs, args.onCalling, onComplete);
                }
                catch (ex) {
                    var status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                    response = { name: Strings.OfficeOM.L_InternalError, message: ex };
                    if (args.onComplete) {
                        args.onComplete(status, response);
                    }
                }
            };
        },
        read: function OSF_DDA_ClientSettingsManager$read(onCalling, onComplete) {
            var keys = [];
            var values = [];
            if (onCalling) {
                onCalling();
            }
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                OSF.DDA._OsfControlContext.GetSettings(OsfOMToken).Read(keys, values);
            }
            else {
                OSF.DDA._OsfControlContext.GetSettings().Read(keys, values);
            }
            var serializedSettings = {};
            for (var index = 0; index < keys.length; index++) {
                serializedSettings[keys[index]] = values[index];
            }
            if (onComplete) {
                onComplete(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, serializedSettings);
            }
            return serializedSettings;
        },
        write: function OSF_DDA_ClientSettingsManager$write(serializedSettings, overwriteIfStale, onCalling, onComplete) {
            var keys = [];
            var values = [];
            for (var key in serializedSettings) {
                keys.push(key);
                values.push(serializedSettings[key]);
            }
            if (onCalling) {
                onCalling();
            }
            var settingObj;
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                settingObj = OSF.DDA._OsfControlContext.GetSettings(OsfOMToken);
            }
            else {
                settingObj = OSF.DDA._OsfControlContext.GetSettings();
            }
            if (typeof settingObj.WriteAsync != 'undefined') {
                settingObj.WriteAsync(keys, values, onComplete);
            }
            else {
                settingObj.Write(keys, values);
                if (onComplete) {
                    onComplete(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                }
            }
        },
        refresh: function OSF_DDA_ClientSettingsManager$refresh(onCalling, onComplete) {
            var keys = [];
            var values = [];
            if (onCalling) {
                onCalling();
            }
            var osfSettingsObj;
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                osfSettingsObj = OSF.DDA._OsfControlContext.GetSettings(OsfOMToken);
            }
            else {
                osfSettingsObj = OSF.DDA._OsfControlContext.GetSettings();
            }
            var readSettingsAndReturn = function () {
                osfSettingsObj.Read(keys, values);
                var serializedSettings = {};
                for (var index = 0; index < keys.length; index++) {
                    serializedSettings[keys[index]] = values[index];
                }
                if (onComplete) {
                    onComplete(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, serializedSettings);
                }
            };
            if (osfSettingsObj.RefreshAsync) {
                osfSettingsObj.RefreshAsync(function () {
                    readSettingsAndReturn();
                });
            }
            else {
                readSettingsAndReturn();
            }
        }
    };
    OSF.InitializationHelper.prototype.initializeSettings = function OSF_InitializationHelper$initializeSettings(refreshSupported) {
        var serializedSettings = OSF.DDA.ClientSettingsManager.read();
        var settings = this.deserializeSettings(serializedSettings, refreshSupported);
        return settings;
    };
    OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
        var returnedContext;
        var context;
        var warningText = "Warning: Office.js is loaded outside of Office client";
        try {
            if (window.external && OSF.OUtil.externalNativeFunctionExists(typeof window.external.GetContext)) {
                context = OSF.DDA._OsfControlContext = window.external.GetContext();
            }
            else {
                OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
                return;
            }
        }
        catch (e) {
            OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
            return;
        }
        var appType = context.GetAppType();
        var id = context.GetSolutionRef();
        var version = context.GetAppVersionMajor();
        var minorVersion = context.GetAppVersionMinor();
        var UILocale = context.GetAppUILocale();
        var dataLocale = context.GetAppDataLocale();
        var docUrl = context.GetDocUrl();
        var clientMode = context.GetAppCapabilities();
        var reason = context.GetActivationMode();
        var osfControlType = context.GetControlIntegrationLevel();
        var settings = [];
        var eToken;
        try {
            eToken = context.GetSolutionToken();
        }
        catch (ex) {
        }
        var correlationId;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetCorrelationId)) {
            correlationId = context.GetCorrelationId();
        }
        var appInstanceId;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetInstanceId)) {
            appInstanceId = context.GetInstanceId();
        }
        var touchEnabled;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetTouchEnabled)) {
            touchEnabled = context.GetTouchEnabled();
        }
        var commerceAllowed;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetCommerceAllowed)) {
            commerceAllowed = context.GetCommerceAllowed();
        }
        var requirementMatrix;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetSupportedMatrix)) {
            requirementMatrix = context.GetSupportedMatrix();
        }
        var hostCustomMessage;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetHostCustomMessage)) {
            hostCustomMessage = context.GetHostCustomMessage();
        }
        var hostFullVersion;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetHostFullVersion)) {
            hostFullVersion = context.GetHostFullVersion();
        }
        var dialogRequirementMatrix;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetDialogRequirementMatrix)) {
            dialogRequirementMatrix = context.GetDialogRequirementMatrix();
        }
        var sdxFeatureGates;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetFeaturesForSolution)) {
            try {
                var sdxFeatureGatesJson = context.GetFeaturesForSolution();
                if (sdxFeatureGatesJson) {
                    sdxFeatureGates = JSON.parse(sdxFeatureGatesJson);
                }
            }
            catch (ex) {
                OsfMsAjaxFactory.msAjaxDebug.trace("Exception while creating the SDX FeatureGates object. Details: " + ex);
            }
        }
        var initialDisplayMode = 0;
        if (OSF.OUtil.externalNativeFunctionExists(typeof context.GetInitialDisplayMode)) {
            initialDisplayMode = context.GetInitialDisplayMode();
        }
        eToken = eToken ? eToken.toString() : "";
        returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, dialogRequirementMatrix, sdxFeatureGates, undefined, initialDisplayMode);
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.initialize(returnedContext);
        }
        gotAppContext(returnedContext);
    };
};
OSF.OUtil.setNamespace("ExtensionLifeCycle", OSF);
OSF.ExtensionLifeCycle.close = function OSF_ExtensionLifeCycle$close(context) {
    OSF.ClientHostController.closeSdxDialog(context);
};
OSF.ClientHostController = new OfficeExt.Win32RichClientHostController();
OSF.initializeRichCommon();
var OSFLog;
(function (OSFLog) {
    var BaseUsageData = (function () {
        function BaseUsageData(table) {
            this._table = table;
            this._fields = {};
        }
        Object.defineProperty(BaseUsageData.prototype, "Fields", {
            get: function () {
                return this._fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BaseUsageData.prototype, "Table", {
            get: function () {
                return this._table;
            },
            enumerable: true,
            configurable: true
        });
        BaseUsageData.prototype.SerializeFields = function () {
        };
        BaseUsageData.prototype.SetSerializedField = function (key, value) {
            if (typeof (value) !== "undefined" && value !== null) {
                this._serializedFields[key] = value.toString();
            }
        };
        BaseUsageData.prototype.SerializeRow = function () {
            this._serializedFields = {};
            this.SetSerializedField("Table", this._table);
            this.SerializeFields();
            return JSON.stringify(this._serializedFields);
        };
        return BaseUsageData;
    }());
    OSFLog.BaseUsageData = BaseUsageData;
    var AppActivatedUsageData = (function (_super) {
        __extends(AppActivatedUsageData, _super);
        function AppActivatedUsageData() {
            return _super.call(this, "AppActivated") || this;
        }
        Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
            get: function () { return this.Fields["AppId"]; },
            set: function (value) { this.Fields["AppId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
            get: function () { return this.Fields["AppInstanceId"]; },
            set: function (value) { this.Fields["AppInstanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
            get: function () { return this.Fields["AppURL"]; },
            set: function (value) { this.Fields["AppURL"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
            get: function () { return this.Fields["Browser"]; },
            set: function (value) { this.Fields["Browser"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
            get: function () { return this.Fields["UserId"]; },
            set: function (value) { this.Fields["UserId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
            get: function () { return this.Fields["Host"]; },
            set: function (value) { this.Fields["Host"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
            get: function () { return this.Fields["HostVersion"]; },
            set: function (value) { this.Fields["HostVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
            get: function () { return this.Fields["ClientId"]; },
            set: function (value) { this.Fields["ClientId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
            get: function () { return this.Fields["AppSizeWidth"]; },
            set: function (value) { this.Fields["AppSizeWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
            get: function () { return this.Fields["AppSizeHeight"]; },
            set: function (value) { this.Fields["AppSizeHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
            get: function () { return this.Fields["DocUrl"]; },
            set: function (value) { this.Fields["DocUrl"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
            get: function () { return this.Fields["OfficeJSVersion"]; },
            set: function (value) { this.Fields["OfficeJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
            get: function () { return this.Fields["HostJSVersion"]; },
            set: function (value) { this.Fields["HostJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
            get: function () { return this.Fields["WacHostEnvironment"]; },
            set: function (value) { this.Fields["WacHostEnvironment"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
            get: function () { return this.Fields["IsFromWacAutomation"]; },
            set: function (value) { this.Fields["IsFromWacAutomation"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "IsMOS", {
            get: function () { return this.Fields["IsMOS"]; },
            set: function (value) { this.Fields["IsMOS"] = value; },
            enumerable: true,
            configurable: true
        });
        AppActivatedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("AppId", this.AppId);
            this.SetSerializedField("AppInstanceId", this.AppInstanceId);
            this.SetSerializedField("AppURL", this.AppURL);
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("Browser", this.Browser);
            this.SetSerializedField("UserId", this.UserId);
            this.SetSerializedField("Host", this.Host);
            this.SetSerializedField("HostVersion", this.HostVersion);
            this.SetSerializedField("ClientId", this.ClientId);
            this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
            this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
            this.SetSerializedField("Message", this.Message);
            this.SetSerializedField("DocUrl", this.DocUrl);
            this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
            this.SetSerializedField("HostJSVersion", this.HostJSVersion);
            this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
            this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
            this.SetSerializedField("IsMOS", this.IsMOS);
        };
        return AppActivatedUsageData;
    }(BaseUsageData));
    OSFLog.AppActivatedUsageData = AppActivatedUsageData;
    var ScriptLoadUsageData = (function (_super) {
        __extends(ScriptLoadUsageData, _super);
        function ScriptLoadUsageData() {
            return _super.call(this, "ScriptLoad") || this;
        }
        Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
            get: function () { return this.Fields["ScriptId"]; },
            set: function (value) { this.Fields["ScriptId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
            get: function () { return this.Fields["StartTime"]; },
            set: function (value) { this.Fields["StartTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        ScriptLoadUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("ScriptId", this.ScriptId);
            this.SetSerializedField("StartTime", this.StartTime);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
        };
        return ScriptLoadUsageData;
    }(BaseUsageData));
    OSFLog.ScriptLoadUsageData = ScriptLoadUsageData;
    var AppClosedUsageData = (function (_super) {
        __extends(AppClosedUsageData, _super);
        function AppClosedUsageData() {
            return _super.call(this, "AppClosed") || this;
        }
        Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
            get: function () { return this.Fields["FocusTime"]; },
            set: function (value) { this.Fields["FocusTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
            get: function () { return this.Fields["AppSizeFinalWidth"]; },
            set: function (value) { this.Fields["AppSizeFinalWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
            get: function () { return this.Fields["AppSizeFinalHeight"]; },
            set: function (value) { this.Fields["AppSizeFinalHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
            get: function () { return this.Fields["OpenTime"]; },
            set: function (value) { this.Fields["OpenTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
            get: function () { return this.Fields["CloseMethod"]; },
            set: function (value) { this.Fields["CloseMethod"] = value; },
            enumerable: true,
            configurable: true
        });
        AppClosedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("FocusTime", this.FocusTime);
            this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
            this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
            this.SetSerializedField("OpenTime", this.OpenTime);
            this.SetSerializedField("CloseMethod", this.CloseMethod);
        };
        return AppClosedUsageData;
    }(BaseUsageData));
    OSFLog.AppClosedUsageData = AppClosedUsageData;
    var APIUsageUsageData = (function (_super) {
        __extends(APIUsageUsageData, _super);
        function APIUsageUsageData() {
            return _super.call(this, "APIUsage") || this;
        }
        Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
            get: function () { return this.Fields["APIType"]; },
            set: function (value) { this.Fields["APIType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
            get: function () { return this.Fields["APIID"]; },
            set: function (value) { this.Fields["APIID"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
            get: function () { return this.Fields["Parameters"]; },
            set: function (value) { this.Fields["Parameters"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
            get: function () { return this.Fields["ErrorType"]; },
            set: function (value) { this.Fields["ErrorType"] = value; },
            enumerable: true,
            configurable: true
        });
        APIUsageUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("APIType", this.APIType);
            this.SetSerializedField("APIID", this.APIID);
            this.SetSerializedField("Parameters", this.Parameters);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("ErrorType", this.ErrorType);
        };
        return APIUsageUsageData;
    }(BaseUsageData));
    OSFLog.APIUsageUsageData = APIUsageUsageData;
    var AppInitializationUsageData = (function (_super) {
        __extends(AppInitializationUsageData, _super);
        function AppInitializationUsageData() {
            return _super.call(this, "AppInitialization") || this;
        }
        Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
            get: function () { return this.Fields["SuccessCode"]; },
            set: function (value) { this.Fields["SuccessCode"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        AppInitializationUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("SuccessCode", this.SuccessCode);
            this.SetSerializedField("Message", this.Message);
        };
        return AppInitializationUsageData;
    }(BaseUsageData));
    OSFLog.AppInitializationUsageData = AppInitializationUsageData;
    var CheckWACHostUsageData = (function (_super) {
        __extends(CheckWACHostUsageData, _super);
        function CheckWACHostUsageData() {
            return _super.call(this, "CheckWACHost") || this;
        }
        Object.defineProperty(CheckWACHostUsageData.prototype, "isWacKnownHost", {
            get: function () { return this.Fields["isWacKnownHost"]; },
            set: function (value) { this.Fields["isWacKnownHost"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "instanceId", {
            get: function () { return this.Fields["instanceId"]; },
            set: function (value) { this.Fields["instanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "hostType", {
            get: function () { return this.Fields["hostType"]; },
            set: function (value) { this.Fields["hostType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "hostPlatform", {
            get: function () { return this.Fields["hostPlatform"]; },
            set: function (value) { this.Fields["hostPlatform"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CheckWACHostUsageData.prototype, "wacDomain", {
            get: function () { return this.Fields["wacDomain"]; },
            set: function (value) { this.Fields["wacDomain"] = value; },
            enumerable: true,
            configurable: true
        });
        CheckWACHostUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("isWacKnownHost", this.isWacKnownHost);
            this.SetSerializedField("instanceId", this.instanceId);
            this.SetSerializedField("hostType", this.hostType);
            this.SetSerializedField("hostPlatform", this.hostPlatform);
            this.SetSerializedField("wacDomain", this.wacDomain);
        };
        return CheckWACHostUsageData;
    }(BaseUsageData));
    OSFLog.CheckWACHostUsageData = CheckWACHostUsageData;
})(OSFLog || (OSFLog = {}));
var Logger;
(function (Logger) {
    "use strict";
    var TraceLevel;
    (function (TraceLevel) {
        TraceLevel[TraceLevel["info"] = 0] = "info";
        TraceLevel[TraceLevel["warning"] = 1] = "warning";
        TraceLevel[TraceLevel["error"] = 2] = "error";
    })(TraceLevel = Logger.TraceLevel || (Logger.TraceLevel = {}));
    var SendFlag;
    (function (SendFlag) {
        SendFlag[SendFlag["none"] = 0] = "none";
        SendFlag[SendFlag["flush"] = 1] = "flush";
    })(SendFlag = Logger.SendFlag || (Logger.SendFlag = {}));
    function allowUploadingData() {
    }
    Logger.allowUploadingData = allowUploadingData;
    function sendLog(traceLevel, message, flag) {
    }
    Logger.sendLog = sendLog;
    function creatULSEndpoint() {
        try {
            return new ULSEndpointProxy();
        }
        catch (e) {
            return null;
        }
    }
    var ULSEndpointProxy = (function () {
        function ULSEndpointProxy() {
        }
        ULSEndpointProxy.prototype.writeLog = function (log) {
        };
        ULSEndpointProxy.prototype.loadProxyFrame = function () {
        };
        return ULSEndpointProxy;
    }());
    if (!OSF.Logger) {
        OSF.Logger = Logger;
    }
    Logger.ulsEndpoint = creatULSEndpoint();
})(Logger || (Logger = {}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
    var TelemetryEventAppActivated = { name: "AppActivated", enabled: true, critical: true, points: [
            { name: "Browser", type: "string" },
            { name: "Message", type: "string" },
            { name: "Host", type: "string" },
            { name: "AppSizeWidth", type: "int64" },
            { name: "AppSizeHeight", type: "int64" },
            { name: "IsFromWacAutomation", type: "string" },
            { name: "IsMOS", type: "int64" },
        ] };
    var TelemetryEventScriptLoad = { name: "ScriptLoad", enabled: true, critical: false, points: [
            { name: "ScriptId", type: "string" },
            { name: "StartTime", type: "double" },
            { name: "ResponseTime", type: "double" },
        ] };
    var enableAPIUsage = shouldAPIUsageBeEnabled();
    var TelemetryEventApiUsage = { name: "APIUsage", enabled: enableAPIUsage, critical: false, points: [
            { name: "APIType", type: "string" },
            { name: "APIID", type: "int64" },
            { name: "Parameters", type: "string" },
            { name: "ResponseTime", type: "int64" },
            { name: "ErrorType", type: "int64" },
        ] };
    var TelemetryEventAppInitialization = { name: "AppInitialization", enabled: true, critical: false, points: [
            { name: "SuccessCode", type: "int64" },
            { name: "Message", type: "string" },
        ] };
    var TelemetryEventAppClosed = { name: "AppClosed", enabled: true, critical: false, points: [
            { name: "FocusTime", type: "int64" },
            { name: "AppSizeFinalWidth", type: "int64" },
            { name: "AppSizeFinalHeight", type: "int64" },
            { name: "OpenTime", type: "int64" },
        ] };
    var TelemetryEventCheckWACHost = { name: "CheckWACHost", enabled: true, critical: false, points: [
            { name: "isWacKnownHost", type: "int64" },
            { name: "solutionId", type: "string" },
            { name: "hostType", type: "string" },
            { name: "hostPlatform", type: "string" },
            { name: "correlationId", type: "string" },
        ] };
    var TelemetryEvents = [
        TelemetryEventAppActivated,
        TelemetryEventScriptLoad,
        TelemetryEventApiUsage,
        TelemetryEventAppInitialization,
        TelemetryEventAppClosed,
        TelemetryEventCheckWACHost,
    ];
    function createDataField(value, point) {
        var key = point.rename === undefined ? point.name : point.rename;
        var type = point.type;
        var field = undefined;
        switch (type) {
            case "string":
                field = oteljs.makeStringDataField(key, value);
                break;
            case "double":
                if (typeof value === "string") {
                    value = parseFloat(value);
                }
                field = oteljs.makeDoubleDataField(key, value);
                break;
            case "int64":
                if (typeof value === "string") {
                    value = parseInt(value);
                }
                field = oteljs.makeInt64DataField(key, value);
                break;
            case "boolean":
                if (typeof value === "string") {
                    value = value === "true";
                }
                field = oteljs.makeBooleanDataField(key, value);
                break;
        }
        return field;
    }
    function getEventDefinition(eventName) {
        for (var _i = 0, TelemetryEvents_1 = TelemetryEvents; _i < TelemetryEvents_1.length; _i++) {
            var event_1 = TelemetryEvents_1[_i];
            if (event_1.name === eventName) {
                return event_1;
            }
        }
        return undefined;
    }
    function eventEnabled(eventName) {
        var eventDefinition = getEventDefinition(eventName);
        if (eventDefinition === undefined) {
            return false;
        }
        return eventDefinition.enabled;
    }
    function shouldAPIUsageBeEnabled() {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo) {
            return false;
        }
        var hostInfo = OSF._OfficeAppFactory.getHostInfo();
        if (!hostInfo) {
            return false;
        }
        switch (hostInfo["hostType"]) {
            case "outlook":
                switch (hostInfo["hostPlatform"]) {
                    case "mac":
                    case "web":
                        return true;
                    default:
                        return false;
                }
            default:
                return false;
        }
    }
    function generateTelemetryEvent(eventName, telemetryData) {
        var eventDefinition = getEventDefinition(eventName);
        if (eventDefinition === undefined) {
            return undefined;
        }
        var dataFields = [];
        for (var _i = 0, _a = eventDefinition.points; _i < _a.length; _i++) {
            var point = _a[_i];
            var key = point.name;
            var value = telemetryData[key];
            if (value === undefined) {
                continue;
            }
            var field = createDataField(value, point);
            if (field !== undefined) {
                dataFields.push(field);
            }
        }
        var flags = { dataCategories: oteljs.DataCategories.ProductServiceUsage };
        if (eventDefinition.critical) {
            flags.samplingPolicy = oteljs.SamplingPolicy.CriticalBusinessImpact;
        }
        flags.diagnosticLevel = oteljs.DiagnosticLevel.NecessaryServiceDataEvent;
        var eventNameFull = "Office.Extensibility.OfficeJs." + eventName + "X";
        var event = { eventName: eventNameFull, dataFields: dataFields, eventFlags: flags };
        return event;
    }
    function sendOtelTelemetryEvent(eventName, telemetryData) {
        if (eventEnabled(eventName)) {
            if (typeof OTel !== "undefined") {
                OTel.OTelLogger.onTelemetryLoaded(function () {
                    var event = generateTelemetryEvent(eventName, telemetryData);
                    if (event === undefined) {
                        return;
                    }
                    Microsoft.Office.WebExtension.sendTelemetryEvent(event);
                });
            }
        }
    }
    var AriaLogger = (function () {
        function AriaLogger() {
        }
        AriaLogger.prototype.getAriaCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + "ariatelemetry/aria-web-telemetry.js");
        };
        AriaLogger.getInstance = function () {
            if (AriaLogger.AriaLoggerObj === undefined) {
                AriaLogger.AriaLoggerObj = new AriaLogger();
            }
            return AriaLogger.AriaLoggerObj;
        };
        AriaLogger.prototype.isIUsageData = function (arg) {
            return arg["Fields"] !== undefined;
        };
        AriaLogger.prototype.shouldSendDirectToAria = function (flavor, version) {
            var BASE10 = 10;
            var MAX_VERSION_WIN32 = [16, 0, 11601];
            var MAX_VERSION_MAC = [16, 28];
            var max_version;
            if (!flavor) {
                return false;
            }
            else if (flavor.toLowerCase() === "win32") {
                max_version = MAX_VERSION_WIN32;
            }
            else if (flavor.toLowerCase() === "mac") {
                max_version = MAX_VERSION_MAC;
            }
            else {
                return true;
            }
            if (!version) {
                return false;
            }
            var versionTokens = version.split('.');
            for (var i = 0; i < max_version.length && i < versionTokens.length; i++) {
                var versionToken = parseInt(versionTokens[i], BASE10);
                if (isNaN(versionToken)) {
                    return false;
                }
                if (versionToken < max_version[i]) {
                    return true;
                }
                if (versionToken > max_version[i]) {
                    return false;
                }
            }
            return false;
        };
        AriaLogger.prototype.isDirectToAriaEnabled = function () {
            if (this.EnableDirectToAria === undefined || this.EnableDirectToAria === null) {
                var flavor = void 0;
                var version = void 0;
                if (OSF._OfficeAppFactory && OSF._OfficeAppFactory.getHostInfo) {
                    flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
                }
                if (window.external && typeof window.external.GetContext !== "undefined" && typeof window.external.GetContext().GetHostFullVersion !== "undefined") {
                    version = window.external.GetContext().GetHostFullVersion();
                }
                this.EnableDirectToAria = this.shouldSendDirectToAria(flavor, version);
            }
            return this.EnableDirectToAria;
        };
        AriaLogger.prototype.sendTelemetry = function (tableName, telemetryData) {
            var startAfterMs = 1000;
            var sendAriaEnabled = AriaLogger.EnableSendingTelemetryWithLegacyAria && this.isDirectToAriaEnabled();
            if (sendAriaEnabled) {
                OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
                    try {
                        if (!this.ALogger) {
                            var OfficeExtensibilityTenantID = "db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
                            this.ALogger = AWTLogManager.initialize(OfficeExtensibilityTenantID);
                        }
                        var eventProperties = new AWTEventProperties();
                        eventProperties.setName("Office.Extensibility.OfficeJS." + tableName);
                        for (var key in telemetryData) {
                            if (key.toLowerCase() !== "table") {
                                eventProperties.setProperty(key, telemetryData[key]);
                            }
                        }
                        var today = new Date();
                        eventProperties.setProperty("Date", today.toISOString());
                        this.ALogger.logEvent(eventProperties);
                    }
                    catch (e) {
                    }
                }, startAfterMs);
            }
            if (AriaLogger.EnableSendingTelemetryWithOTel) {
                sendOtelTelemetryEvent(tableName, telemetryData);
            }
        };
        AriaLogger.prototype.logData = function (data) {
            if (this.isIUsageData(data)) {
                this.sendTelemetry(data["Table"], data["Fields"]);
            }
            else {
                this.sendTelemetry(data["Table"], data);
            }
        };
        AriaLogger.EnableSendingTelemetryWithOTel = true;
        AriaLogger.EnableSendingTelemetryWithLegacyAria = false;
        return AriaLogger;
    }());
    OSFAriaLogger.AriaLogger = AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger = {}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
    "use strict";
    var appInfo;
    var sessionId = OSF.OUtil.Guid.generateNewGuid();
    var osfControlAppCorrelationId = "";
    var omexDomainRegex = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
    var privateAddinId = "PRIVATE";
    OSFAppTelemetry.enableTelemetry = true;
    ;
    var AppInfo = (function () {
        function AppInfo() {
        }
        return AppInfo;
    }());
    OSFAppTelemetry.AppInfo = AppInfo;
    var Event = (function () {
        function Event(name, handler) {
            this.name = name;
            this.handler = handler;
        }
        return Event;
    }());
    var AppStorage = (function () {
        function AppStorage() {
            this.clientIDKey = "Office API client";
            this.logIdSetKey = "Office App Log Id Set";
        }
        AppStorage.prototype.getClientId = function () {
            var clientId = this.getValue(this.clientIDKey);
            if (!clientId || clientId.length <= 0 || clientId.length > 40) {
                clientId = OSF.OUtil.Guid.generateNewGuid();
                this.setValue(this.clientIDKey, clientId);
            }
            return clientId;
        };
        AppStorage.prototype.saveLog = function (logId, log) {
            var logIdSet = this.getValue(this.logIdSetKey);
            logIdSet = ((logIdSet && logIdSet.length > 0) ? (logIdSet + ";") : "") + logId;
            this.setValue(this.logIdSetKey, logIdSet);
            this.setValue(logId, log);
        };
        AppStorage.prototype.enumerateLog = function (callback, clean) {
            var logIdSet = this.getValue(this.logIdSetKey);
            if (logIdSet) {
                var ids = logIdSet.split(";");
                for (var id in ids) {
                    var logId = ids[id];
                    var log = this.getValue(logId);
                    if (log) {
                        if (callback) {
                            callback(logId, log);
                        }
                        if (clean) {
                            this.remove(logId);
                        }
                    }
                }
                if (clean) {
                    this.remove(this.logIdSetKey);
                }
            }
        };
        AppStorage.prototype.getValue = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            var value = "";
            if (osfLocalStorage) {
                value = osfLocalStorage.getItem(key);
            }
            return value;
        };
        AppStorage.prototype.setValue = function (key, value) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                osfLocalStorage.setItem(key, value);
            }
        };
        AppStorage.prototype.remove = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                try {
                    osfLocalStorage.removeItem(key);
                }
                catch (ex) {
                }
            }
        };
        return AppStorage;
    }());
    var AppLogger = (function () {
        function AppLogger() {
        }
        AppLogger.prototype.LogData = function (data) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(data);
            }
            catch (e) {
            }
        };
        AppLogger.prototype.LogRawData = function (log) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
            }
            catch (e) {
            }
        };
        return AppLogger;
    }());
    function trimStringToLowerCase(input) {
        if (input) {
            input = input.replace(/[{}]/g, "").toLowerCase();
        }
        return (input || "");
    }
    function initialize(context) {
        if (!OSFAppTelemetry.enableTelemetry) {
            return;
        }
        if (appInfo) {
            return;
        }
        appInfo = new AppInfo();
        if (context.get_hostFullVersion()) {
            appInfo.hostVersion = context.get_hostFullVersion();
        }
        else {
            appInfo.hostVersion = context.get_appVersion();
        }
        appInfo.appId = canSendAddinId() ? context.get_id() : privateAddinId;
        appInfo.marketplaceType = context._marketplaceType;
        appInfo.browser = window.navigator.userAgent;
        appInfo.correlationId = trimStringToLowerCase(context.get_correlationId());
        appInfo.clientId = (new AppStorage()).getClientId();
        appInfo.appInstanceId = context.get_appInstanceId();
        if (appInfo.appInstanceId) {
            appInfo.appInstanceId = trimStringToLowerCase(appInfo.appInstanceId);
            appInfo.appInstanceId = getCompliantAppInstanceId(context.get_id(), appInfo.appInstanceId);
        }
        appInfo.message = context.get_hostCustomMessage();
        appInfo.officeJSVersion = OSF.ConstantNames.FileVersion;
        appInfo.hostJSVersion = "16.0.15928.10000";
        if (context._wacHostEnvironment) {
            appInfo.wacHostEnvironment = context._wacHostEnvironment;
        }
        if (context._isFromWacAutomation !== undefined && context._isFromWacAutomation !== null) {
            appInfo.isFromWacAutomation = context._isFromWacAutomation.toString().toLowerCase();
        }
        var docUrl = context.get_docUrl();
        appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
        var url = location.href;
        if (url) {
            url = url.split("?")[0].split("#")[0];
        }
        appInfo.isMos = isMos();
        appInfo.appURL = "";
        (function getUserIdAndAssetIdFromToken(token, appInfo) {
            var xmlContent;
            var parser;
            var xmlDoc;
            appInfo.assetId = "";
            appInfo.userId = "";
            try {
                xmlContent = decodeURIComponent(token);
                parser = new DOMParser();
                xmlDoc = parser.parseFromString(xmlContent, "text/xml");
                var cidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
                var oidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                if (cidNode && cidNode.nodeValue) {
                    appInfo.userId = cidNode.nodeValue;
                }
                else if (oidNode && oidNode.nodeValue) {
                    appInfo.userId = oidNode.nodeValue;
                }
                appInfo.assetId = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
            }
            catch (e) {
            }
            finally {
                xmlContent = null;
                xmlDoc = null;
                parser = null;
            }
        })(context.get_eToken(), appInfo);
        appInfo.sessionId = sessionId;
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.initialize(appInfo);
        }
        (function handleLifecycle() {
            var startTime = new Date();
            var lastFocus = null;
            var focusTime = 0;
            var finished = false;
            var adjustFocusTime = function () {
                if (document.hasFocus()) {
                    if (lastFocus == null) {
                        lastFocus = new Date();
                    }
                }
                else if (lastFocus) {
                    focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                    lastFocus = null;
                }
            };
            var eventList = [];
            eventList.push(new Event("focus", adjustFocusTime));
            eventList.push(new Event("blur", adjustFocusTime));
            eventList.push(new Event("focusout", adjustFocusTime));
            eventList.push(new Event("focusin", adjustFocusTime));
            var exitFunction = function () {
                for (var i = 0; i < eventList.length; i++) {
                    OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
                }
                eventList.length = 0;
                if (!finished) {
                    if (document.hasFocus() && lastFocus) {
                        focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                        lastFocus = null;
                    }
                    OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
                    finished = true;
                }
            };
            eventList.push(new Event("beforeunload", exitFunction));
            eventList.push(new Event("unload", exitFunction));
            for (var i = 0; i < eventList.length; i++) {
                OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
            }
            adjustFocusTime();
        })();
        OSFAppTelemetry.onAppActivated();
    }
    OSFAppTelemetry.initialize = initialize;
    function onAppActivated() {
        if (!appInfo) {
            return;
        }
        (new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
        var data = new OSFLog.AppActivatedUsageData();
        data.SessionId = sessionId;
        data.AppId = appInfo.appId;
        data.AssetId = appInfo.assetId;
        data.AppURL = "";
        data.UserId = "";
        data.ClientId = appInfo.clientId;
        data.Browser = appInfo.browser;
        data.HostVersion = appInfo.hostVersion;
        data.CorrelationId = trimStringToLowerCase(appInfo.correlationId);
        data.AppSizeWidth = window.innerWidth;
        data.AppSizeHeight = window.innerHeight;
        data.AppInstanceId = appInfo.appInstanceId;
        data.Message = appInfo.message;
        data.DocUrl = appInfo.docUrl;
        data.OfficeJSVersion = appInfo.officeJSVersion;
        data.HostJSVersion = appInfo.hostJSVersion;
        if (appInfo.wacHostEnvironment) {
            data.WacHostEnvironment = appInfo.wacHostEnvironment;
        }
        if (appInfo.isFromWacAutomation !== undefined && appInfo.isFromWacAutomation !== null) {
            data.IsFromWacAutomation = appInfo.isFromWacAutomation;
        }
        data.IsMOS = appInfo.isMos ? 1 : 0;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onAppActivated = onAppActivated;
    function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
        var data = new OSFLog.ScriptLoadUsageData();
        data.CorrelationId = trimStringToLowerCase(appCorrelationId);
        data.SessionId = sessionId;
        data.ScriptId = scriptId;
        data.StartTime = msStartTime;
        data.ResponseTime = msResponseTime;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onScriptDone = onScriptDone;
    function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
        if (!appInfo) {
            return;
        }
        if (!isAllowedHost() || !isAPIUsageEnabledDispId(id, apiType)) {
            return;
        }
        var data = new OSFLog.APIUsageUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.APIType = apiType;
        data.APIID = id;
        data.Parameters = parameters;
        data.ResponseTime = msResponseTime;
        data.ErrorType = errorType;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onCallDone = onCallDone;
    ;
    function onMethodDone(id, args, msResponseTime, errorType) {
        var parameters = null;
        if (args) {
            if (typeof args == "number") {
                parameters = String(args);
            }
            else if (typeof args === "object") {
                for (var index in args) {
                    if (parameters !== null) {
                        parameters += ",";
                    }
                    else {
                        parameters = "";
                    }
                    if (typeof args[index] == "number") {
                        parameters += String(args[index]);
                    }
                }
            }
            else {
                parameters = "";
            }
        }
        OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
    }
    OSFAppTelemetry.onMethodDone = onMethodDone;
    function onPropertyDone(propertyName, msResponseTime) {
        OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
    }
    OSFAppTelemetry.onPropertyDone = onPropertyDone;
    function onCheckWACHost(isWacKnownHost, instanceId, hostType, hostPlatform, wacDomain) {
        var data = new OSFLog.CheckWACHostUsageData();
        data.isWacKnownHost = isWacKnownHost;
        data.instanceId = instanceId;
        data.hostType = hostType;
        data.hostPlatform = hostPlatform;
        data.wacDomain = "";
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onCheckWACHost = onCheckWACHost;
    function onEventDone(id, errorType) {
        OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
    }
    OSFAppTelemetry.onEventDone = onEventDone;
    function onRegisterDone(register, id, msResponseTime, errorType) {
        OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
    }
    OSFAppTelemetry.onRegisterDone = onRegisterDone;
    function onAppClosed(openTime, focusTime) {
        if (!appInfo) {
            return;
        }
        var data = new OSFLog.AppClosedUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.FocusTime = focusTime;
        data.OpenTime = openTime;
        data.AppSizeFinalWidth = window.innerWidth;
        data.AppSizeFinalHeight = window.innerHeight;
        (new AppStorage()).saveLog(sessionId, data.SerializeRow());
    }
    OSFAppTelemetry.onAppClosed = onAppClosed;
    function setOsfControlAppCorrelationId(correlationId) {
        osfControlAppCorrelationId = trimStringToLowerCase(correlationId);
    }
    OSFAppTelemetry.setOsfControlAppCorrelationId = setOsfControlAppCorrelationId;
    function doAppInitializationLogging(isException, message) {
        var data = new OSFLog.AppInitializationUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.SuccessCode = isException ? 1 : 0;
        data.Message = message;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.doAppInitializationLogging = doAppInitializationLogging;
    function logAppCommonMessage(message) {
        doAppInitializationLogging(false, message);
    }
    OSFAppTelemetry.logAppCommonMessage = logAppCommonMessage;
    function logAppException(errorMessage) {
        doAppInitializationLogging(true, errorMessage);
    }
    OSFAppTelemetry.logAppException = logAppException;
    function isAllowedHost() {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo) {
            return false;
        }
        var hostInfo = OSF._OfficeAppFactory.getHostInfo();
        if (!hostInfo) {
            return false;
        }
        switch (hostInfo["hostType"]) {
            case "outlook":
                switch (hostInfo["hostPlatform"]) {
                    case "mac":
                    case "web":
                        return true;
                    default:
                        return false;
                }
            default:
                return false;
        }
    }
    function isAPIUsageEnabledDispId(dispId, apiType) {
        if (apiType === "method") {
            switch (dispId) {
                case 4:
                case 38:
                case 37:
                case 10:
                case 12:
                    return true;
                default:
                    return false;
            }
        }
        return false;
    }
    function canSendAddinId() {
        var isPublic = (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.PublicAddin) != 0;
        if (isPublic) {
            return isPublic;
        }
        if (!appInfo) {
            return false;
        }
        var hostPlatform = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
        var hostVersion = appInfo.hostVersion;
        return _isComplianceExceptedHost(hostPlatform, hostVersion);
    }
    OSFAppTelemetry.canSendAddinId = canSendAddinId;
    function getCompliantAppInstanceId(addinId, appInstanceId) {
        if (!canSendAddinId() && appInstanceId === addinId) {
            return privateAddinId;
        }
        return appInstanceId;
    }
    OSFAppTelemetry.getCompliantAppInstanceId = getCompliantAppInstanceId;
    function _isComplianceExceptedHost(hostPlatform, hostVersion) {
        var excepted = false;
        var versionExtractor = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/;
        var result = versionExtractor.exec(hostVersion);
        if (result) {
            var major = parseInt(result[1]);
            var minor = parseInt(result[2]);
            var build = parseInt(result[3]);
            if (hostPlatform == "win32") {
                if (major < 16 || major == 16 && build < 14225) {
                    excepted = true;
                }
            }
            else if (hostPlatform == "mac") {
                if (major < 16 || (major == 16 && (minor < 52 || minor == 52 && build < 808))) {
                    excepted = true;
                }
            }
        }
        return excepted;
    }
    OSFAppTelemetry._isComplianceExceptedHost = _isComplianceExceptedHost;
    function isMos() {
        return (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.IsMos) != 0;
    }
    OSFAppTelemetry.isMos = isMos;
    OSF.AppTelemetry = OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry = {}));
Microsoft.Office.WebExtension.FileType = {
    Text: "text",
    Compressed: "compressed",
    Pdf: "pdf"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    FileProperties: "FileProperties",
    FileSliceProperties: "FileSliceProperties"
});
OSF.DDA.FileProperties = {
    Handle: "FileHandle",
    FileSize: "FileSize",
    SliceSize: Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.File = function OSF_DDA_File(handle, fileSize, sliceSize) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "size": {
            value: fileSize
        },
        "sliceCount": {
            value: Math.ceil(fileSize / sliceSize)
        }
    });
    var privateState = {};
    privateState[OSF.DDA.FileProperties.Handle] = handle;
    privateState[OSF.DDA.FileProperties.SliceSize] = sliceSize;
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.GetDocumentCopyChunkAsync,
        am.ReleaseDocumentCopyAsync
    ], privateState);
};
OSF.DDA.FileSliceOffset = "fileSliceoffset";
OSF.DDA.AsyncMethodNames.addNames({
    GetDocumentCopyAsync: "getFileAsync",
    GetDocumentCopyChunkAsync: "getSliceAsync",
    ReleaseDocumentCopyAsync: "closeAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.FileType,
            "enum": Microsoft.Office.WebExtension.FileType
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.SliceSize,
            value: {
                "types": ["number"],
                "defaultValue": 4 * 1024 * 1024
            }
        }
    ],
    checkCallArgs: function (callArgs, caller, stateInfo) {
        var sliceSize = callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize];
        if (sliceSize <= 0 || sliceSize > (4 * 1024 * 1024)) {
            throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;
        }
        return callArgs;
    },
    onSucceeded: function (fileDescriptor, caller, callArgs) {
        return new OSF.DDA.File(fileDescriptor[OSF.DDA.FileProperties.Handle], fileDescriptor[OSF.DDA.FileProperties.FileSize], callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize]);
    }
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.SliceIndex,
            "types": ["number"]
        }
    ],
    privateStateCallbacks: [
        {
            name: OSF.DDA.FileProperties.Handle,
            value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
        },
        {
            name: OSF.DDA.FileProperties.SliceSize,
            value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.SliceSize]; }
        }
    ],
    checkCallArgs: function (callArgs, caller, stateInfo) {
        var index = callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex];
        if (index < 0 || index >= caller.sliceCount) {
            throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
        }
        callArgs[OSF.DDA.FileSliceOffset] = parseInt((index * stateInfo[OSF.DDA.FileProperties.SliceSize]).toString());
        return callArgs;
    },
    onSucceeded: function (sliceDescriptor, caller, callArgs) {
        var slice = {};
        OSF.OUtil.defineEnumerableProperties(slice, {
            "data": {
                value: sliceDescriptor[Microsoft.Office.WebExtension.Parameters.Data]
            },
            "index": {
                value: callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex]
            },
            "size": {
                value: sliceDescriptor[OSF.DDA.FileProperties.SliceSize]
            }
        });
        return slice;
    }
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
    privateStateCallbacks: [
        {
            name: OSF.DDA.FileProperties.Handle,
            value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
        }
    ]
});
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function OSF_EventDispatch(eventTypes) {
    this._eventHandlers = {};
    this._objectEventHandlers = {};
    this._queuedEventsArgs = {};
    if (eventTypes != null) {
        for (var i = 0; i < eventTypes.length; i++) {
            var eventType = eventTypes[i];
            var isObjectEvent = (eventType == "objectDeleted" || eventType == "objectSelectionChanged" || eventType == "objectDataChanged" || eventType == "contentControlAdded");
            if (!isObjectEvent)
                this._eventHandlers[eventType] = [];
            else
                this._objectEventHandlers[eventType] = {};
            this._queuedEventsArgs[eventType] = [];
        }
    }
};
OSF.EventDispatch.prototype = {
    getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
        var events = [];
        for (var eventName in this._eventHandlers)
            events.push(eventName);
        for (var eventName in this._objectEventHandlers)
            events.push(eventName);
        return events;
    },
    supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
        for (var eventName in this._eventHandlers) {
            if (event == eventName)
                return true;
        }
        for (var eventName in this._objectEventHandlers) {
            if (event == eventName)
                return true;
        }
        return false;
    },
    hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var i = 0; i < handlers.length; i++) {
                if (handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._eventHandlers[eventType];
        if (handlers && !this.hasEventHandler(eventType, handler)) {
            handlers.push(handler);
            return true;
        }
        else {
            return false;
        }
    },
    addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._objectEventHandlers[eventType];
        if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
            if (handlers[objectId] == null)
                handlers[objectId] = [];
            handlers[objectId].push(handler);
            return true;
        }
        return false;
    },
    addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        var isFirstHandler = handlers.length == 0;
        var succeed = this.addEventHandler(eventType, handler);
        if (isFirstHandler && succeed) {
            this.fireQueuedEvent(eventType);
        }
        return succeed;
    },
    removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var index = 0; index < handlers.length; index++) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    return true;
                }
            }
        }
        return false;
    },
    removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler) {
                    _handlers.splice(i, 1);
                    return true;
                }
            }
        }
        return false;
    },
    clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
        if (typeof this._eventHandlers[eventType] != "undefined" && this._eventHandlers[eventType].length > 0) {
            this._eventHandlers[eventType] = [];
            return true;
        }
        return false;
    },
    clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
        if (this._objectEventHandlers[eventType] != null && this._objectEventHandlers[eventType][objectId] != null) {
            this._objectEventHandlers[eventType][objectId] = [];
            return true;
        }
        return false;
    },
    getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
        return this._eventHandlers[eventType] != undefined ? this._eventHandlers[eventType].length : -1;
    },
    getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
        if (this._objectEventHandlers[eventType] == null || this._objectEventHandlers[eventType][objectId] == null)
            return 0;
        return this._objectEventHandlers[eventType][objectId].length;
    },
    fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            for (var i = 0; i < eventHandlers.length; i++) {
                eventHandlers[i](eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._objectEventHandlers[eventType]) {
            var eventHandlers = this._objectEventHandlers[eventType];
            var _handlers = eventHandlers[objectId];
            if (_handlers != null) {
                for (var i = 0; i < _handlers.length; i++)
                    _handlers[i](eventArgs);
                return true;
            }
        }
        return false;
    },
    fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length == 0) {
                queuedEvents.push(eventArgs);
            }
            else {
                this.fireEvent(eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length > 0) {
                var eventHandler = eventHandlers[0];
                while (queuedEvents.length > 0) {
                    var eventArgs = queuedEvents.shift();
                    eventHandler(eventArgs);
                }
                return true;
            }
        }
        return false;
    },
    clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (queuedEvents) {
                this._queuedEventsArgs[eventType] = [];
            }
        }
    }
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
    var args;
    switch (eventType) {
        case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
            args = new OSF.DDA.DocumentSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
            args = new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
            args = new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
            break;
        case Microsoft.Office.WebExtension.EventType.SettingsChanged:
            args = new OSF.DDA.SettingsChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
            args = new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
            args = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
            args = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
            args = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook" && OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac") {
                OSF.DDA.convertOlkAppointmentTimeToDateFormat(args);
            }
            break;
        case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
        case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
        case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
        case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
            args = new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
            break;
        case Microsoft.Office.WebExtension.EventType.RichApiMessage:
            args = new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
            args = new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
            args = new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
            args = new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
            args = new OSF.DDA.TaskSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
            args = new OSF.DDA.ResourceSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
            args = new OSF.DDA.ViewSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
            args = new OSF.DDA.DialogEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
            args = new OSF.DDA.DialogParentEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ItemChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
                target.initialize(args["initialData"]);
                if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac") {
                    target.setCurrentItemNumber(args["itemNumber"].itemNumber);
                }
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.AttachmentsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkAttachmentsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkEnhancedLocationsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.InfobarClicked:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkInfobarClickedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.SelectedItemsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkSelectedItemsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.SensitivityLabelChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkSensitivityLabelChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        default:
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
    }
    return args;
};
OSF.DDA.AsyncMethodNames.addNames({
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
    requiredArguments: [{
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        },
        {
            "name": Microsoft.Office.WebExtension.Parameters.Handler,
            "types": ["function"]
        }
    ],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            value: {
                "types": ["function", "object"],
                "defaultValue": null
            }
        }
    ],
    privateStateCallbacks: []
});
OSF.DialogShownStatus = { hasDialogShown: false, isWindowDialog: false };
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    MessageType: "messageType",
    MessageContent: "messageContent",
    MessageOrigin: "messageOrigin"
});
OSF.DDA.DialogEventType = {};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
    DialogClosed: "dialogClosed",
    NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
    DisplayDialogAsync: "displayDialogAsync",
    DisplayModalDialogAsync: "displayModalDialogAsync",
    CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
    MessageParent: "messageParent",
    MessageChild: "messageChild",
    SendMessage: "sendMessage",
    AddMessageHandler: "addEventHandler"
});
OSF.DDA.UI.ParentUI = function OSF_DDA_ParentUI() {
    var eventDispatch;
    if (Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived != null) {
        eventDispatch = new OSF.EventDispatch([
            Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
            Microsoft.Office.WebExtension.EventType.DialogEventReceived,
            Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
        ]);
    }
    else {
        eventDispatch = new OSF.EventDispatch([
            Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
            Microsoft.Office.WebExtension.EventType.DialogEventReceived
        ]);
    }
    var target = this;
    var defineDialogApi = function (apiName, isModalApi) {
        if (!target[apiName]) {
            OSF.OUtil.defineEnumerableProperty(target, apiName, {
                value: function () {
                    var openDialog = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
                    openDialog(arguments, eventDispatch, target, isModalApi);
                }
            });
        }
    };
    defineDialogApi(OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName, false);
    if (Microsoft.Office.WebExtension.FeatureGates["ModalWebDialogAPI"]) {
        defineDialogApi(OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync.displayName, true);
    }
    OSF.OUtil.finalizeProperties(this);
};
OSF.DDA.UI.ChildUI = function OSF_DDA_ChildUI(isPopupWindow) {
    var messageParentName = OSF.DDA.SyncMethodNames.MessageParent.displayName;
    var target = this;
    if (!target[messageParentName]) {
        OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
            value: function () {
                var messageParent = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
                return messageParent(arguments, target);
            }
        });
    }
    var addEventHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
    if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch != "undefined") {
        OSF.DDA.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
    }
    OSF.OUtil.finalizeProperties(this);
};
OSF.DialogHandler = function OSF_DialogHandler() { };
OSF.DDA.DialogEventArgs = function OSF_DDA_DialogEventArgs(message) {
    if (message[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogMessageReceived) {
        OSF.OUtil.defineEnumerableProperties(this, {
            "type": {
                value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
            },
            "message": {
                value: message[OSF.DDA.PropertyDescriptors.MessageContent]
            },
            "origin": {
                value: message[OSF.DDA.PropertyDescriptors.MessageOrigin]
            }
        });
    }
    else {
        OSF.OUtil.defineEnumerableProperties(this, {
            "type": {
                value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
            },
            "error": {
                value: message[OSF.DDA.PropertyDescriptors.MessageType]
            }
        });
    }
};
OSF.DDA.DialogParentEventArgs = function OSF_DDA_DialogParentEventArgs(message) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
        },
        "message": {
            value: message[OSF.DDA.PropertyDescriptors.MessageContent]
        },
        "origin": {
            value: message[OSF.DDA.PropertyDescriptors.MessageOrigin]
        }
    });
};
var DialogApiManager = (function () {
    function DialogApiManager() {
    }
    DialogApiManager.defineApi = function (apiName, supportedOptions) {
        var asyncMethodCalls = OSF.DDA.AsyncMethodCalls;
        asyncMethodCalls.define({
            method: apiName,
            requiredArguments: [
                {
                    "name": Microsoft.Office.WebExtension.Parameters.Url,
                    "types": ["string"]
                }
            ],
            supportedOptions: supportedOptions,
            privateStateCallbacks: [],
            onSucceeded: function (args, caller, callArgs) {
                var targetId = args[Microsoft.Office.WebExtension.Parameters.Id];
                var eventDispatch = args[Microsoft.Office.WebExtension.Parameters.Data];
                var dialog = new OSF.DialogHandler();
                var closeDialog = OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
                OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
                    value: function () {
                        var closeDialogfunction = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
                        closeDialogfunction(arguments, targetId, eventDispatch, dialog);
                    }
                });
                var addHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
                OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
                    value: function () {
                        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
                        var callArgs = syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
                        var eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
                        var handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
                        return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
                    }
                });
                if (OSF.DDA.UI.EnableSendMessageDialogAPI === true) {
                    var sendMessage = OSF.DDA.SyncMethodNames.SendMessage.displayName;
                    OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
                        value: function () {
                            var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                            return execute(arguments, eventDispatch, dialog);
                        }
                    });
                }
                if (OSF.DDA.UI.EnableMessageChildDialogAPI === true) {
                    var messageChild = OSF.DDA.SyncMethodNames.MessageChild.displayName;
                    OSF.OUtil.defineEnumerableProperty(dialog, messageChild, {
                        value: function () {
                            var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                            return execute(arguments, eventDispatch, dialog);
                        }
                    });
                }
                return dialog;
            },
            checkCallArgs: function (callArgs, caller, stateInfo) {
                if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] <= 0) {
                    callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 1;
                }
                if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100) {
                    callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 99;
                }
                if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] <= 0) {
                    callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 1;
                }
                if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100) {
                    callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 99;
                }
                if (!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]) {
                    callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs] = true;
                }
                return callArgs;
            }
        });
    };
    DialogApiManager.messageChildRichApiBridge = function () {
        if (OSF.DDA.UI.EnableMessageChildDialogAPI === true) {
            var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
            return execute(arguments, null, null);
        }
    };
    DialogApiManager.initOnce = function () {
        DialogApiManager.defineApi(OSF.DDA.AsyncMethodNames.DisplayDialogAsync, DialogApiManager.displayDialogAsyncApiSupportedOptions);
        DialogApiManager.defineApi(OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync, DialogApiManager.displayModalDialogAsyncApiSupportedOptions);
    };
    DialogApiManager.displayDialogAsyncApiSupportedOptions = [
        {
            name: Microsoft.Office.WebExtension.Parameters.Width,
            value: {
                "types": ["number"],
                "defaultValue": 99
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.Height,
            value: {
                "types": ["number"],
                "defaultValue": 99
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.HideTitle,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.PromptBeforeOpen,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.EnforceAppDomain,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        },
        {
            name: Microsoft.Office.WebExtension.Parameters.UrlNoHostInfo,
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
    ];
    DialogApiManager.displayModalDialogAsyncApiSupportedOptions = DialogApiManager.displayDialogAsyncApiSupportedOptions.concat([
        {
            name: "abortWhenParentIsMinimized",
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
        {
            name: "abortWhenDocIsInactive",
            value: {
                "types": ["boolean"],
                "defaultValue": false
            }
        },
    ]);
    return DialogApiManager;
}());
DialogApiManager.initOnce();
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.MessageParent,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.MessageToParent,
            "types": ["string", "number", "boolean"]
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
            value: {
                "types": ["string"],
                "defaultValue": ""
            }
        }
    ]
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.AddMessageHandler,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        },
        {
            "name": Microsoft.Office.WebExtension.Parameters.Handler,
            "types": ["function"]
        }
    ],
    supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.SendMessage,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.MessageContent,
            "types": ["string"]
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
            value: {
                "types": ["string"],
                "defaultValue": ""
            }
        }
    ],
    privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.openDialog = function OSF_DDA_SafeArray_Delegate$OpenDialog(args) {
    try {
        if (args.onCalling) {
            args.onCalling();
        }
        var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
        OSF.ClientHostController.openDialog(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
            if (args.onEvent) {
                args.onEvent(payload);
            }
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.onEventDone(args.dispId);
            }
        }, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate.closeDialog = function OSF_DDA_SafeArray_Delegate$CloseDialog(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
    try {
        OSF.ClientHostController.closeDialog(args.dispId, args.targetId, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate.messageParent = function OSF_DDA_SafeArray_Delegate$MessageParent(args) {
    try {
        if (args.onCalling) {
            args.onCalling();
        }
        var startTime = (new Date()).getTime();
        var result = OSF.ClientHostController.messageParent(args.hostCallArgs);
        if (args.onReceiving) {
            args.onReceiving();
        }
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
        }
        return result;
    }
    catch (ex) {
        return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
    }
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.MessageType, value: 0 },
        { name: OSF.DDA.PropertyDescriptors.MessageContent, value: 1 },
        { name: OSF.DDA.PropertyDescriptors.MessageOrigin, value: 2 }
    ],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.sendMessage = function OSF_DDA_SafeArray_Delegate$SendMessage(args) {
    try {
        if (args.onCalling) {
            args.onCalling();
        }
        var startTime = (new Date()).getTime();
        var result = OSF.ClientHostController.sendMessage(args.hostCallArgs);
        if (args.onReceiving) {
            args.onReceiving();
        }
        return result;
    }
    catch (ex) {
        return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
    }
};
OSF.DDA.AsyncMethodNames.addNames({
    GetSelectedDataAsync: "getSelectedDataAsync",
    SetSelectedDataAsync: "setSelectedDataAsync"
});
(function () {
    function processData(dataDescriptor, caller, callArgs) {
        var data = dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
        if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] != undefined || data[OSF.DDA.TableDataProperties.TableHeaders] != undefined)) {
            data = OSF.DDA.OMFactory.manufactureTableData(data);
        }
        data = OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data == undefined ? null : data;
    }
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.CoercionType,
                "enum": Microsoft.Office.WebExtension.CoercionType
            }
        ],
        supportedOptions: [
            {
                name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
                value: {
                    "enum": Microsoft.Office.WebExtension.ValueFormat,
                    "defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.FilterType,
                value: {
                    "enum": Microsoft.Office.WebExtension.FilterType,
                    "defaultValue": Microsoft.Office.WebExtension.FilterType.All
                }
            }
        ],
        privateStateCallbacks: [],
        onSucceeded: processData
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.Data,
                "types": ["string", "object", "number", "boolean"]
            }
        ],
        supportedOptions: [
            {
                name: Microsoft.Office.WebExtension.Parameters.CoercionType,
                value: {
                    "enum": Microsoft.Office.WebExtension.CoercionType,
                    "calculate": function (requiredArgs) {
                        return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
                    }
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ImageLeft,
                value: {
                    "types": ["number", "boolean"],
                    "defaultValue": false
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ImageTop,
                value: {
                    "types": ["number", "boolean"],
                    "defaultValue": false
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
                value: {
                    "types": ["number", "boolean"],
                    "defaultValue": false
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
                value: {
                    "types": ["number", "boolean"],
                    "defaultValue": false
                }
            }
        ],
        privateStateCallbacks: []
    });
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 1 },
        { name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 2 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
        { name: Microsoft.Office.WebExtension.Parameters.ImageLeft, value: 2 },
        { name: Microsoft.Office.WebExtension.Parameters.ImageTop, value: 3 },
        { name: Microsoft.Office.WebExtension.Parameters.ImageWidth, value: 4 },
        { name: Microsoft.Office.WebExtension.Parameters.ImageHeight, value: 5 },
    ]
});
OSF.DDA.SettingsManager = {
    SerializedSettings: "serializedSettings",
    RefreshingSettings: "refreshingSettings",
    DateJSONPrefix: "Date(",
    DataJSONSuffix: ")",
    serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection) {
        return OSF.OUtil.serializeSettings(settingsCollection);
    },
    deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings) {
        return OSF.OUtil.deserializeSettings(serializedSettings);
    }
};
OSF.DDA.Settings = function OSF_DDA_Settings(settings) {
    settings = settings || {};
    var cacheSessionSettings = function (settings) {
        var osfSessionStorage = OSF.OUtil.getSessionStorage();
        if (osfSessionStorage) {
            var serializedSettings = OSF.DDA.SettingsManager.serializeSettings(settings);
            var storageSettings = JSON ? JSON.stringify(serializedSettings) : Sys.Serialization.JavaScriptSerializer.serialize(serializedSettings);
            osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
        }
    };
    OSF.OUtil.defineEnumerableProperties(this, {
        "get": {
            value: function OSF_DDA_Settings$get(name) {
                var e = Function._validateParams(arguments, [
                    { name: "name", type: String, mayBeNull: false }
                ]);
                if (e)
                    throw e;
                var setting = settings[name];
                return typeof (setting) === 'undefined' ? null : setting;
            }
        },
        "set": {
            value: function OSF_DDA_Settings$set(name, value) {
                var e = Function._validateParams(arguments, [
                    { name: "name", type: String, mayBeNull: false },
                    { name: "value", mayBeNull: true }
                ]);
                if (e)
                    throw e;
                settings[name] = value;
                cacheSessionSettings(settings);
            }
        },
        "remove": {
            value: function OSF_DDA_Settings$remove(name) {
                var e = Function._validateParams(arguments, [
                    { name: "name", type: String, mayBeNull: false }
                ]);
                if (e)
                    throw e;
                delete settings[name];
                cacheSessionSettings(settings);
            }
        }
    });
    OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], settings);
};
OSF.DDA.RefreshableSettings = function OSF_DDA_RefreshableSettings(settings) {
    OSF.DDA.RefreshableSettings.uber.constructor.call(this, settings);
    OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], settings);
    OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]));
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    SettingsChanged: "settingsChanged"
});
OSF.DDA.SettingsChangedEventArgs = function OSF_DDA_SettingsChangedEventArgs(settingsInstance) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.EventType.SettingsChanged
        },
        "settings": {
            value: settingsInstance
        }
    });
};
OSF.DDA.AsyncMethodNames.addNames({
    RefreshAsync: "refreshAsync",
    SaveAsync: "saveAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RefreshAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: [
        {
            name: OSF.DDA.SettingsManager.RefreshingSettings,
            value: function getRefreshingSettings(settingsInstance, settingsCollection) {
                return settingsCollection;
            }
        }
    ],
    onSucceeded: function deserializeSettings(serializedSettingsDescriptor, refreshingSettings, refreshingSettingsArgs) {
        var serializedSettings = serializedSettingsDescriptor[OSF.DDA.SettingsManager.SerializedSettings];
        var newSettings = OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
        var oldSettings = refreshingSettingsArgs[OSF.DDA.SettingsManager.RefreshingSettings];
        for (var setting in oldSettings) {
            refreshingSettings.remove(setting);
        }
        for (var setting in newSettings) {
            refreshingSettings.set(setting, newSettings[setting]);
        }
        return refreshingSettings;
    }
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.SaveAsync,
    requiredArguments: [],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
            value: {
                "types": ["boolean"],
                "defaultValue": true
            }
        }
    ],
    privateStateCallbacks: [
        {
            name: OSF.DDA.SettingsManager.SerializedSettings,
            value: function serializeSettings(settingsInstance, settingsCollection) {
                return OSF.DDA.SettingsManager.serializeSettings(settingsCollection);
            }
        }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidLoadSettingsMethod,
    fromHost: [
        { name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidSaveSettingsMethod,
    toHost: [
        { name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
        { name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidSettingsChangedEvent });
Microsoft.Office.WebExtension.BindingType = {
    Table: "table",
    Text: "text",
    Matrix: "matrix"
};
OSF.DDA.BindingProperties = {
    Id: "BindingId",
    Type: Microsoft.Office.WebExtension.Parameters.BindingType
};
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, { BindingList: "BindingList" });
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    Subset: "subset",
    BindingProperties: "BindingProperties"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList, OSF.DDA.PropertyDescriptors.BindingProperties);
OSF.DDA.BindingPromise = function OSF_DDA_BindingPromise(bindingId, errorCallback) {
    this._id = bindingId;
    OSF.OUtil.defineEnumerableProperty(this, "onFail", {
        get: function () {
            return errorCallback;
        },
        set: function (onError) {
            var t = typeof onError;
            if (t != "undefined" && t != "function") {
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, t);
            }
            errorCallback = onError;
        }
    });
};
OSF.DDA.BindingPromise.prototype = {
    _fetch: function OSF_DDA_BindingPromise$_fetch(onComplete) {
        if (this.binding) {
            if (onComplete)
                onComplete(this.binding);
        }
        else {
            if (!this._binding) {
                var me = this;
                Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(this._id, function (asyncResult) {
                    if (asyncResult.status == Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
                        OSF.OUtil.defineEnumerableProperty(me, "binding", {
                            value: asyncResult.value
                        });
                        if (onComplete)
                            onComplete(me.binding);
                    }
                    else {
                        if (me.onFail)
                            me.onFail(asyncResult);
                    }
                });
            }
        }
        return this;
    },
    getDataAsync: function OSF_DDA_BindingPromise$getDataAsync() {
        var args = arguments;
        this._fetch(function onComplete(binding) { binding.getDataAsync.apply(binding, args); });
        return this;
    },
    setDataAsync: function OSF_DDA_BindingPromise$setDataAsync() {
        var args = arguments;
        this._fetch(function onComplete(binding) { binding.setDataAsync.apply(binding, args); });
        return this;
    },
    addHandlerAsync: function OSF_DDA_BindingPromise$addHandlerAsync() {
        var args = arguments;
        this._fetch(function onComplete(binding) { binding.addHandlerAsync.apply(binding, args); });
        return this;
    },
    removeHandlerAsync: function OSF_DDA_BindingPromise$removeHandlerAsync() {
        var args = arguments;
        this._fetch(function onComplete(binding) { binding.removeHandlerAsync.apply(binding, args); });
        return this;
    }
};
OSF.DDA.BindingFacade = function OSF_DDA_BindingFacade(docInstance) {
    this._eventDispatches = [];
    OSF.OUtil.defineEnumerableProperty(this, "document", {
        value: docInstance
    });
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.AddFromSelectionAsync,
        am.AddFromNamedItemAsync,
        am.GetAllAsync,
        am.GetByIdAsync,
        am.ReleaseByIdAsync
    ]);
};
OSF.DDA.UnknownBinding = function OSF_DDA_UknonwnBinding(id, docInstance) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "document": { value: docInstance },
        "id": { value: id }
    });
};
OSF.DDA.Binding = function OSF_DDA_Binding(id, docInstance) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "document": {
            value: docInstance
        },
        "id": {
            value: id
        }
    });
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.GetDataAsync,
        am.SetDataAsync
    ]);
    var et = Microsoft.Office.WebExtension.EventType;
    var bindingEventDispatches = docInstance.bindings._eventDispatches;
    if (!bindingEventDispatches[id]) {
        bindingEventDispatches[id] = new OSF.EventDispatch([
            et.BindingSelectionChanged,
            et.BindingDataChanged
        ]);
    }
    var eventDispatch = bindingEventDispatches[id];
    OSF.DDA.DispIdHost.addEventSupport(this, eventDispatch);
};
OSF.DDA.generateBindingId = function OSF_DDA$GenerateBindingId() {
    return "UnnamedBinding_" + OSF.OUtil.getUniqueId() + "_" + new Date().getTime();
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureBinding = function OSF_DDA_OMFactory$manufactureBinding(bindingProperties, containingDocument) {
    var id = bindingProperties[OSF.DDA.BindingProperties.Id];
    var rows = bindingProperties[OSF.DDA.BindingProperties.RowCount];
    var cols = bindingProperties[OSF.DDA.BindingProperties.ColumnCount];
    var hasHeaders = bindingProperties[OSF.DDA.BindingProperties.HasHeaders];
    var binding;
    switch (bindingProperties[OSF.DDA.BindingProperties.Type]) {
        case Microsoft.Office.WebExtension.BindingType.Text:
            binding = new OSF.DDA.TextBinding(id, containingDocument);
            break;
        case Microsoft.Office.WebExtension.BindingType.Matrix:
            binding = new OSF.DDA.MatrixBinding(id, containingDocument, rows, cols);
            break;
        case Microsoft.Office.WebExtension.BindingType.Table:
            var isExcelApp = function () {
                return (OSF.DDA.ExcelDocument)
                    && (Microsoft.Office.WebExtension.context.document)
                    && (Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument);
            };
            var tableBindingObject;
            if (isExcelApp() && OSF.DDA.ExcelTableBinding) {
                tableBindingObject = OSF.DDA.ExcelTableBinding;
            }
            else {
                tableBindingObject = OSF.DDA.TableBinding;
            }
            binding = new tableBindingObject(id, containingDocument, rows, cols, hasHeaders);
            break;
        default:
            binding = new OSF.DDA.UnknownBinding(id, containingDocument);
    }
    return binding;
};
OSF.DDA.AsyncMethodNames.addNames({
    AddFromSelectionAsync: "addFromSelectionAsync",
    AddFromNamedItemAsync: "addFromNamedItemAsync",
    GetAllAsync: "getAllAsync",
    GetByIdAsync: "getByIdAsync",
    ReleaseByIdAsync: "releaseByIdAsync",
    GetDataAsync: "getDataAsync",
    SetDataAsync: "setDataAsync"
});
(function () {
    function processBinding(bindingDescriptor) {
        return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document);
    }
    function getObjectId(obj) { return obj.id; }
    function processData(dataDescriptor, caller, callArgs) {
        var data = dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
        if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] != undefined || data[OSF.DDA.TableDataProperties.TableHeaders] != undefined)) {
            data = OSF.DDA.OMFactory.manufactureTableData(data);
        }
        data = OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data == undefined ? null : data;
    }
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.BindingType,
                "enum": Microsoft.Office.WebExtension.BindingType
            }
        ],
        supportedOptions: [{
                name: Microsoft.Office.WebExtension.Parameters.Id,
                value: {
                    "types": ["string"],
                    "calculate": OSF.DDA.generateBindingId
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Columns,
                value: {
                    "types": ["object"],
                    "defaultValue": null
                }
            }
        ],
        privateStateCallbacks: [],
        onSucceeded: processBinding
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
        requiredArguments: [{
                "name": Microsoft.Office.WebExtension.Parameters.ItemName,
                "types": ["string"]
            },
            {
                "name": Microsoft.Office.WebExtension.Parameters.BindingType,
                "enum": Microsoft.Office.WebExtension.BindingType
            }
        ],
        supportedOptions: [{
                name: Microsoft.Office.WebExtension.Parameters.Id,
                value: {
                    "types": ["string"],
                    "calculate": OSF.DDA.generateBindingId
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Columns,
                value: {
                    "types": ["object"],
                    "defaultValue": null
                }
            }
        ],
        privateStateCallbacks: [
            {
                name: Microsoft.Office.WebExtension.Parameters.FailOnCollision,
                value: function () { return true; }
            }
        ],
        onSucceeded: processBinding
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.GetAllAsync,
        requiredArguments: [],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.BindingList], processBinding); }
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.GetByIdAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.Id,
                "types": ["string"]
            }
        ],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: processBinding
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.Id,
                "types": ["string"]
            }
        ],
        supportedOptions: [],
        privateStateCallbacks: [],
        onSucceeded: function (response, caller, callArgs) {
            var id = callArgs[Microsoft.Office.WebExtension.Parameters.Id];
            delete caller._eventDispatches[id];
        }
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.GetDataAsync,
        requiredArguments: [],
        supportedOptions: [{
                name: Microsoft.Office.WebExtension.Parameters.CoercionType,
                value: {
                    "enum": Microsoft.Office.WebExtension.CoercionType,
                    "calculate": function (requiredArgs, binding) { return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(binding.type); }
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
                value: {
                    "enum": Microsoft.Office.WebExtension.ValueFormat,
                    "defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.FilterType,
                value: {
                    "enum": Microsoft.Office.WebExtension.FilterType,
                    "defaultValue": Microsoft.Office.WebExtension.FilterType.All
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Rows,
                value: {
                    "types": ["object", "string"],
                    "defaultValue": null
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Columns,
                value: {
                    "types": ["object"],
                    "defaultValue": null
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.StartRow,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.StartColumn,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.RowCount,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.ColumnCount,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            }
        ],
        checkCallArgs: function (callArgs, caller, stateInfo) {
            if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] == 0 &&
                callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] == 0 &&
                callArgs[Microsoft.Office.WebExtension.Parameters.RowCount] == 0 &&
                callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount] == 0) {
                delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
                delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
                delete callArgs[Microsoft.Office.WebExtension.Parameters.RowCount];
                delete callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount];
            }
            if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
                (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
                    callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] ||
                    callArgs[Microsoft.Office.WebExtension.Parameters.RowCount] ||
                    callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount])) {
                throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
            }
            return callArgs;
        },
        privateStateCallbacks: [
            {
                name: Microsoft.Office.WebExtension.Parameters.Id,
                value: getObjectId
            }
        ],
        onSucceeded: processData
    });
    OSF.DDA.AsyncMethodCalls.define({
        method: OSF.DDA.AsyncMethodNames.SetDataAsync,
        requiredArguments: [
            {
                "name": Microsoft.Office.WebExtension.Parameters.Data,
                "types": ["string", "object", "number", "boolean"]
            }
        ],
        supportedOptions: [{
                name: Microsoft.Office.WebExtension.Parameters.CoercionType,
                value: {
                    "enum": Microsoft.Office.WebExtension.CoercionType,
                    "calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Rows,
                value: {
                    "types": ["object", "string"],
                    "defaultValue": null
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.Columns,
                value: {
                    "types": ["object"],
                    "defaultValue": null
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.StartRow,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            },
            {
                name: Microsoft.Office.WebExtension.Parameters.StartColumn,
                value: {
                    "types": ["number"],
                    "defaultValue": 0
                }
            }
        ],
        checkCallArgs: function (callArgs, caller, stateInfo) {
            if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] == 0 &&
                callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] == 0) {
                delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
                delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
            }
            if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
                (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
                    callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn])) {
                throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
            }
            return callArgs;
        },
        privateStateCallbacks: [
            {
                name: Microsoft.Office.WebExtension.Parameters.Id,
                value: getObjectId
            }
        ]
    });
})();
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
    RowCount: "BindingRowCount",
    ColumnCount: "BindingColumnCount",
    HasHeaders: "HasHeaders"
});
OSF.DDA.MatrixBinding = function OSF_DDA_MatrixBinding(id, docInstance, rows, cols) {
    OSF.DDA.MatrixBinding.uber.constructor.call(this, id, docInstance);
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": {
            value: Microsoft.Office.WebExtension.BindingType.Matrix
        },
        "rowCount": {
            value: rows ? rows : 0
        },
        "columnCount": {
            value: cols ? cols : 0
        }
    });
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.BindingProperties,
    fromHost: [
        { name: OSF.DDA.BindingProperties.Id, value: 0 },
        { name: OSF.DDA.BindingProperties.Type, value: 1 },
        { name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
    ],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: Microsoft.Office.WebExtension.Parameters.BindingType,
    toHost: [
        { name: Microsoft.Office.WebExtension.BindingType.Text, value: 0 },
        { name: Microsoft.Office.WebExtension.BindingType.Matrix, value: 1 },
        { name: Microsoft.Office.WebExtension.BindingType.Table, value: 2 }
    ],
    invertible: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.ItemName, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 1 },
        { name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 2 },
        { name: Microsoft.Office.WebExtension.Parameters.FailOnCollision, value: 3 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidReleaseBindingMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetBindingMethod,
    fromHost: [
        { name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,
    fromHost: [
        { name: OSF.DDA.ListDescriptors.BindingList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetBindingDataMethod,
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ],
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
        { name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 2 },
        { name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 3 },
        { name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
        { name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
    fromHost: [
        { name: OSF.DDA.BindingProperties.RowCount, value: 0 },
        { name: OSF.DDA.BindingProperties.ColumnCount, value: 1 },
        { name: OSF.DDA.BindingProperties.HasHeaders, value: 2 }
    ],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.Subset,
    toHost: [
        { name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 0 },
        { name: OSF.DDA.SafeArray.UniqueArguments.Run, value: 1 }
    ],
    canonical: true,
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.SafeArray.UniqueArguments.Offset,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.StartRow, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.StartColumn, value: 1 }
    ],
    canonical: true,
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.SafeArray.UniqueArguments.Run,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.RowCount, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.ColumnCount, value: 1 }
    ],
    canonical: true,
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidAddRowsMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidAddColumnsMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
    ]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidClearAllRowsMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
    ]
});
OSF.DDA.AsyncMethodNames.addNames({
    ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
    requiredArguments: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: ["object"]
        }
    ],
    supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
    ],
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RichApiMessage: "richApiMessage" });
OSF.DDA.RichApiMessageEventArgs = function OSF_DDA_RichApiMessageEventArgs(eventType, eventProperties) {
    var entryArray = eventProperties[Microsoft.Office.WebExtension.Parameters.Data];
    var entries = [];
    if (entryArray) {
        for (var i = 0; i < entryArray.length; i++) {
            var elem = entryArray[i];
            if (elem.toArray) {
                elem = elem.toArray();
            }
            entries.push({
                messageCategory: elem[0],
                messageType: elem[1],
                targetId: elem[2],
                message: elem[3],
                id: elem[4],
                isRemoteOverride: elem[5]
            });
        }
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": { value: Microsoft.Office.WebExtension.EventType.RichApiMessage },
        "entries": { value: entries }
    });
};
(function (OfficeExt) {
    var RichApiMessageManager = (function () {
        function RichApiMessageManager() {
            this._eventDispatch = null;
            this._registerHandlers = [];
            this._eventDispatch = new OSF.EventDispatch([
                Microsoft.Office.WebExtension.EventType.RichApiMessage,
            ]);
            OSF.DDA.DispIdHost.addEventSupport(this, this._eventDispatch);
        }
        RichApiMessageManager.prototype.register = function (handler) {
            var _this = this;
            if (!this._registerWithHostPromise) {
                this._registerWithHostPromise = new Office.Promise(function (resolve, reject) {
                    _this.addHandlerAsync(Microsoft.Office.WebExtension.EventType.RichApiMessage, function (args) {
                        _this._registerHandlers.forEach(function (value) {
                            if (value) {
                                value(args);
                            }
                        });
                    }, function (asyncResult) {
                        if (asyncResult.status == 'failed') {
                            reject(asyncResult.error);
                        }
                        else {
                            resolve();
                        }
                    });
                });
            }
            return this._registerWithHostPromise.then(function () {
                _this._registerHandlers.push(handler);
            });
        };
        return RichApiMessageManager;
    }());
    OfficeExt.RichApiMessageManager = RichApiMessageManager;
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidRichApiMessageEvent,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
    ],
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
    ]
});
(function (OfficeExt) {
    var AppCommand;
    (function (AppCommand) {
        var AppCommandManager = (function () {
            function AppCommandManager() {
                var _this = this;
                this._pseudoDocument = null;
                this._eventDispatch = null;
                this._useAssociatedActionsOnly = null;
                this._processAppCommandInvocation = function (args) {
                    var verifyResult = _this._verifyManifestCallback(args.callbackName);
                    if (verifyResult.errorCode != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                        _this._invokeAppCommandCompletedMethod(args.appCommandId, verifyResult.errorCode, "");
                        return;
                    }
                    var eventObj = _this._constructEventObjectForCallback(args);
                    if (eventObj) {
                        window.setTimeout(function () { verifyResult.callback(eventObj); }, 0);
                    }
                    else {
                        _this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, "");
                    }
                };
            }
            AppCommandManager.initializeOsfDda = function () {
                OSF.DDA.AsyncMethodNames.addNames({
                    AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"
                });
                OSF.DDA.AsyncMethodCalls.define({
                    method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                    requiredArguments: [{
                            "name": Microsoft.Office.WebExtension.Parameters.Id,
                            "types": ["string"]
                        },
                        {
                            "name": Microsoft.Office.WebExtension.Parameters.Status,
                            "types": ["number"]
                        },
                        {
                            "name": Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                            "types": ["string"]
                        }
                    ]
                });
                OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
                    AppCommandInvokedEvent: "AppCommandInvokedEvent"
                });
                OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
                    AppCommandInvoked: "appCommandInvoked"
                });
                OSF.OUtil.setNamespace("AppCommand", OSF.DDA);
                OSF.DDA.AppCommand.AppCommandInvokedEventArgs = OfficeExt.AppCommand.AppCommandInvokedEventArgs;
            };
            AppCommandManager.prototype.initializeAndChangeOnce = function (callback) {
                AppCommand.registerDdaFacade();
                this._pseudoDocument = {};
                OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument, [
                    OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                ]);
                this._eventDispatch = new OSF.EventDispatch([
                    Microsoft.Office.WebExtension.EventType.AppCommandInvoked,
                ]);
                var onRegisterCompleted = function (result) {
                    if (callback) {
                        if (result.status == "succeeded") {
                            callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                        }
                        else {
                            callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                        }
                    }
                };
                OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument, this._eventDispatch);
                this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked, this._processAppCommandInvocation, onRegisterCompleted);
            };
            AppCommandManager.prototype._verifyManifestCallback = function (callbackName) {
                var defaultResult = { callback: null, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback };
                callbackName = callbackName.trim();
                try {
                    var callbackFunc = this._getCallbackFunc(callbackName);
                    if (typeof callbackFunc != "function") {
                        return defaultResult;
                    }
                }
                catch (e) {
                    return defaultResult;
                }
                return { callback: callbackFunc, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess };
            };
            AppCommandManager.prototype._getUseAssociatedActionsOnly = function () {
                if (this._useAssociatedActionsOnly == null) {
                    this._useAssociatedActionsOnly = false;
                    try {
                        if (window["useAssociatedActionsOnly"] === true) {
                            this._useAssociatedActionsOnly = true;
                        }
                        else {
                            this._useAssociatedActionsOnly = OSF._OfficeAppFactory.getLoadScriptHelper().getUseAssociatedActionsOnlyDefined();
                        }
                    }
                    catch (e) { }
                }
                return this._useAssociatedActionsOnly;
            };
            AppCommandManager.prototype._getCallbackFuncFromWindow = function (callbackName) {
                var callList = callbackName.split(".");
                var parentObject = window;
                for (var i = 0; i < callList.length - 1; i++) {
                    if (parentObject[callList[i]] && (typeof parentObject[callList[i]] == "object" || typeof parentObject[callList[i]] == "function")) {
                        parentObject = parentObject[callList[i]];
                    }
                    else {
                        return null;
                    }
                }
                var callbackFunc = parentObject[callList[callList.length - 1]];
                return callbackFunc;
            };
            AppCommandManager.prototype._getCallbackFuncFromActionAssociateTable = function (callbackName) {
                var nameUpperCase = callbackName.toUpperCase();
                return Office.actions._association.mappings[nameUpperCase];
            };
            AppCommandManager.prototype._getCallbackFunc = function (callbackName) {
                var _this = this;
                var callbackFunc = null;
                var useAssociateTable = false;
                if (!this._getUseAssociatedActionsOnly()) {
                    callbackFunc = this._getCallbackFuncFromWindow(callbackName);
                }
                if (!callbackFunc) {
                    callbackFunc = this._getCallbackFuncFromActionAssociateTable(callbackName);
                    if (callbackFunc) {
                        useAssociateTable = true;
                    }
                }
                if (!AppCommandManager.isTelemetrySubmitted) {
                    AppCommandManager.isTelemetrySubmitted = true;
                    try {
                        if (OTel && oteljs && Microsoft.Office.WebExtension.sendTelemetryEvent) {
                            OTel.OTelLogger.onTelemetryLoaded(function () {
                                var dataFields = [
                                    oteljs.makeBooleanDataField("UseAction", _this._useAssociatedActionsOnly === true),
                                    oteljs.makeBooleanDataField("UseAssociateTable", useAssociateTable)
                                ];
                                Microsoft.Office.WebExtension.sendTelemetryEvent({
                                    eventName: "Office.Extensibility.OfficeJs.AppCommandDefinition",
                                    dataFields: dataFields,
                                    eventFlags: {
                                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                                    }
                                });
                            });
                        }
                    }
                    catch (e) { }
                }
                return callbackFunc;
            };
            AppCommandManager.prototype._invokeAppCommandCompletedMethod = function (appCommandId, resultCode, data) {
                this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId, resultCode, data);
            };
            AppCommandManager.prototype._constructEventObjectForCallback = function (args) {
                var _this = this;
                var eventObj = new AppCommandCallbackEventArgs();
                try {
                    var jsonData = JSON.parse(args.eventObjStr);
                    this._translateEventObjectInternal(jsonData, eventObj);
                    Object.defineProperty(eventObj, 'completed', {
                        value: function (completedContext) {
                            eventObj.completedContext = completedContext;
                            var jsonString = JSON.stringify(eventObj);
                            _this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, jsonString);
                        },
                        enumerable: true
                    });
                }
                catch (e) {
                    eventObj = null;
                }
                return eventObj;
            };
            AppCommandManager.prototype._translateEventObjectInternal = function (input, output) {
                for (var key in input) {
                    if (!input.hasOwnProperty(key))
                        continue;
                    var inputChild = input[key];
                    if (typeof inputChild == "object" && inputChild != null) {
                        OSF.OUtil.defineEnumerableProperty(output, key, {
                            value: {}
                        });
                        this._translateEventObjectInternal(inputChild, output[key]);
                    }
                    else {
                        Object.defineProperty(output, key, {
                            value: inputChild,
                            enumerable: true,
                            writable: true
                        });
                    }
                }
            };
            AppCommandManager.prototype._constructObjectByTemplate = function (template, input) {
                var output = {};
                if (!template || !input)
                    return output;
                for (var key in template) {
                    if (template.hasOwnProperty(key)) {
                        output[key] = null;
                        if (input[key] != null) {
                            var templateChild = template[key];
                            var inputChild = input[key];
                            var inputChildType = typeof inputChild;
                            if (typeof templateChild == "object" && templateChild != null) {
                                output[key] = this._constructObjectByTemplate(templateChild, inputChild);
                            }
                            else if (inputChildType == "number" || inputChildType == "string" || inputChildType == "boolean") {
                                output[key] = inputChild;
                            }
                        }
                    }
                }
                return output;
            };
            AppCommandManager.instance = function () {
                if (AppCommandManager._instance == null) {
                    AppCommandManager._instance = new AppCommandManager();
                }
                return AppCommandManager._instance;
            };
            AppCommandManager.isTelemetrySubmitted = false;
            AppCommandManager._instance = null;
            return AppCommandManager;
        }());
        AppCommand.AppCommandManager = AppCommandManager;
        var AppCommandInvokedEventArgs = (function () {
            function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr) {
                this.type = Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
                this.appCommandId = appCommandId;
                this.callbackName = callbackName;
                this.eventObjStr = eventObjStr;
            }
            AppCommandInvokedEventArgs.create = function (eventProperties) {
                return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId], eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName], eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr]);
            };
            return AppCommandInvokedEventArgs;
        }());
        AppCommand.AppCommandInvokedEventArgs = AppCommandInvokedEventArgs;
        var AppCommandCallbackEventArgs = (function () {
            function AppCommandCallbackEventArgs() {
            }
            return AppCommandCallbackEventArgs;
        }());
        AppCommand.AppCommandCallbackEventArgs = AppCommandCallbackEventArgs;
        AppCommand.AppCommandInvokedEventEnums = {
            AppCommandId: "appCommandId",
            CallbackName: "callbackName",
            EventObjStr: "eventObjStr"
        };
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}));
})(OfficeExt || (OfficeExt = {}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
(function (OfficeExt) {
    var AppCommand;
    (function (AppCommand) {
        function registerDdaFacade() {
            if (OSF.DDA.SafeArray) {
                var parameterMap = OSF.DDA.SafeArray.Delegate.ParameterMap;
                parameterMap.define({
                    type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
                    toHost: [
                        { name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
                        { name: Microsoft.Office.WebExtension.Parameters.Status, value: 1 },
                        { name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData, value: 2 }
                    ]
                });
                parameterMap.define({
                    type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                    fromHost: [
                        { name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent, value: parameterMap.self }
                    ],
                    isComplexType: true
                });
                parameterMap.define({
                    type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                    fromHost: [
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId, value: 0 },
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName, value: 1 },
                        { name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr, value: 2 },
                    ],
                    isComplexType: true
                });
            }
        }
        AppCommand.registerDdaFacade = registerDdaFacade;
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
    var appCommandHandler = OfficeExt.AppCommand.AppCommandManager.instance();
    appCommandHandler.initializeAndChangeOnce();
};
OSF.DDA.VisioDocument = function OSF_DDA_VisioDocument(officeAppContext, settings) {
    OSF.DDA.VisioDocument.uber.constructor.call(this, officeAppContext, new OSF.DDA.BindingFacade(this), settings);
    OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.VisioDocument, OSF.DDA.JsomDocument);
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    appContext.doc = new OSF.DDA.VisioDocument(appContext, this._initializeSettings(true));
    OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
    OSF.DDA.RichApi.richApiMessageManager = new OfficeExt.RichApiMessageManager();
    appReady();
};
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b)
                if (b.hasOwnProperty(p))
                    d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OfficeExtension;
(function (OfficeExtension) {
    var _Internal;
    (function (_Internal) {
        _Internal.OfficeRequire = function () {
            return null;
        }();
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = window.setImmediate;
                        }
                        return function () {
                            nextTick(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useVertxTimer() {
                        return function () {
                            lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useMutationObserver() {
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$selfFullfillment() {
                        return new TypeError("You cannot resolve a promise with itself");
                    }
                    function lib$es6$promise$$internal$$cannotReturnOwn() {
                        return new TypeError('A promises callback cannot return that same promise.');
                    }
                    function lib$es6$promise$$internal$$getThen(promise) {
                        try {
                            return promise.then;
                        }
                        catch (error) {
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                            return lib$es6$promise$$internal$$GET_THEN_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                        try {
                            then.call(value, fulfillmentHandler, rejectionHandler);
                        }
                        catch (e) {
                            return e;
                        }
                    }
                    function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                        lib$es6$promise$asap$$asap(function (promise) {
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, thenable._result);
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function (reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                    }
                    function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                            else if (lib$es6$promise$utils$$isFunction(then)) {
                                lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                        }
                    }
                    function lib$es6$promise$$internal$$resolve(promise, value) {
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                        }
                        else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                            lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$publishRejection(promise) {
                        if (promise._onerror) {
                            promise._onerror(promise._result);
                        }
                        lib$es6$promise$$internal$$publish(promise);
                    }
                    function lib$es6$promise$$internal$$fulfill(promise, value) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                        try {
                            resolver(function resolvePromise(value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function rejectPromise(reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$reject(promise, e);
                        }
                    }
                    function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new _Internal.Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
                                lib$es6$promise$asap$$asap(function () {
                                    lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                                });
                            }
                            else {
                                lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                            }
                            return child;
                        },
                        'catch': function (onRejection) {
                            return this.then(null, onRejection);
                        }
                    };
                    return lib$es6$promise$promise$$default;
                }).call(this);
            }
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
                if (isEdgeLessThan14()) {
                    return _Internal.PromiseImpl.Init();
                }
                else {
                    return window.Promise;
                }
            }
            else {
                return _Internal.PromiseImpl.Init();
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    OfficeExtension.Promise = OfficePromise;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension_1) {
    var SessionBase = (function () {
        function SessionBase() {
        }
        SessionBase.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        SessionBase.prototype._createRequestExecutorOrNull = function () {
            return null;
        };
        SessionBase.prototype.getEventRegistration = function (controlId) {
            return null;
        };
        return SessionBase;
    }());
    OfficeExtension_1.SessionBase = SessionBase;
    var HttpUtility = (function () {
        function HttpUtility() {
        }
        HttpUtility.setCustomSendRequestFunc = function (func) {
            HttpUtility.s_customSendRequestFunc = func;
        };
        HttpUtility.xhrSendRequestFunc = function (request) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var xhr = new XMLHttpRequest();
                xhr.open(request.method, request.url);
                xhr.onload = function () {
                    var resp = {
                        statusCode: xhr.status,
                        headers: CoreUtility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
                        body: xhr.responseText
                    };
                    resolve(resp);
                };
                xhr.onerror = function () {
                    reject(new _Internal.RuntimeError({
                        code: CoreErrorCodes.connectionFailure,
                        httpStatusCode: xhr.status,
                        message: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, xhr.statusText)
                    }));
                };
                if (request.headers) {
                    for (var key in request.headers) {
                        xhr.setRequestHeader(key, request.headers[key]);
                    }
                }
                xhr.send(CoreUtility._getRequestBodyText(request));
            });
        };
        HttpUtility.fetchSendRequestFunc = function (request) {
            var requestBodyText = CoreUtility._getRequestBodyText(request);
            if (requestBodyText === '') {
                requestBodyText = undefined;
            }
            return fetch(request.url, {
                method: request.method,
                headers: request.headers,
                body: requestBodyText
            })
                .then(function (resp) {
                return resp.text()
                    .then(function (body) {
                    var statusCode = resp.status;
                    var headers = {};
                    resp.headers.forEach(function (value, name) {
                        headers[name] = value;
                    });
                    var ret = { statusCode: statusCode, headers: headers, body: body };
                    return ret;
                });
            });
        };
        HttpUtility.sendRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func = HttpUtility.s_customSendRequestFunc;
            if (!func) {
                if (typeof (fetch) !== 'undefined') {
                    func = HttpUtility.fetchSendRequestFunc;
                }
                else {
                    func = HttpUtility.xhrSendRequestFunc;
                }
            }
            return func(request);
        };
        HttpUtility.setCustomSendLocalDocumentRequestFunc = function (func) {
            HttpUtility.s_customSendLocalDocumentRequestFunc = func;
        };
        HttpUtility.sendLocalDocumentRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func;
            func = HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
            return func(request);
        };
        HttpUtility.officeJsSendLocalDocumentRequestFunc = function (request) {
            request = CoreUtility._validateLocalDocumentRequest(request);
            var requestSafeArray = CoreUtility._buildRequestMessageSafeArray(request);
            return CoreUtility.createPromise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
                    var response;
                    if (asyncResult.status == 'succeeded') {
                        response = {
                            statusCode: RichApiMessageUtility.getResponseStatusCode(asyncResult),
                            headers: RichApiMessageUtility.getResponseHeaders(asyncResult),
                            body: RichApiMessageUtility.getResponseBody(asyncResult)
                        };
                    }
                    else {
                        response = RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
                    }
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(response));
                    resolve(response);
                });
            });
        };
        HttpUtility.validateAndNormalizeRequest = function (request) {
            if (CoreUtility.isNullOrUndefined(request)) {
                throw _Internal.RuntimeError._createInvalidArgError({
                    argumentName: 'request'
                });
            }
            if (CoreUtility.isNullOrEmptyString(request.method)) {
                request.method = 'GET';
            }
            request.method = request.method.toUpperCase();
            var alreadyHasTestName = false;
            if (typeof (request.headers) === 'object' && request.headers[CoreConstants.testRequestNameHeader]) {
                alreadyHasTestName = true;
            }
            if (!alreadyHasTestName) {
                var currentTestName = TestUtility._getCurrentTestNameWithSequenceId();
                if (currentTestName) {
                    if (!request.headers) {
                        request.headers = {};
                    }
                    request.headers[CoreConstants.testRequestNameHeader] = currentTestName;
                }
            }
        };
        HttpUtility.logRequest = function (request) {
            if (CoreUtility._logEnabled) {
                CoreUtility.log('---HTTP Request---');
                CoreUtility.log(request.method + ' ' + request.url);
                if (request.headers) {
                    for (var key in request.headers) {
                        CoreUtility.log(key + ': ' + request.headers[key]);
                    }
                }
                if (HttpUtility._logBodyEnabled) {
                    CoreUtility.log(CoreUtility._getRequestBodyText(request));
                }
            }
        };
        HttpUtility.logResponse = function (response) {
            if (CoreUtility._logEnabled) {
                CoreUtility.log('---HTTP Response---');
                CoreUtility.log('' + response.statusCode);
                if (response.headers) {
                    for (var key in response.headers) {
                        CoreUtility.log(key + ': ' + response.headers[key]);
                    }
                }
                if (HttpUtility._logBodyEnabled) {
                    CoreUtility.log(response.body);
                }
            }
        };
        HttpUtility._logBodyEnabled = false;
        return HttpUtility;
    }());
    OfficeExtension_1.HttpUtility = HttpUtility;
    var HostBridge = (function () {
        function HostBridge(m_bridge) {
            var _this = this;
            this.m_bridge = m_bridge;
            this.m_promiseResolver = {};
            this.m_handlers = [];
            this.m_bridge.onMessageFromHost = function (messageText) {
                var message = JSON.parse(messageText);
                if (message.type == 3) {
                    var genericMessageBody = message.message;
                    if (genericMessageBody && genericMessageBody.entries) {
                        for (var i = 0; i < genericMessageBody.entries.length; i++) {
                            var entryObjectOrArray = genericMessageBody.entries[i];
                            if (Array.isArray(entryObjectOrArray)) {
                                var entry = {
                                    messageCategory: entryObjectOrArray[0],
                                    messageType: entryObjectOrArray[1],
                                    targetId: entryObjectOrArray[2],
                                    message: entryObjectOrArray[3],
                                    id: entryObjectOrArray[4]
                                };
                                genericMessageBody.entries[i] = entry;
                            }
                        }
                    }
                }
                _this.dispatchMessage(message);
            };
        }
        HostBridge.init = function (bridge) {
            if (typeof bridge !== 'object' || !bridge) {
                return;
            }
            var instance = new HostBridge(bridge);
            HostBridge.s_instance = instance;
            HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
                request = CoreUtility._validateLocalDocumentRequest(request);
                var requestFlags = 0;
                if (!CoreUtility.isReadonlyRestRequest(request.method)) {
                    requestFlags = 1;
                }
                var index = request.url.indexOf('?');
                if (index >= 0) {
                    var query = request.url.substr(index + 1);
                    var flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(query);
                    if (flagsAndCustomData.flags >= 0) {
                        requestFlags = flagsAndCustomData.flags;
                    }
                }
                if (typeof (request.body) === "string") {
                    request.body = JSON.parse(request.body);
                }
                var bridgeMessage = {
                    id: HostBridge.nextId(),
                    type: 1,
                    flags: requestFlags,
                    message: request
                };
                return instance.sendMessageToHostAndExpectResponse(bridgeMessage).then(function (bridgeResponse) {
                    var responseInfo = bridgeResponse.message;
                    return responseInfo;
                });
            });
            for (var i = 0; i < HostBridge.s_onInitedHandlers.length; i++) {
                HostBridge.s_onInitedHandlers[i](instance);
            }
        };
        Object.defineProperty(HostBridge, "instance", {
            get: function () {
                return HostBridge.s_instance;
            },
            enumerable: true,
            configurable: true
        });
        HostBridge.prototype.sendMessageToHost = function (message) {
            this.m_bridge.sendMessageToHost(JSON.stringify(message));
        };
        HostBridge.prototype.sendMessageToHostAndExpectResponse = function (message) {
            var _this = this;
            var ret = CoreUtility.createPromise(function (resolve, reject) {
                _this.m_promiseResolver[message.id] = resolve;
            });
            this.m_bridge.sendMessageToHost(JSON.stringify(message));
            return ret;
        };
        HostBridge.prototype.addHostMessageHandler = function (handler) {
            this.m_handlers.push(handler);
        };
        HostBridge.prototype.removeHostMessageHandler = function (handler) {
            var index = this.m_handlers.indexOf(handler);
            if (index >= 0) {
                this.m_handlers.splice(index, 1);
            }
        };
        HostBridge.onInited = function (handler) {
            HostBridge.s_onInitedHandlers.push(handler);
            if (HostBridge.s_instance) {
                handler(HostBridge.s_instance);
            }
        };
        HostBridge.prototype.dispatchMessage = function (message) {
            if (typeof message.id === 'number') {
                var resolve = this.m_promiseResolver[message.id];
                if (resolve) {
                    resolve(message);
                    delete this.m_promiseResolver[message.id];
                    return;
                }
            }
            for (var i = 0; i < this.m_handlers.length; i++) {
                this.m_handlers[i](message);
            }
        };
        HostBridge.nextId = function () {
            return HostBridge.s_nextId++;
        };
        HostBridge.s_onInitedHandlers = [];
        HostBridge.s_nextId = 1;
        return HostBridge;
    }());
    OfficeExtension_1.HostBridge = HostBridge;
    if (typeof _richApiNativeBridge === 'object' && _richApiNativeBridge) {
        HostBridge.init(_richApiNativeBridge);
    }
    var _Internal;
    (function (_Internal) {
        var RuntimeError = (function (_super) {
            __extends(RuntimeError, _super);
            function RuntimeError(error) {
                var _this = _super.call(this, typeof error === 'string' ? error : error.message) || this;
                Object.setPrototypeOf(_this, RuntimeError.prototype);
                _this.name = 'RichApi.Error';
                if (typeof error === 'string') {
                    _this.message = error;
                }
                else {
                    _this.code = error.code;
                    _this.message = error.message;
                    _this.traceMessages = error.traceMessages || [];
                    _this.innerError = error.innerError || null;
                    _this.debugInfo = _this._createDebugInfo(error.debugInfo || {});
                    _this.httpStatusCode = error.httpStatusCode;
                    _this.data = error.data;
                }
                if (CoreUtility.isNullOrUndefined(_this.httpStatusCode) || _this.httpStatusCode === 200) {
                    var mapping = {};
                    mapping[CoreErrorCodes.accessDenied] = 401;
                    mapping[CoreErrorCodes.connectionFailure] = 500;
                    mapping[CoreErrorCodes.generalException] = 500;
                    mapping[CoreErrorCodes.invalidArgument] = 400;
                    mapping[CoreErrorCodes.invalidObjectPath] = 400;
                    mapping[CoreErrorCodes.invalidOrTimedOutSession] = 408;
                    mapping[CoreErrorCodes.invalidRequestContext] = 400;
                    mapping[CoreErrorCodes.timeout] = 408;
                    mapping[CoreErrorCodes.valueNotLoaded] = 400;
                    _this.httpStatusCode = mapping[_this.code];
                }
                if (CoreUtility.isNullOrUndefined(_this.httpStatusCode)) {
                    _this.httpStatusCode = 500;
                }
                return _this;
            }
            RuntimeError.prototype.toString = function () {
                return this.code + ': ' + this.message;
            };
            RuntimeError.prototype._createDebugInfo = function (partialDebugInfo) {
                var debugInfo = {
                    code: this.code,
                    message: this.message
                };
                debugInfo.toString = function () {
                    return JSON.stringify(this);
                };
                for (var key in partialDebugInfo) {
                    debugInfo[key] = partialDebugInfo[key];
                }
                if (this.innerError) {
                    if (this.innerError instanceof _Internal.RuntimeError) {
                        debugInfo.innerError = this.innerError.debugInfo;
                    }
                    else {
                        debugInfo.innerError = this.innerError;
                    }
                }
                return debugInfo;
            };
            RuntimeError._createInvalidArgError = function (error) {
                return new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: CoreUtility.isNullOrEmptyString(error.argumentName)
                        ? CoreUtility._getResourceString(CoreResourceStrings.invalidArgumentGeneric)
                        : CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, error.argumentName),
                    debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
                    innerError: error.innerError
                });
            };
            return RuntimeError;
        }(Error));
        _Internal.RuntimeError = RuntimeError;
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    OfficeExtension_1.Error = _Internal.RuntimeError;
    var CoreErrorCodes = (function () {
        function CoreErrorCodes() {
        }
        CoreErrorCodes.apiNotFound = 'ApiNotFound';
        CoreErrorCodes.accessDenied = 'AccessDenied';
        CoreErrorCodes.generalException = 'GeneralException';
        CoreErrorCodes.activityLimitReached = 'ActivityLimitReached';
        CoreErrorCodes.invalidArgument = 'InvalidArgument';
        CoreErrorCodes.connectionFailure = 'ConnectionFailure';
        CoreErrorCodes.timeout = 'Timeout';
        CoreErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
        CoreErrorCodes.invalidObjectPath = 'InvalidObjectPath';
        CoreErrorCodes.invalidRequestContext = 'InvalidRequestContext';
        CoreErrorCodes.valueNotLoaded = 'ValueNotLoaded';
        CoreErrorCodes.requestPayloadSizeLimitExceeded = 'RequestPayloadSizeLimitExceeded';
        CoreErrorCodes.responsePayloadSizeLimitExceeded = 'ResponsePayloadSizeLimitExceeded';
        CoreErrorCodes.writeNotSupportedWhenModalDialogOpen = 'WriteNotSupportedWhenModalDialogOpen';
        return CoreErrorCodes;
    }());
    OfficeExtension_1.CoreErrorCodes = CoreErrorCodes;
    var CoreResourceStrings = (function () {
        function CoreResourceStrings() {
        }
        CoreResourceStrings.apiNotFoundDetails = 'ApiNotFoundDetails';
        CoreResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
        CoreResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
        CoreResourceStrings.invalidArgument = 'InvalidArgument';
        CoreResourceStrings.invalidArgumentGeneric = 'InvalidArgumentGeneric';
        CoreResourceStrings.timeout = 'Timeout';
        CoreResourceStrings.invalidOrTimedOutSessionMessage = 'InvalidOrTimedOutSessionMessage';
        CoreResourceStrings.invalidSheetName = 'InvalidSheetName';
        CoreResourceStrings.invalidObjectPath = 'InvalidObjectPath';
        CoreResourceStrings.invalidRequestContext = 'InvalidRequestContext';
        CoreResourceStrings.valueNotLoaded = 'ValueNotLoaded';
        return CoreResourceStrings;
    }());
    OfficeExtension_1.CoreResourceStrings = CoreResourceStrings;
    var CoreConstants = (function () {
        function CoreConstants() {
        }
        CoreConstants.flags = 'flags';
        CoreConstants.sourceLibHeader = 'SdkVersion';
        CoreConstants.processQuery = 'ProcessQuery';
        CoreConstants.localDocument = 'http://document.localhost/';
        CoreConstants.localDocumentApiPrefix = 'http://document.localhost/_api/';
        CoreConstants.customData = 'customdata';
        CoreConstants.testRequestNameHeader = 'x-test-request-name';
        return CoreConstants;
    }());
    OfficeExtension_1.CoreConstants = CoreConstants;
    var RichApiMessageUtility = (function () {
        function RichApiMessageUtility() {
        }
        RichApiMessageUtility.buildMessageArrayForIRequestExecutor = function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            CoreUtility.log('Request:');
            CoreUtility.log(requestMessageText);
            var headers = {};
            CoreUtility._copyHeaders(requestMessage.Headers, headers);
            headers[CoreConstants.sourceLibHeader] = sourceLibHeaderValue;
            var messageSafearray = RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, 'POST', CoreConstants.processQuery, headers, requestMessageText);
            return messageSafearray;
        };
        RichApiMessageUtility.buildResponseOnSuccess = function (responseBody, responseHeaders) {
            var response = { HttpStatusCode: 200, ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.Body = JSON.parse(responseBody);
            response.Headers = responseHeaders;
            return response;
        };
        RichApiMessageUtility.buildResponseOnError = function (errorCode, message) {
            var response = { HttpStatusCode: 500, ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.ErrorCode = CoreErrorCodes.generalException;
            response.ErrorMessage = message;
            if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                response.ErrorCode = CoreErrorCodes.accessDenied;
                response.HttpStatusCode = 401;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                response.ErrorCode = CoreErrorCodes.activityLimitReached;
                response.HttpStatusCode = 429;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
                response.ErrorCode = CoreErrorCodes.invalidOrTimedOutSession;
                response.HttpStatusCode = 408;
                response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage);
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded) {
                response.ErrorCode = CoreErrorCodes.requestPayloadSizeLimitExceeded;
                response.HttpStatusCode = 400;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeResponsePayloadSizeLimitExceeded) {
                response.ErrorCode = CoreErrorCodes.responsePayloadSizeLimitExceeded;
                response.HttpStatusCode = 400;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeWriteNotSupportedWhenModalDialogOpen) {
                response.ErrorCode = CoreErrorCodes.writeNotSupportedWhenModalDialogOpen;
                response.HttpStatusCode = 400;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidSheetName) {
                response.ErrorCode = CoreErrorCodes.invalidRequestContext;
                response.HttpStatusCode = 400;
                response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidSheetName);
            }
            return response;
        };
        RichApiMessageUtility.buildHttpResponseFromOfficeJsError = function (errorCode, message) {
            var statusCode = 500;
            var errorBody = {};
            errorBody['error'] = {};
            errorBody['error']['code'] = CoreErrorCodes.generalException;
            errorBody['error']['message'] = message;
            if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                statusCode = 403;
                errorBody['error']['code'] = CoreErrorCodes.accessDenied;
            }
            else if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                statusCode = 429;
                errorBody['error']['code'] = CoreErrorCodes.activityLimitReached;
            }
            return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
        };
        RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) {
                for (var headerName in headers) {
                    headerArray.push(headerName);
                    headerArray.push(headers[headerName]);
                }
            }
            var appPermission = 0;
            var solutionId = '';
            var instanceId = '';
            var marketplaceType = '';
            var solutionVersion = '';
            var storeLocation = '';
            var compliantSolutionId = '';
            return [
                customData,
                method,
                path,
                headerArray,
                body,
                appPermission,
                requestFlags,
                solutionId,
                instanceId,
                marketplaceType,
                solutionVersion,
                storeLocation,
                compliantSolutionId
            ];
        };
        RichApiMessageUtility.getResponseBody = function (result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseHeaders = function (result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
            var ret = data[2];
            if (typeof ret === 'string') {
                return ret;
            }
            var arr = ret;
            return arr.join('');
        };
        RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
            var arrayHeader = data[1];
            if (!arrayHeader) {
                return null;
            }
            var headers = {};
            for (var i = 0; i < arrayHeader.length - 1; i += 2) {
                headers[arrayHeader[i]] = arrayHeader[i + 1];
            }
            return headers;
        };
        RichApiMessageUtility.getResponseStatusCode = function (result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
            return data[0];
        };
        RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession = 5012;
        RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached = 5102;
        RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability = 7000;
        RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded = 5103;
        RichApiMessageUtility.OfficeJsErrorCode_ooeResponsePayloadSizeLimitExceeded = 5104;
        RichApiMessageUtility.OfficeJsErrorCode_ooeWriteNotSupportedWhenModalDialogOpen = 5016;
        RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidSheetName = 1014;
        return RichApiMessageUtility;
    }());
    OfficeExtension_1.RichApiMessageUtility = RichApiMessageUtility;
    (function (_Internal) {
        function getPromiseType() {
            if (typeof Promise !== 'undefined') {
                return Promise;
            }
            if (typeof Office !== 'undefined') {
                if (Office.Promise) {
                    return Office.Promise;
                }
            }
            if (typeof OfficeExtension !== 'undefined') {
                if (OfficeExtension.Promise) {
                    return OfficeExtension.Promise;
                }
            }
            throw new _Internal.Error('No Promise implementation found');
        }
        _Internal.getPromiseType = getPromiseType;
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    var CoreUtility = (function () {
        function CoreUtility() {
        }
        CoreUtility.log = function (message) {
            if (CoreUtility._logEnabled && typeof console !== 'undefined' && console.log) {
                console.log(message);
            }
        };
        CoreUtility.checkArgumentNull = function (value, name) {
            if (CoreUtility.isNullOrUndefined(value)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: name });
            }
        };
        CoreUtility.isNullOrUndefined = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof value === 'undefined') {
                return true;
            }
            return false;
        };
        CoreUtility.isUndefined = function (value) {
            if (typeof value === 'undefined') {
                return true;
            }
            return false;
        };
        CoreUtility.isNullOrEmptyString = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof value === 'undefined') {
                return true;
            }
            if (value.length == 0) {
                return true;
            }
            return false;
        };
        CoreUtility.isPlainJsonObject = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return false;
            }
            if (typeof value !== 'object') {
                return false;
            }
            if (Object.prototype.toString.apply(value) !== '[object Object]') {
                return false;
            }
            if (value.constructor &&
                !Object.prototype.hasOwnProperty.call(value, 'constructor') &&
                !Object.prototype.hasOwnProperty.call(value.constructor.prototype, 'hasOwnProperty')) {
                return false;
            }
            for (var key in value) {
                if (!Object.prototype.hasOwnProperty.call(value, key)) {
                    return false;
                }
            }
            return true;
        };
        CoreUtility.trim = function (str) {
            return str.replace(new RegExp('^\\s+|\\s+$', 'g'), '');
        };
        CoreUtility.caseInsensitiveCompareString = function (str1, str2) {
            if (CoreUtility.isNullOrUndefined(str1)) {
                return CoreUtility.isNullOrUndefined(str2);
            }
            else {
                if (CoreUtility.isNullOrUndefined(str2)) {
                    return false;
                }
                else {
                    return str1.toUpperCase() == str2.toUpperCase();
                }
            }
        };
        CoreUtility.isReadonlyRestRequest = function (method) {
            return CoreUtility.caseInsensitiveCompareString(method, 'GET');
        };
        CoreUtility._getResourceString = function (resourceId, arg) {
            var ret;
            if (typeof window !== 'undefined' && window.Strings && window.Strings.OfficeOM) {
                var stringName = 'L_' + resourceId;
                var stringValue = window.Strings.OfficeOM[stringName];
                if (stringValue) {
                    ret = stringValue;
                }
            }
            if (!ret) {
                ret = CoreUtility.s_resourceStringValues[resourceId];
            }
            if (!ret) {
                ret = resourceId;
            }
            if (!CoreUtility.isNullOrUndefined(arg)) {
                if (Array.isArray(arg)) {
                    var arrArg = arg;
                    ret = CoreUtility._formatString(ret, arrArg);
                }
                else {
                    ret = ret.replace('{0}', arg);
                }
            }
            return ret;
        };
        CoreUtility._formatString = function (format, arrArg) {
            return format.replace(/\{\d\}/g, function (v) {
                var position = parseInt(v.substr(1, v.length - 2));
                if (position < arrArg.length) {
                    return arrArg[position];
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'format' });
                }
            });
        };
        Object.defineProperty(CoreUtility, "Promise", {
            get: function () {
                return _Internal.getPromiseType();
            },
            enumerable: true,
            configurable: true
        });
        CoreUtility.createPromise = function (executor) {
            var ret = new CoreUtility.Promise(executor);
            return ret;
        };
        CoreUtility._createPromiseFromResult = function (value) {
            return CoreUtility.createPromise(function (resolve, reject) {
                resolve(value);
            });
        };
        CoreUtility._createPromiseFromException = function (reason) {
            return CoreUtility.createPromise(function (resolve, reject) {
                reject(reason);
            });
        };
        CoreUtility._createTimeoutPromise = function (timeout) {
            return CoreUtility.createPromise(function (resolve, reject) {
                setTimeout(function () {
                    resolve(null);
                }, timeout);
            });
        };
        CoreUtility._createInvalidArgError = function (error) {
            return _Internal.RuntimeError._createInvalidArgError(error);
        };
        CoreUtility._isLocalDocumentUrl = function (url) {
            return CoreUtility._getLocalDocumentUrlPrefixLength(url) > 0;
        };
        CoreUtility._getLocalDocumentUrlPrefixLength = function (url) {
            var localDocumentPrefixes = [
                'http://document.localhost',
                'https://document.localhost',
                '//document.localhost'
            ];
            var urlLower = url.toLowerCase().trim();
            for (var i = 0; i < localDocumentPrefixes.length; i++) {
                if (urlLower === localDocumentPrefixes[i]) {
                    return localDocumentPrefixes[i].length;
                }
                else if (urlLower.substr(0, localDocumentPrefixes[i].length + 1) === localDocumentPrefixes[i] + '/') {
                    return localDocumentPrefixes[i].length + 1;
                }
            }
            return 0;
        };
        CoreUtility._validateLocalDocumentRequest = function (request) {
            var index = CoreUtility._getLocalDocumentUrlPrefixLength(request.url);
            if (index <= 0) {
                throw _Internal.RuntimeError._createInvalidArgError({
                    argumentName: 'request'
                });
            }
            var path = request.url.substr(index);
            var pathLower = path.toLowerCase();
            if (pathLower === '_api') {
                path = '';
            }
            else if (pathLower.substr(0, '_api/'.length) === '_api/') {
                path = path.substr('_api/'.length);
            }
            return {
                method: request.method,
                url: path,
                headers: request.headers,
                body: request.body
            };
        };
        CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny = function (queryString) {
            var ret = { flags: -1, customData: '' };
            var parts = queryString.split('&');
            for (var i = 0; i < parts.length; i++) {
                var keyvalue = parts[i].split('=');
                if (keyvalue[0].toLowerCase() === CoreConstants.flags) {
                    var flags = parseInt(keyvalue[1]);
                    flags = flags & 8191;
                    ret.flags = flags;
                }
                else if (keyvalue[0].toLowerCase() === CoreConstants.customData) {
                    ret.customData = decodeURIComponent(keyvalue[1]);
                }
            }
            return ret;
        };
        CoreUtility._getRequestBodyText = function (request) {
            var body = '';
            if (typeof request.body === 'string') {
                body = request.body;
            }
            else if (request.body && typeof request.body === 'object') {
                body = JSON.stringify(request.body);
            }
            return body;
        };
        CoreUtility._parseResponseBody = function (response) {
            if (typeof response.body === 'string') {
                var bodyText = CoreUtility.trim(response.body);
                return JSON.parse(bodyText);
            }
            else {
                return response.body;
            }
        };
        CoreUtility._buildRequestMessageSafeArray = function (request) {
            var requestFlags = 0;
            if (!CoreUtility.isReadonlyRestRequest(request.method)) {
                requestFlags = 1;
            }
            var customData = '';
            if (request.url.substr(0, CoreConstants.processQuery.length).toLowerCase() ===
                CoreConstants.processQuery.toLowerCase()) {
                var index = request.url.indexOf('?');
                if (index > 0) {
                    var queryString = request.url.substr(index + 1);
                    var flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(queryString);
                    if (flagsAndCustomData.flags >= 0) {
                        requestFlags = flagsAndCustomData.flags;
                    }
                    customData = flagsAndCustomData.customData;
                }
            }
            return RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, request.method, request.url, request.headers, CoreUtility._getRequestBodyText(request));
        };
        CoreUtility._parseHttpResponseHeaders = function (allResponseHeaders) {
            var responseHeaders = {};
            if (!CoreUtility.isNullOrEmptyString(allResponseHeaders)) {
                var regex = new RegExp('\r?\n');
                var entries = allResponseHeaders.split(regex);
                for (var i = 0; i < entries.length; i++) {
                    var entry = entries[i];
                    if (entry != null) {
                        var index = entry.indexOf(':');
                        if (index > 0) {
                            var key = entry.substr(0, index);
                            var value = entry.substr(index + 1);
                            key = CoreUtility.trim(key);
                            value = CoreUtility.trim(value);
                            responseHeaders[key.toUpperCase()] = value;
                        }
                    }
                }
            }
            return responseHeaders;
        };
        CoreUtility._parseErrorResponse = function (responseInfo) {
            var errorObj = CoreUtility._parseErrorResponseBody(responseInfo);
            var statusCode = responseInfo.statusCode.toString();
            if (CoreUtility.isNullOrUndefined(errorObj) || typeof errorObj !== 'object' || !errorObj.error) {
                return CoreUtility._createDefaultErrorResponse(statusCode);
            }
            var error = errorObj.error;
            var innerError = error.innerError;
            if (innerError && innerError.code) {
                return CoreUtility._createErrorResponse(innerError.code, statusCode, innerError.message);
            }
            if (error.code) {
                return CoreUtility._createErrorResponse(error.code, statusCode, error.message);
            }
            return CoreUtility._createDefaultErrorResponse(statusCode);
        };
        CoreUtility._parseErrorResponseBody = function (responseInfo) {
            if (CoreUtility.isPlainJsonObject(responseInfo.body)) {
                return responseInfo.body;
            }
            else if (!CoreUtility.isNullOrEmptyString(responseInfo.body)) {
                var errorResponseBody = CoreUtility.trim(responseInfo.body);
                try {
                    return JSON.parse(errorResponseBody);
                }
                catch (e) {
                    CoreUtility.log('Error when parse ' + errorResponseBody);
                }
            }
        };
        CoreUtility._createDefaultErrorResponse = function (statusCode) {
            return {
                errorCode: CoreErrorCodes.connectionFailure,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, statusCode)
            };
        };
        CoreUtility._createErrorResponse = function (code, statusCode, message) {
            return {
                errorCode: code,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [
                    statusCode,
                    code,
                    message
                ])
            };
        };
        CoreUtility._copyHeaders = function (src, dest) {
            if (src && dest) {
                for (var key in src) {
                    dest[key] = src[key];
                }
            }
        };
        CoreUtility.addResourceStringValues = function (values) {
            for (var key in values) {
                CoreUtility.s_resourceStringValues[key] = values[key];
            }
        };
        CoreUtility._logEnabled = false;
        CoreUtility.s_resourceStringValues = {
            ApiNotFoundDetails: 'The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.',
            ConnectionFailureWithStatus: 'The request failed with status code of {0}.',
            ConnectionFailureWithDetails: 'The request failed with status code of {0}, error code {1} and the following error message: {2}',
            InvalidArgument: "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",
            InvalidObjectPath: 'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',
            InvalidRequestContext: 'Cannot use the object across different request contexts.',
            Timeout: 'The operation has timed out.',
            ValueNotLoaded: 'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'
        };
        return CoreUtility;
    }());
    OfficeExtension_1.CoreUtility = CoreUtility;
    var TestUtility = (function () {
        function TestUtility() {
        }
        TestUtility.setMock = function (value) {
            TestUtility.s_isMock = value;
        };
        TestUtility.isMock = function () {
            return TestUtility.s_isMock;
        };
        TestUtility._setCurrentTestName = function (value) {
            TestUtility.s_currentTestName = value;
            TestUtility.s_currentTestSequenceId = 0;
        };
        TestUtility._getCurrentTestNameWithSequenceId = function () {
            if (TestUtility.s_currentTestName) {
                TestUtility.s_currentTestSequenceId++;
                return TestUtility.s_currentTestName + "." + TestUtility.s_currentTestSequenceId;
            }
            return null;
        };
        return TestUtility;
    }());
    OfficeExtension_1.TestUtility = TestUtility;
    OfficeExtension_1._internalConfig = {
        showDisposeInfoInDebugInfo: false,
        showInternalApiInDebugInfo: false,
        enableEarlyDispose: true,
        alwaysPolyfillClientObjectUpdateMethod: false,
        alwaysPolyfillClientObjectRetrieveMethod: false,
        enableConcurrentFlag: true,
        enableUndoableFlag: true,
        appendTypeNameToObjectPathInfo: false,
        enablePreviewExecution: false
    };
    OfficeExtension_1.config = {
        extendedErrorLogging: false
    };
    var CommonActionFactory = (function () {
        function CommonActionFactory() {
        }
        CommonActionFactory.createSetPropertyAction = function (context, parent, propertyName, value, flags) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 4,
                Name: propertyName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var args = [value];
            var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            CommonUtility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var action = new Action(actionInfo, 0, flags);
            action.referencedObjectPath = parent._objectPath;
            action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
            if (OfficeExtension_1._internalConfig.enablePreviewExecution && (flags & 16) !== 0) {
                var previewExecutionAction = {
                    Id: context._nextId(),
                    ActionType: 4,
                    Name: propertyName,
                    ObjectId: '',
                    ObjectType: '',
                    Arguments: [value]
                };
                parent._addPreviewExecutionAction(previewExecutionAction);
            }
            return parent._addAction(action);
        };
        CommonActionFactory.createQueryAction = function (context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 2,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action, resultHandler);
        };
        CommonActionFactory.createQueryAsJsonAction = function (context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 7,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action, resultHandler);
        };
        CommonActionFactory.createUpdateAction = function (context, parent, objectState) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 9,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            };
            var action = new Action(actionInfo, 0, 0);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action);
        };
        return CommonActionFactory;
    }());
    OfficeExtension_1.CommonActionFactory = CommonActionFactory;
    var ClientObjectBase = (function () {
        function ClientObjectBase(contextBase, objectPath) {
            this.m_contextBase = contextBase;
            this.m_objectPath = objectPath;
        }
        Object.defineProperty(ClientObjectBase.prototype, "_objectPath", {
            get: function () {
                return this.m_objectPath;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObjectBase.prototype, "_context", {
            get: function () {
                return this.m_contextBase;
            },
            enumerable: true,
            configurable: true
        });
        ClientObjectBase.prototype._addAction = function (action, resultHandler) {
            var _this = this;
            if (resultHandler === void 0) {
                resultHandler = null;
            }
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._context._addServiceApiAction(action, resultHandler, resolve, reject);
            });
        };
        ClientObjectBase.prototype._addPreviewExecutionAction = function (action) {
        };
        ClientObjectBase.prototype._retrieve = function (option, resultHandler) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.1');
            }
            var queryOption = ClientRequestContextBase._parseQueryOption(option);
            if (shouldPolyfill) {
                return CommonActionFactory.createQueryAction(this._context, this, queryOption, resultHandler);
            }
            return CommonActionFactory.createQueryAsJsonAction(this._context, this, queryOption, resultHandler);
        };
        ClientObjectBase.prototype._recursivelyUpdate = function (properties) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.2');
            }
            try {
                var scalarPropNames = this[CommonConstants.scalarPropertyNames];
                if (!scalarPropNames) {
                    scalarPropNames = [];
                }
                var scalarPropUpdatable = this[CommonConstants.scalarPropertyUpdateable];
                if (!scalarPropUpdatable) {
                    scalarPropUpdatable = [];
                    for (var i = 0; i < scalarPropNames.length; i++) {
                        scalarPropUpdatable.push(false);
                    }
                }
                var navigationPropNames = this[CommonConstants.navigationPropertyNames];
                if (!navigationPropNames) {
                    navigationPropNames = [];
                }
                var scalarProps = {};
                var navigationProps = {};
                var scalarPropCount = 0;
                for (var propName in properties) {
                    var index = scalarPropNames.indexOf(propName);
                    if (index >= 0) {
                        if (!scalarPropUpdatable[index]) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, propName),
                                debugInfo: {
                                    errorLocation: propName
                                }
                            });
                        }
                        scalarProps[propName] = properties[propName];
                        ++scalarPropCount;
                    }
                    else if (navigationPropNames.indexOf(propName) >= 0) {
                        navigationProps[propName] = properties[propName];
                    }
                    else {
                        throw new _Internal.RuntimeError({
                            code: CoreErrorCodes.invalidArgument,
                            httpStatusCode: 400,
                            message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, propName),
                            debugInfo: {
                                errorLocation: propName
                            }
                        });
                    }
                }
                if (scalarPropCount > 0) {
                    if (shouldPolyfill) {
                        for (var i = 0; i < scalarPropNames.length; i++) {
                            var propName = scalarPropNames[i];
                            var propValue = scalarProps[propName];
                            if (!CommonUtility.isUndefined(propValue)) {
                                CommonActionFactory.createSetPropertyAction(this._context, this, propName, propValue);
                            }
                        }
                    }
                    else {
                        CommonActionFactory.createUpdateAction(this._context, this, scalarProps);
                    }
                }
                for (var propName in navigationProps) {
                    var navigationPropProxy = this[propName];
                    var navigationPropValue = navigationProps[propName];
                    navigationPropProxy._recursivelyUpdate(navigationPropValue);
                }
            }
            catch (innerError) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
                    debugInfo: {
                        errorLocation: this._className + '.update'
                    },
                    innerError: innerError
                });
            }
        };
        return ClientObjectBase;
    }());
    OfficeExtension_1.ClientObjectBase = ClientObjectBase;
    var Action = (function () {
        function Action(actionInfo, operationType, flags) {
            this.m_actionInfo = actionInfo;
            this.m_operationType = operationType;
            this.m_flags = flags;
        }
        Object.defineProperty(Action.prototype, "actionInfo", {
            get: function () {
                return this.m_actionInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "operationType", {
            get: function () {
                return this.m_operationType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });
        return Action;
    }());
    OfficeExtension_1.Action = Action;
    var ObjectPath = (function () {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest, operationType, flags) {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
            this.m_operationType = operationType;
            this.m_flags = flags;
        }
        Object.defineProperty(ObjectPath.prototype, "id", {
            get: function () {
                var argumentInfo = this.m_objectPathInfo.ArgumentInfo;
                if (!argumentInfo) {
                    return undefined;
                }
                var argument = argumentInfo.Arguments;
                if (!argument) {
                    return undefined;
                }
                return argument[0];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parent", {
            get: function () {
                var parent = this.m_parentObjectPath;
                if (!parent) {
                    return undefined;
                }
                return parent;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parentId", {
            get: function () {
                return this.parent ? this.parent.id : undefined;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function () {
                return this.m_objectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "operationType", {
            get: function () {
                return this.m_operationType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isCollection", {
            get: function () {
                return this.m_isCollection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
            get: function () {
                return this.m_isInvalidAfterRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
            get: function () {
                return this.m_parentObjectPath;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
            get: function () {
                return this.m_argumentObjectPaths;
            },
            set: function (value) {
                this.m_argumentObjectPaths = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function () {
                return this.m_isValid;
            },
            set: function (value) {
                this.m_isValid = value;
                if (!value &&
                    this.m_objectPathInfo.ObjectPathType === 6 &&
                    this.m_savedObjectPathInfo) {
                    ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
                    this.m_parentObjectPath = this.m_savedObjectPathInfo.parent;
                    this.m_isValid = true;
                    this.m_savedObjectPathInfo = null;
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
            get: function () {
                return this.m_originalObjectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
            get: function () {
                return this.m_getByIdMethodName;
            },
            set: function (value) {
                this.m_getByIdMethodName = value;
            },
            enumerable: true,
            configurable: true
        });
        ObjectPath.prototype._updateAsNullObject = function () {
            this.resetForUpdateUsingObjectData();
            this.m_objectPathInfo.ObjectPathType = 7;
            this.m_objectPathInfo.Name = '';
            this.m_parentObjectPath = null;
        };
        ObjectPath.prototype.saveOriginalObjectPathInfo = function () {
            if (OfficeExtension_1.config.extendedErrorLogging && !this.m_originalObjectPathInfo) {
                this.m_originalObjectPathInfo = {};
                ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo);
            }
        };
        ObjectPath.prototype.updateUsingObjectData = function (value, clientObject) {
            var referenceId = value[CommonConstants.referenceId];
            if (!CoreUtility.isNullOrEmptyString(referenceId)) {
                if (!this.m_savedObjectPathInfo &&
                    !this.isInvalidAfterRequest &&
                    ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
                    var pathInfo = {};
                    ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
                    this.m_savedObjectPathInfo = {
                        pathInfo: pathInfo,
                        parent: this.m_parentObjectPath
                    };
                }
                this.saveOriginalObjectPathInfo();
                this.resetForUpdateUsingObjectData();
                this.m_objectPathInfo.ObjectPathType = 6;
                this.m_objectPathInfo.Name = referenceId;
                delete this.m_objectPathInfo.ParentObjectPathId;
                this.m_parentObjectPath = null;
                return;
            }
            if (clientObject) {
                var collectionPropertyPath = clientObject[CommonConstants.collectionPropertyPath];
                if (!CoreUtility.isNullOrEmptyString(collectionPropertyPath) && clientObject.context) {
                    var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                    if (!CoreUtility.isNullOrUndefined(id)) {
                        var propNames = collectionPropertyPath.split('.');
                        var parent_1 = clientObject.context[propNames[0]];
                        for (var i = 1; i < propNames.length; i++) {
                            parent_1 = parent_1[propNames[i]];
                        }
                        this.saveOriginalObjectPathInfo();
                        this.resetForUpdateUsingObjectData();
                        this.m_parentObjectPath = parent_1._objectPath;
                        this.m_objectPathInfo.ParentObjectPathId = this.m_parentObjectPath.objectPathInfo.Id;
                        this.m_objectPathInfo.ObjectPathType = 5;
                        this.m_objectPathInfo.Name = '';
                        this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                        return;
                    }
                }
            }
            var parentIsCollection = this.parentObjectPath && this.parentObjectPath.isCollection;
            var getByIdMethodName = this.getByIdMethodName;
            if (parentIsCollection || !CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                if (!CoreUtility.isNullOrUndefined(id)) {
                    this.saveOriginalObjectPathInfo();
                    this.resetForUpdateUsingObjectData();
                    if (!CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                        this.m_objectPathInfo.ObjectPathType = 3;
                        this.m_objectPathInfo.Name = getByIdMethodName;
                    }
                    else {
                        this.m_objectPathInfo.ObjectPathType = 5;
                        this.m_objectPathInfo.Name = '';
                    }
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    return;
                }
            }
        };
        ObjectPath.prototype.resetForUpdateUsingObjectData = function () {
            this.m_isInvalidAfterRequest = false;
            this.m_isValid = true;
            this.m_operationType = 1;
            this.m_flags = 4;
            this.m_objectPathInfo.ArgumentInfo = {};
            this.m_argumentObjectPaths = null;
            this.m_getByIdMethodName = null;
        };
        ObjectPath.isRestorableObjectPath = function (objectPathType) {
            return (objectPathType === 1 ||
                objectPathType === 5 ||
                objectPathType === 3 ||
                objectPathType === 4);
        };
        ObjectPath.copyObjectPathInfo = function (src, dest) {
            dest.Id = src.Id;
            dest.ArgumentInfo = src.ArgumentInfo;
            dest.Name = src.Name;
            dest.ObjectPathType = src.ObjectPathType;
            dest.ParentObjectPathId = src.ParentObjectPathId;
        };
        return ObjectPath;
    }());
    OfficeExtension_1.ObjectPath = ObjectPath;
    var ClientRequestContextBase = (function () {
        function ClientRequestContextBase() {
            this.m_nextId = 0;
        }
        ClientRequestContextBase.prototype._nextId = function () {
            return ++this.m_nextId;
        };
        ClientRequestContextBase.prototype._addServiceApiAction = function (action, resultHandler, resolve, reject) {
            if (!this.m_serviceApiQueue) {
                this.m_serviceApiQueue = new ServiceApiQueue(this);
            }
            this.m_serviceApiQueue.add(action, resultHandler, resolve, reject);
        };
        ClientRequestContextBase._parseQueryOption = function (option) {
            var queryOption = {};
            if (typeof option === 'string') {
                var select = option;
                queryOption.Select = CommonUtility._parseSelectExpand(select);
            }
            else if (Array.isArray(option)) {
                queryOption.Select = option;
            }
            else if (typeof option === 'object') {
                var loadOption = option;
                if (ClientRequestContextBase.isLoadOption(loadOption)) {
                    if (typeof loadOption.select === 'string') {
                        queryOption.Select = CommonUtility._parseSelectExpand(loadOption.select);
                    }
                    else if (Array.isArray(loadOption.select)) {
                        queryOption.Select = loadOption.select;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.select)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.select' });
                    }
                    if (typeof loadOption.expand === 'string') {
                        queryOption.Expand = CommonUtility._parseSelectExpand(loadOption.expand);
                    }
                    else if (Array.isArray(loadOption.expand)) {
                        queryOption.Expand = loadOption.expand;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.expand)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.expand' });
                    }
                    if (typeof loadOption.top === 'number') {
                        queryOption.Top = loadOption.top;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.top)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.top' });
                    }
                    if (typeof loadOption.skip === 'number') {
                        queryOption.Skip = loadOption.skip;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.skip)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.skip' });
                    }
                }
                else {
                    queryOption = ClientRequestContextBase.parseStrictLoadOption(option);
                }
            }
            else if (!CommonUtility.isNullOrUndefined(option)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option' });
            }
            return queryOption;
        };
        ClientRequestContextBase.isLoadOption = function (loadOption) {
            if (!CommonUtility.isUndefined(loadOption.select) &&
                (typeof loadOption.select === 'string' || Array.isArray(loadOption.select)))
                return true;
            if (!CommonUtility.isUndefined(loadOption.expand) &&
                (typeof loadOption.expand === 'string' || Array.isArray(loadOption.expand)))
                return true;
            if (!CommonUtility.isUndefined(loadOption.top) && typeof loadOption.top === 'number')
                return true;
            if (!CommonUtility.isUndefined(loadOption.skip) && typeof loadOption.skip === 'number')
                return true;
            for (var i in loadOption) {
                return false;
            }
            return true;
        };
        ClientRequestContextBase.parseStrictLoadOption = function (option) {
            var ret = { Select: [] };
            ClientRequestContextBase.parseStrictLoadOptionHelper(ret, '', 'option', option);
            return ret;
        };
        ClientRequestContextBase.combineQueryPath = function (pathPrefix, key, separator) {
            if (pathPrefix.length === 0) {
                return key;
            }
            else {
                return pathPrefix + separator + key;
            }
        };
        ClientRequestContextBase.parseStrictLoadOptionHelper = function (queryInfo, pathPrefix, argPrefix, option) {
            for (var key in option) {
                var value = option[key];
                if (key === '$all') {
                    if (typeof value !== 'boolean') {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    if (value) {
                        queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, '*', '/'));
                    }
                }
                else if (key === '$top') {
                    if (typeof value !== 'number' || pathPrefix.length > 0) {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    queryInfo.Top = value;
                }
                else if (key === '$skip') {
                    if (typeof value !== 'number' || pathPrefix.length > 0) {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    queryInfo.Skip = value;
                }
                else {
                    if (typeof value === 'boolean') {
                        if (value) {
                            queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'));
                        }
                    }
                    else if (typeof value === 'object') {
                        ClientRequestContextBase.parseStrictLoadOptionHelper(queryInfo, ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'), ClientRequestContextBase.combineQueryPath(argPrefix, key, '.'), value);
                    }
                    else {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                }
            }
        };
        return ClientRequestContextBase;
    }());
    OfficeExtension_1.ClientRequestContextBase = ClientRequestContextBase;
    var InstantiateActionUpdateObjectPathHandler = (function () {
        function InstantiateActionUpdateObjectPathHandler(m_objectPath) {
            this.m_objectPath = m_objectPath;
        }
        InstantiateActionUpdateObjectPathHandler.prototype._handleResult = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                this.m_objectPath._updateAsNullObject();
            }
            else {
                this.m_objectPath.updateUsingObjectData(value, null);
            }
        };
        return InstantiateActionUpdateObjectPathHandler;
    }());
    var ClientRequestBase = (function () {
        function ClientRequestBase(context) {
            this.m_contextBase = context;
            this.m_actions = [];
            this.m_actionResultHandler = {};
            this.m_referencedObjectPaths = {};
            this.m_instantiatedObjectPaths = {};
            this.m_preSyncPromises = [];
            this.m_previewExecutionActions = [];
        }
        ClientRequestBase.prototype.addAction = function (action) {
            this.m_actions.push(action);
            if (action.actionInfo.ActionType == 1) {
                this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId] = action;
            }
        };
        ClientRequestBase.prototype.addPreviewExecutionAction = function (action) {
            this.m_previewExecutionActions.push(action);
        };
        Object.defineProperty(ClientRequestBase.prototype, "hasActions", {
            get: function () {
                return this.m_actions.length > 0;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestBase.prototype._getLastAction = function () {
            return this.m_actions[this.m_actions.length - 1];
        };
        ClientRequestBase.prototype.ensureInstantiateObjectPath = function (objectPath) {
            if (objectPath) {
                if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                    return;
                }
                this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
                this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
                if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                    var actionInfo = {
                        Id: this.m_contextBase._nextId(),
                        ActionType: 1,
                        Name: '',
                        ObjectPathId: objectPath.objectPathInfo.Id
                    };
                    var instantiateAction = new Action(actionInfo, 1, 4);
                    instantiateAction.referencedObjectPath = objectPath;
                    this.addReferencedObjectPath(objectPath);
                    this.addAction(instantiateAction);
                    var resultHandler = new InstantiateActionUpdateObjectPathHandler(objectPath);
                    this.addActionResultHandler(instantiateAction, resultHandler);
                }
            }
        };
        ClientRequestBase.prototype.ensureInstantiateObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.ensureInstantiateObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequestBase.prototype.addReferencedObjectPath = function (objectPath) {
            if (!objectPath || this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }
            if (!objectPath.isValid) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidObjectPath,
                    httpStatusCode: 400,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                    }
                });
            }
            while (objectPath) {
                this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;
                if (objectPath.objectPathInfo.ObjectPathType == 3) {
                    this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequestBase.prototype.addReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.addReferencedObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequestBase.prototype.addActionResultHandler = function (action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        };
        ClientRequestBase.prototype.aggregrateRequestFlags = function (requestFlags, operationType, flags) {
            if (operationType === 0) {
                requestFlags = requestFlags | 1;
                if ((flags & 2) === 0) {
                    requestFlags = requestFlags & ~16;
                }
                if ((flags & 8) === 0) {
                    requestFlags = requestFlags & ~256;
                }
                requestFlags = requestFlags & ~4;
            }
            if (flags & 1) {
                requestFlags = requestFlags | 2;
            }
            if ((flags & 4) === 0) {
                requestFlags = requestFlags & ~4;
            }
            return requestFlags;
        };
        ClientRequestBase.prototype.finallyNormalizeFlags = function (requestFlags) {
            if ((requestFlags & 1) === 0) {
                requestFlags = requestFlags & ~16;
                requestFlags = requestFlags & ~256;
            }
            if (!OfficeExtension_1._internalConfig.enableConcurrentFlag) {
                requestFlags = requestFlags & ~4;
            }
            if (!OfficeExtension_1._internalConfig.enableUndoableFlag) {
                requestFlags = requestFlags & ~16;
            }
            if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.1')) {
                requestFlags = requestFlags & ~4;
                requestFlags = requestFlags & ~16;
            }
            if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.2')) {
                requestFlags = requestFlags & ~256;
            }
            if (typeof this.m_flagsForTesting === 'number') {
                requestFlags = this.m_flagsForTesting;
            }
            return requestFlags;
        };
        ClientRequestBase.prototype.buildRequestMessageBodyAndRequestFlags = function () {
            if (OfficeExtension_1._internalConfig.enableEarlyDispose) {
                ClientRequestBase._calculateLastUsedObjectPathIds(this.m_actions);
            }
            var requestFlags = 4 |
                16 |
                256;
            var objectPaths = {};
            for (var i in this.m_referencedObjectPaths) {
                requestFlags = this.aggregrateRequestFlags(requestFlags, this.m_referencedObjectPaths[i].operationType, this.m_referencedObjectPaths[i].flags);
                objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            }
            var actions = [];
            var hasKeepReference = false;
            for (var index = 0; index < this.m_actions.length; index++) {
                var action = this.m_actions[index];
                if (action.actionInfo.ActionType === 3 &&
                    action.actionInfo.Name === CommonConstants.keepReference) {
                    hasKeepReference = true;
                }
                requestFlags = this.aggregrateRequestFlags(requestFlags, action.operationType, action.flags);
                actions.push(action.actionInfo);
            }
            requestFlags = this.finallyNormalizeFlags(requestFlags);
            var body = {
                AutoKeepReference: this.m_contextBase._autoCleanup && hasKeepReference,
                Actions: actions,
                ObjectPaths: objectPaths
            };
            if (this.m_previewExecutionActions.length > 0) {
                body.PreviewExecutionActions = this.m_previewExecutionActions;
                requestFlags = requestFlags | 4096;
            }
            return {
                body: body,
                flags: requestFlags
            };
        };
        ClientRequestBase.prototype.processResponse = function (actionResults) {
            if (actionResults) {
                for (var i = 0; i < actionResults.length; i++) {
                    var actionResult = actionResults[i];
                    var handler = this.m_actionResultHandler[actionResult.ActionId];
                    if (handler) {
                        handler._handleResult(actionResult.Value);
                    }
                }
            }
        };
        ClientRequestBase.prototype.invalidatePendingInvalidObjectPaths = function () {
            for (var i in this.m_referencedObjectPaths) {
                if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                    this.m_referencedObjectPaths[i].isValid = false;
                }
            }
        };
        ClientRequestBase.prototype._addPreSyncPromise = function (value) {
            this.m_preSyncPromises.push(value);
        };
        Object.defineProperty(ClientRequestBase.prototype, "_preSyncPromises", {
            get: function () {
                return this.m_preSyncPromises;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestBase.prototype, "_actions", {
            get: function () {
                return this.m_actions;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestBase.prototype, "_objectPaths", {
            get: function () {
                return this.m_referencedObjectPaths;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestBase.prototype._removeKeepReferenceAction = function (objectPathId) {
            for (var i = this.m_actions.length - 1; i >= 0; i--) {
                var actionInfo = this.m_actions[i].actionInfo;
                if (actionInfo.ObjectPathId === objectPathId &&
                    actionInfo.ActionType === 3 &&
                    actionInfo.Name === CommonConstants.keepReference) {
                    this.m_actions.splice(i, 1);
                    break;
                }
            }
        };
        ClientRequestBase._updateLastUsedActionIdOfObjectPathId = function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
            while (objectPath) {
                if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
                    return;
                }
                lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id] = actionId;
                var argumentObjectPaths = objectPath.argumentObjectPaths;
                if (argumentObjectPaths) {
                    var argumentObjectPathsLength = argumentObjectPaths.length;
                    for (var i = 0; i < argumentObjectPathsLength; i++) {
                        ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
                    }
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequestBase._calculateLastUsedObjectPathIds = function (actions) {
            var lastUsedActionIdOfObjectPathId = {};
            var actionsLength = actions.length;
            for (var index = actionsLength - 1; index >= 0; --index) {
                var action = actions[index];
                var actionId = action.actionInfo.Id;
                if (action.referencedObjectPath) {
                    ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
                }
                var referencedObjectPaths = action.referencedArgumentObjectPaths;
                if (referencedObjectPaths) {
                    var referencedObjectPathsLength = referencedObjectPaths.length;
                    for (var refIndex = 0; refIndex < referencedObjectPathsLength; refIndex++) {
                        ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
                    }
                }
            }
            var lastUsedObjectPathIdsOfAction = {};
            for (var key in lastUsedActionIdOfObjectPathId) {
                var actionId = lastUsedActionIdOfObjectPathId[key];
                var objectPathIds = lastUsedObjectPathIdsOfAction[actionId];
                if (!objectPathIds) {
                    objectPathIds = [];
                    lastUsedObjectPathIdsOfAction[actionId] = objectPathIds;
                }
                objectPathIds.push(parseInt(key));
            }
            for (var index = 0; index < actionsLength; index++) {
                var action = actions[index];
                var lastUsedObjectPathIds = lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
                if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                    action.actionInfo.L = lastUsedObjectPathIds;
                }
                else if (action.actionInfo.L) {
                    delete action.actionInfo.L;
                }
            }
        };
        return ClientRequestBase;
    }());
    OfficeExtension_1.ClientRequestBase = ClientRequestBase;
    var ClientResult = (function () {
        function ClientResult(m_type) {
            this.m_type = m_type;
        }
        Object.defineProperty(ClientResult.prototype, "value", {
            get: function () {
                if (!this.m_isLoaded) {
                    throw new _Internal.RuntimeError({
                        code: CoreErrorCodes.valueNotLoaded,
                        httpStatusCode: 400,
                        message: CoreUtility._getResourceString(CoreResourceStrings.valueNotLoaded),
                        debugInfo: {
                            errorLocation: 'clientResult.value'
                        }
                    });
                }
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ClientResult.prototype._handleResult = function (value) {
            this.m_isLoaded = true;
            if (typeof value === 'object' && value && value._IsNull) {
                return;
            }
            if (this.m_type === 1) {
                this.m_value = CommonUtility.adjustToDateTime(value);
            }
            else {
                this.m_value = value;
            }
        };
        return ClientResult;
    }());
    OfficeExtension_1.ClientResult = ClientResult;
    var ServiceApiQueue = (function () {
        function ServiceApiQueue(m_context) {
            this.m_context = m_context;
            this.m_actions = [];
        }
        ServiceApiQueue.prototype.add = function (action, resultHandler, resolve, reject) {
            var _this = this;
            this.m_actions.push({ action: action, resultHandler: resultHandler, resolve: resolve, reject: reject });
            if (this.m_actions.length === 1) {
                setTimeout(function () { return _this.processActions(); }, 0);
            }
        };
        ServiceApiQueue.prototype.processActions = function () {
            var _this = this;
            if (this.m_actions.length === 0) {
                return;
            }
            var actions = this.m_actions;
            this.m_actions = [];
            var request = new ClientRequestBase(this.m_context);
            for (var i = 0; i < actions.length; i++) {
                var action = actions[i];
                request.ensureInstantiateObjectPath(action.action.referencedObjectPath);
                request.ensureInstantiateObjectPaths(action.action.referencedArgumentObjectPaths);
                request.addAction(action.action);
                request.addReferencedObjectPath(action.action.referencedObjectPath);
                request.addReferencedObjectPaths(action.action.referencedArgumentObjectPaths);
            }
            var _a = request.buildRequestMessageBodyAndRequestFlags(), body = _a.body, flags = _a.flags;
            var requestMessage = {
                Url: CoreConstants.localDocumentApiPrefix,
                Headers: null,
                Body: body
            };
            CoreUtility.log('Request:');
            CoreUtility.log(JSON.stringify(body));
            var executor = new HttpRequestExecutor();
            executor
                .executeAsync(this.m_context._customData, flags, requestMessage)
                .then(function (response) {
                _this.processResponse(request, actions, response);
            })["catch"](function (ex) {
                for (var i = 0; i < actions.length; i++) {
                    var action = actions[i];
                    action.reject(ex);
                }
            });
        };
        ServiceApiQueue.prototype.processResponse = function (request, actions, response) {
            var error = this.getErrorFromResponse(response);
            var actionResults = null;
            if (response.Body.Results) {
                actionResults = response.Body.Results;
            }
            else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                actionResults = response.Body.ProcessedResults.Results;
            }
            if (!actionResults) {
                actionResults = [];
            }
            this.processActionResults(request, actions, actionResults, error);
        };
        ServiceApiQueue.prototype.getErrorFromResponse = function (response) {
            if (!CoreUtility.isNullOrEmptyString(response.ErrorCode)) {
                return new _Internal.RuntimeError({
                    code: response.ErrorCode,
                    httpStatusCode: response.HttpStatusCode,
                    message: response.ErrorMessage
                });
            }
            if (response.Body && response.Body.Error) {
                return new _Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    httpStatusCode: response.Body.Error.HttpStatusCode,
                    message: response.Body.Error.Message
                });
            }
            return null;
        };
        ServiceApiQueue.prototype.processActionResults = function (request, actions, actionResults, err) {
            request.processResponse(actionResults);
            for (var i = 0; i < actions.length; i++) {
                var action = actions[i];
                var actionId = action.action.actionInfo.Id;
                var hasResult = false;
                for (var j = 0; j < actionResults.length; j++) {
                    if (actionId == actionResults[j].ActionId) {
                        var resultValue = actionResults[j].Value;
                        if (action.resultHandler) {
                            action.resultHandler._handleResult(resultValue);
                            resultValue = action.resultHandler.value;
                        }
                        if (action.resolve) {
                            action.resolve(resultValue);
                        }
                        hasResult = true;
                        break;
                    }
                }
                if (!hasResult && action.reject) {
                    if (err) {
                        action.reject(err);
                    }
                    else {
                        action.reject('No response for the action.');
                    }
                }
            }
        };
        return ServiceApiQueue;
    }());
    var HttpRequestExecutor = (function () {
        function HttpRequestExecutor() {
        }
        HttpRequestExecutor.prototype.getRequestUrl = function (baseUrl, requestFlags) {
            if (baseUrl.charAt(baseUrl.length - 1) != '/') {
                baseUrl = baseUrl + '/';
            }
            baseUrl = baseUrl + CoreConstants.processQuery;
            baseUrl = baseUrl + '?' + CoreConstants.flags + '=' + requestFlags.toString();
            return baseUrl;
        };
        HttpRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var url = this.getRequestUrl(requestMessage.Url, requestFlags);
            var requestInfo = {
                method: 'POST',
                url: url,
                headers: {},
                body: requestMessage.Body
            };
            requestInfo.headers[CoreConstants.sourceLibHeader] = HttpRequestExecutor.SourceLibHeaderValue;
            requestInfo.headers['CONTENT-TYPE'] = 'application/json';
            if (requestMessage.Headers) {
                for (var key in requestMessage.Headers) {
                    requestInfo.headers[key] = requestMessage.Headers[key];
                }
            }
            var sendRequestFunc = CoreUtility._isLocalDocumentUrl(requestInfo.url)
                ? HttpUtility.sendLocalDocumentRequest
                : HttpUtility.sendRequest;
            return sendRequestFunc(requestInfo).then(function (responseInfo) {
                var response;
                if (responseInfo.statusCode === 200) {
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: null,
                        ErrorMessage: null,
                        Headers: responseInfo.headers,
                        Body: CoreUtility._parseResponseBody(responseInfo)
                    };
                }
                else {
                    CoreUtility.log('Error Response:' + responseInfo.body);
                    var error = CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null,
                        RawErrorResponseBody: CoreUtility._parseErrorResponseBody(responseInfo)
                    };
                }
                return response;
            });
        };
        HttpRequestExecutor.SourceLibHeaderValue = 'officejs-rest';
        return HttpRequestExecutor;
    }());
    OfficeExtension_1.HttpRequestExecutor = HttpRequestExecutor;
    var CommonConstants = (function (_super) {
        __extends(CommonConstants, _super);
        function CommonConstants() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonConstants.collectionPropertyPath = '_collectionPropertyPath';
        CommonConstants.id = 'Id';
        CommonConstants.idLowerCase = 'id';
        CommonConstants.idPrivate = '_Id';
        CommonConstants.keepReference = '_KeepReference';
        CommonConstants.objectPathIdPrivate = '_ObjectPathId';
        CommonConstants.referenceId = '_ReferenceId';
        CommonConstants.items = '_Items';
        CommonConstants.itemsLowerCase = 'items';
        CommonConstants.scalarPropertyNames = '_scalarPropertyNames';
        CommonConstants.scalarPropertyOriginalNames = '_scalarPropertyOriginalNames';
        CommonConstants.navigationPropertyNames = '_navigationPropertyNames';
        CommonConstants.scalarPropertyUpdateable = '_scalarPropertyUpdateable';
        CommonConstants.previewExecutionObjectId = '_previewExecutionObjectId';
        return CommonConstants;
    }(CoreConstants));
    OfficeExtension_1.CommonConstants = CommonConstants;
    var CommonUtility = (function (_super) {
        __extends(CommonUtility, _super);
        function CommonUtility() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonUtility.validateObjectPath = function (clientObject) {
            var objectPath = clientObject._objectPath;
            while (objectPath) {
                if (!objectPath.isValid) {
                    throw new _Internal.RuntimeError({
                        code: CoreErrorCodes.invalidObjectPath,
                        httpStatusCode: 400,
                        message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                        debugInfo: {
                            errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                        }
                    });
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        CommonUtility.validateReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    var objectPath = objectPaths[i];
                    while (objectPath) {
                        if (!objectPath.isValid) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidObjectPath,
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath))
                            });
                        }
                        objectPath = objectPath.parentObjectPath;
                    }
                }
            }
        };
        CommonUtility._toCamelLowerCase = function (name) {
            if (CoreUtility.isNullOrEmptyString(name)) {
                return name;
            }
            var index = 0;
            while (index < name.length && name.charCodeAt(index) >= 65 && name.charCodeAt(index) <= 90) {
                index++;
            }
            if (index < name.length) {
                return name.substr(0, index).toLowerCase() + name.substr(index);
            }
            else {
                return name.toLowerCase();
            }
        };
        CommonUtility.adjustToDateTime = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return null;
            }
            if (typeof value === 'string') {
                return new Date(value);
            }
            if (Array.isArray(value)) {
                var arr = value;
                for (var i = 0; i < arr.length; i++) {
                    arr[i] = CommonUtility.adjustToDateTime(arr[i]);
                }
                return arr;
            }
            throw CoreUtility._createInvalidArgError({ argumentName: 'date' });
        };
        CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult = function (value) {
            var id = value[CommonConstants.id];
            if (CoreUtility.isNullOrUndefined(id)) {
                id = value[CommonConstants.idLowerCase];
            }
            if (CoreUtility.isNullOrUndefined(id)) {
                id = value[CommonConstants.idPrivate];
            }
            return id;
        };
        CommonUtility.getObjectPathExpression = function (objectPath) {
            var ret = '';
            while (objectPath) {
                switch (objectPath.objectPathInfo.ObjectPathType) {
                    case 1:
                        ret = ret;
                        break;
                    case 2:
                        ret = 'new()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 3:
                        ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + '()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 4:
                        ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 5:
                        ret = 'getItem()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 6:
                        ret = '_reference()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                }
                objectPath = objectPath.parentObjectPath;
            }
            return ret;
        };
        CommonUtility.setMethodArguments = function (context, argumentInfo, args) {
            if (CoreUtility.isNullOrUndefined(args)) {
                return null;
            }
            var referencedObjectPaths = new Array();
            var referencedObjectPathIds = new Array();
            var hasOne = CommonUtility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
            argumentInfo.Arguments = args;
            if (hasOne) {
                argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
            }
            return referencedObjectPaths;
        };
        CommonUtility.validateContext = function (context, obj) {
            if (context && obj && obj._context !== context) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidRequestContext,
                    httpStatusCode: 400,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidRequestContext)
                });
            }
        };
        CommonUtility.isSetSupported = function (apiSetName, apiSetVersion) {
            if (typeof window !== 'undefined' &&
                window.Office &&
                window.Office.context &&
                window.Office.context.requirements) {
                return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
            }
            return true;
        };
        CommonUtility.throwIfApiNotSupported = function (apiFullName, apiSetName, apiSetVersion, hostName) {
            if (!CommonUtility._doApiNotSupportedCheck) {
                return;
            }
            if (!CommonUtility.isSetSupported(apiSetName, apiSetVersion)) {
                var message = CoreUtility._getResourceString(CoreResourceStrings.apiNotFoundDetails, [
                    apiFullName,
                    apiSetName + ' ' + apiSetVersion,
                    hostName
                ]);
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.apiNotFound,
                    httpStatusCode: 404,
                    message: message,
                    debugInfo: { errorLocation: apiFullName }
                });
            }
        };
        CommonUtility.calculateApiFlags = function (apiFlags, undoableApiSetName, undoableApiSetVersion) {
            if (!CommonUtility.isSetSupported(undoableApiSetName, undoableApiSetVersion)) {
                apiFlags = apiFlags & (~2);
            }
            return apiFlags;
        };
        CommonUtility._parseSelectExpand = function (select) {
            var args = [];
            if (!CoreUtility.isNullOrEmptyString(select)) {
                var propertyNames = select.split(',');
                for (var i = 0; i < propertyNames.length; i++) {
                    var propertyName = propertyNames[i];
                    propertyName = sanitizeForAnyItemsSlash(propertyName.trim());
                    if (propertyName.length > 0) {
                        args.push(propertyName);
                    }
                }
            }
            return args;
            function sanitizeForAnyItemsSlash(propertyName) {
                var propertyNameLower = propertyName.toLowerCase();
                if (propertyNameLower === 'items' || propertyNameLower === 'items/') {
                    return '*';
                }
                var itemsSlashLength = 6;
                var isItemsSlashOrItemsDot = propertyNameLower.substr(0, itemsSlashLength) === 'items/' ||
                    propertyNameLower.substr(0, itemsSlashLength) === 'items.';
                if (isItemsSlashOrItemsDot) {
                    propertyName = propertyName.substr(itemsSlashLength);
                }
                return propertyName.replace(new RegExp('[/.]items[/.]', 'gi'), '/');
            }
        };
        CommonUtility.changePropertyNameToCamelLowerCase = function (value) {
            var charCodeUnderscore = 95;
            if (Array.isArray(value)) {
                var ret = [];
                for (var i = 0; i < value.length; i++) {
                    ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
                }
                return ret;
            }
            else if (typeof value === 'object' && value !== null) {
                var ret = {};
                for (var key in value) {
                    var propValue = value[key];
                    if (key === CommonConstants.items) {
                        ret = {};
                        ret[CommonConstants.itemsLowerCase] = this.changePropertyNameToCamelLowerCase(propValue);
                        break;
                    }
                    else {
                        var propName = CommonUtility._toCamelLowerCase(key);
                        ret[propName] = this.changePropertyNameToCamelLowerCase(propValue);
                    }
                }
                return ret;
            }
            else {
                return value;
            }
        };
        CommonUtility.purifyJson = function (value) {
            var charCodeUnderscore = 95;
            if (Array.isArray(value)) {
                var ret = [];
                for (var i = 0; i < value.length; i++) {
                    ret.push(this.purifyJson(value[i]));
                }
                return ret;
            }
            else if (typeof value === 'object' && value !== null) {
                var ret = {};
                for (var key in value) {
                    if (key.charCodeAt(0) !== charCodeUnderscore) {
                        var propValue = value[key];
                        if (typeof propValue === 'object' && propValue !== null && Array.isArray(propValue['items'])) {
                            propValue = propValue['items'];
                        }
                        ret[key] = this.purifyJson(propValue);
                    }
                }
                return ret;
            }
            else {
                return value;
            }
        };
        CommonUtility.collectObjectPathInfos = function (context, args, referencedObjectPaths, referencedObjectPathIds) {
            var hasOne = false;
            for (var i = 0; i < args.length; i++) {
                if (args[i] instanceof ClientObjectBase) {
                    var clientObject = args[i];
                    CommonUtility.validateContext(context, clientObject);
                    args[i] = clientObject._objectPath.objectPathInfo.Id;
                    referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
                    referencedObjectPaths.push(clientObject._objectPath);
                    hasOne = true;
                }
                else if (Array.isArray(args[i])) {
                    var childArrayObjectPathIds = new Array();
                    var childArrayHasOne = CommonUtility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
                    if (childArrayHasOne) {
                        referencedObjectPathIds.push(childArrayObjectPathIds);
                        hasOne = true;
                    }
                    else {
                        referencedObjectPathIds.push(0);
                    }
                }
                else if (CoreUtility.isPlainJsonObject(args[i])) {
                    referencedObjectPathIds.push(0);
                    CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths);
                }
                else {
                    referencedObjectPathIds.push(0);
                }
            }
            return hasOne;
        };
        CommonUtility.replaceClientObjectPropertiesWithObjectPathIds = function (value, referencedObjectPaths) {
            var _a, _b;
            for (var key in value) {
                var propValue = value[key];
                if (propValue instanceof ClientObjectBase) {
                    referencedObjectPaths.push(propValue._objectPath);
                    value[key] = (_a = {}, _a[CommonConstants.objectPathIdPrivate] = propValue._objectPath.objectPathInfo.Id, _a);
                }
                else if (Array.isArray(propValue)) {
                    for (var i = 0; i < propValue.length; i++) {
                        if (propValue[i] instanceof ClientObjectBase) {
                            var elem = propValue[i];
                            referencedObjectPaths.push(elem._objectPath);
                            propValue[i] = (_b = {}, _b[CommonConstants.objectPathIdPrivate] = elem._objectPath.objectPathInfo.Id, _b);
                        }
                        else if (CoreUtility.isPlainJsonObject(propValue[i])) {
                            CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths);
                        }
                    }
                }
                else if (CoreUtility.isPlainJsonObject(propValue)) {
                    CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
                }
                else {
                }
            }
        };
        CommonUtility.normalizeName = function (name) {
            return name.substr(0, 1).toLowerCase() + name.substr(1);
        };
        CommonUtility._doApiNotSupportedCheck = false;
        return CommonUtility;
    }(CoreUtility));
    OfficeExtension_1.CommonUtility = CommonUtility;
    var CommonResourceStrings = (function (_super) {
        __extends(CommonResourceStrings, _super);
        function CommonResourceStrings() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonResourceStrings.propertyDoesNotExist = 'PropertyDoesNotExist';
        CommonResourceStrings.attemptingToSetReadOnlyProperty = 'AttemptingToSetReadOnlyProperty';
        return CommonResourceStrings;
    }(CoreResourceStrings));
    OfficeExtension_1.CommonResourceStrings = CommonResourceStrings;
    var ClientRetrieveResult = (function (_super) {
        __extends(ClientRetrieveResult, _super);
        function ClientRetrieveResult(m_shouldPolyfill) {
            var _this = _super.call(this) || this;
            _this.m_shouldPolyfill = m_shouldPolyfill;
            return _this;
        }
        ClientRetrieveResult.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (this.m_shouldPolyfill) {
                this.m_value = CommonUtility.changePropertyNameToCamelLowerCase(this.m_value);
            }
            this.m_value = this.removeItemNodes(this.m_value);
        };
        ClientRetrieveResult.prototype.removeItemNodes = function (value) {
            if (typeof value === 'object' && value !== null && value[CommonConstants.itemsLowerCase]) {
                value = value[CommonConstants.itemsLowerCase];
            }
            return CommonUtility.purifyJson(value);
        };
        return ClientRetrieveResult;
    }(ClientResult));
    OfficeExtension_1.ClientRetrieveResult = ClientRetrieveResult;
    var TraceActionResultHandler = (function () {
        function TraceActionResultHandler(callback) {
            this.callback = callback;
        }
        TraceActionResultHandler.prototype._handleResult = function (value) {
            if (this.callback) {
                this.callback();
            }
        };
        return TraceActionResultHandler;
    }());
    var ClientResultCallback = (function (_super) {
        __extends(ClientResultCallback, _super);
        function ClientResultCallback(callback) {
            var _this = _super.call(this) || this;
            _this.callback = callback;
            return _this;
        }
        ClientResultCallback.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            this.callback();
        };
        return ClientResultCallback;
    }(ClientResult));
    OfficeExtension_1.ClientResultCallback = ClientResultCallback;
    var OperationalApiHelper = (function () {
        function OperationalApiHelper() {
        }
        OperationalApiHelper.invokeMethod = function (obj, methodName, operationType, args, flags, resultProcessType) {
            if (operationType === void 0) {
                operationType = 0;
            }
            if (args === void 0) {
                args = [];
            }
            if (flags === void 0) {
                flags = 0;
            }
            if (resultProcessType === void 0) {
                resultProcessType = 0;
            }
            return CoreUtility.createPromise(function (resolve, reject) {
                var result = new ClientResult();
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 3,
                    Name: methodName,
                    ObjectPathId: obj._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(obj._context, actionInfo.ArgumentInfo, args);
                var action = new Action(actionInfo, operationType, flags);
                action.referencedObjectPath = obj._objectPath;
                action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
                obj._context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.invokeMethodWithClientResultCallback = function (callback, obj, methodName) {
            var operationType = 0;
            var args = [];
            var flags = 0;
            return CoreUtility.createPromise(function (resolve, reject) {
                var result = new ClientResultCallback(callback);
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 3,
                    Name: methodName,
                    ObjectPathId: obj._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(obj._context, actionInfo.ArgumentInfo, args);
                var action = new Action(actionInfo, operationType, flags);
                action.referencedObjectPath = obj._objectPath;
                action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
                obj._context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.invokeRetrieve = function (obj, select) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.1');
            }
            var option;
            if (typeof select[0] === 'object' && select[0].hasOwnProperty('$all')) {
                if (!select[0]['$all']) {
                    throw OfficeExtension_1.Error._createInvalidArgError({});
                }
                option = select[0];
            }
            else {
                option = OperationalApiHelper._parseSelectOption(select);
            }
            return obj._retrieve(option, new ClientRetrieveResult(shouldPolyfill));
        };
        OperationalApiHelper._parseSelectOption = function (select) {
            if (!select || !select[0]) {
                throw OfficeExtension_1.Error._createInvalidArgError({});
            }
            var parsedSelect = select[0] && typeof select[0] !== 'string' ? select[0] : select;
            return Array.isArray(parsedSelect) ? parsedSelect : OperationalApiHelper.parseRecursiveSelect(parsedSelect);
        };
        OperationalApiHelper.parseRecursiveSelect = function (select) {
            var deconstruct = function (selectObj) {
                return Object.keys(selectObj).reduce(function (scalars, name) {
                    var value = selectObj[name];
                    if (typeof value === 'object') {
                        return scalars.concat(deconstruct(value).map(function (postfix) { return name + "/" + postfix; }));
                    }
                    if (value) {
                        return scalars.concat(name);
                    }
                    return scalars;
                }, []);
            };
            return deconstruct(select);
        };
        OperationalApiHelper.invokeRecursiveUpdate = function (obj, properties) {
            return CoreUtility.createPromise(function (resolve, reject) {
                obj._recursivelyUpdate(properties);
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 5,
                    Name: 'Trace',
                    ObjectPathId: 0
                };
                var action = new Action(actionInfo, 1, 4);
                obj._context._addServiceApiAction(action, null, resolve, reject);
            });
        };
        OperationalApiHelper.createRootServiceObject = function (type, context) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 1,
                Name: ''
            };
            var objectPath = new ObjectPath(objectPathInfo, null, false, false, 1, 4);
            return new type(context, objectPath);
        };
        OperationalApiHelper.createTopLevelServiceObject = function (type, context, typeName, isCollection, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 2,
                Name: typeName
            };
            var objectPath = new ObjectPath(objectPathInfo, null, isCollection, false, 1, flags | 4);
            return new type(context, objectPath);
        };
        OperationalApiHelper.createPropertyObject = function (type, parent, propertyName, isCollection, flags) {
            var objectPathInfo = {
                Id: parent._context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id
            };
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, false, 1, flags | 4);
            return new type(parent._context, objectPath);
        };
        OperationalApiHelper.createIndexerObject = function (type, parent, args) {
            var objectPathInfo = {
                Id: parent._context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
            return new type(parent._context, objectPath);
        };
        OperationalApiHelper.createMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var id = parent._context._nextId();
            var objectPathInfo = {
                Id: id,
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
            objectPath.argumentObjectPaths = argumentObjectPaths;
            objectPath.getByIdMethodName = getByIdMethodName;
            var o = new type(parent._context, objectPath);
            return o;
        };
        OperationalApiHelper.createAndInstantiateMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var objectPathInfo = {
                    Id: parent._context._nextId(),
                    ObjectPathType: 3,
                    Name: methodName,
                    ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var argumentObjectPaths = CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
                var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
                objectPath.argumentObjectPaths = argumentObjectPaths;
                objectPath.getByIdMethodName = getByIdMethodName;
                var result = new ClientResult();
                var actionInfo = {
                    Id: parent._context._nextId(),
                    ActionType: 1,
                    Name: '',
                    ObjectPathId: objectPath.objectPathInfo.Id,
                    QueryInfo: {}
                };
                var action = new Action(actionInfo, 1, 4);
                action.referencedObjectPath = objectPath;
                parent._context._addServiceApiAction(action, result, function () { return resolve(new type(parent._context, objectPath)); }, reject);
            });
        };
        OperationalApiHelper.createTraceAction = function (context, callback) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var actionInfo = {
                    Id: context._nextId(),
                    ActionType: 5,
                    Name: 'Trace',
                    ObjectPathId: 0
                };
                var action = new Action(actionInfo, 1, 4);
                var result = new TraceActionResultHandler(callback);
                context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.localDocumentContext = new ClientRequestContextBase();
        return OperationalApiHelper;
    }());
    OfficeExtension_1.OperationalApiHelper = OperationalApiHelper;
    var GenericEventRegistryOperational = (function () {
        function GenericEventRegistryOperational(eventId, targetId, eventArgumentTransform) {
            this.eventId = eventId;
            this.targetId = targetId;
            this.eventArgumentTransform = eventArgumentTransform;
            this.registeredCallbacks = [];
        }
        GenericEventRegistryOperational.prototype.add = function (callback) {
            if (this.hasZero()) {
                GenericEventRegistration.getGenericEventRegistration('').register(this.eventId, this.targetId, this.registerCallback);
            }
            this.registeredCallbacks.push(callback);
        };
        GenericEventRegistryOperational.prototype.remove = function (callback) {
            var index = this.registeredCallbacks.lastIndexOf(callback);
            if (index !== -1) {
                this.registeredCallbacks.splice(index, 1);
            }
        };
        GenericEventRegistryOperational.prototype.removeAll = function () {
            this.registeredCallbacks = [];
            GenericEventRegistration.getGenericEventRegistration('').unregister(this.eventId, this.targetId, this.registerCallback);
        };
        GenericEventRegistryOperational.prototype.hasZero = function () {
            return this.registeredCallbacks.length === 0;
        };
        Object.defineProperty(GenericEventRegistryOperational.prototype, "registerCallback", {
            get: function () {
                var i = this;
                if (!this.outsideCallback) {
                    this.outsideCallback = function (argument) {
                        i.call(argument);
                    };
                }
                return this.outsideCallback;
            },
            enumerable: true,
            configurable: true
        });
        GenericEventRegistryOperational.prototype.call = function (rawEventArguments) {
            var _this = this;
            this.eventArgumentTransform(rawEventArguments).then(function (eventArguments) {
                var promises = _this.registeredCallbacks.map(function (callback) { return GenericEventRegistryOperational.callCallback(callback, eventArguments); });
                CoreUtility.Promise.all(promises);
            });
        };
        GenericEventRegistryOperational.callCallback = function (callback, eventArguments) {
            return CoreUtility._createPromiseFromResult(null)
                .then(GenericEventRegistryOperational.wrapCallbackInFunction(callback, eventArguments))["catch"](function (e) {
                CoreUtility.log('Error when invoke handler: ' + JSON.stringify(e));
            });
        };
        GenericEventRegistryOperational.wrapCallbackInFunction = function (callback, args) {
            return function () { return callback(args); };
        };
        return GenericEventRegistryOperational;
    }());
    OfficeExtension_1.GenericEventRegistryOperational = GenericEventRegistryOperational;
    var GlobalEventRegistryOperational = (function () {
        function GlobalEventRegistryOperational() {
            this.eventToTargetToHandlerMap = {};
        }
        Object.defineProperty(GlobalEventRegistryOperational, "globalEventRegistry", {
            get: function () {
                if (!GlobalEventRegistryOperational.singleton) {
                    GlobalEventRegistryOperational.singleton = new GlobalEventRegistryOperational();
                }
                return GlobalEventRegistryOperational.singleton;
            },
            enumerable: true,
            configurable: true
        });
        GlobalEventRegistryOperational.getGlobalEventRegistry = function (eventId, targetId, eventArgumentTransform) {
            var global = GlobalEventRegistryOperational.globalEventRegistry;
            var mapGlobal = global.eventToTargetToHandlerMap;
            if (!mapGlobal.hasOwnProperty(eventId)) {
                mapGlobal[eventId] = {};
            }
            var mapEvent = mapGlobal[eventId];
            if (!mapEvent.hasOwnProperty(targetId)) {
                mapEvent[targetId] = new GenericEventRegistryOperational(eventId, targetId, eventArgumentTransform);
            }
            var target = mapEvent[targetId];
            return target;
        };
        GlobalEventRegistryOperational.singleton = undefined;
        return GlobalEventRegistryOperational;
    }());
    OfficeExtension_1.GlobalEventRegistryOperational = GlobalEventRegistryOperational;
    var GenericEventHandlerOperational = (function () {
        function GenericEventHandlerOperational(genericEventInfo) {
            this.genericEventInfo = genericEventInfo;
        }
        GenericEventHandlerOperational.prototype.add = function (callback) {
            var _this = this;
            var eventRegistered = undefined;
            var promise = CoreUtility.createPromise(function (resolve) {
                eventRegistered = resolve;
            });
            var addCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.add(callback);
                eventRegistered();
            };
            this.register();
            this.createTrace(addCallback);
            return promise;
        };
        GenericEventHandlerOperational.prototype.remove = function (callback) {
            var _this = this;
            var removeCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.remove(callback);
            };
            this.register();
            this.createTrace(removeCallback);
        };
        GenericEventHandlerOperational.prototype.removeAll = function () {
            var _this = this;
            var removeAllCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.removeAll();
            };
            this.unregister();
            this.createTrace(removeAllCallback);
        };
        GenericEventHandlerOperational.prototype.createTrace = function (callback) {
            OperationalApiHelper.createTraceAction(this.genericEventInfo.object._context, callback);
        };
        GenericEventHandlerOperational.prototype.register = function () {
            var operationType = 0;
            var args = [];
            var flags = 0;
            OperationalApiHelper.invokeMethod(this.genericEventInfo.object, this.genericEventInfo.register, operationType, args, flags);
            if (!GenericEventRegistration.getGenericEventRegistration('').isReady) {
                GenericEventRegistration.getGenericEventRegistration('').ready();
            }
        };
        GenericEventHandlerOperational.prototype.unregister = function () {
            OperationalApiHelper.invokeMethod(this.genericEventInfo.object, this.genericEventInfo.unregister);
        };
        return GenericEventHandlerOperational;
    }());
    OfficeExtension_1.GenericEventHandlerOperational = GenericEventHandlerOperational;
    var EventHelper = (function () {
        function EventHelper() {
        }
        EventHelper.invokeOn = function (eventHandler, callback, options) {
            var promiseResolve = undefined;
            var promise = CoreUtility.createPromise(function (resolve, reject) {
                promiseResolve = resolve;
            });
            eventHandler.add(callback).then(function () {
                promiseResolve({});
            });
            return promise;
        };
        EventHelper.invokeOff = function (genericEventHandlersOpObj, eventHandler, eventName, callback) {
            if (!eventName && !callback) {
                var allGenericEventHandlersOp = Object.keys(genericEventHandlersOpObj).map(function (eventName) { return genericEventHandlersOpObj[eventName]; });
                return EventHelper.invokeAllOff(allGenericEventHandlersOp);
            }
            if (!eventName) {
                return CoreUtility._createPromiseFromException(eventName + " must be supplied if handler is supplied.");
            }
            if (callback) {
                eventHandler.remove(callback);
            }
            else {
                eventHandler.removeAll();
            }
            return CoreUtility.createPromise(function (resolve, reject) { return resolve(); });
        };
        EventHelper.invokeAllOff = function (allGenericEventHandlersOperational) {
            allGenericEventHandlersOperational.forEach(function (genericEventHandlerOperational) {
                genericEventHandlerOperational.removeAll();
            });
            return CoreUtility.createPromise(function (resolve, reject) { return resolve(); });
        };
        return EventHelper;
    }());
    OfficeExtension_1.EventHelper = EventHelper;
    var ErrorCodes = (function (_super) {
        __extends(ErrorCodes, _super);
        function ErrorCodes() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ErrorCodes.propertyNotLoaded = 'PropertyNotLoaded';
        ErrorCodes.runMustReturnPromise = 'RunMustReturnPromise';
        ErrorCodes.cannotRegisterEvent = 'CannotRegisterEvent';
        ErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
        ErrorCodes.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
        return ErrorCodes;
    }(CoreErrorCodes));
    OfficeExtension_1.ErrorCodes = ErrorCodes;
    var TraceMarkerActionResultHandler = (function () {
        function TraceMarkerActionResultHandler(callback) {
            this.m_callback = callback;
        }
        TraceMarkerActionResultHandler.prototype._handleResult = function (value) {
            if (this.m_callback) {
                this.m_callback();
            }
        };
        return TraceMarkerActionResultHandler;
    }());
    var ActionFactory = (function (_super) {
        __extends(ActionFactory, _super);
        function ActionFactory() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args, flags) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 3,
                Name: methodName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var fixedFlags = Utility._fixupApiFlags(flags);
            var action = new Action(actionInfo, operationType, fixedFlags);
            action.referencedObjectPath = parent._objectPath;
            action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
            parent._addAction(action);
            if (OfficeExtension_1._internalConfig.enablePreviewExecution && (fixedFlags & 16) !== 0) {
                var previewExecutionAction = {
                    Id: context._nextId(),
                    ActionType: 3,
                    Name: methodName,
                    Arguments: args,
                    ObjectId: '',
                    ObjectType: ''
                };
                parent._addPreviewExecutionAction(previewExecutionAction);
            }
            return action;
        };
        ActionFactory.createRecursiveQueryAction = function (context, parent, query) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 6,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                RecursiveQueryInfo: query
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            parent._addAction(action);
            return action;
        };
        ActionFactory.createEnsureUnchangedAction = function (context, parent, objectState) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 8,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            parent._addAction(action);
            return action;
        };
        ActionFactory.createInstantiateAction = function (context, obj) {
            Utility.validateObjectPath(obj);
            context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
            context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 1,
                Name: '',
                ObjectPathId: obj._objectPath.objectPathInfo.Id
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = obj._objectPath;
            obj._addAction(action, new InstantiateActionResultHandler(obj), true);
            return action;
        };
        ActionFactory.createTraceAction = function (context, message, addTraceMessage) {
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 5,
                Name: 'Trace',
                ObjectPathId: 0
            };
            var ret = new Action(actionInfo, 1, 4);
            context._pendingRequest.addAction(ret);
            if (addTraceMessage) {
                context._pendingRequest.addTrace(actionInfo.Id, message);
            }
            return ret;
        };
        ActionFactory.createTraceMarkerForCallback = function (context, callback) {
            var action = ActionFactory.createTraceAction(context, null, false);
            context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
        };
        return ActionFactory;
    }(CommonActionFactory));
    OfficeExtension_1.ActionFactory = ActionFactory;
    var ClientObject = (function (_super) {
        __extends(ClientObject, _super);
        function ClientObject(context, objectPath) {
            var _this = _super.call(this, context, objectPath) || this;
            Utility.checkArgumentNull(context, 'context');
            _this.m_context = context;
            if (_this._objectPath) {
                if (!context._processingResult && context._pendingRequest) {
                    ActionFactory.createInstantiateAction(context, _this);
                    if (context._autoCleanup && _this._KeepReference) {
                        context.trackedObjects._autoAdd(_this);
                    }
                }
                if (OfficeExtension_1._internalConfig.appendTypeNameToObjectPathInfo && _this._objectPath.objectPathInfo && _this._className) {
                    _this._objectPath.objectPathInfo.T = _this._className;
                }
            }
            return _this;
        }
        Object.defineProperty(ClientObject.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNull", {
            get: function () {
                if (typeof (this.m_isNull) === 'undefined' && TestUtility.isMock()) {
                    return false;
                }
                Utility.throwIfNotLoaded('isNull', this._isNull, null, this._isNull);
                return this._isNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNullObject", {
            get: function () {
                if (typeof (this.m_isNull) === 'undefined' && TestUtility.isMock()) {
                    return false;
                }
                Utility.throwIfNotLoaded('isNullObject', this._isNull, null, this._isNull);
                return this._isNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "_isNull", {
            get: function () {
                return this.m_isNull;
            },
            set: function (value) {
                this.m_isNull = value;
                if (value && this._objectPath) {
                    this._objectPath._updateAsNullObject();
                }
            },
            enumerable: true,
            configurable: true
        });
        ClientObject.prototype._addAction = function (action, resultHandler, isInstantiationEnsured) {
            if (resultHandler === void 0) {
                resultHandler = null;
            }
            if (!isInstantiationEnsured) {
                this.context._pendingRequest.ensureInstantiateObjectPath(this._objectPath);
                this.context._pendingRequest.ensureInstantiateObjectPaths(action.referencedArgumentObjectPaths);
            }
            this.context._pendingRequest.addAction(action);
            this.context._pendingRequest.addReferencedObjectPath(this._objectPath);
            this.context._pendingRequest.addReferencedObjectPaths(action.referencedArgumentObjectPaths);
            this.context._pendingRequest.addActionResultHandler(action, resultHandler);
            return CoreUtility._createPromiseFromResult(null);
        };
        ClientObject.prototype._addPreviewExecutionAction = function (action) {
            if (!Utility.isUndefined(this._className)) {
                action.ObjectType = this._className;
                var objectId = Utility._getPropertyValueWithoutCheckLoaded(this, Constants.idLowerCase);
                if (Utility.isUndefined(objectId)) {
                    objectId = Utility._getPropertyValueWithoutCheckLoaded(this, Constants.idPrivate);
                }
                if (Utility.isUndefined(objectId)) {
                    objectId = Utility._getPropertyValueWithoutCheckLoaded(this, Constants.previewExecutionObjectId);
                }
                action.ObjectId = objectId;
                this.context._pendingRequest.addPreviewExecutionAction(action);
            }
        };
        ClientObject.prototype._handleResult = function (value) {
            this._isNull = Utility.isNullOrUndefined(value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._handleIdResult = function (value) {
            this._isNull = Utility.isNullOrUndefined(value);
            Utility.fixObjectPathIfNecessary(this, value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._handleRetrieveResult = function (value, result) {
            this._handleIdResult(value);
        };
        ClientObject.prototype._recursivelySet = function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
            var isClientObject = input instanceof ClientObject;
            var originalInput = input;
            if (isClientObject) {
                if (Object.getPrototypeOf(this) === Object.getPrototypeOf(input)) {
                    input = JSON.parse(JSON.stringify(input));
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({
                        argumentName: 'properties',
                        errorLocation: this._className + '.set'
                    });
                }
            }
            try {
                var prop;
                for (var i = 0; i < scalarWriteablePropertyNames.length; i++) {
                    prop = scalarWriteablePropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined') {
                            this[prop] = input[prop];
                        }
                    }
                }
                for (var i = 0; i < objectPropertyNames.length; i++) {
                    prop = objectPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined') {
                            var dataToPassToSet = isClientObject ? originalInput[prop] : input[prop];
                            this[prop].set(dataToPassToSet, options);
                        }
                    }
                }
                var throwOnReadOnly = !isClientObject;
                if (options && !Utility.isNullOrUndefined(throwOnReadOnly)) {
                    throwOnReadOnly = options.throwOnReadOnly;
                }
                for (var i = 0; i < notAllowedToBeSetPropertyNames.length; i++) {
                    prop = notAllowedToBeSetPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined' && throwOnReadOnly) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                    }
                }
                for (prop in input) {
                    if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
                        var propertyDescriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
                        if (!propertyDescriptor) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                        if (throwOnReadOnly && !propertyDescriptor.set) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                    }
                }
            }
            catch (innerError) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    httpStatusCode: 400,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
                    debugInfo: {
                        errorLocation: this._className + '.set'
                    },
                    innerError: innerError
                });
            }
        };
        return ClientObject;
    }(ClientObjectBase));
    OfficeExtension_1.ClientObject = ClientObject;
    var HostBridgeRequestExecutor = (function () {
        function HostBridgeRequestExecutor(session) {
            this.m_session = session;
        }
        HostBridgeRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var httpRequestInfo = {
                url: CoreConstants.processQuery,
                method: 'POST',
                headers: requestMessage.Headers,
                body: requestMessage.Body
            };
            var controlId = '';
            if (requestMessage.Headers) {
                controlId = requestMessage.Headers[Constants.officeControlId];
            }
            var message = {
                id: HostBridge.nextId(),
                type: 1,
                flags: requestFlags,
                controlId: controlId,
                message: httpRequestInfo
            };
            CoreUtility.log(JSON.stringify(message));
            return this.m_session.sendMessageToHost(message).then(function (nativeBridgeResponse) {
                CoreUtility.log('Received response: ' + JSON.stringify(nativeBridgeResponse));
                var responseInfo = nativeBridgeResponse.message;
                var response;
                if (responseInfo.statusCode === 200) {
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: null,
                        ErrorMessage: null,
                        Headers: responseInfo.headers,
                        Body: CoreUtility._parseResponseBody(responseInfo)
                    };
                }
                else {
                    CoreUtility.log('Error Response:' + responseInfo.body);
                    var error = CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        HttpStatusCode: responseInfo.statusCode,
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            });
        };
        return HostBridgeRequestExecutor;
    }());
    var HostBridgeSession = (function (_super) {
        __extends(HostBridgeSession, _super);
        function HostBridgeSession(m_bridge) {
            var _this = _super.call(this) || this;
            _this.m_bridge = m_bridge;
            _this.m_bridge.addHostMessageHandler(function (message) {
                if (message.type === 3) {
                    var controlId = message.controlId;
                    if (CoreUtility.isNullOrEmptyString(controlId)) {
                        GenericEventRegistration.getGenericEventRegistration(controlId)._handleRichApiMessage(message.message);
                    }
                    else {
                        var eventRegistration = GenericEventRegistration.peekGenericEventRegistrationOrNull(controlId);
                        if (eventRegistration) {
                            eventRegistration._handleRichApiMessage(message.message);
                        }
                        eventRegistration = GenericEventRegistration.peekGenericEventRegistrationOrNull('');
                        if (eventRegistration) {
                            eventRegistration._handleRichApiMessage(message.message);
                        }
                    }
                }
            });
            return _this;
        }
        HostBridgeSession.getInstanceIfHostBridgeInited = function () {
            if (HostBridge.instance) {
                if (CoreUtility.isNullOrUndefined(HostBridgeSession.s_instance) ||
                    HostBridgeSession.s_instance.m_bridge !== HostBridge.instance) {
                    HostBridgeSession.s_instance = new HostBridgeSession(HostBridge.instance);
                }
                return HostBridgeSession.s_instance;
            }
            return null;
        };
        HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        HostBridgeSession.prototype._createRequestExecutorOrNull = function () {
            CoreUtility.log('NativeBridgeSession::CreateRequestExecutor');
            return new HostBridgeRequestExecutor(this);
        };
        HostBridgeSession.prototype.getEventRegistration = function (controlId) {
            return GenericEventRegistration.getGenericEventRegistration(controlId);
        };
        HostBridgeSession.prototype.sendMessageToHost = function (message) {
            return this.m_bridge.sendMessageToHostAndExpectResponse(message);
        };
        return HostBridgeSession;
    }(SessionBase));
    OfficeExtension_1.HostBridgeSession = HostBridgeSession;
    var ClientRequestContext = (function (_super) {
        __extends(ClientRequestContext, _super);
        function ClientRequestContext(url) {
            var _this = _super.call(this) || this;
            _this.m_customRequestHeaders = {};
            _this.m_batchMode = 0;
            _this._onRunFinishedNotifiers = [];
            if (SessionBase._overrideSession) {
                _this.m_requestUrlAndHeaderInfoResolver = SessionBase._overrideSession;
            }
            else {
                if (Utility.isNullOrUndefined(url) || (typeof url === 'string' && url.length === 0)) {
                    url = ClientRequestContext.defaultRequestUrlAndHeaders;
                    if (!url) {
                        url = { url: CoreConstants.localDocument, headers: {} };
                    }
                }
                if (typeof url === 'string') {
                    _this.m_requestUrlAndHeaderInfo = { url: url, headers: {} };
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
                    _this.m_requestUrlAndHeaderInfoResolver = url;
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
                    var requestInfo = url;
                    _this.m_requestUrlAndHeaderInfo = { url: requestInfo.url, headers: {} };
                    CoreUtility._copyHeaders(requestInfo.headers, _this.m_requestUrlAndHeaderInfo.headers);
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'url' });
                }
            }
            if (!_this.m_requestUrlAndHeaderInfoResolver &&
                _this.m_requestUrlAndHeaderInfo &&
                CoreUtility._isLocalDocumentUrl(_this.m_requestUrlAndHeaderInfo.url) &&
                HostBridgeSession.getInstanceIfHostBridgeInited()) {
                _this.m_requestUrlAndHeaderInfo = null;
                _this.m_requestUrlAndHeaderInfoResolver = HostBridgeSession.getInstanceIfHostBridgeInited();
            }
            if (_this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
                _this.m_session = _this.m_requestUrlAndHeaderInfoResolver;
            }
            _this._processingResult = false;
            _this._customData = Constants.iterativeExecutor;
            _this.sync = _this.sync.bind(_this);
            return _this;
        }
        Object.defineProperty(ClientRequestContext.prototype, "session", {
            get: function () {
                return this.m_session;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
            get: function () {
                if (this.m_session) {
                    return this.m_session.getEventRegistration(this._controlId);
                }
                return _Internal.officeJsEventRegistration;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "_url", {
            get: function () {
                if (this.m_requestUrlAndHeaderInfo) {
                    return this.m_requestUrlAndHeaderInfo.url;
                }
                return null;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
            get: function () {
                if (this.m_pendingRequest == null) {
                    this.m_pendingRequest = new ClientRequest(this);
                }
                return this.m_pendingRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "_controlId", {
            get: function () {
                var ret = this.m_customRequestHeaders[Constants.officeControlId];
                if (CoreUtility.isNullOrUndefined(ret)) {
                    ret = '';
                }
                return ret;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
            get: function () {
                var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension_1._internalConfig.showDisposeInfoInDebugInfo);
                var statements = prettyPrinter.process();
                return { pendingStatements: statements };
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
            get: function () {
                if (!this.m_trackedObjects) {
                    this.m_trackedObjects = new TrackedObjects(this);
                }
                return this.m_trackedObjects;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
            get: function () {
                return this.m_customRequestHeaders;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
            get: function () {
                return this.m_batchMode;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode = function () {
            if (this.m_batchMode === 1 && !this.m_explicitBatchInProgress) {
                throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.notInsideBatch), null);
            }
        };
        ClientRequestContext.prototype.load = function (clientObj, option) {
            Utility.validateContext(this, clientObj);
            var queryOption = ClientRequestContext._parseQueryOption(option);
            CommonActionFactory.createQueryAction(this, clientObj, queryOption, clientObj);
        };
        ClientRequestContext.prototype.loadRecursive = function (clientObj, options, maxDepth) {
            if (!Utility.isPlainJsonObject(options)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'options' });
            }
            var quries = {};
            for (var key in options) {
                quries[key] = ClientRequestContext._parseQueryOption(options[key]);
            }
            var action = ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
            this._pendingRequest.addActionResultHandler(action, clientObj);
        };
        ClientRequestContext.prototype.trace = function (message) {
            ActionFactory.createTraceAction(this, message, true);
        };
        ClientRequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) { };
        ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo = function () {
            var _this = this;
            return Utility._createPromiseFromResult(null).then(function () {
                if (!_this.m_requestUrlAndHeaderInfo) {
                    return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then(function (value) {
                        _this.m_requestUrlAndHeaderInfo = value;
                        if (!_this.m_requestUrlAndHeaderInfo) {
                            _this.m_requestUrlAndHeaderInfo = { url: CoreConstants.localDocument, headers: {} };
                        }
                        if (Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
                            _this.m_requestUrlAndHeaderInfo.url = CoreConstants.localDocument;
                        }
                        if (!_this.m_requestUrlAndHeaderInfo.headers) {
                            _this.m_requestUrlAndHeaderInfo.headers = {};
                        }
                        if (typeof _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull === 'function') {
                            var executor = _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
                            if (executor) {
                                _this._requestExecutor = executor;
                            }
                        }
                    });
                }
            });
        };
        ClientRequestContext.prototype.syncPrivateMain = function () {
            var _this = this;
            return this.ensureRequestUrlAndHeaderInfo().then(function () {
                var req = _this._pendingRequest;
                _this.m_pendingRequest = null;
                return _this.processPreSyncPromises(req).then(function () { return _this.syncPrivate(req); });
            });
        };
        ClientRequestContext.prototype.syncPrivate = function (req) {
            var _this = this;
            if (TestUtility.isMock()) {
                return CoreUtility._createPromiseFromResult(null);
            }
            if (!req.hasActions) {
                return this.processPendingEventHandlers(req);
            }
            var _a = req.buildRequestMessageBodyAndRequestFlags(), msgBody = _a.body, requestFlags = _a.flags;
            if (this._requestFlagModifier) {
                requestFlags |= this._requestFlagModifier;
            }
            if (!this._requestExecutor) {
                if (CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                    this._requestExecutor = new OfficeJsRequestExecutor(this);
                }
                else {
                    this._requestExecutor = new HttpRequestExecutor();
                }
            }
            var requestExecutor = this._requestExecutor;
            var headers = {};
            CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
            CoreUtility._copyHeaders(this.m_customRequestHeaders, headers);
            delete this.m_customRequestHeaders[Constants.officeScriptEventId];
            var testNameWithSequenceId = TestUtility._getCurrentTestNameWithSequenceId();
            if (testNameWithSequenceId) {
                headers[CoreConstants.testRequestNameHeader] = testNameWithSequenceId;
            }
            var requestExecutorRequestMessage = {
                Url: this.m_requestUrlAndHeaderInfo.url,
                Headers: headers,
                Body: msgBody
            };
            req.invalidatePendingInvalidObjectPaths();
            var errorFromResponse = null;
            var errorFromProcessEventHandlers = null;
            this._lastSyncStart = typeof performance === 'undefined' ? Date.now() : performance.now();
            this._lastRequestFlags = requestFlags;
            return requestExecutor
                .executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
                .then(function (response) {
                _this._lastSyncEnd = typeof performance === 'undefined' ? Date.now() : performance.now();
                if (OfficeExtension_1.config.executePerfLogFunc) {
                    OfficeExtension_1.config.executePerfLogFunc({ syncStart: _this._lastSyncStart,
                        syncEnd: _this._lastSyncEnd
                    });
                }
                errorFromResponse = _this.processRequestExecutorResponseMessage(req, response);
                return _this.processPendingEventHandlers(req)["catch"](function (ex) {
                    CoreUtility.log('Error in processPendingEventHandlers');
                    CoreUtility.log(JSON.stringify(ex));
                    errorFromProcessEventHandlers = ex;
                });
            })
                .then(function () {
                if (errorFromResponse) {
                    CoreUtility.log('Throw error from response: ' + JSON.stringify(errorFromResponse));
                    throw errorFromResponse;
                }
                if (errorFromProcessEventHandlers) {
                    CoreUtility.log('Throw error from ProcessEventHandler: ' + JSON.stringify(errorFromProcessEventHandlers));
                    var transformedError = null;
                    if (errorFromProcessEventHandlers instanceof _Internal.RuntimeError) {
                        transformedError = errorFromProcessEventHandlers;
                        transformedError.traceMessages = req._responseTraceMessages;
                    }
                    else {
                        var message = null;
                        if (typeof errorFromProcessEventHandlers === 'string') {
                            message = errorFromProcessEventHandlers;
                        }
                        else {
                            message = errorFromProcessEventHandlers.message;
                        }
                        if (Utility.isNullOrEmptyString(message)) {
                            message = CoreUtility._getResourceString(ResourceStrings.cannotRegisterEvent);
                        }
                        transformedError = new _Internal.RuntimeError({
                            code: ErrorCodes.cannotRegisterEvent,
                            httpStatusCode: 400,
                            message: message,
                            traceMessages: req._responseTraceMessages
                        });
                    }
                    throw transformedError;
                }
            });
        };
        ClientRequestContext.prototype.processRequestExecutorResponseMessage = function (req, response) {
            if (response.Body && response.Body.TraceIds) {
                req._setResponseTraceIds(response.Body.TraceIds);
            }
            var traceMessages = req._responseTraceMessages;
            var errorStatementInfo = null;
            if (response.Body) {
                if (response.Body.Error && response.Body.Error.ActionIndex >= 0) {
                    var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, true);
                    var debugInfoStatementInfo = prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
                    errorStatementInfo = {
                        statement: debugInfoStatementInfo.statement,
                        surroundingStatements: debugInfoStatementInfo.surroundingStatements,
                        fullStatements: ['Please enable config.extendedErrorLogging to see full statements.']
                    };
                    if (OfficeExtension_1.config.extendedErrorLogging) {
                        prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, false);
                        errorStatementInfo.fullStatements = prettyPrinter.process();
                    }
                }
                var actionResults = null;
                if (response.Body.Results) {
                    actionResults = response.Body.Results;
                }
                else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                    actionResults = response.Body.ProcessedResults.Results;
                }
                if (actionResults) {
                    this._processingResult = true;
                    try {
                        req.processResponse(actionResults);
                    }
                    finally {
                        this._processingResult = false;
                    }
                }
            }
            if (!Utility.isNullOrEmptyString(response.ErrorCode)) {
                return new _Internal.RuntimeError({
                    code: response.ErrorCode,
                    httpStatusCode: response.HttpStatusCode,
                    message: response.ErrorMessage,
                    traceMessages: traceMessages,
                    data: { responseBody: response.RawErrorResponseBody }
                });
            }
            else if (response.Body && response.Body.Error) {
                var debugInfo = {
                    errorLocation: response.Body.Error.Location
                };
                if (errorStatementInfo) {
                    debugInfo.statement = errorStatementInfo.statement;
                    debugInfo.surroundingStatements = errorStatementInfo.surroundingStatements;
                    debugInfo.fullStatements = errorStatementInfo.fullStatements;
                }
                return new _Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    httpStatusCode: response.Body.Error.HttpStatusCode,
                    message: response.Body.Error.Message,
                    traceMessages: traceMessages,
                    debugInfo: debugInfo
                });
            }
            return null;
        };
        ClientRequestContext.prototype.processPendingEventHandlers = function (req) {
            var ret = Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._pendingProcessEventHandlers.length; i++) {
                var eventHandlers = req._pendingProcessEventHandlers[i];
                ret = ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneEventHandlersFunc = function (eventHandlers, req) {
            return function () { return eventHandlers._processRegistration(req); };
        };
        ClientRequestContext.prototype.processPreSyncPromises = function (req) {
            var ret = Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._preSyncPromises.length; i++) {
                var p = req._preSyncPromises[i];
                ret = ret.then(this.createProcessOneProSyncFunc(p));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneProSyncFunc = function (p) {
            return function () { return p; };
        };
        ClientRequestContext.prototype.sync = function (passThroughValue) {
            if (TestUtility.isMock()) {
                return CoreUtility._createPromiseFromResult(passThroughValue);
            }
            return this.syncPrivateMain().then(function () { return passThroughValue; });
        };
        ClientRequestContext.prototype.batch = function (batchBody) {
            var _this = this;
            if (this.m_batchMode !== 1) {
                return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, null, null));
            }
            if (this.m_explicitBatchInProgress) {
                return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.pendingBatchInProgress), null));
            }
            if (Utility.isNullOrUndefined(batchBody)) {
                return Utility._createPromiseFromResult(null);
            }
            this.m_explicitBatchInProgress = true;
            var previousRequest = this.m_pendingRequest;
            this.m_pendingRequest = new ClientRequest(this);
            var batchBodyResult;
            try {
                batchBodyResult = batchBody(this._rootObject, this);
            }
            catch (ex) {
                this.m_explicitBatchInProgress = false;
                this.m_pendingRequest = previousRequest;
                return CoreUtility._createPromiseFromException(ex);
            }
            var request;
            var batchBodyResultPromise;
            if (typeof batchBodyResult === 'object' && batchBodyResult && typeof batchBodyResult.then === 'function') {
                batchBodyResultPromise = Utility._createPromiseFromResult(null)
                    .then(function () {
                    return batchBodyResult;
                })
                    .then(function (result) {
                    _this.m_explicitBatchInProgress = false;
                    request = _this.m_pendingRequest;
                    _this.m_pendingRequest = previousRequest;
                    return result;
                })["catch"](function (ex) {
                    _this.m_explicitBatchInProgress = false;
                    request = _this.m_pendingRequest;
                    _this.m_pendingRequest = previousRequest;
                    return CoreUtility._createPromiseFromException(ex);
                });
            }
            else {
                this.m_explicitBatchInProgress = false;
                request = this.m_pendingRequest;
                this.m_pendingRequest = previousRequest;
                batchBodyResultPromise = Utility._createPromiseFromResult(batchBodyResult);
            }
            return batchBodyResultPromise.then(function (result) {
                return _this.ensureRequestUrlAndHeaderInfo()
                    .then(function () {
                    return _this.syncPrivate(request);
                })
                    .then(function () {
                    return result;
                });
            });
        };
        ClientRequestContext._run = function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            return ClientRequestContext._runCommon('run', null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, null, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.isValidRequestInfo = function (value) {
            return (typeof value === 'string' ||
                ClientRequestContext.isRequestUrlAndHeaderInfo(value) ||
                ClientRequestContext.isRequestUrlAndHeaderInfoResolver(value));
        };
        ClientRequestContext.isRequestUrlAndHeaderInfo = function (value) {
            return (typeof value === 'object' &&
                value !== null &&
                Object.getPrototypeOf(value) === Object.getPrototypeOf({}) &&
                !Utility.isNullOrUndefined(value.url));
        };
        ClientRequestContext.isRequestUrlAndHeaderInfoResolver = function (value) {
            return typeof value === 'object' && value !== null && typeof value._resolveRequestUrlAndHeaderInfo === 'function';
        };
        ClientRequestContext._runBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runExplicitBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runBatchCommon = function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            var ctxRetriever;
            var batch;
            var requestInfo = null;
            var previousObjects = null;
            var argOffset = 0;
            var options = null;
            if (receivedRunArgs.length > 0) {
                if (ClientRequestContext.isValidRequestInfo(receivedRunArgs[0])) {
                    requestInfo = receivedRunArgs[0];
                    argOffset = 1;
                }
                else if (Utility.isPlainJsonObject(receivedRunArgs[0])) {
                    options = receivedRunArgs[0];
                    requestInfo = options.session;
                    if (requestInfo != null && !ClientRequestContext.isValidRequestInfo(requestInfo)) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    previousObjects = options.previousObjects;
                    argOffset = 1;
                }
            }
            if (receivedRunArgs.length == argOffset + 1) {
                batch = receivedRunArgs[argOffset + 0];
            }
            else if (options == null && receivedRunArgs.length == argOffset + 2) {
                previousObjects = receivedRunArgs[argOffset + 0];
                batch = receivedRunArgs[argOffset + 1];
            }
            else {
                return ClientRequestContext.createErrorPromise(functionName);
            }
            if (previousObjects != null) {
                if (previousObjects instanceof ClientObject) {
                    ctxRetriever = function () { return previousObjects.context; };
                }
                else if (previousObjects instanceof ClientRequestContext) {
                    ctxRetriever = function () { return previousObjects; };
                }
                else if (Array.isArray(previousObjects)) {
                    var array = previousObjects;
                    if (array.length == 0) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    for (var i = 0; i < array.length; i++) {
                        if (!(array[i] instanceof ClientObject)) {
                            return ClientRequestContext.createErrorPromise(functionName);
                        }
                        if (array[i].context != array[0].context) {
                            return ClientRequestContext.createErrorPromise(functionName, ResourceStrings.invalidRequestContext);
                        }
                    }
                    ctxRetriever = function () { return array[0].context; };
                }
                else {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
            }
            else {
                ctxRetriever = ctxInitializer;
            }
            var onBeforeRunWithOptions = null;
            if (onBeforeRun) {
                onBeforeRunWithOptions = function (context) { return onBeforeRun(options || {}, context); };
            }
            return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onBeforeRunWithOptions, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.createErrorPromise = function (functionName, code) {
            if (code === void 0) {
                code = CoreResourceStrings.invalidArgument;
            }
            return CoreUtility._createPromiseFromException(Utility.createRuntimeError(code, CoreUtility._getResourceString(code), functionName));
        };
        ClientRequestContext._runCommon = function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            if (SessionBase._overrideSession) {
                requestInfo = SessionBase._overrideSession;
            }
            var starterPromise = CoreUtility.createPromise(function (resolve, reject) {
                resolve();
            });
            var ctx;
            var succeeded = false;
            var resultOrError;
            var previousBatchMode;
            return starterPromise
                .then(function () {
                ctx = ctxRetriever(requestInfo);
                if (ctx._autoCleanup) {
                    return new OfficeExtension_1.Promise(function (resolve, reject) {
                        ctx._onRunFinishedNotifiers.push(function () {
                            ctx._autoCleanup = true;
                            resolve();
                        });
                    });
                }
                else {
                    ctx._autoCleanup = true;
                }
            })
                .then(function () {
                if (typeof runBody !== 'function') {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                previousBatchMode = ctx.m_batchMode;
                ctx.m_batchMode = batchMode;
                if (onBeforeRun) {
                    onBeforeRun(ctx);
                }
                var runBodyResult;
                if (batchMode == 1) {
                    runBodyResult = runBody(ctx.batch.bind(ctx));
                }
                else {
                    runBodyResult = runBody(ctx);
                }
                if (Utility.isNullOrUndefined(runBodyResult) || typeof runBodyResult.then !== 'function') {
                    Utility.throwError(ResourceStrings.runMustReturnPromise);
                }
                return runBodyResult;
            })
                .then(function (runBodyResult) {
                if (batchMode === 1) {
                    return runBodyResult;
                }
                else {
                    return ctx.sync(runBodyResult);
                }
            })
                .then(function (result) {
                succeeded = true;
                resultOrError = result;
            })["catch"](function (error) {
                resultOrError = error;
            })
                .then(function () {
                var itemsToRemove = ctx.trackedObjects._retrieveAndClearAutoCleanupList();
                ctx._autoCleanup = false;
                ctx.m_batchMode = previousBatchMode;
                for (var key in itemsToRemove) {
                    itemsToRemove[key]._objectPath.isValid = false;
                }
                var cleanupCounter = 0;
                if (Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
                    return attemptCleanup();
                }
                else {
                    attemptCleanup();
                }
                function attemptCleanup() {
                    cleanupCounter++;
                    var savedPendingRequest = ctx.m_pendingRequest;
                    var savedBatchMode = ctx.m_batchMode;
                    var request = new ClientRequest(ctx);
                    ctx.m_pendingRequest = request;
                    ctx.m_batchMode = 0;
                    try {
                        for (var key in itemsToRemove) {
                            ctx.trackedObjects.remove(itemsToRemove[key]);
                        }
                    }
                    finally {
                        ctx.m_batchMode = savedBatchMode;
                        ctx.m_pendingRequest = savedPendingRequest;
                    }
                    return ctx
                        .syncPrivate(request)
                        .then(function () {
                        if (onCleanupSuccess) {
                            onCleanupSuccess(cleanupCounter);
                        }
                    })["catch"](function () {
                        if (onCleanupFailure) {
                            onCleanupFailure(cleanupCounter);
                        }
                        if (cleanupCounter < numCleanupAttempts) {
                            setTimeout(function () {
                                attemptCleanup();
                            }, retryDelay);
                        }
                    });
                }
            })
                .then(function () {
                if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
                    var func = ctx._onRunFinishedNotifiers.shift();
                    func();
                }
                if (succeeded) {
                    return resultOrError;
                }
                else {
                    throw resultOrError;
                }
            });
        };
        return ClientRequestContext;
    }(ClientRequestContextBase));
    OfficeExtension_1.ClientRequestContext = ClientRequestContext;
    var RetrieveResultImpl = (function () {
        function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
            this.m_proxy = m_proxy;
            this.m_shouldPolyfill = m_shouldPolyfill;
            var scalarPropertyNames = m_proxy[Constants.scalarPropertyNames];
            var navigationPropertyNames = m_proxy[Constants.navigationPropertyNames];
            var typeName = m_proxy[Constants.className];
            var isCollection = m_proxy[Constants.isCollection];
            if (scalarPropertyNames) {
                for (var i = 0; i < scalarPropertyNames.length; i++) {
                    Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
                }
            }
            if (navigationPropertyNames) {
                for (var i = 0; i < navigationPropertyNames.length; i++) {
                    Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
                }
            }
            if (isCollection) {
                Utility.definePropertyThrowUnloadedException(this, typeName, Constants.itemsLowerCase);
            }
        }
        Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
            get: function () {
                return this.m_proxy;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
            get: function () {
                if (!this.m_isLoaded) {
                    throw new _Internal.RuntimeError({
                        code: ErrorCodes.valueNotLoaded,
                        httpStatusCode: 400,
                        message: CoreUtility._getResourceString(ResourceStrings.valueNotLoaded),
                        debugInfo: {
                            errorLocation: 'retrieveResult.$isNullObject'
                        }
                    });
                }
                return this.m_isNullObject;
            },
            enumerable: true,
            configurable: true
        });
        RetrieveResultImpl.prototype.toJSON = function () {
            if (!this.m_isLoaded) {
                return undefined;
            }
            if (this.m_isNullObject) {
                return null;
            }
            if (Utility.isUndefined(this.m_json)) {
                this.m_json = Utility.purifyJson(this.m_value);
            }
            return this.m_json;
        };
        RetrieveResultImpl.prototype.toString = function () {
            return JSON.stringify(this.toJSON());
        };
        RetrieveResultImpl.prototype._handleResult = function (value) {
            this.m_isLoaded = true;
            if (value === null || (typeof value === 'object' && value && value._IsNull)) {
                this.m_isNullObject = true;
                value = null;
            }
            else {
                this.m_isNullObject = false;
            }
            if (this.m_shouldPolyfill) {
                value = Utility.changePropertyNameToCamelLowerCase(value);
            }
            this.m_value = value;
            this.m_proxy._handleRetrieveResult(value, this);
        };
        return RetrieveResultImpl;
    }());
    var Constants = (function (_super) {
        __extends(Constants, _super);
        function Constants() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Constants.getItemAt = 'GetItemAt';
        Constants.index = '_Index';
        Constants.iterativeExecutor = 'IterativeExecutor';
        Constants.isTracked = '_IsTracked';
        Constants.eventMessageCategory = 65536;
        Constants.eventWorkbookId = 'Workbook';
        Constants.eventSourceRemote = 'Remote';
        Constants.proxy = '$proxy';
        Constants.className = '_className';
        Constants.isCollection = '_isCollection';
        Constants.collectionPropertyPath = '_collectionPropertyPath';
        Constants.objectPathInfoDoNotKeepReferenceFieldName = 'D';
        Constants.officeScriptEventId = 'X-OfficeScriptEventId';
        Constants.officeScriptFireRecordingEvent = 'X-OfficeScriptFireRecordingEvent';
        Constants.officeControlId = 'X-OfficeControlId';
        return Constants;
    }(CommonConstants));
    OfficeExtension_1.Constants = Constants;
    var ClientRequest = (function (_super) {
        __extends(ClientRequest, _super);
        function ClientRequest(context) {
            var _this = _super.call(this, context) || this;
            _this.m_context = context;
            _this.m_pendingProcessEventHandlers = [];
            _this.m_pendingEventHandlerActions = {};
            _this.m_traceInfos = {};
            _this.m_responseTraceIds = {};
            _this.m_responseTraceMessages = [];
            return _this;
        }
        Object.defineProperty(ClientRequest.prototype, "traceInfos", {
            get: function () {
                return this.m_traceInfos;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
            get: function () {
                return this.m_responseTraceMessages;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
            get: function () {
                return this.m_responseTraceIds;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype._setResponseTraceIds = function (value) {
            if (value) {
                for (var i = 0; i < value.length; i++) {
                    var traceId = value[i];
                    this.m_responseTraceIds[traceId] = traceId;
                    var message = this.m_traceInfos[traceId];
                    if (!CoreUtility.isNullOrUndefined(message)) {
                        this.m_responseTraceMessages.push(message);
                    }
                }
            }
        };
        ClientRequest.prototype.addTrace = function (actionId, message) {
            this.m_traceInfos[actionId] = message;
        };
        ClientRequest.prototype._addPendingEventHandlerAction = function (eventHandlers, action) {
            if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
                this.m_pendingEventHandlerActions[eventHandlers._id] = [];
                this.m_pendingProcessEventHandlers.push(eventHandlers);
            }
            this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
        };
        Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
            get: function () {
                return this.m_pendingProcessEventHandlers;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype._getPendingEventHandlerActions = function (eventHandlers) {
            return this.m_pendingEventHandlerActions[eventHandlers._id];
        };
        return ClientRequest;
    }(ClientRequestBase));
    OfficeExtension_1.ClientRequest = ClientRequest;
    var EventHandlers = (function () {
        function EventHandlers(context, parentObject, name, eventInfo) {
            var _this = this;
            this.m_id = context._nextId();
            this.m_context = context;
            this.m_name = name;
            this.m_handlers = [];
            this.m_registered = false;
            this.m_eventInfo = eventInfo;
            this.m_callback = function (args) {
                _this.m_eventInfo.eventArgsTransformFunc(args).then(function (newArgs) { return _this.fireEvent(newArgs); });
            };
        }
        Object.defineProperty(EventHandlers.prototype, "_registered", {
            get: function () {
                return this.m_registered;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_handlers", {
            get: function () {
                return this.m_handlers;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(EventHandlers.prototype, "_callback", {
            get: function () {
                return this.m_callback;
            },
            enumerable: true,
            configurable: true
        });
        EventHandlers.prototype.add = function (handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 0
            });
            return new EventHandlerResult(this.m_context, this, handler);
        };
        EventHandlers.prototype.remove = function (handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 1
            });
        };
        EventHandlers.prototype.removeAll = function () {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: null,
                operation: 2
            });
        };
        EventHandlers.prototype._processRegistration = function (req) {
            var _this = this;
            var ret = CoreUtility._createPromiseFromResult(null);
            var actions = req._getPendingEventHandlerActions(this);
            if (!actions) {
                return ret;
            }
            var handlersResult = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                handlersResult.push(this.m_handlers[i]);
            }
            var hasChange = false;
            for (var i = 0; i < actions.length; i++) {
                if (req._responseTraceIds[actions[i].id]) {
                    hasChange = true;
                    switch (actions[i].operation) {
                        case 0:
                            handlersResult.push(actions[i].handler);
                            break;
                        case 1:
                            for (var index = handlersResult.length - 1; index >= 0; index--) {
                                if (handlersResult[index] === actions[i].handler) {
                                    handlersResult.splice(index, 1);
                                    break;
                                }
                            }
                            break;
                        case 2:
                            handlersResult = [];
                            break;
                    }
                }
            }
            if (hasChange) {
                if (!this.m_registered && handlersResult.length > 0) {
                    ret = ret.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); }).then(function () { return (_this.m_registered = true); });
                }
                else if (this.m_registered && handlersResult.length == 0) {
                    ret = ret
                        .then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })["catch"](function (ex) {
                        CoreUtility.log('Error when unregister event: ' + JSON.stringify(ex));
                    })
                        .then(function () { return (_this.m_registered = false); });
                }
                ret = ret.then(function () { return (_this.m_handlers = handlersResult); });
            }
            return ret;
        };
        EventHandlers.prototype.fireEvent = function (args) {
            var promises = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                var handler = this.m_handlers[i];
                var p = CoreUtility._createPromiseFromResult(null)
                    .then(this.createFireOneEventHandlerFunc(handler, args))["catch"](function (ex) {
                    CoreUtility.log('Error when invoke handler: ' + JSON.stringify(ex));
                });
                promises.push(p);
            }
            CoreUtility.Promise.all(promises);
        };
        EventHandlers.prototype.createFireOneEventHandlerFunc = function (handler, args) {
            return function () { return handler(args); };
        };
        return EventHandlers;
    }());
    OfficeExtension_1.EventHandlers = EventHandlers;
    var EventHandlerResult = (function () {
        function EventHandlerResult(context, handlers, handler) {
            this.m_context = context;
            this.m_allHandlers = handlers;
            this.m_handler = handler;
        }
        Object.defineProperty(EventHandlerResult.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        EventHandlerResult.prototype.remove = function () {
            if (this.m_allHandlers && this.m_handler) {
                this.m_allHandlers.remove(this.m_handler);
                this.m_allHandlers = null;
                this.m_handler = null;
            }
        };
        return EventHandlerResult;
    }());
    OfficeExtension_1.EventHandlerResult = EventHandlerResult;
    (function (_Internal) {
        var OfficeJsEventRegistration = (function () {
            function OfficeJsEventRegistration() {
            }
            OfficeJsEventRegistration.prototype.register = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback);
                            });
                        });
                    case 3:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback);
                            });
                        });
                    case 2:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback);
                        });
                    case 1:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback);
                        });
                    case 5:
                        return OSF.DDA.RichApi.richApiMessageManager.register(handler);
                    case 13:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback);
                        });
                    case 14:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback);
                        });
                    case 15:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback);
                        });
                    case 16:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback);
                        });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
                }
            };
            OfficeJsEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback);
                            });
                        });
                    case 3:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback);
                            });
                        });
                    case 2:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback);
                        });
                    case 1:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback);
                        });
                    case 5:
                        return Utility.promisify(function (callback) {
                            return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync('richApiMessage', { handler: handler }, callback);
                        });
                    case 13:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback);
                        });
                    case 14:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback);
                        });
                    case 15:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback);
                        });
                    case 16:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback);
                        });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
                }
            };
            return OfficeJsEventRegistration;
        }());
        _Internal.officeJsEventRegistration = new OfficeJsEventRegistration();
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    var EventRegistration = (function () {
        function EventRegistration(registerEventImpl, unregisterEventImpl) {
            this.m_handlersByEventByTarget = {};
            this.m_registerEventImpl = registerEventImpl;
            this.m_unregisterEventImpl = unregisterEventImpl;
        }
        EventRegistration.getTargetIdOrDefault = function (targetId) {
            if (Utility.isNullOrUndefined(targetId)) {
                return '';
            }
            return targetId;
        };
        EventRegistration.prototype.getHandlers = function (eventId, targetId) {
            targetId = EventRegistration.getTargetIdOrDefault(targetId);
            var handlersById = this.m_handlersByEventByTarget[eventId];
            if (!handlersById) {
                handlersById = {};
                this.m_handlersByEventByTarget[eventId] = handlersById;
            }
            var handlers = handlersById[targetId];
            if (!handlers) {
                handlers = [];
                handlersById[targetId] = handlers;
            }
            return handlers;
        };
        EventRegistration.prototype.callHandlers = function (eventId, targetId, argument) {
            var funcs = this.getHandlers(eventId, targetId);
            for (var i = 0; i < funcs.length; i++) {
                funcs[i](argument);
            }
        };
        EventRegistration.prototype.hasHandlers = function (eventId, targetId) {
            return this.getHandlers(eventId, targetId).length > 0;
        };
        EventRegistration.prototype.register = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
            }
            var handlers = this.getHandlers(eventId, targetId);
            handlers.push(handler);
            if (handlers.length === 1) {
                return this.m_registerEventImpl(eventId, targetId);
            }
            return Utility._createPromiseFromResult(null);
        };
        EventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
            }
            var handlers = this.getHandlers(eventId, targetId);
            for (var index = handlers.length - 1; index >= 0; index--) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    break;
                }
            }
            if (handlers.length === 0) {
                return this.m_unregisterEventImpl(eventId, targetId);
            }
            return Utility._createPromiseFromResult(null);
        };
        return EventRegistration;
    }());
    OfficeExtension_1.EventRegistration = EventRegistration;
    var GenericEventRegistration = (function () {
        function GenericEventRegistration() {
            this.m_eventRegistration = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
            this.m_richApiMessageHandler = this._handleRichApiMessage.bind(this);
        }
        GenericEventRegistration.prototype.ready = function () {
            var _this = this;
            if (!this.m_ready) {
                if (GenericEventRegistration._testReadyImpl) {
                    this.m_ready = GenericEventRegistration._testReadyImpl().then(function () {
                        _this.m_isReady = true;
                    });
                }
                else if (HostBridge.instance) {
                    this.m_ready = Utility._createPromiseFromResult(null).then(function () {
                        _this.m_isReady = true;
                    });
                }
                else {
                    this.m_ready = _Internal.officeJsEventRegistration
                        .register(5, '', this.m_richApiMessageHandler)
                        .then(function () {
                        _this.m_isReady = true;
                    });
                }
            }
            return this.m_ready;
        };
        Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
            get: function () {
                return this.m_isReady;
            },
            enumerable: true,
            configurable: true
        });
        GenericEventRegistration.prototype.register = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready().then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready().then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype._registerEventImpl = function (eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._unregisterEventImpl = function (eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._handleRichApiMessage = function (msg) {
            if (msg && msg.entries) {
                for (var entryIndex = 0; entryIndex < msg.entries.length; entryIndex++) {
                    var entry = msg.entries[entryIndex];
                    if (entry.messageCategory == Constants.eventMessageCategory) {
                        if (CoreUtility._logEnabled) {
                            CoreUtility.log(JSON.stringify(entry));
                        }
                        var eventId = entry.messageType;
                        var targetId = entry.targetId;
                        var hasHandlers = this.m_eventRegistration.hasHandlers(eventId, targetId);
                        if (hasHandlers) {
                            var arg = JSON.parse(entry.message);
                            if (entry.isRemoteOverride) {
                                arg.source = Constants.eventSourceRemote;
                            }
                            this.m_eventRegistration.callHandlers(eventId, targetId, arg);
                        }
                    }
                }
            }
        };
        GenericEventRegistration.getGenericEventRegistration = function (controlId) {
            if (CoreUtility.isNullOrUndefined(controlId)) {
                controlId = '';
            }
            var ret = GenericEventRegistration.s_genericEventRegistrations[controlId];
            if (!ret) {
                ret = new GenericEventRegistration();
                GenericEventRegistration.s_genericEventRegistrations[controlId] = ret;
            }
            return ret;
        };
        GenericEventRegistration.peekGenericEventRegistrationOrNull = function (controlId) {
            if (CoreUtility.isNullOrUndefined(controlId)) {
                controlId = '';
            }
            var ret = GenericEventRegistration.s_genericEventRegistrations[controlId];
            return ret;
        };
        GenericEventRegistration.richApiMessageEventCategory = 65536;
        GenericEventRegistration.s_genericEventRegistrations = {};
        return GenericEventRegistration;
    }());
    OfficeExtension_1.GenericEventRegistration = GenericEventRegistration;
    function _testSetRichApiMessageReadyImpl(impl) {
        GenericEventRegistration._testReadyImpl = impl;
    }
    OfficeExtension_1._testSetRichApiMessageReadyImpl = _testSetRichApiMessageReadyImpl;
    function _testTriggerRichApiMessageEvent(msg) {
        GenericEventRegistration.getGenericEventRegistration('')._handleRichApiMessage(msg);
    }
    OfficeExtension_1._testTriggerRichApiMessageEvent = _testTriggerRichApiMessageEvent;
    var GenericEventHandlers = (function (_super) {
        __extends(GenericEventHandlers, _super);
        function GenericEventHandlers(context, parentObject, name, eventInfo) {
            var _this = _super.call(this, context, parentObject, name, eventInfo) || this;
            _this.m_genericEventInfo = eventInfo;
            return _this;
        }
        GenericEventHandlers.prototype.add = function (handler) {
            var _this = this;
            if (this._handlers.length == 0 && this.m_genericEventInfo.registerFunc) {
                this.m_genericEventInfo.registerFunc();
            }
            var controlId = this._context._controlId;
            if (!GenericEventRegistration.getGenericEventRegistration(controlId).isReady) {
                this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration(controlId).ready());
            }
            ActionFactory.createTraceMarkerForCallback(this._context, function () {
                _this._handlers.push(handler);
                if (_this._handlers.length == 1) {
                    GenericEventRegistration.getGenericEventRegistration(controlId).register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
            return new EventHandlerResult(this._context, this, handler);
        };
        GenericEventHandlers.prototype.remove = function (handler) {
            var _this = this;
            if (this._handlers.length == 1 && this.m_genericEventInfo.unregisterFunc) {
                this.m_genericEventInfo.unregisterFunc();
            }
            var controlId = this._context._controlId;
            ActionFactory.createTraceMarkerForCallback(this._context, function () {
                var handlers = _this._handlers;
                for (var index = handlers.length - 1; index >= 0; index--) {
                    if (handlers[index] === handler) {
                        handlers.splice(index, 1);
                        break;
                    }
                }
                if (handlers.length == 0) {
                    GenericEventRegistration.getGenericEventRegistration(controlId).unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
        };
        GenericEventHandlers.prototype.removeAll = function () { };
        return GenericEventHandlers;
    }(EventHandlers));
    OfficeExtension_1.GenericEventHandlers = GenericEventHandlers;
    var InstantiateActionResultHandler = (function () {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        InstantiateActionResultHandler.prototype._handleResult = function (value) {
            this.m_clientObject._handleIdResult(value);
        };
        return InstantiateActionResultHandler;
    }());
    var ObjectPathFactory = (function () {
        function ObjectPathFactory() {
        }
        ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 1,
                Name: ''
            };
            return new ObjectPath(objectPathInfo, null, false, false, 1, 4);
        };
        ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 2,
                Name: typeName
            };
            var ret = new ObjectPath(objectPathInfo, null, isCollection, false, 1, Utility._fixupApiFlags(flags));
            return ret;
        };
        ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id
            };
            var ret = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, 1, Utility._fixupApiFlags(flags));
            return ret;
        };
        ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new ObjectPath(objectPathInfo, parentObjectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
            var ret = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, Utility._fixupApiFlags(flags));
            ret.argumentObjectPaths = argumentObjectPaths;
            ret.getByIdMethodName = getByIdMethodName;
            return ret;
        };
        ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 6,
                Name: referenceId,
                ArgumentInfo: {}
            };
            var ret = new ObjectPath(objectPathInfo, null, false, false, 1, 4);
            return ret;
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt = function (hasIndexerMethod, context, parent, childItem, index) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
            if (hasIndexerMethod && !Utility.isNullOrUndefined(id)) {
                return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
            }
            else {
                return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
            }
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
            var objectPathInfo = (objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            });
            objectPathInfo.ArgumentInfo.Arguments = [id];
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
            var indexFromServer = childItem[Constants.index];
            if (indexFromServer) {
                index = indexFromServer;
            }
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: Constants.getItemAt,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [index];
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        return ObjectPathFactory;
    }());
    OfficeExtension_1.ObjectPathFactory = ObjectPathFactory;
    var OfficeJsRequestExecutor = (function () {
        function OfficeJsRequestExecutor(context) {
            this.m_context = context;
        }
        OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
            return new OfficeExtension_1.Promise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(result));
                    var response;
                    if (result.status == 'succeeded') {
                        response = RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBody(result), RichApiMessageUtility.getResponseHeaders(result));
                    }
                    else {
                        response = RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
                        _this.m_context._processOfficeJsErrorResponse(result.error.code, response);
                    }
                    resolve(response);
                });
            });
        };
        OfficeJsRequestExecutor.SourceLibHeaderValue = 'officejs';
        return OfficeJsRequestExecutor;
    }());
    var TrackedObjects = (function () {
        function TrackedObjects(context) {
            this._autoCleanupList = {};
            this.m_context = context;
        }
        TrackedObjects.prototype.add = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._addCommon(item, true); });
            }
            else {
                this._addCommon(param, true);
            }
        };
        TrackedObjects.prototype._autoAdd = function (object) {
            this._addCommon(object, false);
            this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
        };
        TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue = function (object, resultValue) {
            var shouldAutoTrack = this.m_context._autoCleanup &&
                !object[Constants.isTracked] &&
                object !== this.m_context._rootObject &&
                resultValue &&
                !Utility.isNullOrEmptyString(resultValue[Constants.referenceId]);
            if (shouldAutoTrack) {
                this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
                object[Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype._addCommon = function (object, isExplicitlyAdded) {
            if (object[Constants.isTracked]) {
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                return;
            }
            var referenceId = object[Constants.referenceId];
            var donotKeepReference = object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName];
            if (donotKeepReference) {
                throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.objectIsUntracked), null);
            }
            if (Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
                object._KeepReference();
                ActionFactory.createInstantiateAction(this.m_context, object);
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                object[Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype.remove = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._removeCommon(item); });
            }
            else {
                this._removeCommon(param);
            }
        };
        TrackedObjects.prototype._removeCommon = function (object) {
            object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName] = true;
            object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
            var referenceId = object[Constants.referenceId];
            if (!Utility.isNullOrEmptyString(referenceId)) {
                var rootObject = this.m_context._rootObject;
                if (rootObject._RemoveReference) {
                    rootObject._RemoveReference(referenceId);
                }
            }
            delete object[Constants.isTracked];
        };
        TrackedObjects.prototype._retrieveAndClearAutoCleanupList = function () {
            var list = this._autoCleanupList;
            this._autoCleanupList = {};
            return list;
        };
        return TrackedObjects;
    }());
    OfficeExtension_1.TrackedObjects = TrackedObjects;
    var RequestPrettyPrinter = (function () {
        function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
            if (!globalObjName) {
                globalObjName = 'root';
            }
            this.m_globalObjName = globalObjName;
            this.m_referencedObjectPaths = referencedObjectPaths;
            this.m_actions = actions;
            this.m_statements = [];
            this.m_variableNameForObjectPathMap = {};
            this.m_variableNameToObjectPathMap = {};
            this.m_declaredObjectPathMap = {};
            this.m_showDispose = showDispose;
            this.m_removePII = removePII;
        }
        RequestPrettyPrinter.prototype.process = function () {
            if (this.m_showDispose) {
                ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            }
            for (var i = 0; i < this.m_actions.length; i++) {
                this.processOneAction(this.m_actions[i]);
            }
            return this.m_statements;
        };
        RequestPrettyPrinter.prototype.processForDebugStatementInfo = function (actionIndex) {
            if (this.m_showDispose) {
                ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            }
            var surroundingCount = 5;
            this.m_statements = [];
            var oneStatement = '';
            var statementIndex = -1;
            for (var i = 0; i < this.m_actions.length; i++) {
                this.processOneAction(this.m_actions[i]);
                if (actionIndex == i) {
                    statementIndex = this.m_statements.length - 1;
                }
                if (statementIndex >= 0 && this.m_statements.length > statementIndex + surroundingCount + 1) {
                    break;
                }
            }
            if (statementIndex < 0) {
                return null;
            }
            var startIndex = statementIndex - surroundingCount;
            if (startIndex < 0) {
                startIndex = 0;
            }
            var endIndex = statementIndex + 1 + surroundingCount;
            if (endIndex > this.m_statements.length) {
                endIndex = this.m_statements.length;
            }
            var surroundingStatements = [];
            if (startIndex != 0) {
                surroundingStatements.push('...');
            }
            for (var i_1 = startIndex; i_1 < statementIndex; i_1++) {
                surroundingStatements.push(this.m_statements[i_1]);
            }
            surroundingStatements.push('// >>>>>');
            surroundingStatements.push(this.m_statements[statementIndex]);
            surroundingStatements.push('// <<<<<');
            for (var i_2 = statementIndex + 1; i_2 < endIndex; i_2++) {
                surroundingStatements.push(this.m_statements[i_2]);
            }
            if (endIndex < this.m_statements.length) {
                surroundingStatements.push('...');
            }
            return {
                statement: this.m_statements[statementIndex],
                surroundingStatements: surroundingStatements
            };
        };
        RequestPrettyPrinter.prototype.processOneAction = function (action) {
            var actionInfo = action.actionInfo;
            switch (actionInfo.ActionType) {
                case 1:
                    this.processInstantiateAction(action);
                    break;
                case 3:
                    this.processMethodAction(action);
                    break;
                case 2:
                    this.processQueryAction(action);
                    break;
                case 7:
                    this.processQueryAsJsonAction(action);
                    break;
                case 6:
                    this.processRecursiveQueryAction(action);
                    break;
                case 4:
                    this.processSetPropertyAction(action);
                    break;
                case 5:
                    this.processTraceAction(action);
                    break;
                case 8:
                    this.processEnsureUnchangedAction(action);
                    break;
                case 9:
                    this.processUpdateAction(action);
                    break;
            }
        };
        RequestPrettyPrinter.prototype.processInstantiateAction = function (action) {
            var objId = action.actionInfo.ObjectPathId;
            var objPath = this.m_referencedObjectPaths[objId];
            var varName = this.getObjVarName(objId);
            if (!this.m_declaredObjectPathMap[objId]) {
                var statement = 'var ' + varName + ' = ' + this.buildObjectPathExpressionWithParent(objPath) + ';';
                statement = this.appendDisposeCommentIfRelevant(statement, action);
                this.m_statements.push(statement);
                this.m_declaredObjectPathMap[objId] = varName;
            }
            else {
                var statement = '// Instantiate {' + varName + '}';
                statement = this.appendDisposeCommentIfRelevant(statement, action);
                this.m_statements.push(statement);
            }
        };
        RequestPrettyPrinter.prototype.processMethodAction = function (action) {
            var methodName = action.actionInfo.Name;
            if (methodName === '_KeepReference') {
                if (!OfficeExtension_1._internalConfig.showInternalApiInDebugInfo) {
                    return;
                }
                methodName = 'track';
            }
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.' +
                Utility._toCamelLowerCase(methodName) +
                '(' +
                this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processQueryAction = function (action) {
            var queryExp = this.buildQueryExpression(action);
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.load(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processQueryAsJsonAction = function (action) {
            var queryExp = this.buildQueryExpression(action);
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.retrieve(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processRecursiveQueryAction = function (action) {
            var queryExp = '';
            if (action.actionInfo.RecursiveQueryInfo) {
                queryExp = JSON.stringify(action.actionInfo.RecursiveQueryInfo);
            }
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.loadRecursive(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processSetPropertyAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.' +
                Utility._toCamelLowerCase(action.actionInfo.Name) +
                ' = ' +
                this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
                ';';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processTraceAction = function (action) {
            var statement = 'context.trace();';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processEnsureUnchangedAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.ensureUnchanged(' +
                JSON.stringify(action.actionInfo.ObjectState) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processUpdateAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.update(' +
                JSON.stringify(action.actionInfo.ObjectState) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant = function (statement, action) {
            var _this = this;
            if (this.m_showDispose) {
                var lastUsedObjectPathIds = action.actionInfo.L;
                if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                    var objectNamesToDispose = lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(', ');
                    return statement + ' // And then dispose {' + objectNamesToDispose + '}';
                }
            }
            return statement;
        };
        RequestPrettyPrinter.prototype.buildQueryExpression = function (action) {
            if (action.actionInfo.QueryInfo) {
                var option = {};
                option.select = action.actionInfo.QueryInfo.Select;
                option.expand = action.actionInfo.QueryInfo.Expand;
                option.skip = action.actionInfo.QueryInfo.Skip;
                option.top = action.actionInfo.QueryInfo.Top;
                if (typeof option.top === 'undefined' &&
                    typeof option.skip === 'undefined' &&
                    typeof option.expand === 'undefined') {
                    if (typeof option.select === 'undefined') {
                        return '';
                    }
                    else {
                        return JSON.stringify(option.select);
                    }
                }
                else {
                    return JSON.stringify(option);
                }
            }
            return '';
        };
        RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent = function (objPath) {
            var hasParent = objPath.objectPathInfo.ObjectPathType == 5 ||
                objPath.objectPathInfo.ObjectPathType == 3 ||
                objPath.objectPathInfo.ObjectPathType == 4;
            if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
                return (this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId) + '.' + this.buildObjectPathExpression(objPath));
            }
            return this.buildObjectPathExpression(objPath);
        };
        RequestPrettyPrinter.prototype.buildObjectPathExpression = function (objPath) {
            var expr = this.buildObjectPathInfoExpression(objPath.objectPathInfo);
            var originalObjectPathInfo = objPath.originalObjectPathInfo;
            if (originalObjectPathInfo) {
                expr = expr + ' /* originally ' + this.buildObjectPathInfoExpression(originalObjectPathInfo) + ' */';
            }
            return expr;
        };
        RequestPrettyPrinter.prototype.buildObjectPathInfoExpression = function (objectPathInfo) {
            switch (objectPathInfo.ObjectPathType) {
                case 1:
                    return 'context.' + this.m_globalObjName;
                case 5:
                    return 'getItem(' + this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) + ')';
                case 3:
                    return (Utility._toCamelLowerCase(objectPathInfo.Name) +
                        '(' +
                        this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) +
                        ')');
                case 2:
                    return objectPathInfo.Name + '.newObject()';
                case 7:
                    return 'null';
                case 4:
                    return Utility._toCamelLowerCase(objectPathInfo.Name);
                case 6:
                    return ('context.' + this.m_globalObjName + '._getObjectByReferenceId(' + JSON.stringify(objectPathInfo.Name) + ')');
            }
        };
        RequestPrettyPrinter.prototype.buildArgumentsExpression = function (args) {
            var ret = '';
            if (!args.Arguments || args.Arguments.length === 0) {
                return ret;
            }
            if (this.m_removePII) {
                if (typeof args.Arguments[0] === 'undefined') {
                    return ret;
                }
                return '...';
            }
            for (var i = 0; i < args.Arguments.length; i++) {
                if (i > 0) {
                    ret = ret + ', ';
                }
                ret =
                    ret +
                        this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
            }
            if (ret === 'undefined') {
                ret = '';
            }
            return ret;
        };
        RequestPrettyPrinter.prototype.buildArgumentLiteral = function (value, objectPathId) {
            if (typeof value == 'number' && value === objectPathId) {
                return this.getObjVarName(objectPathId);
            }
            else {
                return JSON.stringify(value);
            }
        };
        RequestPrettyPrinter.prototype.getObjVarNameBase = function (objectPathId) {
            var ret = 'v';
            var objPath = this.m_referencedObjectPaths[objectPathId];
            if (objPath) {
                switch (objPath.objectPathInfo.ObjectPathType) {
                    case 1:
                        ret = this.m_globalObjName;
                        break;
                    case 4:
                        ret = Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
                        break;
                    case 3:
                        var methodName = objPath.objectPathInfo.Name;
                        if (methodName.length > 3 && methodName.substr(0, 3) === 'Get') {
                            methodName = methodName.substr(3);
                        }
                        ret = Utility._toCamelLowerCase(methodName);
                        break;
                    case 5:
                        var parentName = this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
                        if (parentName.charAt(parentName.length - 1) === 's') {
                            ret = parentName.substr(0, parentName.length - 1);
                        }
                        else {
                            ret = parentName + 'Item';
                        }
                        break;
                }
            }
            return ret;
        };
        RequestPrettyPrinter.prototype.getObjVarName = function (objectPathId) {
            if (this.m_variableNameForObjectPathMap[objectPathId]) {
                return this.m_variableNameForObjectPathMap[objectPathId];
            }
            var ret = this.getObjVarNameBase(objectPathId);
            if (!this.m_variableNameToObjectPathMap[ret]) {
                this.m_variableNameForObjectPathMap[objectPathId] = ret;
                this.m_variableNameToObjectPathMap[ret] = objectPathId;
                return ret;
            }
            var i = 1;
            while (this.m_variableNameToObjectPathMap[ret + i.toString()]) {
                i++;
            }
            ret = ret + i.toString();
            this.m_variableNameForObjectPathMap[objectPathId] = ret;
            this.m_variableNameToObjectPathMap[ret] = objectPathId;
            return ret;
        };
        return RequestPrettyPrinter;
    }());
    var ResourceStrings = (function (_super) {
        __extends(ResourceStrings, _super);
        function ResourceStrings() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ResourceStrings.cannotRegisterEvent = 'CannotRegisterEvent';
        ResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
        ResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
        ResourceStrings.propertyNotLoaded = 'PropertyNotLoaded';
        ResourceStrings.runMustReturnPromise = 'RunMustReturnPromise';
        ResourceStrings.moreInfoInnerError = 'MoreInfoInnerError';
        ResourceStrings.cannotApplyPropertyThroughSetMethod = 'CannotApplyPropertyThroughSetMethod';
        ResourceStrings.invalidOperationInCellEditMode = 'InvalidOperationInCellEditMode';
        ResourceStrings.objectIsUntracked = 'ObjectIsUntracked';
        ResourceStrings.customFunctionDefintionMissing = 'CustomFunctionDefintionMissing';
        ResourceStrings.customFunctionImplementationMissing = 'CustomFunctionImplementationMissing';
        ResourceStrings.customFunctionNameContainsBadChars = 'CustomFunctionNameContainsBadChars';
        ResourceStrings.customFunctionNameCannotSplit = 'CustomFunctionNameCannotSplit';
        ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch = 'CustomFunctionUnexpectedNumberOfEntriesInResultBatch';
        ResourceStrings.customFunctionCancellationHandlerMissing = 'CustomFunctionCancellationHandlerMissing';
        ResourceStrings.customFunctionInvalidFunction = 'CustomFunctionInvalidFunction';
        ResourceStrings.customFunctionInvalidFunctionMapping = 'CustomFunctionInvalidFunctionMapping';
        ResourceStrings.customFunctionWindowMissing = 'CustomFunctionWindowMissing';
        ResourceStrings.customFunctionDefintionMissingOnWindow = 'CustomFunctionDefintionMissingOnWindow';
        ResourceStrings.pendingBatchInProgress = 'PendingBatchInProgress';
        ResourceStrings.notInsideBatch = 'NotInsideBatch';
        ResourceStrings.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
        return ResourceStrings;
    }(CommonResourceStrings));
    OfficeExtension_1.ResourceStrings = ResourceStrings;
    CoreUtility.addResourceStringValues({
        CannotRegisterEvent: 'The event handler cannot be registered.',
        PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",
        RunMustReturnPromise: 'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',
        InvalidOrTimedOutSessionMessage: 'Your Office Online session has expired or is invalid. To continue, refresh the page.',
        InvalidOperationInCellEditMode: 'Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.',
        InvalidSheetName: 'The request cannot be processed because the specified worksheet cannot be found. Please try again.',
        CustomFunctionDefintionMissing: "A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",
        CustomFunctionDefintionMissingOnWindow: "A property with the name '{0}' that represents the function's definition must exist on the window object.",
        CustomFunctionImplementationMissing: "The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",
        CustomFunctionNameContainsBadChars: 'The function name may only contain letters, digits, underscores, and periods.',
        CustomFunctionNameCannotSplit: 'The function name must contain a non-empty namespace and a non-empty short name.',
        CustomFunctionUnexpectedNumberOfEntriesInResultBatch: "The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",
        CustomFunctionCancellationHandlerMissing: 'The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.',
        CustomFunctionInvalidFunction: "The property with the name '{0}' that represents the function's definition is not a valid function.",
        CustomFunctionInvalidFunctionMapping: "The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",
        CustomFunctionWindowMissing: 'The window object was not found.',
        PendingBatchInProgress: 'There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.',
        NotInsideBatch: 'Operations may not be invoked outside of a batch method.',
        CannotUpdateReadOnlyProperty: "The property '{0}' is read-only and it cannot be updated.",
        ObjectIsUntracked: 'The object is untracked.'
    });
    var Utility = (function (_super) {
        __extends(Utility, _super);
        function Utility() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Utility.fixObjectPathIfNecessary = function (clientObject, value) {
            if (clientObject && clientObject._objectPath && value) {
                clientObject._objectPath.updateUsingObjectData(value, clientObject);
            }
        };
        Utility.load = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
            return clientObj;
        };
        Utility.loadAndSync = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
            return clientObj.context.sync().then(function () { return clientObj; });
        };
        Utility.retrieve = function (clientObj, option) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !Utility.isSetSupported('RichApiRuntime', '1.1');
            }
            var result = new RetrieveResultImpl(clientObj, shouldPolyfill);
            clientObj._retrieve(option, result);
            return result;
        };
        Utility.retrieveAndSync = function (clientObj, option) {
            var result = Utility.retrieve(clientObj, option);
            return clientObj.context.sync().then(function () { return result; });
        };
        Utility.toJson = function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
            var result = {};
            for (var prop in scalarProperties) {
                var value = scalarProperties[prop];
                if (typeof value !== 'undefined') {
                    result[prop] = value;
                }
            }
            for (var prop in navigationProperties) {
                var value = navigationProperties[prop];
                if (typeof value !== 'undefined') {
                    if (value[Utility.fieldName_isCollection] && typeof value[Utility.fieldName_m__items] !== 'undefined') {
                        result[prop] = value.toJSON()['items'];
                    }
                    else {
                        result[prop] = value.toJSON();
                    }
                }
            }
            if (collectionItemsIfAny) {
                result['items'] = collectionItemsIfAny.map(function (item) { return item.toJSON(); });
            }
            return result;
        };
        Utility.throwError = function (resourceId, arg, errorLocation) {
            throw new _Internal.RuntimeError({
                code: resourceId,
                httpStatusCode: 400,
                message: CoreUtility._getResourceString(resourceId, arg),
                debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
            });
        };
        Utility.createRuntimeError = function (code, message, location, httpStatusCode, data) {
            return new _Internal.RuntimeError({
                code: code,
                httpStatusCode: httpStatusCode,
                message: message,
                debugInfo: { errorLocation: location },
                data: data
            });
        };
        Utility.throwIfNotLoaded = function (propertyName, fieldValue, entityName, isNull) {
            if (!isNull &&
                CoreUtility.isUndefined(fieldValue) &&
                propertyName.charCodeAt(0) != Utility.s_underscoreCharCode &&
                !Utility.s_suppressPropertyNotLoadedException) {
                throw Utility.createPropertyNotLoadedException(entityName, propertyName);
            }
        };
        Utility._getPropertyValueWithoutCheckLoaded = function (object, propertyName) {
            Utility.s_suppressPropertyNotLoadedException = true;
            try {
                return object[propertyName];
            }
            finally {
                Utility.s_suppressPropertyNotLoadedException = false;
            }
        };
        Utility.createPropertyNotLoadedException = function (entityName, propertyName) {
            return new _Internal.RuntimeError({
                code: ErrorCodes.propertyNotLoaded,
                httpStatusCode: 400,
                message: CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
                debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
            });
        };
        Utility.createCannotUpdateReadOnlyPropertyException = function (entityName, propertyName) {
            return new _Internal.RuntimeError({
                code: ErrorCodes.cannotUpdateReadOnlyProperty,
                httpStatusCode: 400,
                message: CoreUtility._getResourceString(ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
                debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
            });
        };
        Utility.promisify = function (action) {
            return new OfficeExtension_1.Promise(function (resolve, reject) {
                var callback = function (result) {
                    if (result.status == 'failed') {
                        reject(result.error);
                    }
                    else {
                        resolve(result.value);
                    }
                };
                action(callback);
            });
        };
        Utility._addActionResultHandler = function (clientObj, action, resultHandler) {
            clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
        };
        Utility._handleNavigationPropertyResults = function (clientObj, objectValue, propertyNames) {
            for (var i = 0; i < propertyNames.length - 1; i += 2) {
                if (!CoreUtility.isUndefined(objectValue[propertyNames[i + 1]])) {
                    clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
                }
            }
        };
        Utility._fixupApiFlags = function (flags) {
            if (typeof flags === 'boolean') {
                if (flags) {
                    flags = 1;
                }
                else {
                    flags = 0;
                }
            }
            return flags;
        };
        Utility.definePropertyThrowUnloadedException = function (obj, typeName, propertyName) {
            Object.defineProperty(obj, propertyName, {
                configurable: true,
                enumerable: true,
                get: function () {
                    throw Utility.createPropertyNotLoadedException(typeName, propertyName);
                },
                set: function () {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
                }
            });
        };
        Utility.defineReadOnlyPropertyWithValue = function (obj, propertyName, value) {
            Object.defineProperty(obj, propertyName, {
                configurable: true,
                enumerable: true,
                get: function () {
                    return value;
                },
                set: function () {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
                }
            });
        };
        Utility.processRetrieveResult = function (proxy, value, result, childItemCreateFunc) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return;
            }
            if (childItemCreateFunc) {
                var data = value[Constants.itemsLowerCase];
                if (Array.isArray(data)) {
                    var itemsResult = [];
                    for (var i = 0; i < data.length; i++) {
                        var itemProxy = childItemCreateFunc(data[i], i);
                        var itemResult = {};
                        itemResult[Constants.proxy] = itemProxy;
                        itemProxy._handleRetrieveResult(data[i], itemResult);
                        itemsResult.push(itemResult);
                    }
                    Utility.defineReadOnlyPropertyWithValue(result, Constants.itemsLowerCase, itemsResult);
                }
            }
            else {
                var scalarPropertyNames = proxy[Constants.scalarPropertyNames];
                var navigationPropertyNames = proxy[Constants.navigationPropertyNames];
                var typeName = proxy[Constants.className];
                if (scalarPropertyNames) {
                    for (var i = 0; i < scalarPropertyNames.length; i++) {
                        var propName = scalarPropertyNames[i];
                        var propValue = value[propName];
                        if (CoreUtility.isUndefined(propValue)) {
                            Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                        }
                        else {
                            Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
                        }
                    }
                }
                if (navigationPropertyNames) {
                    for (var i = 0; i < navigationPropertyNames.length; i++) {
                        var propName = navigationPropertyNames[i];
                        var propValue = value[propName];
                        if (CoreUtility.isUndefined(propValue)) {
                            Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                        }
                        else {
                            var propProxy = proxy[propName];
                            var propResult = {};
                            propProxy._handleRetrieveResult(propValue, propResult);
                            propResult[Constants.proxy] = propProxy;
                            if (Array.isArray(propResult[Constants.itemsLowerCase])) {
                                propResult = propResult[Constants.itemsLowerCase];
                            }
                            Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
                        }
                    }
                }
            }
        };
        Utility.setMockData = function (clientObj, value, childItemCreateFunc, setItemsFunc) {
            if (CoreUtility.isNullOrUndefined(value)) {
                clientObj._handleResult(value);
                return;
            }
            if (clientObj[Constants.scalarPropertyOriginalNames]) {
                var result = {};
                var scalarPropertyOriginalNames = clientObj[Constants.scalarPropertyOriginalNames];
                var scalarPropertyNames = clientObj[Constants.scalarPropertyNames];
                for (var i = 0; i < scalarPropertyNames.length; i++) {
                    if (typeof (value[scalarPropertyNames[i]]) !== 'undefined') {
                        result[scalarPropertyOriginalNames[i]] = value[scalarPropertyNames[i]];
                    }
                }
                clientObj._handleResult(result);
            }
            if (clientObj[Constants.navigationPropertyNames]) {
                var navigationPropertyNames = clientObj[Constants.navigationPropertyNames];
                for (var i = 0; i < navigationPropertyNames.length; i++) {
                    if (typeof (value[navigationPropertyNames[i]]) !== 'undefined') {
                        var navigationPropValue = clientObj[navigationPropertyNames[i]];
                        if (navigationPropValue.setMockData) {
                            navigationPropValue.setMockData(value[navigationPropertyNames[i]]);
                        }
                    }
                }
            }
            if (clientObj[Constants.isCollection] && childItemCreateFunc) {
                var itemsData = Array.isArray(value) ? value : value[Constants.itemsLowerCase];
                if (Array.isArray(itemsData)) {
                    var items = [];
                    for (var i = 0; i < itemsData.length; i++) {
                        var item = childItemCreateFunc(itemsData, i);
                        Utility.setMockData(item, itemsData[i]);
                        items.push(item);
                    }
                    setItemsFunc(items);
                }
            }
        };
        Utility.applyMixin = function (derived, base) {
            Object.getOwnPropertyNames(base.prototype).forEach(function (name) {
                if (name !== 'constructor') {
                    Object.defineProperty(derived.prototype, name, Object.getOwnPropertyDescriptor(base.prototype, name));
                }
            });
        };
        Utility.ensureTypeInitialized = function (type) {
            var context = new ClientRequestContext();
            var objectPath = ObjectPathFactory.createNewObjectObjectPath(context, "Temp", false, 0);
            new type(context, objectPath);
        };
        Utility.fieldName_m__items = 'm__items';
        Utility.fieldName_isCollection = '_isCollection';
        Utility._synchronousCleanup = false;
        Utility.s_underscoreCharCode = '_'.charCodeAt(0);
        Utility.s_suppressPropertyNotLoadedException = false;
        return Utility;
    }(CommonUtility));
    OfficeExtension_1.Utility = Utility;
    var BatchApiHelper = (function () {
        function BatchApiHelper() {
        }
        BatchApiHelper.invokeMethod = function (obj, methodName, operationType, args, flags, resultProcessType) {
            var action = ActionFactory.createMethodAction(obj.context, obj, methodName, operationType, args, flags);
            var result = new ClientResult(resultProcessType);
            Utility._addActionResultHandler(obj, action, result);
            return result;
        };
        BatchApiHelper.invokeEnsureUnchanged = function (obj, objectState) {
            ActionFactory.createEnsureUnchangedAction(obj.context, obj, objectState);
        };
        BatchApiHelper.invokeSetProperty = function (obj, propName, propValue, flags) {
            ActionFactory.createSetPropertyAction(obj.context, obj, propName, propValue, flags);
        };
        BatchApiHelper.createRootServiceObject = function (type, context) {
            var objectPath = ObjectPathFactory.createGlobalObjectObjectPath(context);
            return new type(context, objectPath);
        };
        BatchApiHelper.createObjectFromReferenceId = function (type, context, referenceId) {
            var objectPath = ObjectPathFactory.createReferenceIdObjectPath(context, referenceId);
            return new type(context, objectPath);
        };
        BatchApiHelper.createTopLevelServiceObject = function (type, context, typeName, isCollection, flags) {
            var objectPath = ObjectPathFactory.createNewObjectObjectPath(context, typeName, isCollection, flags);
            return new type(context, objectPath);
        };
        BatchApiHelper.createPropertyObject = function (type, parent, propertyName, isCollection, flags) {
            var objectPath = ObjectPathFactory.createPropertyObjectPath(parent.context, parent, propertyName, isCollection, false, flags);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createIndexerObject = function (type, parent, args) {
            var objectPath = ObjectPathFactory.createIndexerObjectPath(parent.context, parent, args);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPath = ObjectPathFactory.createMethodObjectPath(parent.context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createChildItemObject = function (type, hasIndexerMethod, parent, chileItem, index) {
            var objectPath = ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, parent.context, parent, chileItem, index);
            return new type(parent.context, objectPath);
        };
        return BatchApiHelper;
    }());
    OfficeExtension_1.BatchApiHelper = BatchApiHelper;
    var LibraryBuilder = (function () {
        function LibraryBuilder(options) {
            this.m_namespaceMap = {};
            this.m_namespace = options.metadata.name;
            this.m_targetNamespaceObject = options.targetNamespaceObject;
            this.m_namespaceMap[this.m_namespace] = options.targetNamespaceObject;
            if (options.namespaceMap) {
                for (var ns in options.namespaceMap) {
                    this.m_namespaceMap[ns] = options.namespaceMap[ns];
                }
            }
            this.m_defaultApiSetName = options.metadata.defaultApiSetName;
            this.m_hostName = options.metadata.hostName;
            var metadata = options.metadata;
            if (metadata.enumTypes) {
                for (var i = 0; i < metadata.enumTypes.length; i++) {
                    this.buildEnumType(metadata.enumTypes[i]);
                }
            }
            if (metadata.apiSets) {
                for (var i = 0; i < metadata.apiSets.length; i++) {
                    var elem = metadata.apiSets[i];
                    if (Array.isArray(elem)) {
                        metadata.apiSets[i] = {
                            version: elem[0],
                            name: elem[1] || this.m_defaultApiSetName
                        };
                    }
                }
                this.m_apiSets = metadata.apiSets;
            }
            this.m_strings = metadata.strings;
            if (metadata.clientObjectTypes) {
                for (var i = 0; i < metadata.clientObjectTypes.length; i++) {
                    var elem = metadata.clientObjectTypes[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 11);
                        metadata.clientObjectTypes[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            collectionPropertyPath: this.getString(elem[6]),
                            newObjectServerTypeFullName: this.getString(elem[9]),
                            newObjectApiFlags: elem[10],
                            childItemTypeFullName: this.getString(elem[7]),
                            scalarProperties: elem[2],
                            navigationProperties: elem[3],
                            scalarMethods: elem[4],
                            navigationMethods: elem[5],
                            events: elem[8]
                        };
                    }
                    this.buildClientObjectType(metadata.clientObjectTypes[i], options.fullyInitialize);
                }
            }
        }
        LibraryBuilder.prototype.ensureArraySize = function (value, size) {
            var count = size - value.length;
            while (count > 0) {
                value.push(0);
                count--;
            }
        };
        LibraryBuilder.prototype.getString = function (ordinalOrValue) {
            if (typeof (ordinalOrValue) === "number") {
                if (ordinalOrValue > 0) {
                    return this.m_strings[ordinalOrValue - 1];
                }
                return null;
            }
            return ordinalOrValue;
        };
        LibraryBuilder.prototype.buildEnumType = function (elem) {
            var enumType;
            if (Array.isArray(elem)) {
                enumType = {
                    name: elem[0],
                    fields: elem[2]
                };
                if (!enumType.fields) {
                    enumType.fields = {};
                }
                var fieldsWithCamelUpperCaseValue = elem[1];
                if (Array.isArray(fieldsWithCamelUpperCaseValue)) {
                    for (var index = 0; index < fieldsWithCamelUpperCaseValue.length; index++) {
                        enumType.fields[fieldsWithCamelUpperCaseValue[index]] = this.toSimpleCamelUpperCase(fieldsWithCamelUpperCaseValue[index]);
                    }
                }
            }
            else {
                enumType = elem;
            }
            this.m_targetNamespaceObject[enumType.name] = enumType.fields;
        };
        LibraryBuilder.prototype.buildClientObjectType = function (typeInfo, fullyInitialize) {
            var thisBuilder = this;
            var type = function (context, objectPath) {
                ClientObject.apply(this, arguments);
                if (!thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited) {
                    thisBuilder.buildPrototype(thisBuilder.m_targetNamespaceObject[typeInfo.name], typeInfo);
                    thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited = true;
                }
                if (OfficeExtension_1._internalConfig.appendTypeNameToObjectPathInfo) {
                    if (this._objectPath && this._objectPath.objectPathInfo && this._className) {
                        this._objectPath.objectPathInfo.T = this._className;
                    }
                }
            };
            this.m_targetNamespaceObject[typeInfo.name] = type;
            this.extendsType(type, ClientObject);
            this.buildNewObject(type, typeInfo);
            if ((typeInfo.behaviorFlags & 2) !== 0) {
                type.prototype._KeepReference = function () {
                    BatchApiHelper.invokeMethod(this, "_KeepReference", 1, [], 0, 0);
                };
            }
            if ((typeInfo.behaviorFlags & 32) !== 0) {
                var func = this.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_StaticCustomize");
                func.call(null, type);
            }
            if (fullyInitialize) {
                this.buildPrototype(type, typeInfo);
                type._typeInited = true;
            }
        };
        LibraryBuilder.prototype.extendsType = function (d, b) {
            function __() { this.constructor = d; }
            d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
        };
        LibraryBuilder.prototype.findObjectUnderPath = function (top, paths, pathStartIndex) {
            var obj = top;
            for (var i = pathStartIndex; i < paths.length; i++) {
                if (typeof (obj) !== 'object') {
                    throw new OfficeExtension_1.Error("Cannot find " + paths.join("."));
                }
                obj = obj[paths[i]];
            }
            return obj;
        };
        LibraryBuilder.prototype.getFunction = function (fullName) {
            var ret = this.resolveObjectByFullName(fullName);
            if (typeof (ret) !== 'function') {
                throw new OfficeExtension_1.Error("Cannot find function or type: " + fullName);
            }
            return ret;
        };
        LibraryBuilder.prototype.resolveObjectByFullName = function (fullName) {
            var parts = fullName.split('.');
            if (parts.length === 1) {
                return this.m_targetNamespaceObject[parts[0]];
            }
            var rootName = parts[0];
            if (rootName === this.m_namespace) {
                return this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 1);
            }
            if (this.m_namespaceMap[rootName]) {
                return this.findObjectUnderPath(this.m_namespaceMap[rootName], parts, 1);
            }
            return this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 0);
        };
        LibraryBuilder.prototype.evaluateSimpleExpression = function (expression, thisObj) {
            if (Utility.isNullOrUndefined(expression)) {
                return null;
            }
            var paths = expression.split('.');
            if (paths.length === 3 && paths[0] === 'OfficeExtension' && paths[1] === 'Constants') {
                return Constants[paths[2]];
            }
            if (paths[0] === 'this') {
                var obj = thisObj;
                for (var i = 1; i < paths.length; i++) {
                    if (paths[i] == 'toString()') {
                        obj = obj.toString();
                    }
                    else if (paths[i].substr(paths[i].length - 2) === "()") {
                        obj = obj[paths[i].substr(0, paths[i].length - 2)]();
                    }
                    else {
                        obj = obj[paths[i]];
                    }
                }
                return obj;
            }
            throw new OfficeExtension_1.Error("Cannot evaluate: " + expression);
        };
        LibraryBuilder.prototype.evaluateEventTargetId = function (targetIdExpression, thisObj) {
            if (Utility.isNullOrEmptyString(targetIdExpression)) {
                return '';
            }
            return this.evaluateSimpleExpression(targetIdExpression, thisObj);
        };
        LibraryBuilder.prototype.isAllDigits = function (expression) {
            var charZero = '0'.charCodeAt(0);
            var charNine = '9'.charCodeAt(0);
            for (var i = 0; i < expression.length; i++) {
                if (expression.charCodeAt(i) < charZero ||
                    expression.charCodeAt(i) > charNine) {
                    return false;
                }
            }
            return true;
        };
        LibraryBuilder.prototype.evaluateEventType = function (eventTypeExpression) {
            if (Utility.isNullOrEmptyString(eventTypeExpression)) {
                return 0;
            }
            if (this.isAllDigits(eventTypeExpression)) {
                return parseInt(eventTypeExpression);
            }
            var ret = this.resolveObjectByFullName(eventTypeExpression);
            if (typeof (ret) !== 'number') {
                throw new OfficeExtension_1.Error("Invalid event type: " + eventTypeExpression);
            }
            return ret;
        };
        LibraryBuilder.prototype.buildPrototype = function (type, typeInfo) {
            this.buildScalarProperties(type, typeInfo);
            this.buildNavigationProperties(type, typeInfo);
            this.buildScalarMethods(type, typeInfo);
            this.buildNavigationMethods(type, typeInfo);
            this.buildEvents(type, typeInfo);
            this.buildHandleResult(type, typeInfo);
            this.buildHandleIdResult(type, typeInfo);
            this.buildHandleRetrieveResult(type, typeInfo);
            this.buildLoad(type, typeInfo);
            this.buildRetrieve(type, typeInfo);
            this.buildSetMockData(type, typeInfo);
            this.buildEnsureUnchanged(type, typeInfo);
            this.buildUpdate(type, typeInfo);
            this.buildSet(type, typeInfo);
            this.buildToJSON(type, typeInfo);
            this.buildItems(type, typeInfo);
            this.buildTypeMetadataInfo(type, typeInfo);
            this.buildTrackUntrack(type, typeInfo);
            this.buildMixin(type, typeInfo);
        };
        LibraryBuilder.prototype.toSimpleCamelUpperCase = function (name) {
            return name.substr(0, 1).toUpperCase() + name.substr(1);
        };
        LibraryBuilder.prototype.ensureOriginalName = function (member) {
            if (member.originalName === null) {
                member.originalName = this.toSimpleCamelUpperCase(member.name);
            }
        };
        LibraryBuilder.prototype.getFieldName = function (member) {
            return "m_" + member.name;
        };
        LibraryBuilder.prototype.throwIfApiNotSupported = function (typeInfo, member) {
            if (this.m_apiSets && member.apiSetInfoOrdinal > 0) {
                var apiSetInfo = this.m_apiSets[member.apiSetInfoOrdinal - 1];
                if (apiSetInfo) {
                    Utility.throwIfApiNotSupported(typeInfo.name + "." + member.name, apiSetInfo.name, apiSetInfo.version, this.m_hostName);
                }
            }
        };
        LibraryBuilder.prototype.buildScalarProperties = function (type, typeInfo) {
            if (Array.isArray(typeInfo.scalarProperties)) {
                for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                    var elem = typeInfo.scalarProperties[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 6);
                        typeInfo.scalarProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            apiSetInfoOrdinal: elem[2],
                            originalName: this.getString(elem[3]),
                            setMethodApiFlags: elem[4],
                            undoableApiSetInfoOrdinal: elem[5]
                        };
                    }
                    this.buildScalarProperty(type, typeInfo, typeInfo.scalarProperties[i]);
                }
            }
        };
        LibraryBuilder.prototype.calculateApiFlags = function (apiFlags, undoableApiSetInfoOrdinal) {
            if (undoableApiSetInfoOrdinal > 0) {
                var undoableApiSetInfo = this.m_apiSets[undoableApiSetInfoOrdinal - 1];
                if (undoableApiSetInfo) {
                    apiFlags = CommonUtility.calculateApiFlags(apiFlags, undoableApiSetInfo.name, undoableApiSetInfo.version);
                }
            }
            return apiFlags;
        };
        LibraryBuilder.prototype.buildScalarProperty = function (type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this;
            var fieldName = this.getFieldName(propInfo);
            var descriptor = {
                get: function () {
                    Utility.throwIfNotLoaded(propInfo.name, this[fieldName], typeInfo.name, this._isNull);
                    thisBuilder.throwIfApiNotSupported(typeInfo, propInfo);
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            };
            if ((propInfo.behaviorFlags & 2) === 0) {
                descriptor.set = function (value) {
                    if (propInfo.behaviorFlags & 4) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set");
                        var handled = customizationFunc.call(this, this, value).handled;
                        if (handled) {
                            return;
                        }
                    }
                    this[fieldName] = value;
                    var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildNavigationProperties = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationProperties)) {
                for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                    var elem = typeInfo.navigationProperties[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 8);
                        typeInfo.navigationProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[4]),
                            getMethodApiFlags: elem[5],
                            setMethodApiFlags: elem[6],
                            propertyTypeFullName: this.getString(elem[1]),
                            undoableApiSetInfoOrdinal: elem[7]
                        };
                    }
                    this.buildNavigationProperty(type, typeInfo, typeInfo.navigationProperties[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildNavigationProperty = function (type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this;
            var fieldName = this.getFieldName(propInfo);
            var descriptor = {
                get: function () {
                    if (!this[thisBuilder.getFieldName(propInfo)]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, propInfo);
                        this[fieldName] = BatchApiHelper.createPropertyObject(thisBuilder.getFunction(propInfo.propertyTypeFullName), this, propInfo.originalName, (propInfo.behaviorFlags & 16) !== 0, propInfo.getMethodApiFlags);
                    }
                    if (propInfo.behaviorFlags & 64) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Get");
                        customizationFunc.call(this, this, this[fieldName]);
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            };
            if ((propInfo.behaviorFlags & 2) === 0) {
                descriptor.set = function (value) {
                    if (propInfo.behaviorFlags & 4) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set");
                        var handled = customizationFunc.call(this, this, value).handled;
                        if (handled) {
                            return;
                        }
                    }
                    this[fieldName] = value;
                    var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildScalarMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.scalarMethods)) {
                for (var i = 0; i < typeInfo.scalarMethods.length; i++) {
                    var elem = typeInfo.scalarMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 7);
                        typeInfo.scalarMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[5]),
                            apiFlags: elem[4],
                            parameterCount: elem[1],
                            undoableApiSetInfoOrdinal: elem[6]
                        };
                    }
                    this.buildScalarMethod(type, typeInfo, typeInfo.scalarMethods[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildScalarMethod = function (type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function () {
                var args = [];
                if ((methodInfo.behaviorFlags & 64) && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) {
                        args.push(arguments[i]);
                    }
                    var rest = [];
                    for (var i = methodInfo.parameterCount - 1; i < arguments.length; i++) {
                        rest.push(arguments[i]);
                    }
                    args.push(rest);
                }
                else {
                    for (var i = 0; i < arguments.length; i++) {
                        args.push(arguments[i]);
                    }
                }
                if (methodInfo.behaviorFlags & 1) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName);
                    var applyArgs = [this];
                    for (var i = 0; i < args.length; i++) {
                        applyArgs.push(args[i]);
                    }
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) {
                        return result;
                    }
                }
                thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo);
                var resultProcessType = 0;
                if (methodInfo.behaviorFlags & 32) {
                    resultProcessType = 1;
                }
                var operationType = 0;
                if (methodInfo.behaviorFlags & 2) {
                    operationType = 1;
                }
                var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                return BatchApiHelper.invokeMethod(this, methodInfo.originalName, operationType, args, apiFlags, resultProcessType);
            };
        };
        LibraryBuilder.prototype.buildNavigationMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationMethods)) {
                for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                    var elem = typeInfo.navigationMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 9);
                        typeInfo.navigationMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[3],
                            apiSetInfoOrdinal: elem[4],
                            originalName: this.getString(elem[6]),
                            apiFlags: elem[5],
                            parameterCount: elem[2],
                            returnTypeFullName: this.getString(elem[1]),
                            returnObjectGetByIdMethodName: this.getString(elem[7]),
                            undoableApiSetInfoOrdinal: elem[8]
                        };
                    }
                    this.buildNavigationMethod(type, typeInfo, typeInfo.navigationMethods[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildNavigationMethod = function (type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function () {
                var args = [];
                if ((methodInfo.behaviorFlags & 64) && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) {
                        args.push(arguments[i]);
                    }
                    var rest = [];
                    for (var i = methodInfo.parameterCount - 1; i < arguments.length; i++) {
                        rest.push(arguments[i]);
                    }
                    args.push(rest);
                }
                else {
                    for (var i = 0; i < arguments.length; i++) {
                        args.push(arguments[i]);
                    }
                }
                if (methodInfo.behaviorFlags & 1) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName);
                    var applyArgs = [this];
                    for (var i = 0; i < args.length; i++) {
                        applyArgs.push(args[i]);
                    }
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) {
                        return result;
                    }
                }
                thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo);
                if ((methodInfo.behaviorFlags & 16) !== 0) {
                    return BatchApiHelper.createIndexerObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, args);
                }
                else {
                    var operationType = 0;
                    if (methodInfo.behaviorFlags & 2) {
                        operationType = 1;
                    }
                    var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                    return BatchApiHelper.createMethodObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, methodInfo.originalName, operationType, args, (methodInfo.behaviorFlags & 4) !== 0, (methodInfo.behaviorFlags & 8) !== 0, methodInfo.returnObjectGetByIdMethodName, apiFlags);
                }
            };
        };
        LibraryBuilder.prototype.buildHandleResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleResult = function (value) {
                ClientObject.prototype._handleResult.call(this, value);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                Utility.fixObjectPathIfNecessary(this, value);
                if (typeInfo.behaviorFlags & 8) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleResult");
                    customizationFunc.call(this, this, value);
                }
                if (typeInfo.scalarProperties) {
                    for (var i_3 = 0; i_3 < typeInfo.scalarProperties.length; i_3++) {
                        if (!Utility.isUndefined(value[typeInfo.scalarProperties[i_3].originalName])) {
                            if ((typeInfo.scalarProperties[i_3].behaviorFlags & 8) !== 0) {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_3])] = Utility.adjustToDateTime(value[typeInfo.scalarProperties[i_3].originalName]);
                            }
                            else {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_3])] = value[typeInfo.scalarProperties[i_3].originalName];
                            }
                        }
                    }
                }
                if (typeInfo.navigationProperties) {
                    var propNames = [];
                    for (var i_4 = 0; i_4 < typeInfo.navigationProperties.length; i_4++) {
                        propNames.push(typeInfo.navigationProperties[i_4].name);
                        propNames.push(typeInfo.navigationProperties[i_4].originalName);
                    }
                    Utility._handleNavigationPropertyResults(this, value, propNames);
                }
                if ((typeInfo.behaviorFlags & 1) !== 0) {
                    var hasIndexerMethod = thisBuilder.hasIndexMethod(typeInfo);
                    if (!Utility.isNullOrUndefined(value[Constants.items])) {
                        this.m__items = [];
                        var _data = value[Constants.items];
                        var childItemType = thisBuilder.getFunction(typeInfo.childItemTypeFullName);
                        for (var i = 0; i < _data.length; i++) {
                            var _item = BatchApiHelper.createChildItemObject(childItemType, hasIndexerMethod, this, _data[i], i);
                            _item._handleResult(_data[i]);
                            this.m__items.push(_item);
                        }
                    }
                }
            };
        };
        LibraryBuilder.prototype.buildHandleRetrieveResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleRetrieveResult = function (value, result) {
                ClientObject.prototype._handleRetrieveResult.call(this, value, result);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                if (typeInfo.scalarProperties) {
                    for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                        if (typeInfo.scalarProperties[i].behaviorFlags & 8) {
                            if (!Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].name])) {
                                value[typeInfo.scalarProperties[i].name] = Utility.adjustToDateTime(value[typeInfo.scalarProperties[i].name]);
                            }
                        }
                    }
                }
                if (typeInfo.behaviorFlags & 1) {
                    var hasIndexerMethod_1 = thisBuilder.hasIndexMethod(typeInfo);
                    var childItemType_1 = thisBuilder.getFunction(typeInfo.childItemTypeFullName);
                    var thisObj_1 = this;
                    Utility.processRetrieveResult(thisObj_1, value, result, function (childItemData, index) { return BatchApiHelper.createChildItemObject(childItemType_1, hasIndexerMethod_1, thisObj_1, childItemData, index); });
                }
                else {
                    Utility.processRetrieveResult(this, value, result);
                }
            };
        };
        LibraryBuilder.prototype.buildHandleIdResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleIdResult = function (value) {
                ClientObject.prototype._handleIdResult.call(this, value);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                if (typeInfo.behaviorFlags & 16) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleIdResult");
                    customizationFunc.call(this, this, value);
                }
                if (typeInfo.scalarProperties) {
                    for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                        var propName = typeInfo.scalarProperties[i].originalName;
                        if (propName === "Id" || propName === "_Id" || propName === "_ReferenceId") {
                            if (!Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].originalName])) {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])] = value[typeInfo.scalarProperties[i].originalName];
                            }
                        }
                    }
                }
            };
        };
        LibraryBuilder.prototype.buildLoad = function (type, typeInfo) {
            type.prototype.load = function (options) {
                return Utility.load(this, options);
            };
        };
        LibraryBuilder.prototype.buildRetrieve = function (type, typeInfo) {
            type.prototype.retrieve = function (options) {
                return Utility.retrieve(this, options);
            };
        };
        LibraryBuilder.prototype.buildNewObject = function (type, typeInfo) {
            if (!Utility.isNullOrEmptyString(typeInfo.newObjectServerTypeFullName)) {
                type.newObject = function (context) {
                    return BatchApiHelper.createTopLevelServiceObject(type, context, typeInfo.newObjectServerTypeFullName, (typeInfo.behaviorFlags & 1) !== 0, typeInfo.newObjectApiFlags);
                };
            }
        };
        LibraryBuilder.prototype.buildSetMockData = function (type, typeInfo) {
            var thisBuilder = this;
            if (typeInfo.behaviorFlags & 1) {
                var hasIndexMethod_1 = thisBuilder.hasIndexMethod(typeInfo);
                type.prototype.setMockData = function (data) {
                    var thisObj = this;
                    Utility.setMockData(thisObj, data, function (childItemData, index) {
                        return BatchApiHelper.createChildItemObject(thisBuilder.getFunction(typeInfo.childItemTypeFullName), hasIndexMethod_1, thisObj, childItemData, index);
                    }, function (items) {
                        thisObj.m__items = items;
                    });
                };
            }
            else {
                type.prototype.setMockData = function (data) {
                    Utility.setMockData(this, data);
                };
            }
        };
        LibraryBuilder.prototype.buildEnsureUnchanged = function (type, typeInfo) {
            type.prototype.ensureUnchanged = function (data) {
                BatchApiHelper.invokeEnsureUnchanged(this, data);
            };
        };
        LibraryBuilder.prototype.buildUpdate = function (type, typeInfo) {
            type.prototype.update = function (properties) {
                this._recursivelyUpdate(properties);
            };
        };
        LibraryBuilder.prototype.buildSet = function (type, typeInfo) {
            if ((typeInfo.behaviorFlags & 1) !== 0) {
                return;
            }
            var notAllowedToBeSetPropertyNames = [];
            var allowedScalarPropertyNames = [];
            if (typeInfo.scalarProperties) {
                for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                    if ((typeInfo.scalarProperties[i].behaviorFlags & 2) === 0 &&
                        (typeInfo.scalarProperties[i].behaviorFlags & 1) !== 0) {
                        allowedScalarPropertyNames.push(typeInfo.scalarProperties[i].name);
                    }
                    else {
                        notAllowedToBeSetPropertyNames.push(typeInfo.scalarProperties[i].name);
                    }
                }
            }
            var allowedNavigationPropertyNames = [];
            if (typeInfo.navigationProperties) {
                for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                    if ((typeInfo.navigationProperties[i].behaviorFlags & 16) !== 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else if ((typeInfo.navigationProperties[i].behaviorFlags & 1) === 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else if ((typeInfo.navigationProperties[i].behaviorFlags & 32) === 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else {
                        allowedNavigationPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                }
            }
            if (allowedNavigationPropertyNames.length === 0 && allowedScalarPropertyNames.length === 0) {
                return;
            }
            type.prototype.set = function (properties, options) {
                this._recursivelySet(properties, options, allowedScalarPropertyNames, allowedNavigationPropertyNames, notAllowedToBeSetPropertyNames);
            };
        };
        LibraryBuilder.prototype.buildItems = function (type, typeInfo) {
            if ((typeInfo.behaviorFlags & 1) === 0) {
                return;
            }
            Object.defineProperty(type.prototype, "items", {
                get: function () {
                    Utility.throwIfNotLoaded("items", this.m__items, typeInfo.name, this._isNull);
                    return this.m__items;
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.buildToJSON = function (type, typeInfo) {
            var thisBuilder = this;
            if ((typeInfo.behaviorFlags & 1) !== 0) {
                type.prototype.toJSON = function () {
                    return Utility.toJson(this, {}, {}, this.m__items);
                };
                return;
            }
            else {
                type.prototype.toJSON = function () {
                    var scalarProperties = {};
                    if (typeInfo.scalarProperties) {
                        for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                            if ((typeInfo.scalarProperties[i].behaviorFlags & 1) !== 0) {
                                scalarProperties[typeInfo.scalarProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])];
                            }
                        }
                    }
                    var navProperties = {};
                    if (typeInfo.navigationProperties) {
                        for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                            if ((typeInfo.navigationProperties[i].behaviorFlags & 1) !== 0) {
                                navProperties[typeInfo.navigationProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.navigationProperties[i])];
                            }
                        }
                    }
                    return Utility.toJson(this, scalarProperties, navProperties);
                };
            }
        };
        LibraryBuilder.prototype.buildTypeMetadataInfo = function (type, typeInfo) {
            Object.defineProperty(type.prototype, "_className", {
                get: function () {
                    return typeInfo.name;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(type.prototype, "_isCollection", {
                get: function () {
                    return (typeInfo.behaviorFlags & 1) !== 0;
                },
                enumerable: true,
                configurable: true
            });
            if (!Utility.isNullOrEmptyString(typeInfo.collectionPropertyPath)) {
                Object.defineProperty(type.prototype, "_collectionPropertyPath", {
                    get: function () {
                        return typeInfo.collectionPropertyPath;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
            if (typeInfo.scalarProperties && typeInfo.scalarProperties.length > 0) {
                Object.defineProperty(type.prototype, "_scalarPropertyNames", {
                    get: function () {
                        if (!this.m__scalarPropertyNames) {
                            this.m__scalarPropertyNames = typeInfo.scalarProperties.map(function (p) { return p.name; });
                        }
                        return this.m__scalarPropertyNames;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(type.prototype, "_scalarPropertyOriginalNames", {
                    get: function () {
                        if (!this.m__scalarPropertyOriginalNames) {
                            this.m__scalarPropertyOriginalNames = typeInfo.scalarProperties.map(function (p) { return p.originalName; });
                        }
                        return this.m__scalarPropertyOriginalNames;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(type.prototype, "_scalarPropertyUpdateable", {
                    get: function () {
                        if (!this.m__scalarPropertyUpdateable) {
                            this.m__scalarPropertyUpdateable = typeInfo.scalarProperties.map(function (p) { return (p.behaviorFlags & 2) === 0; });
                        }
                        return this.m__scalarPropertyUpdateable;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
            if (typeInfo.navigationProperties && typeInfo.navigationProperties.length > 0) {
                Object.defineProperty(type.prototype, "_navigationPropertyNames", {
                    get: function () {
                        if (!this.m__navigationPropertyNames) {
                            this.m__navigationPropertyNames = typeInfo.navigationProperties.map(function (p) { return p.name; });
                        }
                        return this.m__navigationPropertyNames;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
        };
        LibraryBuilder.prototype.buildTrackUntrack = function (type, typeInfo) {
            if (typeInfo.behaviorFlags & 2) {
                type.prototype.track = function () {
                    this.context.trackedObjects.add(this);
                    return this;
                };
                type.prototype.untrack = function () {
                    this.context.trackedObjects.remove(this);
                    return this;
                };
            }
        };
        LibraryBuilder.prototype.buildMixin = function (type, typeInfo) {
            if (typeInfo.behaviorFlags & 4) {
                var mixinType = this.getFunction(typeInfo.name + 'Custom');
                Utility.applyMixin(type, mixinType);
            }
        };
        LibraryBuilder.prototype.getOnEventName = function (name) {
            if (name[0] === '_') {
                return '_on' + name.substr(1);
            }
            return 'on' + name;
        };
        LibraryBuilder.prototype.buildEvents = function (type, typeInfo) {
            if (typeInfo.events) {
                for (var i = 0; i < typeInfo.events.length; i++) {
                    var elem = typeInfo.events[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 7);
                        typeInfo.events[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            apiSetInfoOrdinal: elem[2],
                            typeExpression: this.getString(elem[3]),
                            targetIdExpression: this.getString(elem[4]),
                            register: this.getString(elem[5]),
                            unregister: this.getString(elem[6])
                        };
                    }
                    this.buildEvent(type, typeInfo, typeInfo.events[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildEvent = function (type, typeInfo, evt) {
            if (evt.behaviorFlags & 1) {
                this.buildV0Event(type, typeInfo, evt);
            }
            else {
                this.buildV2Event(type, typeInfo, evt);
            }
        };
        LibraryBuilder.prototype.buildV2Event = function (type, typeInfo, evt) {
            var thisBuilder = this;
            var eventName = this.getOnEventName(evt.name);
            var fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function () {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this;
                        var registerFunc = null;
                        if (evt.register !== 'null') {
                            registerFunc = this[evt.register].bind(this);
                        }
                        var unregisterFunc = null;
                        if (evt.unregister !== 'null') {
                            unregisterFunc = this[evt.unregister].bind(this);
                        }
                        var getTargetIdFunc = function () {
                            return thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                        };
                        var func = null;
                        if (evt.behaviorFlags & 2) {
                            func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform");
                        }
                        var eventArgsTransformFunc = function (value) {
                            if (func) {
                                value = func.call(thisObj, thisObj, value);
                            }
                            return Utility._createPromiseFromResult(value);
                        };
                        var eventType = thisBuilder.evaluateEventType(evt.typeExpression);
                        this[fieldName] = new GenericEventHandlers(this.context, this, evt.name, {
                            eventType: eventType,
                            getTargetIdFunc: getTargetIdFunc,
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: eventArgsTransformFunc
                        });
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.buildV0Event = function (type, typeInfo, evt) {
            var thisBuilder = this;
            var eventName = this.getOnEventName(evt.name);
            var fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function () {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this;
                        var registerFunc = null;
                        if (Utility.isNullOrEmptyString(evt.register)) {
                            var eventType_1 = thisBuilder.evaluateEventType(evt.typeExpression);
                            registerFunc =
                                function (handlerCallback) {
                                    var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                    return thisObj.context.eventRegistration.register(eventType_1, targetId, handlerCallback);
                                };
                        }
                        else if (evt.register !== 'null') {
                            var func_1 = thisBuilder.getFunction(evt.register);
                            registerFunc =
                                function (handlerCallback) {
                                    return func_1.call(thisObj, thisObj, handlerCallback);
                                };
                        }
                        var unregisterFunc = null;
                        if (Utility.isNullOrEmptyString(evt.unregister)) {
                            var eventType_2 = thisBuilder.evaluateEventType(evt.typeExpression);
                            unregisterFunc =
                                function (handlerCallback) {
                                    var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                    return thisObj.context.eventRegistration.unregister(eventType_2, targetId, handlerCallback);
                                };
                        }
                        else if (evt.unregister !== 'null') {
                            var func_2 = thisBuilder.getFunction(evt.unregister);
                            unregisterFunc =
                                function (handlerCallback) {
                                    return func_2.call(thisObj, thisObj, handlerCallback);
                                };
                        }
                        var func = null;
                        if (evt.behaviorFlags & 2) {
                            func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform");
                        }
                        var eventArgsTransformFunc = function (value) {
                            if (func) {
                                value = func.call(thisObj, thisObj, value);
                            }
                            return Utility._createPromiseFromResult(value);
                        };
                        this[fieldName] = new EventHandlers(this.context, this, evt.name, {
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: eventArgsTransformFunc
                        });
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.hasIndexMethod = function (typeInfo) {
            var ret = false;
            if (typeInfo.navigationMethods) {
                for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                    if ((typeInfo.navigationMethods[i].behaviorFlags & 16) !== 0) {
                        ret = true;
                        break;
                    }
                }
            }
            return ret;
        };
        LibraryBuilder.CustomizationCodeNamespace = "_CC";
        return LibraryBuilder;
    }());
    OfficeExtension_1.LibraryBuilder = LibraryBuilder;
    var versionToken = 1;
    var internalConfiguration = {
        invokeRequestModifier: function (request) {
            request.DdaMethod.Version = versionToken;
            return request;
        },
        invokeResponseModifier: function (args) {
            versionToken = args.Version;
            if (args.Error) {
                args.error = {};
                args.error.Code = args.Error;
            }
            return args;
        }
    };
    var CommunicationConstants;
    (function (CommunicationConstants) {
        CommunicationConstants["SendingId"] = "sId";
        CommunicationConstants["RespondingId"] = "rId";
        CommunicationConstants["CommandKey"] = "command";
        CommunicationConstants["SessionInfoKey"] = "sessionInfo";
        CommunicationConstants["ParamsKey"] = "params";
        CommunicationConstants["ApiReadyCommand"] = "apiready";
        CommunicationConstants["ExecuteMethodCommand"] = "executeMethod";
        CommunicationConstants["GetAppContextCommand"] = "getAppContext";
        CommunicationConstants["RegisterEventCommand"] = "registerEvent";
        CommunicationConstants["UnregisterEventCommand"] = "unregisterEvent";
        CommunicationConstants["FireEventCommand"] = "fireEvent";
    })(CommunicationConstants || (CommunicationConstants = {}));
    var EmbeddedConstants = (function () {
        function EmbeddedConstants() {
        }
        EmbeddedConstants.sessionContext = 'sc';
        EmbeddedConstants.embeddingPageOrigin = 'EmbeddingPageOrigin';
        EmbeddedConstants.embeddingPageSessionInfo = 'EmbeddingPageSessionInfo';
        return EmbeddedConstants;
    }());
    OfficeExtension_1.EmbeddedConstants = EmbeddedConstants;
    var EmbeddedSession = (function (_super) {
        __extends(EmbeddedSession, _super);
        function EmbeddedSession(url, options) {
            var _this = _super.call(this) || this;
            _this.m_chosenWindow = null;
            _this.m_chosenOrigin = null;
            _this.m_enabled = true;
            _this.m_onMessageHandler = _this._onMessage.bind(_this);
            _this.m_callbackList = {};
            _this.m_id = 0;
            _this.m_timeoutId = -1;
            _this.m_appContext = null;
            _this.m_url = url;
            _this.m_options = options;
            if (!_this.m_options) {
                _this.m_options = { sessionKey: Math.random().toString() };
            }
            if (!_this.m_options.sessionKey) {
                _this.m_options.sessionKey = Math.random().toString();
            }
            if (!_this.m_options.container) {
                _this.m_options.container = document.body;
            }
            if (!_this.m_options.timeoutInMilliseconds) {
                _this.m_options.timeoutInMilliseconds = 60000;
            }
            if (!_this.m_options.height) {
                _this.m_options.height = '400px';
            }
            if (!_this.m_options.width) {
                _this.m_options.width = '100%';
            }
            if (!(_this.m_options.webApplication &&
                _this.m_options.webApplication.accessToken &&
                _this.m_options.webApplication.accessTokenTtl)) {
                _this.m_options.webApplication = null;
            }
            return _this;
        }
        EmbeddedSession.prototype._getIFrameSrc = function () {
            var origin = window.location.protocol + '//' + window.location.host;
            var toAppend = EmbeddedConstants.embeddingPageOrigin +
                '=' +
                encodeURIComponent(origin) +
                '&' +
                EmbeddedConstants.embeddingPageSessionInfo +
                '=' +
                encodeURIComponent(this.m_options.sessionKey);
            var useHash = false;
            if (this.m_url.toLowerCase().indexOf('/_layouts/preauth.aspx') > 0 ||
                this.m_url.toLowerCase().indexOf('/_layouts/15/preauth.aspx') > 0) {
                useHash = true;
            }
            var a = document.createElement('a');
            a.href = this.m_url;
            if (this.m_options.webApplication) {
                var toAppendWAC = EmbeddedConstants.embeddingPageOrigin +
                    '=' +
                    origin +
                    '&' +
                    EmbeddedConstants.embeddingPageSessionInfo +
                    '=' +
                    this.m_options.sessionKey;
                if (a.search.length === 0 || a.search === '?') {
                    a.search = '?' + EmbeddedConstants.sessionContext + '=' + encodeURIComponent(toAppendWAC);
                }
                else {
                    a.search = a.search + '&' + EmbeddedConstants.sessionContext + '=' + encodeURIComponent(toAppendWAC);
                }
            }
            else if (useHash) {
                if (a.hash.length === 0 || a.hash === '#') {
                    a.hash = '#' + toAppend;
                }
                else {
                    a.hash = a.hash + '&' + toAppend;
                }
            }
            else {
                if (a.search.length === 0 || a.search === '?') {
                    a.search = '?' + toAppend;
                }
                else {
                    a.search = a.search + '&' + toAppend;
                }
            }
            var iframeSrc = a.href;
            return iframeSrc;
        };
        EmbeddedSession.prototype.init = function () {
            var _this = this;
            window.addEventListener('message', this.m_onMessageHandler);
            var iframeSrc = this._getIFrameSrc();
            return CoreUtility.createPromise(function (resolve, reject) {
                var iframeElement = document.createElement('iframe');
                if (_this.m_options.id) {
                    iframeElement.id = _this.m_options.id;
                    iframeElement.name = _this.m_options.id;
                }
                iframeElement.style.height = _this.m_options.height;
                iframeElement.style.width = _this.m_options.width;
                if (!_this.m_options.webApplication) {
                    iframeElement.src = iframeSrc;
                    _this.m_options.container.appendChild(iframeElement);
                }
                else {
                    var webApplicationForm = document.createElement('form');
                    webApplicationForm.setAttribute('action', iframeSrc);
                    webApplicationForm.setAttribute('method', 'post');
                    webApplicationForm.setAttribute('target', iframeElement.name);
                    _this.m_options.container.appendChild(webApplicationForm);
                    var token_input = document.createElement('input');
                    token_input.setAttribute('type', 'hidden');
                    token_input.setAttribute('name', 'access_token');
                    token_input.setAttribute('value', _this.m_options.webApplication.accessToken);
                    webApplicationForm.appendChild(token_input);
                    var token_ttl_input = document.createElement('input');
                    token_ttl_input.setAttribute('type', 'hidden');
                    token_ttl_input.setAttribute('name', 'access_token_ttl');
                    token_ttl_input.setAttribute('value', _this.m_options.webApplication.accessTokenTtl);
                    webApplicationForm.appendChild(token_ttl_input);
                    _this.m_options.container.appendChild(iframeElement);
                    webApplicationForm.submit();
                }
                _this.m_timeoutId = window.setTimeout(function () {
                    _this.close();
                    var err = Utility.createRuntimeError(CoreErrorCodes.timeout, CoreUtility._getResourceString(CoreResourceStrings.timeout), 'EmbeddedSession.init');
                    reject(err);
                }, _this.m_options.timeoutInMilliseconds);
                _this.m_promiseResolver = resolve;
            });
        };
        EmbeddedSession.prototype._invoke = function (method, callback, params) {
            if (!this.m_enabled) {
                callback(5001, null);
                return;
            }
            if (internalConfiguration.invokeRequestModifier) {
                params = internalConfiguration.invokeRequestModifier(params);
            }
            this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
                if (internalConfiguration.invokeResponseModifier) {
                    args = internalConfiguration.invokeResponseModifier(args);
                }
                var errorCode = args['Error'];
                delete args['Error'];
                callback(errorCode || 0, args);
            });
        };
        EmbeddedSession.prototype.close = function () {
            window.removeEventListener('message', this.m_onMessageHandler);
            window.clearTimeout(this.m_timeoutId);
            this.m_enabled = false;
        };
        EmbeddedSession.prototype.getEventRegistration = function (controlId) {
            if (!this.m_sessionEventManager) {
                this.m_sessionEventManager = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
            }
            return this.m_sessionEventManager;
        };
        EmbeddedSession.prototype._createRequestExecutorOrNull = function () {
            return new EmbeddedRequestExecutor(this);
        };
        EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        EmbeddedSession.prototype._registerEventImpl = function (eventId, targetId) {
            var _this = this;
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve(null);
                });
            });
        };
        EmbeddedSession.prototype._unregisterEventImpl = function (eventId, targetId) {
            var _this = this;
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve();
                });
            });
        };
        EmbeddedSession.prototype._onMessage = function (event) {
            var _this = this;
            if (!this.m_enabled) {
                return;
            }
            if (this.m_chosenWindow && (this.m_chosenWindow !== event.source || this.m_chosenOrigin !== event.origin)) {
                return;
            }
            var eventData = event.data;
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.ApiReadyCommand) {
                if (!this.m_chosenWindow &&
                    this._isValidDescendant(event.source) &&
                    eventData[CommunicationConstants.SessionInfoKey] === this.m_options.sessionKey) {
                    this.m_chosenWindow = event.source;
                    this.m_chosenOrigin = event.origin;
                    this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
                        _this._setupContext(appContext);
                        window.clearTimeout(_this.m_timeoutId);
                        _this.m_promiseResolver();
                    });
                }
                return;
            }
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.FireEventCommand) {
                var msg = eventData[CommunicationConstants.ParamsKey];
                var eventId = msg['EventId'];
                var targetId = msg['TargetId'];
                var data = msg['Data'];
                if (this.m_sessionEventManager) {
                    var handlers = this.m_sessionEventManager.getHandlers(eventId, targetId);
                    for (var i = 0; i < handlers.length; i++) {
                        handlers[i](data);
                    }
                }
                return;
            }
            if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
                var rId = eventData[CommunicationConstants.RespondingId];
                if (this.m_callbackList.hasOwnProperty(rId)) {
                    var callback = this.m_callbackList[rId];
                    if (typeof callback === 'function') {
                        callback(eventData[CommunicationConstants.ParamsKey]);
                    }
                    delete this.m_callbackList[rId];
                }
            }
        };
        EmbeddedSession.prototype._sendMessageWithCallback = function (id, command, data, callback) {
            this.m_callbackList[id] = callback;
            var message = {};
            message[CommunicationConstants.SendingId] = id;
            message[CommunicationConstants.CommandKey] = command;
            message[CommunicationConstants.ParamsKey] = data;
            this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
        };
        EmbeddedSession.prototype._isValidDescendant = function (wnd) {
            var container = this.m_options.container || document.body;
            function doesFrameWindow(containerWindow) {
                if (containerWindow === wnd) {
                    return true;
                }
                for (var i = 0, len = containerWindow.frames.length; i < len; i++) {
                    if (doesFrameWindow(containerWindow.frames[i])) {
                        return true;
                    }
                }
                return false;
            }
            var iframes = container.getElementsByTagName('iframe');
            for (var i = 0, len = iframes.length; i < len; i++) {
                if (doesFrameWindow(iframes[i].contentWindow)) {
                    return true;
                }
            }
            return false;
        };
        EmbeddedSession.prototype._setupContext = function (appContext) {
            if (!(this.m_appContext = appContext)) {
                return;
            }
        };
        return EmbeddedSession;
    }(SessionBase));
    OfficeExtension_1.EmbeddedSession = EmbeddedSession;
    var EmbeddedRequestExecutor = (function () {
        function EmbeddedRequestExecutor(session) {
            this.m_session = session;
        }
        EmbeddedRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
            return CoreUtility.createPromise(function (resolve, reject) {
                _this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(result));
                    var response;
                    if (status == 0) {
                        response = RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
                    }
                    else {
                        response = RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
                    }
                    resolve(response);
                }, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
            });
        };
        EmbeddedRequestExecutor._transformMessageArrayIntoParams = function (msgArray) {
            return {
                ArrayData: msgArray,
                DdaMethod: {
                    DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
                }
            };
        };
        EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod = 93;
        EmbeddedRequestExecutor.SourceLibHeaderValue = 'Embedded';
        return EmbeddedRequestExecutor;
    }());
})(OfficeExtension || (OfficeExtension = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b)
                if (b.hasOwnProperty(p))
                    d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try {
            step(generator.next(value));
        }
        catch (e) {
            reject(e);
        } }
        function rejected(value) { try {
            step(generator["throw"](value));
        }
        catch (e) {
            reject(e);
        } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function () { if (t[0] & 1)
            throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function () { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f)
            throw new TypeError("Generator is already executing.");
        while (_)
            try {
                if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done)
                    return t;
                if (y = 0, t)
                    op = [op[0] & 2, t.value];
                switch (op[0]) {
                    case 0:
                    case 1:
                        t = op;
                        break;
                    case 4:
                        _.label++;
                        return { value: op[1], done: false };
                    case 5:
                        _.label++;
                        y = op[1];
                        op = [0];
                        continue;
                    case 7:
                        op = _.ops.pop();
                        _.trys.pop();
                        continue;
                    default:
                        if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
                            _ = 0;
                            continue;
                        }
                        if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) {
                            _.label = op[1];
                            break;
                        }
                        if (op[0] === 6 && _.label < t[1]) {
                            _.label = t[1];
                            t = op;
                            break;
                        }
                        if (t && _.label < t[2]) {
                            _.label = t[2];
                            _.ops.push(op);
                            break;
                        }
                        if (t[2])
                            _.ops.pop();
                        _.trys.pop();
                        continue;
                }
                op = body.call(thisArg, _);
            }
            catch (e) {
                op = [6, e];
                y = 0;
            }
            finally {
                f = t = 0;
            }
        if (op[0] & 5)
            throw op[1];
        return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "AgaveVisualApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _typeBiShim = "BiShim";
    var BiShim = (function (_super) {
        __extends(BiShim, _super);
        function BiShim() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(BiShim.prototype, "_className", {
            get: function () {
                return "BiShim";
            },
            enumerable: true,
            configurable: true
        });
        BiShim.prototype.initialize = function (capabilities) {
            _invokeMethod(this, "Initialize", 0, [capabilities], 0, 0);
        };
        BiShim.prototype.getData = function () {
            return _invokeMethod(this, "getData", 1, [], 4, 0);
        };
        BiShim.prototype.setVisualObjects = function (visualObjects) {
            _invokeMethod(this, "setVisualObjects", 0, [visualObjects], 2, 0);
        };
        BiShim.prototype.setVisualObjectsToPersist = function (visualObjectsToPersist) {
            _invokeMethod(this, "setVisualObjectsToPersist", 0, [visualObjectsToPersist], 2, 0);
        };
        BiShim.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        BiShim.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        BiShim.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.BiShim, context, "Microsoft.AgaveVisual.BiShim", false, 4);
        };
        BiShim.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return BiShim;
    }(OfficeExtension.ClientObject));
    OfficeCore.BiShim = BiShim;
    var AgaveVisualErrorCodes;
    (function (AgaveVisualErrorCodes) {
        AgaveVisualErrorCodes["generalException1"] = "GeneralException";
    })(AgaveVisualErrorCodes = OfficeCore.AgaveVisualErrorCodes || (OfficeCore.AgaveVisualErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "ExperimentApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _typeFlightingService = "FlightingService";
    var FlightingService = (function (_super) {
        __extends(FlightingService, _super);
        function FlightingService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(FlightingService.prototype, "_className", {
            get: function () {
                return "FlightingService";
            },
            enumerable: true,
            configurable: true
        });
        FlightingService.prototype.getClientSessionId = function () {
            return _invokeMethod(this, "GetClientSessionId", 1, [], 4, 0);
        };
        FlightingService.prototype.getDeferredFlights = function () {
            return _invokeMethod(this, "GetDeferredFlights", 1, [], 4, 0);
        };
        FlightingService.prototype.getFeature = function (featureName, type, defaultValue, possibleValues) {
            return _createMethodObject(OfficeCore.ABType, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null, 4);
        };
        FlightingService.prototype.getFeatureGate = function (featureName, scope) {
            return _createMethodObject(OfficeCore.ABType, this, "GetFeatureGate", 1, [featureName, scope], false, false, null, 4);
        };
        FlightingService.prototype.resetOverride = function (featureName) {
            _invokeMethod(this, "ResetOverride", 0, [featureName], 0, 0);
        };
        FlightingService.prototype.setOverride = function (featureName, type, value) {
            _invokeMethod(this, "SetOverride", 0, [featureName, type, value], 0, 0);
        };
        FlightingService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        FlightingService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        FlightingService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.FlightingService, context, "Microsoft.Experiment.FlightingService", false, 4);
        };
        FlightingService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return FlightingService;
    }(OfficeExtension.ClientObject));
    OfficeCore.FlightingService = FlightingService;
    var _typeABType = "ABType";
    var ABType = (function (_super) {
        __extends(ABType, _super);
        function ABType() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ABType.prototype, "_className", {
            get: function () {
                return "ABType";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ABType.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["value"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ABType.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this._V, _typeABType, this._isNull);
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        ABType.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Value"])) {
                this._V = obj["Value"];
            }
        };
        ABType.prototype.load = function (option) {
            return _load(this, option);
        };
        ABType.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ABType.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ABType.prototype.toJSON = function () {
            return _toJson(this, {
                "value": this._V
            }, {});
        };
        ABType.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return ABType;
    }(OfficeExtension.ClientObject));
    OfficeCore.ABType = ABType;
    var FeatureType;
    (function (FeatureType) {
        FeatureType["boolean"] = "Boolean";
        FeatureType["integer"] = "Integer";
        FeatureType["string"] = "String";
    })(FeatureType = OfficeCore.FeatureType || (OfficeCore.FeatureType = {}));
    var ExperimentErrorCodes;
    (function (ExperimentErrorCodes) {
        ExperimentErrorCodes["generalException"] = "GeneralException";
    })(ExperimentErrorCodes = OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeFirstPartyDialog;
(function (OfficeFirstPartyDialog) {
    var Dialog = (function () {
        function Dialog(_dialogService) {
            this._dialogService = _dialogService;
        }
        Dialog.prototype.close = function () {
            this._dialogService.close();
            return this._dialogService.context.sync();
        };
        Dialog.prototype.messageChild = function (message, options) {
            if (DialogApiManager && DialogApiManager.messageChildRichApiBridge) {
                DialogApiManager.messageChildRichApiBridge(message, options);
            }
        };
        return Dialog;
    }());
    OfficeFirstPartyDialog.Dialog = Dialog;
    function lookupErrorCodeAndMessage(internalCode) {
        var _a;
        var table = (_a = {},
            _a[12002] = { code: "InvalidUrl", message: "Cannot load URL, no such page or bad URL syntax." },
            _a[12003] = { code: "InvalidUrl", message: "HTTPS is required." },
            _a[12004] = { code: "Untrusted", message: "Domain is not trusted." },
            _a[12005] = { code: "InvalidUrl", message: "HTTPS is required." },
            _a[12007] = { code: "FailedToOpen", message: "Another dialog is already opened." },
            _a);
        if (table[internalCode]) {
            return table[internalCode];
        }
        else {
            return { code: "Unknown", message: "An unknown error has occured with code: " + internalCode };
        }
    }
    function displayWebDialog(url, options) {
        if (options === void 0) {
            options = {};
        }
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (options.width && options.height && (!isInt(options.width) || !isInt(options.height))) {
                throw new OfficeExtension.Error({ code: "InvalidArgument", message: 'Dimensions must be "number%" or number.' });
            }
            var ctx = new OfficeExtension.ClientRequestContext();
            var dialogService = OfficeCore.FirstPartyDialogService.newObject(ctx);
            var dialog = new OfficeFirstPartyDialog.Dialog(dialogService);
            var eventResult = dialogService.onDialogMessage.add(function (args) {
                OfficeExtension.Utility.log("dialogMessageHandler:" + JSON.stringify(args));
                var parsedMessage = JSON.parse(args.message);
                var eventError = parsedMessage.errorCode
                    ? new OfficeExtension.Error(lookupErrorCodeAndMessage(parsedMessage.errorCode))
                    : null;
                var messageType = parsedMessage.type;
                var origin = parsedMessage.origin;
                var messageContent = parsedMessage.message;
                switch (messageType) {
                    case 17:
                        if (eventError) {
                            reject(eventError);
                        }
                        else {
                            resolve(dialog);
                        }
                        break;
                    case 12:
                        if (options.onMessage) {
                            options.onMessage(messageContent, dialog, origin);
                        }
                        break;
                    case 10:
                    default:
                        if (parsedMessage.errorCode === 12006) {
                            if (eventResult) {
                                eventResult.remove();
                                ctx.sync();
                            }
                            if (options.onClose) {
                                options.onClose();
                            }
                        }
                        else {
                            if (options.onRuntimeError) {
                                options.onRuntimeError(eventError, dialog);
                                reject(eventError);
                            }
                        }
                }
                return OfficeExtension.CoreUtility.Promise.resolve();
            });
            return ctx.sync()
                .then(function () {
                var dialogOptions = {
                    width: options.width ? parseInt(options.width) : 50,
                    height: options.height ? parseInt(options.height) : 50,
                    displayInIFrame: options.displayInIFrame,
                    dialogTitle: options.dialogTitle
                };
                dialogService.displayDialog(url, dialogOptions);
                return ctx.sync();
            })["catch"](function (e) {
                reject(e);
            });
        });
        function isInt(value) {
            return (/^(\-|\+)?([0-9]+)%?$/.test(value));
        }
    }
    OfficeFirstPartyDialog.displayWebDialog = displayWebDialog;
    var DialogEventType;
    (function (DialogEventType) {
        DialogEventType[DialogEventType["dialogMessageReceived"] = 0] = "dialogMessageReceived";
        DialogEventType[DialogEventType["dialogEventReceived"] = 1] = "dialogEventReceived";
    })(DialogEventType || (DialogEventType = {}));
})(OfficeFirstPartyDialog || (OfficeFirstPartyDialog = {}));
var OfficeCore;
(function (OfficeCore) {
    OfficeCore.OfficeOnlineDomainList = [
        "*.dod.online.office365.us",
        "*.gov.online.office365.us",
        "*.officeapps-df.live.com",
        "*.officeapps.live.com",
        "*.online.office.de",
        "*.partner.officewebapps.cn"
    ];
    function isHostOriginTrusted() {
        if (typeof window.external === 'undefined' ||
            typeof window.external.GetContext === 'undefined') {
            var hostUrl = OSF.getClientEndPoint()._targetUrl;
            var hostname_1 = getHostNameFromUrl(hostUrl);
            if (hostUrl.indexOf("https:") != 0) {
                return false;
            }
            OfficeCore.OfficeOnlineDomainList.forEach(function (domain) {
                if (domain.indexOf("*.") == 0) {
                    domain = domain.substring(2);
                }
                if (hostname_1.indexOf(domain) == hostname_1.length - domain.length) {
                    return true;
                }
            });
            return false;
        }
        return true;
    }
    OfficeCore.isHostOriginTrusted = isHostOriginTrusted;
    function getHostNameFromUrl(url) {
        var hostName = "";
        hostName = url.split("/")[2];
        hostName = hostName.split(":")[0];
        hostName = hostName.split("?")[0];
        return hostName;
    }
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var FirstPartyApis = (function () {
        function FirstPartyApis(context) {
            this.context = context;
        }
        Object.defineProperty(FirstPartyApis.prototype, "roamingSettings", {
            get: function () {
                if (!this.m_roamingSettings) {
                    this.m_roamingSettings = OfficeCore.AuthenticationService.newObject(this.context).roamingSettings;
                }
                return this.m_roamingSettings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FirstPartyApis.prototype, "tap", {
            get: function () {
                if (!this.m_tap) {
                    this.m_tap = OfficeCore.Tap.newObject(this.context);
                }
                return this.m_tap;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FirstPartyApis.prototype, "skill", {
            get: function () {
                if (!this.m_skill) {
                    this.m_skill = OfficeCore.Skill.newObject(this.context);
                }
                return this.m_skill;
            },
            enumerable: true,
            configurable: true
        });
        return FirstPartyApis;
    }());
    OfficeCore.FirstPartyApis = FirstPartyApis;
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            return _super.call(this, url) || this;
        }
        Object.defineProperty(RequestContext.prototype, "firstParty", {
            get: function () {
                if (!this.m_firstPartyApis) {
                    this.m_firstPartyApis = new FirstPartyApis(this);
                }
                return this.m_firstPartyApis;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "flighting", {
            get: function () {
                return this.flightingService;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "telemetry", {
            get: function () {
                if (!this.m_telemetry) {
                    this.m_telemetry = OfficeCore.TelemetryService.newObject(this);
                }
                return this.m_telemetry;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "ribbon", {
            get: function () {
                if (!this.m_ribbon) {
                    this.m_ribbon = OfficeCore.DynamicRibbon.newObject(this);
                }
                return this.m_ribbon;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "bi", {
            get: function () {
                if (!this.m_biShim) {
                    this.m_biShim = OfficeCore.BiShim.newObject(this);
                }
                return this.m_biShim;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "flightingService", {
            get: function () {
                if (!this.m_flightingService) {
                    this.m_flightingService = OfficeCore.FlightingService.newObject(this);
                }
                return this.m_flightingService;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeExtension.ClientRequestContext));
    OfficeCore.RequestContext = RequestContext;
    function run(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runBatch("OfficeCore.run", arguments, function (requestInfo) { return new OfficeCore.RequestContext(requestInfo); });
    }
    OfficeCore.run = run;
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    var license;
    (function (license_1) {
        function _createRequestContext() {
            var context = new OfficeCore.RequestContext();
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function isFeatureEnabled(feature, fallbackValue) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, isEnabled;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            isEnabled = license.isFeatureEnabled(feature, fallbackValue);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, isEnabled.value];
                    }
                });
            });
        }
        license_1.isFeatureEnabled = isFeatureEnabled;
        function getFeatureTier(feature, fallbackValue) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, tier;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            tier = license.getFeatureTier(feature, fallbackValue);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, tier.value];
                    }
                });
            });
        }
        license_1.getFeatureTier = getFeatureTier;
        function isFreemiumUpsellEnabled() {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, isFreemiumUpsellEnabled;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            isFreemiumUpsellEnabled = license.isFreemiumUpsellEnabled();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, isFreemiumUpsellEnabled.value];
                    }
                });
            });
        }
        license_1.isFreemiumUpsellEnabled = isFreemiumUpsellEnabled;
        function launchUpsellExperience(experienceId) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            license.launchUpsellExperience(experienceId);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        license_1.launchUpsellExperience = launchUpsellExperience;
        function onFeatureStateChanged(feature, listener) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, licenseFeature, removeListener;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            licenseFeature = license.getLicenseFeature(feature);
                            licenseFeature.onStateChanged.add(listener);
                            removeListener = function () {
                                licenseFeature.onStateChanged.remove(listener);
                                return null;
                            };
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, removeListener];
                    }
                });
            });
        }
        license_1.onFeatureStateChanged = onFeatureStateChanged;
        function getMsaDeviceTicket(resource, policy, options) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, msaDeviceTicket;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            msaDeviceTicket = license.getMsaDeviceTicket(resource, policy, options);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, msaDeviceTicket.value];
                    }
                });
            });
        }
        license_1.getMsaDeviceTicket = getMsaDeviceTicket;
    })(license = Office.license || (Office.license = {}));
})(Office || (Office = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "Office";
    var _defaultApiSetName = "OfficeSharedApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var _typeSkill = "Skill";
    var Skill = (function (_super) {
        __extends(Skill, _super);
        function Skill() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Skill.prototype, "_className", {
            get: function () {
                return "Skill";
            },
            enumerable: true,
            configurable: true
        });
        Skill.prototype.executeAction = function (paneId, actionId, actionDescriptor) {
            return _invokeMethod(this, "ExecuteAction", 1, [paneId, actionId, actionDescriptor], 4 | 1, 0);
        };
        Skill.prototype.notifyPaneEvent = function (paneId, eventDescriptor) {
            _invokeMethod(this, "NotifyPaneEvent", 1, [paneId, eventDescriptor], 4 | 1, 0);
        };
        Skill.prototype.registerHostSkillEvent = function () {
            _invokeMethod(this, "RegisterHostSkillEvent", 0, [], 1, 0);
        };
        Skill.prototype.testFireEvent = function () {
            _invokeMethod(this, "TestFireEvent", 0, [], 1, 0);
        };
        Skill.prototype.unregisterHostSkillEvent = function () {
            _invokeMethod(this, "UnregisterHostSkillEvent", 0, [], 1, 0);
        };
        Skill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        Skill.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Skill.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.Skill, context, "Microsoft.SkillApi.Skill", false, 4);
        };
        Object.defineProperty(Skill.prototype, "onHostSkillEvent", {
            get: function () {
                var _this = this;
                if (!this.m_hostSkillEvent) {
                    this.m_hostSkillEvent = new OfficeExtension.GenericEventHandlers(this.context, this, "HostSkillEvent", {
                        eventType: 65538,
                        registerFunc: function () { return _this.registerHostSkillEvent(); },
                        unregisterFunc: function () { return _this.unregisterHostSkillEvent(); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var event = _CC.Skill_HostSkillEvent_EventArgsTransform(_this, value);
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_hostSkillEvent;
            },
            enumerable: true,
            configurable: true
        });
        Skill.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return Skill;
    }(OfficeExtension.ClientObject));
    OfficeCore.Skill = Skill;
    var _CC;
    (function (_CC) {
        function Skill_HostSkillEvent_EventArgsTransform(thisObj, args) {
            var transformedArgs = {
                type: args.type,
                data: args.data
            };
            return transformedArgs;
        }
        _CC.Skill_HostSkillEvent_EventArgsTransform = Skill_HostSkillEvent_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var SkillErrorCodes;
    (function (SkillErrorCodes) {
        SkillErrorCodes["generalException"] = "GeneralException";
    })(SkillErrorCodes = OfficeCore.SkillErrorCodes || (OfficeCore.SkillErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "TelemetryApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var _typeTelemetryService = "TelemetryService";
    var TelemetryService = (function (_super) {
        __extends(TelemetryService, _super);
        function TelemetryService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TelemetryService.prototype, "_className", {
            get: function () {
                return "TelemetryService";
            },
            enumerable: true,
            configurable: true
        });
        TelemetryService.prototype.sendCustomerContent = function (telemetryProperties, eventName, eventContract, eventFlags, value) {
            _throwIfApiNotSupported("TelemetryService.sendCustomerContent", "Telemetry", "1.3", _hostName);
            _invokeMethod(this, "SendCustomerContent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], 4, 0);
        };
        TelemetryService.prototype.sendTelemetryEvent = function (telemetryProperties, eventName, eventContract, eventFlags, value) {
            _invokeMethod(this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], 4, 0);
        };
        TelemetryService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        TelemetryService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        TelemetryService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.TelemetryService, context, "Microsoft.Telemetry.TelemetryService", false, 4);
        };
        TelemetryService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return TelemetryService;
    }(OfficeExtension.ClientObject));
    OfficeCore.TelemetryService = TelemetryService;
    var DataFieldType;
    (function (DataFieldType) {
        DataFieldType["unset"] = "Unset";
        DataFieldType["string"] = "String";
        DataFieldType["boolean"] = "Boolean";
        DataFieldType["int64"] = "Int64";
        DataFieldType["double"] = "Double";
    })(DataFieldType = OfficeCore.DataFieldType || (OfficeCore.DataFieldType = {}));
    var TelemetryErrorCodes;
    (function (TelemetryErrorCodes) {
        TelemetryErrorCodes["generalException"] = "GeneralException";
    })(TelemetryErrorCodes = OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeFirstPartyAuth;
(function (OfficeFirstPartyAuth) {
    var WebAuthReplyUrlsStorageKey = "officeWebAuthReplyUrls";
    var loaded = false;
    OfficeFirstPartyAuth.authFlow = "authcode";
    OfficeFirstPartyAuth.autoPopup = false;
    OfficeFirstPartyAuth.upnCheck = true;
    OfficeFirstPartyAuth.msal = "https://alcdn.msauth.net/browser-1p/2.28.1/js/msal-browser-1p.min.js";
    OfficeFirstPartyAuth.debugging = false;
    OfficeFirstPartyAuth.delay = 0;
    OfficeFirstPartyAuth.delayMsal = 0;
    function load(replyUrl, prefetch) {
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF.WebAuth && OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                var retrievedAuthContext = false;
                try {
                    if (!Office || !Office.context || !Office.context.webAuth) {
                        reject({
                            code: "GetAuthContextAsyncMissing",
                            message: "Office:[" + !Office + "],Office.context:[" + !Office.context + "],Office.context.webAuth:[" + !Office.context.webAuth + "]"
                        });
                        return;
                    }
                    Office.context.webAuth.getAuthContextAsync(function (result) {
                        if (result.status === "succeeded") {
                            retrievedAuthContext = true;
                            var authContext = result.value;
                            if (!authContext || authContext.isAnonymous) {
                                reject({
                                    code: "CannotGetAuthContext",
                                    message: authContext.Error
                                });
                                return;
                            }
                            var isMsa = authContext.authorityType.toLowerCase() === 'msa';
                            OSF.WebAuth.config = {
                                authFlow: OfficeFirstPartyAuth.authFlow,
                                authVersion: (OfficeFirstPartyAuth.authVersion) ? OfficeFirstPartyAuth.authVersion : null,
                                msal: OfficeFirstPartyAuth.msal,
                                delayWebAuth: OfficeFirstPartyAuth.delay,
                                delayMsal: OfficeFirstPartyAuth.delayMsal,
                                debugging: OfficeFirstPartyAuth.debugging,
                                authority: (OfficeFirstPartyAuth.authorityOverride) ? OfficeFirstPartyAuth.authorityOverride : authContext.authority,
                                idp: authContext.authorityType.toLowerCase(),
                                appIds: [isMsa ? (authContext.msaAppId) ? authContext.msaAppId : authContext.appId : authContext.appId],
                                redirectUri: (replyUrl) ? replyUrl : null,
                                upn: authContext.upn,
                                puid: authContext.userId,
                                prefetch: prefetch,
                                telemetryInstance: 'otel',
                                autoPopup: OfficeFirstPartyAuth.autoPopup,
                                enableUpnCheck: OfficeFirstPartyAuth.upnCheck,
                                enableConsoleLogging: OfficeFirstPartyAuth.debugging
                            };
                            OSF.WebAuth.load().then(function (result) {
                                loaded = true;
                                logLoadEvent(result, loaded);
                                resolve();
                            })["catch"](function (result) {
                                logLoadEvent(result, loaded);
                                reject({
                                    code: "PackageNotLoaded",
                                    message: (result instanceof Event) ? result.type : result
                                });
                            });
                            if (OfficeFirstPartyAuth.authFlow === "implicit") {
                                var finalReplyUrl = (replyUrl) ? replyUrl : window.location.href.split("?")[0];
                                var replyUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
                                if (replyUrls || replyUrls === "") {
                                    replyUrls = finalReplyUrl;
                                }
                                else {
                                    replyUrls += ", " + finalReplyUrl;
                                }
                                if (replyUrls)
                                    sessionStorage.setItem(WebAuthReplyUrlsStorageKey, replyUrls);
                            }
                        }
                        else {
                            OSF.WebAuth.config = null;
                            reject({
                                code: "CannotGetAuthContext",
                                message: result.status
                            });
                        }
                    });
                }
                catch (e) {
                    OSF.WebAuth.config = null;
                    OSF.WebAuth.load().then(function () {
                        resolve();
                    })["catch"](function () {
                        reject({
                            code: retrievedAuthContext ? "CannotGetAuthContext" : "FailedToLoad",
                            message: e
                        });
                    });
                }
            }
            else {
                resolve();
            }
        });
    }
    OfficeFirstPartyAuth.load = load;
    function getAccessToken(options, behaviorOption) {
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                Office.context.webAuth.getAuthContextAsync(function (result) {
                    var supportsAuthToken = false;
                    if (result.status === "succeeded") {
                        var authContext = result.value;
                        if (authContext.supportsAuthToken) {
                            supportsAuthToken = true;
                        }
                    }
                    if (!supportsAuthToken) {
                        if (OSF.WebAuth && loaded) {
                            if (OSF.WebAuth.config.appIds[0]) {
                                OSF.WebAuth.getToken(options.resource, OSF.WebAuth.config.appIds[0], OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null, (options && options.authChallenge) ? options.authChallenge : null).then(function (result) {
                                    logAcquireEvent(result, true, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : false);
                                    resolve({
                                        accessToken: result.Token,
                                        tokenIdenityType: (OSF.WebAuth.config.idp.toLowerCase() == "msa")
                                            ? OfficeCore.IdentityType.microsoftAccount
                                            : OfficeCore.IdentityType.organizationAccount
                                    });
                                })["catch"](function (result) {
                                    logAcquireEvent(result, false, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : false);
                                    reject({
                                        code: result.ErrorCode,
                                        message: (result instanceof Event) ? result.type : result.ErrorMessage
                                    });
                                });
                            }
                        }
                        else {
                            logUnexpectedAcquire(loaded, OSF.WebAuth.loadAttempts);
                        }
                    }
                    else {
                        var context = new OfficeCore.RequestContext();
                        var auth = OfficeCore.AuthenticationService.newObject(context);
                        context._customData = "WacPartition";
                        var result_1 = auth.getAccessToken(options, null);
                        context.sync().then(function () {
                            resolve(result_1.value);
                        });
                    }
                });
            }
            else {
                var context_1 = new OfficeCore.RequestContext();
                var auth_1 = OfficeCore.AuthenticationService.newObject(context_1);
                var handler_1 = auth_1.onTokenReceived.add(function (arg) {
                    if (!OfficeExtension.CoreUtility.isNullOrUndefined(arg)) {
                        handler_1.remove();
                        context_1.sync()["catch"](function () {
                        });
                        if (arg.code == 0) {
                            resolve(arg.tokenValue);
                        }
                        else {
                            if (OfficeExtension.CoreUtility.isNullOrUndefined(arg.errorInfo)) {
                                reject({ code: arg.code });
                            }
                            else {
                                try {
                                    reject(JSON.parse(arg.errorInfo));
                                }
                                catch (e) {
                                    reject({ code: arg.code, message: arg.errorInfo });
                                }
                            }
                        }
                    }
                    return null;
                });
                context_1.sync()
                    .then(function () {
                    var apiResult = auth_1.getAccessToken(options, auth_1._targetId);
                    return context_1.sync()
                        .then(function () {
                        if (OfficeExtension.CoreUtility.isNullOrUndefined(apiResult.value)) {
                            return null;
                        }
                        var tokenValue = apiResult.value.accessToken;
                        if (!OfficeExtension.CoreUtility.isNullOrUndefined(tokenValue)) {
                            resolve(apiResult.value);
                        }
                    });
                })["catch"](function (e) {
                    reject(e);
                });
            }
        });
    }
    OfficeFirstPartyAuth.getAccessToken = getAccessToken;
    function getPrimaryIdentityInfo() {
        var context = new OfficeCore.RequestContext();
        var auth = OfficeCore.AuthenticationService.newObject(context);
        context._customData = "WacPartition";
        var result = auth.getPrimaryIdentityInfo();
        return context.sync().then(function () { return result.value; });
    }
    OfficeFirstPartyAuth.getPrimaryIdentityInfo = getPrimaryIdentityInfo;
    function getIdentities() {
        var context = new OfficeCore.RequestContext();
        var auth_service = OfficeCore.AuthenticationService.newObject(context);
        var result = auth_service.getIdentities();
        return context.sync().then(function () { return result.value; });
    }
    OfficeFirstPartyAuth.getIdentities = getIdentities;
    function logLoadEvent(result, succeeded) {
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeStringDataField('IdentityProvider', OSF.WebAuth.config.idp),
                    oteljs.makeStringDataField('AppId', OSF.WebAuth.config.appIds[0]),
                    oteljs.makeStringDataField('Target', OSF.WebAuth.config.authFlow),
                    oteljs.makeBooleanDataField('Result', succeeded),
                    oteljs.makeStringDataField('Error', (result instanceof Event) ? result.type : "")
                ];
                if (result && !(result instanceof Event) && result.Telemetry) {
                    for (var key in result.Telemetry) {
                        if (!result.Telemetry[key]) {
                            continue;
                        }
                        switch (key) {
                            case 'succeeded':
                                telemetryData.push(oteljs.makeBooleanDataField(key, result.Telemetry[key]));
                                break;
                            case 'loadedApplicationCount':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            case 'timeToLoad':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            default:
                                telemetryData.push(oteljs.makeStringDataField(key, result.Telemetry[key]));
                        }
                    }
                }
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.Load",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function logAcquireEvent(result, succeeded, popup) {
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeStringDataField('IdentityProvider', OSF.WebAuth.config.idp),
                    oteljs.makeStringDataField('AppId', OSF.WebAuth.config.appIds[0]),
                    oteljs.makeStringDataField('Target', OSF.WebAuth.config.authFlow),
                    oteljs.makeBooleanDataField('Result', succeeded),
                    oteljs.makeStringDataField('Error', (result instanceof Event) ? result.type : result.ErrorCode),
                    oteljs.makeBooleanDataField('Popup', (typeof popup === "boolean") ? popup : false),
                ];
                if (result && !(result instanceof Event) && result.Telemetry) {
                    for (var key in result.Telemetry) {
                        if (!result.Telemetry[key]) {
                            continue;
                        }
                        switch (key) {
                            case 'succeeded':
                                telemetryData.push(oteljs.makeBooleanDataField(key, result.Telemetry[key]));
                                break;
                            case 'timeToGetToken':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            default:
                                telemetryData.push(oteljs.makeStringDataField(key, result.Telemetry[key]));
                        }
                    }
                }
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.GetAccessToken",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function logUnexpectedAcquire(loadResult, loadAttempts) {
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeBooleanDataField('Loaded', loadResult),
                    oteljs.makeInt64DataField('LoadAttempts', (typeof loadAttempts === "number") ? loadAttempts : 0)
                ];
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.UnexpectedAcquire",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function loadWebAuthForReplyPage() {
        try {
            if (typeof (window) === "undefined" || !window.sessionStorage) {
                return;
            }
            var webAuthRedirectUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
            if (webAuthRedirectUrls !== null && webAuthRedirectUrls.indexOf(window.location.origin + window.location.pathname) !== -1) {
                load();
            }
        }
        catch (ex) {
            console.error(ex);
        }
    }
    if (typeof (window) !== "undefined" && window.OSF) {
        loadWebAuthForReplyPage();
    }
})(OfficeFirstPartyAuth || (OfficeFirstPartyAuth = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "Office";
    var _defaultApiSetName = "OfficeSharedApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var AuthenticationServiceCustom = (function () {
        function AuthenticationServiceCustom() {
        }
        Object.defineProperty(AuthenticationServiceCustom.prototype, "_targetId", {
            get: function () {
                if (this.m_targetId == undefined) {
                    if (typeof (OSF) !== 'undefined' && OSF.OUtil) {
                        this.m_targetId = OSF.OUtil.Guid.generateNewGuid();
                    }
                    else {
                        this.m_targetId = "" + this.context._nextId();
                    }
                }
                return this.m_targetId;
            },
            enumerable: true,
            configurable: true
        });
        return AuthenticationServiceCustom;
    }());
    OfficeCore.AuthenticationServiceCustom = AuthenticationServiceCustom;
    var _CC;
    (function (_CC) {
        function AuthenticationService_TokenReceived_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                tokenValue: value.tokenValue,
                code: value.code,
                errorInfo: value.errorInfo
            };
            return newArgs;
        }
        _CC.AuthenticationService_TokenReceived_EventArgsTransform = AuthenticationService_TokenReceived_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    (function (_CC) {
        function FirstPartyDialogService_DialogMessage_EventArgsTransform(thisObj, args) {
            return {
                message: args.message
            };
        }
        _CC.FirstPartyDialogService_DialogMessage_EventArgsTransform = FirstPartyDialogService_DialogMessage_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var PersonaPromiseType;
    (function (PersonaPromiseType) {
        PersonaPromiseType[PersonaPromiseType["immediate"] = 0] = "immediate";
        PersonaPromiseType[PersonaPromiseType["load"] = 3] = "load";
    })(PersonaPromiseType = OfficeCore.PersonaPromiseType || (OfficeCore.PersonaPromiseType = {}));
    var PersonaInfoAndSource = (function () {
        function PersonaInfoAndSource() {
        }
        return PersonaInfoAndSource;
    }());
    OfficeCore.PersonaInfoAndSource = PersonaInfoAndSource;
    ;
    var PersonaCustom = (function () {
        function PersonaCustom() {
        }
        PersonaCustom.prototype.performAsyncOperation = function (type, waitFor, action, check) {
            var _this = this;
            if (type == PersonaPromiseType.immediate) {
                action();
                return;
            }
            check().then(function (isWarmedUp) {
                if (isWarmedUp) {
                    action();
                }
                else {
                    var persona = _this;
                    persona.load("hostId");
                    persona.context.sync().then(function () {
                        var hostId = persona.hostId;
                        _this.getPersonaLifetime().then(function (personaLifetime) {
                            var eventHandler = function (args) {
                                return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                                    if (args.sendingPersonaHostId == hostId) {
                                        for (var index = 0; index < args.dataUpdated.length; ++index) {
                                            var updated = args.dataUpdated[index];
                                            if (waitFor == updated) {
                                                check().then(function (isWarmedUp) {
                                                    if (isWarmedUp) {
                                                        action();
                                                        personaLifetime.onPersonaUpdated.remove(eventHandler);
                                                        persona.context.sync();
                                                    }
                                                    resolve(isWarmedUp);
                                                });
                                                return;
                                            }
                                        }
                                    }
                                    resolve(false);
                                });
                            };
                            personaLifetime.onPersonaUpdated.add(eventHandler);
                            persona.context.sync();
                        });
                    });
                }
            });
        };
        PersonaCustom.prototype.getOrganizationAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var organization = persona.organization;
                    organization.load("*");
                    persona.context.sync().then(function () {
                        resolve(organization);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve, isWarmedUpReject) {
                        var organization = persona.organization;
                        organization.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(organization.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.organization, action, check);
            });
        };
        PersonaCustom.prototype.getIsPersonaInfoResolvedCheck = function () {
            var persona = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var info = persona.personaInfo;
                info.load("isPersonResolved");
                persona.context.sync().then(function () {
                    resolve(info.isPersonResolved);
                });
            });
        };
        PersonaCustom.prototype.getPersonaInfoAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var info = persona.personaInfo;
                    info.load();
                    persona.context.sync().then(function () {
                        resolve(info);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getPersonaInfoWithSourceAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var result = new PersonaInfoAndSource();
                    result.info = persona.personaInfo;
                    result.info.load();
                    result.source = persona.personaInfo.sources;
                    result.source.load();
                    persona.context.sync().then(function () {
                        resolve(result);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getUnifiedCommunicationInfo = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var ucInfo = persona.unifiedCommunicationInfo;
                    ucInfo.load("*");
                    persona.context.sync().then(function () {
                        resolve(ucInfo);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getUnifiedGroupInfoAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var group = persona.unifiedGroupInfo;
                    group.load("*");
                    persona.context.sync().then(function () {
                        resolve(group);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getTypeAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    persona.load("type");
                    persona.context.sync().then(function () {
                        resolve(OfficeCore.PersonaType[persona.type.valueOf()]);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getCustomizationsAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var customizations = persona.getCustomizations();
                    persona.context.sync().then(function () {
                        resolve(customizations.value);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getMembersAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, rejcet) {
                var persona = _this;
                var action = function () {
                    var members = persona.getMembers();
                    members.load("isWarmedUp");
                    persona.context.sync().then(function () {
                        resolve(members);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve, isWarmedUpReject) {
                        var members = persona.getMembers();
                        members.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(members.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.members, action, check);
            });
        };
        PersonaCustom.prototype.getMembershipAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var membership = persona.getMembership();
                    membership.load("*");
                    persona.context.sync().then(function () {
                        resolve(membership);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve) {
                        var membership = persona.getMembership();
                        membership.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(membership.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.membership, action, check);
            });
        };
        PersonaCustom.prototype.getPersonaLifetime = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                persona.load("instanceId");
                persona.context.sync().then(function () {
                    var peopleApi = new PeopleApiContext(persona.context, persona.instanceId);
                    peopleApi.getPersonaLifetime().then(function (lifetime) {
                        resolve(lifetime);
                    });
                });
            });
        };
        return PersonaCustom;
    }());
    OfficeCore.PersonaCustom = PersonaCustom;
    var PeopleApiContext = (function () {
        function PeopleApiContext(context, instanceId) {
            this.context = context;
            this.instanceId = instanceId;
        }
        Object.defineProperty(PeopleApiContext.prototype, "serviceContext", {
            get: function () {
                if (!this.m_serviceConext) {
                    this.m_serviceConext = OfficeCore.ServiceContext.newObject(this.context);
                }
                return this.m_serviceConext;
            },
            enumerable: true,
            configurable: true
        });
        PeopleApiContext.prototype.getPersonaLifetime = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var lifetime = _this.serviceContext.getPersonaLifetime(_this.instanceId);
                _this.context.sync().then(function () {
                    lifetime.load("instanceId");
                    _this.context.sync().then(function () {
                        resolve(lifetime);
                    });
                });
            });
        };
        PeopleApiContext.prototype.getInitialPersona = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this.serviceContext.getInitialPersona(_this.instanceId);
                _this.context.sync().then(function () {
                    resolve(persona);
                });
            });
        };
        PeopleApiContext.prototype.getLokiTokenProvider = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var provider = _this.serviceContext.getLokiTokenProvider(_this.instanceId);
                _this.context.sync().then(function () {
                    provider.load("instanceId");
                    _this.context.sync().then(function () {
                        resolve(provider);
                    });
                });
            });
        };
        return PeopleApiContext;
    }());
    OfficeCore.PeopleApiContext = PeopleApiContext;
    (function (_CC) {
        function LicenseFeature_StateChanged_EventArgsTransform(thisObj, args) {
            var newArgs = {
                feature: args.featureName,
                isEnabled: args.isEnabled,
                tier: args.tierName
            };
            if (args.tierName) {
                newArgs.tier = args.tierName == 0 ? LicenseFeatureTier.unknown :
                    args.tierName == 1 ? LicenseFeatureTier.basic :
                        args.tierName == 2 ? LicenseFeatureTier.premium :
                            args.tierName;
            }
            return newArgs;
        }
        _CC.LicenseFeature_StateChanged_EventArgsTransform = LicenseFeature_StateChanged_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["apiNotAvailable"] = "ApiNotAvailable";
        ErrorCodes["clientError"] = "ClientError";
        ErrorCodes["controlIdNotFound"] = "ControlIdNotFound";
        ErrorCodes["entryIdRequired"] = "EntryIdRequired";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["hostRestartNeeded"] = "HostRestartNeeded";
        ErrorCodes["instanceNotFound"] = "InstanceNotFound";
        ErrorCodes["interactiveFlowAborted"] = "InteractiveFlowAborted";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["invalidGrant"] = "InvalidGrant";
        ErrorCodes["invalidResourceUrl"] = "InvalidResourceUrl";
        ErrorCodes["invalidRibbonDefinition"] = "InvalidRibbonDefinition";
        ErrorCodes["objectNotFound"] = "ObjectNotFound";
        ErrorCodes["resourceNotSupported"] = "ResourceNotSupported";
        ErrorCodes["serverError"] = "ServerError";
        ErrorCodes["serviceUrlNotFound"] = "ServiceUrlNotFound";
        ErrorCodes["sharedRuntimeNotAvailable"] = "SharedRuntimeNotAvailable";
        ErrorCodes["ticketInvalidParams"] = "TicketInvalidParams";
        ErrorCodes["ticketNetworkError"] = "TicketNetworkError";
        ErrorCodes["ticketUnauthorized"] = "TicketUnauthorized";
        ErrorCodes["ticketUninitialized"] = "TicketUninitialized";
        ErrorCodes["ticketUnknownError"] = "TicketUnknownError";
        ErrorCodes["unexpectedError"] = "UnexpectedError";
        ErrorCodes["unsupportedUserIdentity"] = "UnsupportedUserIdentity";
        ErrorCodes["userNotSignedIn"] = "UserNotSignedIn";
    })(ErrorCodes = OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = OfficeCore.Interfaces || (OfficeCore.Interfaces = {}));
    var _libraryMetadataOfficeSharedApi = { "version": "1.0.0",
        "name": "OfficeCore",
        "defaultApiSetName": "OfficeSharedApi",
        "hostName": "Office",
        "apiSets": [["1.2", "FirstPartyAuthentication"], ["1.3", "FirstPartyAuthentication"], ["1.2", "DynamicRibbon"], ["1.2", "SharedRuntimeInternal"]],
        "strings": ["AuthenticationService", "RoamingSetting", "RoamingSettingCollection", "BeforeDocumentCloseNotification", "ServiceUrlProvider", "LinkedIn", "NetworkUsage", "DynamicRibbon", "RibbonTab", "RibbonButton", "RibbonButtonCollection", "FirstPartyDialogService", "LocaleApi", "OfficeServicesManagerApi", "Comment", "CommentCollection", "ExtensionLifeCycle", "MemberInfoList", "PersonaActions", "PersonaInfoSource", "PersonaInfo", "PersonaUnifiedCommunicationInfo", "PersonaPhotoInfo", "PersonaCollection", "PersonaOrganizationInfo", "UnifiedGroupInfo", "Persona", "PersonaLifetime", "LokiTokenProvider", "LokiTokenProviderFactory", "ServiceContext", "RichapiPcxFeatureChecks", "Tap", "AppRuntimePersistenceService", "AppRuntimeService", "License", "LicenseFeature", "MsaDeviceTicketOptions", "DialogPage", "SharedFilePicker", "ActionService", "null", "id", "getItem", "", "getCount", "close", "isWarmedUp", "isWarmingUp", "displayName", "email", "emailAddresses", "sipAddresses", "birthday", "birthdays", "title", "jobInfoDepartment", "companyName", "office", "linkedTitles", "linkedDepartments", "linkedCompanyNames", "linkedOffices", "webSites", "notes", "getImageUri", "setPlaceholderColor", "getPlaceholderUri", "getImageUriWithMetadata", "instanceId", "dispose", "_RegisterPersonaUpdatedEvent", "_UnregisterPersonaUpdatedEvent", "this.instanceId", "_RegisterLokiTokenAvailableEvent", "_UnregisterLokiTokenAvailableEvent", "_RegisterIdentityUniqueIdAvailableEvent", "_UnregisterIdentityUniqueIdAvailableEvent", "_RegisterClientAccessTokenAvailableEvent", "_UnregisterClientAccessTokenAvailableEvent", "getLokiTokenProvider", "_RegisterStateChange", "_UnregisterStateChange", "registerOnShow", "unregisterOnShow"],
        "enumTypes": [["IdentityType", ["organizationAccount", "microsoftAccount", "unsupported"]],
            ["ServiceProvider", ["ariaBrowserPipeUrl", "ariaUploadUrl", "ariaVNextUploadUrl", "lokiAutoDiscoverUrl"]],
            ["TimeStringFormat", ["shortTime", "longTime", "shortDate", "longDate"]],
            ["CommentTextFormat", ["plain", "markdown", "delta"]],
            ["PersonaCardPerfPoint", ["placeHolderRendered", "initialCardRendered"]],
            ["MessageType", [], { "personaLifetimePersonaUpdatedEvent": 3502, "lokiTokenProviderLokiTokenAvailableEvent": 3503, "lokiTokenProviderIdentityUniqueIdAvailableEvent": 3504, "lokiTokenProviderClientAccessTokenAvailableEvent": 3505 }],
            ["UnifiedCommunicationAvailability", ["notSet", "free", "idle", "busy", "idleBusy", "doNotDisturb", "unalertable", "unavailable"]],
            ["UnifiedCommunicationStatus", ["online", "notOnline", "away", "busy", "beRightBack", "onThePhone", "outToLunch", "inAMeeting", "outOfOffice", "doNotDisturb", "inAConference", "getting", "notABuddy", "disconnected", "notInstalled", "urgentInterruptionsOnly", "mayBeAvailable", "idle", "inPresentation"]],
            ["UnifiedCommunicationPresence", ["free", "busy", "idle", "doNotDistrub", "blocked", "notSet", "outOfOffice"]],
            ["FreeBusyCalendarState", ["unknown", "free", "busy", "elsewhere", "tentative", "outOfOffice"]],
            ["PersonaType", ["unknown", "enterprise", "contact", "bot", "phoneOnly", "oneOff", "distributionList", "personalDistributionList", "anonymous", "unifiedGroup"]],
            ["PhoneType", ["workPhone", "homePhone", "mobilePhone", "businessFax", "otherPhone"]],
            ["AddressType", ["workAddress", "homeAddress", "otherAddress"]],
            ["MemberType", ["unknown", "individual", "group"]],
            ["PersonaDataUpdated", ["hostId", "type", "photo", "personaInfo", "unifiedCommunicationInfo", "organization", "unifiedGroupInfo", "members", "membership", "capabilities", "customizations", "viewableSources", "placeholder"]],
            ["CustomizedData", ["email", "workPhone", "workPhone2", "workFax", "mobilePhone", "homePhone", "homePhone2", "otherPhone", "sipAddress", "profile", "office", "company", "workAddress", "homeAddress", "otherAddress", "birthday"]],
            ["ObjectType", ["unknown", "chart", "smartArt", "table", "image", "slide", "text"], { "ole": "OLE" }],
            ["AppRuntimeState", ["inactive", "background", "visible"]],
            ["Visibility", ["hidden", "visible"]],
            ["LicenseFeatureTier", ["unknown", "basic", "premium"]],
            ["LicenseEventType", [], { "featureStateChanged": 1 }],
            ["DialogPageEventType", [], { "onShow": 1 }]],
        "clientObjectTypes": [[1,
                4,
                0,
                [["roamingSettings",
                        3,
                        2,
                        0,
                        0,
                        4]],
                [["getAccessToken",
                        2,
                        2,
                        0,
                        5],
                    ["getPrimaryIdentityInfo",
                        0,
                        2,
                        1,
                        5],
                    ["getIdentities",
                        0,
                        2,
                        2,
                        5]],
                0,
                0,
                0,
                [["TokenReceived",
                        2,
                        1,
                        "3001",
                        "this._targetId",
                        42,
                        42]],
                "Microsoft.Authentication.AuthenticationService",
                4],
            [2,
                0,
                [[43,
                        3],
                    ["value",
                        1]]],
            [3,
                0,
                0,
                0,
                0,
                [[44,
                        2,
                        1,
                        2,
                        0,
                        4],
                    ["getItemOrNullObject",
                        2,
                        1,
                        2,
                        0,
                        4]]],
            [4,
                0,
                0,
                0,
                [["enable",
                        0,
                        2,
                        0,
                        4],
                    ["disable",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["BeforeCloseNotificationCancelled",
                        0,
                        0,
                        "65540",
                        45,
                        42,
                        42]],
                "Microsoft.BeforeDocumentCloseNotification.BeforeDocumentCloseNotification",
                4],
            [5,
                0,
                0,
                0,
                [["getServiceUrl",
                        2,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.ServiceUrlProvider",
                4],
            [6,
                0,
                0,
                0,
                [["isEnabledForOffice",
                        0,
                        2,
                        0,
                        4],
                    ["recordLinkedInSettingsCompliance",
                        2]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.LinkedIn",
                4],
            [7,
                0,
                0,
                0,
                [["isInOnlineMode",
                        0,
                        2,
                        0,
                        4],
                    ["isInDisconnectedMode",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.NetworkUsage",
                4],
            [8,
                0,
                0,
                [["buttons",
                        11,
                        19,
                        0,
                        0,
                        4]],
                [["executeRequestUpdate",
                        1,
                        2,
                        0,
                        4],
                    ["executeRequestCreate",
                        1,
                        2,
                        3,
                        4]],
                [["getButton",
                        10,
                        1,
                        2,
                        0,
                        4],
                    ["getTab",
                        9,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.DynamicRibbon.DynamicRibbon",
                4],
            [9,
                0,
                [[43,
                        3]],
                0,
                [["setVisibility",
                        1]]],
            [10,
                0,
                [[43,
                        3],
                    ["enabled",
                        1],
                    ["label",
                        3]],
                0,
                [["setEnabled",
                        1]]],
            [11,
                1,
                0,
                0,
                [[46,
                        0,
                        2,
                        0,
                        4]],
                [[44,
                        10,
                        1,
                        18,
                        0,
                        4]],
                0,
                10],
            [12,
                0,
                0,
                0,
                [["displayDialog",
                        2,
                        2,
                        0,
                        5],
                    [47,
                        0,
                        2,
                        0,
                        5]],
                0,
                0,
                0,
                [["DialogMessage",
                        2,
                        0,
                        "65536",
                        45,
                        42,
                        42]],
                "Microsoft.FirstPartyDialog.FirstPartyDialogService",
                4],
            [13,
                0,
                0,
                0,
                [["getLocaleDateTimeFormattingInfo",
                        1,
                        2,
                        0,
                        4],
                    ["formatDateTimeString",
                        3,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.LocaleApi.LocaleApi",
                4],
            [14,
                0,
                0,
                0,
                [["bindServiceToProfile",
                        3]],
                0,
                0,
                0,
                0,
                "Microsoft.OfficeServicesManager.OfficeServicesManagerApi",
                4],
            [15,
                0,
                [[43,
                        3],
                    ["text",
                        1],
                    ["created",
                        11],
                    ["level",
                        3],
                    ["resolved",
                        1],
                    ["author",
                        3],
                    ["mentions",
                        3]],
                [["parent",
                        15,
                        2,
                        0,
                        0,
                        4],
                    ["parentOrNullObject",
                        15,
                        2,
                        0,
                        0,
                        4],
                    ["replies",
                        16,
                        19,
                        0,
                        0,
                        4]],
                [["getRichText",
                        1,
                        2,
                        0,
                        4],
                    ["setRichText",
                        2],
                    ["delete"]],
                [["getParentOrSelf",
                        15,
                        0,
                        2,
                        0,
                        4],
                    ["reply",
                        15,
                        2]]],
            [16,
                1,
                0,
                0,
                [[46,
                        0,
                        2,
                        0,
                        4]],
                [[44,
                        15,
                        1,
                        18,
                        0,
                        4]],
                0,
                15],
            [17,
                0,
                0,
                0,
                [["launchExtensionComponent",
                        3,
                        2,
                        0,
                        5]],
                0,
                0,
                0,
                0,
                "Microsoft.OfficeSharedApi.ExtensionLifeCycle",
                4],
            [18,
                0,
                [[48,
                        3],
                    [49,
                        3]],
                0,
                [["items",
                        0,
                        2,
                        0,
                        4]],
                [["getPersonaForMember",
                        27,
                        1,
                        2,
                        0,
                        4]]],
            [19,
                0,
                0,
                0,
                [["addContact"],
                    ["editContact"],
                    ["composeEmail",
                        1],
                    ["composeInstantMessage",
                        1],
                    ["callPhoneNumber",
                        1],
                    ["pinPersonaToQuickContacts"],
                    ["toggleTagForAlerts"],
                    ["scheduleMeeting"],
                    ["openLinkContactUx"],
                    ["editContactByIdentifier",
                        1],
                    ["showHoverCardForPersona",
                        6],
                    ["hideHoverCardForPersona"],
                    ["showContextMenu",
                        6],
                    ["showContactCard",
                        6],
                    ["showExpandedCard",
                        6],
                    ["openGroupCalendar"],
                    ["subscribeToGroup"],
                    ["unsubscribeFromGroup"],
                    ["getChangePhotoUrlAndOpenInBrowser"],
                    ["startAudioCall"],
                    ["startVideoCall"],
                    ["openOutlookProperties"],
                    ["editUnifiedGroup"],
                    ["joinUnifiedGroup"],
                    ["leaveUnifiedGroup"]]],
            [20,
                0,
                [[50,
                        3],
                    [51,
                        3],
                    [52,
                        3],
                    [53,
                        3],
                    [54,
                        3],
                    [55,
                        3],
                    [56,
                        3],
                    [57,
                        3],
                    [58,
                        3],
                    [59,
                        3],
                    [60,
                        3],
                    [61,
                        3],
                    [62,
                        3],
                    [63,
                        3],
                    ["phones",
                        3],
                    ["addresses",
                        3],
                    [64,
                        3],
                    [65,
                        3]]],
            [21,
                0,
                [[50,
                        3],
                    [51,
                        3],
                    [52,
                        3],
                    [53,
                        3],
                    [54,
                        11],
                    [55,
                        11],
                    [56,
                        3],
                    [57,
                        3],
                    [58,
                        3],
                    [59,
                        3],
                    [60,
                        3],
                    [61,
                        3],
                    [62,
                        3],
                    [63,
                        3],
                    [64,
                        3],
                    [65,
                        3],
                    ["isPersonResolved",
                        3]],
                [["sources",
                        20,
                        3,
                        0,
                        0,
                        4]],
                [["getPhones",
                        0,
                        2,
                        0,
                        4],
                    ["getAddresses",
                        0,
                        2,
                        0,
                        4]]],
            [22,
                0,
                [["availability",
                        3],
                    ["status",
                        3],
                    ["isSelf",
                        3],
                    ["isTagged",
                        3],
                    ["customStatusString",
                        3],
                    ["isBlocked",
                        3],
                    ["presenceTooltip",
                        3],
                    ["isOutOfOffice",
                        3],
                    ["outOfOfficeNote",
                        3],
                    ["timezone",
                        3],
                    ["meetingLocation",
                        3],
                    ["meetingSubject",
                        3],
                    ["timezoneBias",
                        3],
                    ["idleStartTime",
                        11],
                    ["overallCapability",
                        3],
                    ["isOnBuddyList",
                        3],
                    ["presenceNote",
                        3],
                    ["voiceMailUri",
                        3],
                    ["availabilityText",
                        3],
                    ["availabilityTooltip",
                        3],
                    ["isDurationInAvailabilityText",
                        3],
                    ["freeBusyStatus",
                        3],
                    ["calendarState",
                        3],
                    ["presence",
                        3]]],
            [23,
                0,
                0,
                0,
                [[66,
                        1,
                        2,
                        0,
                        4,
                        66],
                    [67,
                        1,
                        0,
                        0,
                        0,
                        67],
                    [68,
                        1,
                        2,
                        0,
                        4,
                        68],
                    [69,
                        1,
                        2,
                        0,
                        4,
                        69]]],
            [24,
                1,
                0,
                0,
                [[46,
                        0,
                        2,
                        0,
                        4]],
                [[44,
                        27,
                        1,
                        18,
                        0,
                        4]],
                0,
                27],
            [25,
                0,
                [[48,
                        3],
                    [49,
                        3]],
                [["hierarchy",
                        24,
                        18,
                        0,
                        0,
                        4],
                    ["manager",
                        27,
                        2,
                        0,
                        0,
                        4],
                    ["directReports",
                        24,
                        18,
                        0,
                        0,
                        4]]],
            [26,
                0,
                [["description",
                        1],
                    ["oneDrive",
                        1],
                    ["oneNote",
                        1],
                    ["isPublic",
                        1],
                    ["amIOwner",
                        1],
                    ["amIMember",
                        1],
                    ["amISubscribed",
                        1],
                    ["memberCount",
                        1],
                    ["ownerCount",
                        1],
                    ["hasGuests",
                        1],
                    ["site",
                        1],
                    ["planner",
                        1],
                    ["classification",
                        1],
                    ["subscriptionEnabled",
                        1]]],
            [27,
                4,
                [["hostId",
                        3],
                    ["type",
                        3],
                    ["capabilities",
                        3],
                    ["diagnosticId",
                        3],
                    [70,
                        3]],
                [["photo",
                        23,
                        3,
                        0,
                        0,
                        4],
                    ["personaInfo",
                        21,
                        3,
                        0,
                        0,
                        4],
                    ["unifiedCommunicationInfo",
                        22,
                        3,
                        0,
                        0,
                        4],
                    ["organization",
                        25,
                        3,
                        0,
                        0,
                        4],
                    ["unifiedGroupInfo",
                        26,
                        35,
                        0,
                        0,
                        4],
                    ["actions",
                        19,
                        2,
                        0,
                        0,
                        4]],
                [["getCustomizations",
                        0,
                        2,
                        0,
                        4],
                    ["warmup",
                        1],
                    [71],
                    ["getViewableSources",
                        0,
                        2,
                        0,
                        4],
                    ["reportTimeForRender",
                        2]],
                [["getMembers",
                        18,
                        0,
                        2,
                        0,
                        4],
                    ["getMembership",
                        18,
                        0,
                        2,
                        0,
                        4]]],
            [28,
                0,
                [[70,
                        3]],
                0,
                [["getPolicies",
                        0,
                        2,
                        0,
                        4],
                    [72],
                    [73],
                    ["getTextScaleFactor",
                        0,
                        2,
                        0,
                        4]],
                [["getPersona",
                        27,
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaForOrgEntry",
                        27,
                        4,
                        2,
                        0,
                        4],
                    ["getPersonaForOrgByEntryId",
                        27,
                        4,
                        2,
                        0,
                        4]],
                0,
                0,
                [["PersonaUpdated",
                        0,
                        0,
                        "MessageType.personaLifetimePersonaUpdatedEvent",
                        74,
                        72,
                        73]]],
            [29,
                0,
                [["emailOrUpn",
                        3],
                    [70,
                        3]],
                0,
                [["requestToken"],
                    [75],
                    [76],
                    ["requestIdentityUniqueId"],
                    [77],
                    [78],
                    ["requestClientAccessToken"],
                    [79],
                    [80]],
                0,
                0,
                0,
                [["ClientAccessTokenAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderClientAccessTokenAvailableEvent",
                        74,
                        79,
                        80],
                    ["IdentityUniqueIdAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderIdentityUniqueIdAvailableEvent",
                        74,
                        77,
                        78],
                    ["LokiTokenAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderLokiTokenAvailableEvent",
                        74,
                        75,
                        76]]],
            [30,
                0,
                0,
                0,
                0,
                [[81,
                        29,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.People.LokiTokenProviderFactory",
                4],
            [31,
                0,
                0,
                0,
                [[71,
                        1],
                    ["accountEmailOrUpn",
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaPolicies",
                        0,
                        2,
                        0,
                        4]],
                [[81,
                        29,
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaLifetime",
                        28,
                        1,
                        2,
                        0,
                        4],
                    ["getInitialPersona",
                        27,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.People.ServiceContext",
                4],
            [32,
                0,
                0,
                0,
                [["isAddChangePhotoLinkOnLpcPersonaImageFlightEnabled",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.People.RichapiPcxFeatureChecks",
                4],
            [33,
                0,
                0,
                0,
                [["getEnterpriseUserInfo",
                        0,
                        2,
                        0,
                        5],
                    ["getMruFriendlyPath",
                        1,
                        2,
                        0,
                        5],
                    ["launchFileUrlInOfficeApp",
                        2,
                        2,
                        0,
                        5],
                    ["performLocalSearch",
                        4,
                        2,
                        0,
                        5],
                    ["readSearchCache",
                        3,
                        2,
                        0,
                        5],
                    ["writeSearchCache",
                        3,
                        2,
                        0,
                        5]],
                0,
                0,
                0,
                0,
                "Microsoft.TapRichApi.Tap",
                4],
            [34,
                0,
                0,
                0,
                [["setAppRuntimeStartState",
                        1,
                        0,
                        0,
                        2,
                        0,
                        4],
                    ["getAppRuntimeStartState",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.AppRuntime.AppRuntimePersistenceService",
                4],
            [35,
                0,
                0,
                0,
                [["setAppRuntimeState",
                        1,
                        0,
                        0,
                        2,
                        0,
                        4],
                    ["getAppRuntimeState",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["VisibilityChanged",
                        0,
                        0,
                        "65539",
                        45,
                        42,
                        42]],
                "Microsoft.AppRuntime.AppRuntimeService",
                4],
            [36,
                0,
                0,
                0,
                [["isFeatureEnabled",
                        2,
                        2,
                        0,
                        4],
                    ["getFeatureTier",
                        2,
                        2,
                        0,
                        4],
                    ["isFreemiumUpsellEnabled",
                        0,
                        2,
                        0,
                        4],
                    ["launchUpsellExperience",
                        1,
                        2,
                        0,
                        4],
                    ["_TestFireStateChangedEvent",
                        1,
                        0,
                        0,
                        1],
                    ["getMsaDeviceTicket",
                        3,
                        2,
                        0,
                        5]],
                [["getLicenseFeature",
                        37,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.Office.Licensing.License",
                4],
            [37,
                0,
                [[43,
                        3]],
                0,
                [[82,
                        0,
                        2,
                        0,
                        4],
                    [83,
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["StateChanged",
                        2,
                        0,
                        "LicenseEventType.featureStateChanged",
                        "this.id",
                        82,
                        83]]],
            [38,
                0,
                [["scopes",
                        1]],
                0,
                0,
                0,
                0,
                0,
                0,
                "Microsoft.Office.Licensing.MsaDeviceTicketOptions",
                4],
            [39,
                0,
                [["_Id",
                        2]],
                0,
                [[47,
                        0,
                        2,
                        0,
                        4],
                    ["readyToShow",
                        0,
                        2,
                        0,
                        4],
                    [84,
                        0,
                        2,
                        0,
                        4],
                    [85,
                        0,
                        2,
                        0,
                        4],
                    ["sendMessageToHost",
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["OnShowEvent",
                        0,
                        0,
                        "DialogPageEventType.onShow",
                        "this._Id",
                        84,
                        85]],
                "Microsoft.Office.DialogPage.DialogPage",
                4],
            [40,
                0,
                0,
                0,
                [["getSharedFilePickerResponse",
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.Office.SharedFilePicker",
                4],
            [41,
                0,
                0,
                0,
                [["getShortcuts",
                        0,
                        2,
                        0,
                        4],
                    ["replaceShortcuts",
                        1],
                    ["areShortcutsInUse",
                        1]],
                0,
                0,
                0,
                0,
                "Microsoft.Office.ActionService",
                4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataOfficeSharedApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    var VisibilityMode;
    (function (VisibilityMode) {
        VisibilityMode["hidden"] = "Hidden";
        VisibilityMode["taskpane"] = "Taskpane";
    })(VisibilityMode = Office.VisibilityMode || (Office.VisibilityMode = {}));
    var StartupBehavior;
    (function (StartupBehavior) {
        StartupBehavior["none"] = "None";
        StartupBehavior["load"] = "Load";
    })(StartupBehavior = Office.StartupBehavior || (Office.StartupBehavior = {}));
    var addin;
    (function (addin) {
        function _createRequestContext(wacPartition) {
            var context = new OfficeCore.RequestContext();
            context._requestFlagModifier |= 64;
            if (wacPartition) {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function setStartupBehavior(behavior) {
            return __awaiter(this, void 0, void 0, function () {
                var state, context, appRuntimePersistenceService;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            if (behavior !== StartupBehavior.load && behavior !== StartupBehavior.none) {
                                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.invalidArgument, null, null);
                            }
                            state = (behavior == StartupBehavior.load ? OfficeCore.AppRuntimeState.background : OfficeCore.AppRuntimeState.inactive);
                            context = _createRequestContext(false);
                            appRuntimePersistenceService = OfficeCore.AppRuntimePersistenceService.newObject(context);
                            appRuntimePersistenceService.setAppRuntimeStartState(state);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        addin.setStartupBehavior = setStartupBehavior;
        function getStartupBehavior() {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimePersistenceService, stateResult, state, ret;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(false);
                            appRuntimePersistenceService = OfficeCore.AppRuntimePersistenceService.newObject(context);
                            stateResult = appRuntimePersistenceService.getAppRuntimeStartState();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            state = stateResult.value;
                            ret = (state == OfficeCore.AppRuntimeState.inactive ? StartupBehavior.none : StartupBehavior.load);
                            return [2, ret];
                    }
                });
            });
        }
        addin.getStartupBehavior = getStartupBehavior;
        function _setState(state) {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimeService;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(true);
                            appRuntimeService = OfficeCore.AppRuntimeService.newObject(context);
                            appRuntimeService.setAppRuntimeState(state);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        function _getState() {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimeService, stateResult;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(true);
                            appRuntimeService = OfficeCore.AppRuntimeService.newObject(context);
                            stateResult = appRuntimeService.getAppRuntimeState();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, stateResult.value];
                    }
                });
            });
        }
        addin._getState = _getState;
        function showAsTaskpane() {
            return _setState(OfficeCore.AppRuntimeState.visible);
        }
        addin.showAsTaskpane = showAsTaskpane;
        function hide() {
            return _setState(OfficeCore.AppRuntimeState.background);
        }
        addin.hide = hide;
        var _appRuntimeEvent;
        function _getAppRuntimeEventService() {
            if (!_appRuntimeEvent) {
                var context = _createRequestContext(true);
                _appRuntimeEvent = OfficeCore.AppRuntimeService.newObject(context);
            }
            return _appRuntimeEvent;
        }
        function _convertVisibilityToVisibilityMode(visibility) {
            if (visibility === OfficeCore.Visibility.visible) {
                return VisibilityMode.taskpane;
            }
            return VisibilityMode.hidden;
        }
        function onVisibilityModeChanged(listener) {
            return __awaiter(this, void 0, void 0, function () {
                var eventService, registrationToken, ret;
                var _this = this;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            eventService = _getAppRuntimeEventService();
                            registrationToken = eventService.onVisibilityChanged.add(function (args) {
                                if (listener) {
                                    var msg = {
                                        visibilityMode: _convertVisibilityToVisibilityMode(args.visibility)
                                    };
                                    listener(msg);
                                }
                                return null;
                            });
                            return [4, eventService.context.sync()];
                        case 1:
                            _a.sent();
                            ret = function () {
                                return __awaiter(_this, void 0, void 0, function () {
                                    return __generator(this, function (_a) {
                                        switch (_a.label) {
                                            case 0:
                                                registrationToken.remove();
                                                return [4, eventService.context.sync()];
                                            case 1:
                                                _a.sent();
                                                return [2];
                                        }
                                    });
                                });
                            };
                            return [2, ret];
                    }
                });
            });
        }
        addin.onVisibilityModeChanged = onVisibilityModeChanged;
        var beforeDocumentCloseNotification;
        (function (beforeDocumentCloseNotification) {
            function _createRequestContext(wacPartition) {
                var context = new OfficeCore.RequestContext();
                context._requestFlagModifier |= 64;
                if (wacPartition) {
                    context._customData = 'WacPartition';
                }
                return context;
            }
            function enable() {
                return __awaiter(this, void 0, void 0, function () {
                    var isWac, context, beforeCloseNotification;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                isWac = false;
                                if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                                    isWac = true;
                                }
                                context = _createRequestContext(isWac);
                                beforeCloseNotification = OfficeCore.BeforeDocumentCloseNotification.newObject(context);
                                beforeCloseNotification.enable();
                                return [4, context.sync()];
                            case 1:
                                _a.sent();
                                return [2];
                        }
                    });
                });
            }
            beforeDocumentCloseNotification.enable = enable;
            function disable() {
                return __awaiter(this, void 0, void 0, function () {
                    var isWac, context, beforeCloseNotification;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                isWac = false;
                                if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                                    isWac = true;
                                }
                                context = _createRequestContext(isWac);
                                beforeCloseNotification = OfficeCore.BeforeDocumentCloseNotification.newObject(context);
                                beforeCloseNotification.disable();
                                return [4, context.sync()];
                            case 1:
                                _a.sent();
                                return [2];
                        }
                    });
                });
            }
            beforeDocumentCloseNotification.disable = disable;
            function onCloseActionCancelled(listener) {
                return __awaiter(this, void 0, void 0, function () {
                    var isWac, context, beforeCloseNotification, registrationToken, ret;
                    var _this = this;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                isWac = false;
                                if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                                    isWac = true;
                                }
                                context = _createRequestContext(isWac);
                                beforeCloseNotification = OfficeCore.BeforeDocumentCloseNotification.newObject(context);
                                registrationToken = beforeCloseNotification.onBeforeCloseNotificationCancelled.add(function (args) {
                                    if (listener) {
                                        listener();
                                    }
                                    return null;
                                });
                                return [4, context.sync()];
                            case 1:
                                _a.sent();
                                ret = function () {
                                    return __awaiter(_this, void 0, void 0, function () {
                                        return __generator(this, function (_a) {
                                            switch (_a.label) {
                                                case 0:
                                                    registrationToken.remove();
                                                    return [4, context.sync()];
                                                case 1:
                                                    _a.sent();
                                                    return [2];
                                            }
                                        });
                                    });
                                };
                                return [2, ret];
                        }
                    });
                });
            }
            beforeDocumentCloseNotification.onCloseActionCancelled = onCloseActionCancelled;
        })(beforeDocumentCloseNotification = addin.beforeDocumentCloseNotification || (addin.beforeDocumentCloseNotification = {}));
    })(addin = Office.addin || (Office.addin = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var ExtensionComponentType;
    (function (ExtensionComponentType) {
        ExtensionComponentType["taskpane"] = "Taskpane";
    })(ExtensionComponentType || (ExtensionComponentType = {}));
    var extensionLifeCycle;
    (function (extensionLifeCycle) {
        function _createRequestContext(wacPartition) {
            var context = new OfficeCore.RequestContext();
            return context;
        }
        function launchTaskpane(launchOptions) {
            return __awaiter(this, void 0, void 0, function () {
                var context, extensionLifecycle, settings;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(false);
                            extensionLifecycle = OfficeCore.ExtensionLifeCycle.newObject(context);
                            settings = launchOptions.settings;
                            if (settings != undefined) {
                                launchOptions.settings = OSF.OUtil.serializeSettings(settings);
                            }
                            extensionLifecycle.launchExtensionComponent("", ExtensionComponentType.taskpane, launchOptions);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        extensionLifeCycle.launchTaskpane = launchTaskpane;
    })(extensionLifeCycle = Office.extensionLifeCycle || (Office.extensionLifeCycle = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var ribbon;
    (function (ribbon_1) {
        function _createRequestContext() {
            var context = new OfficeCore.RequestContext();
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function requestUpdate(input) {
            var requestContext = _createRequestContext();
            var ribbon = requestContext.ribbon;
            function processControls(parent) {
                if (parent.controls !== undefined
                    && parent.controls.length !== undefined
                    && !!parent.controls.length) {
                    parent.controls
                        .filter(function (control) { return !(!control.id); })
                        .forEach(function (control) {
                        var ribbonControl = ribbon.getButton(control.id);
                        if (control.enabled !== undefined && control.enabled !== null) {
                            ribbonControl.enabled = control.enabled;
                        }
                    });
                }
            }
            input.tabs
                .filter(function (tab) { return !(!tab.id); })
                .forEach(function (tab) {
                var ribbonTab = ribbon.getTab(tab.id);
                if (tab.visible !== undefined && tab.visible !== null) {
                    ribbonTab.setVisibility(tab.visible);
                }
                if (!!tab.groups && !!tab.groups.length) {
                    tab.groups
                        .filter(function (group) { return !(!group.id); })
                        .forEach(function (group) {
                        processControls(group);
                    });
                }
                else {
                    processControls(tab);
                }
            });
            return requestContext.sync();
        }
        ribbon_1.requestUpdate = requestUpdate;
        function requestCreateControls(input) {
            var requestContext = _createRequestContext();
            var ribbon = requestContext.ribbon;
            var delay = function (milliseconds) {
                return new Promise(function (resolve, _) { return setTimeout(function () { return resolve(); }, milliseconds); });
            };
            ribbon.executeRequestCreate(JSON.stringify(input));
            return delay(250)
                .then(function () { return requestContext.sync(); });
        }
        ribbon_1.requestCreateControls = requestCreateControls;
    })(ribbon = Office.ribbon || (Office.ribbon = {}));
})(Office || (Office = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "Office";
    var _defaultApiSetName = "OfficeSharedApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var AddinInternalServiceErrorCodes;
    (function (AddinInternalServiceErrorCodes) {
        AddinInternalServiceErrorCodes["generalException"] = "GeneralException";
    })(AddinInternalServiceErrorCodes || (AddinInternalServiceErrorCodes = {}));
    var _libraryMetadataInternalServiceApi = { "version": "1.0.0",
        "name": "OfficeCore",
        "defaultApiSetName": "OfficeSharedApi",
        "hostName": "Office",
        "apiSets": [],
        "strings": ["AddinInternalService"],
        "enumTypes": [],
        "clientObjectTypes": [[1,
                0,
                0,
                0,
                [["notifyActionHandlerReady",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.InternalService.AddinInternalService",
                4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataInternalServiceApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    var actionProxy;
    (function (actionProxy) {
        var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
        var _association;
        var ActionMessageCategory = 2;
        var ActionDispatchMessageType = 1000;
        function init() {
            if (typeof (OSF) !== "undefined" && OSF.DDA && OSF.DDA.RichApi && OSF.DDA.RichApi.richApiMessageManager) {
                var context = new OfficeExtension.ClientRequestContext();
                return context.eventRegistration.register(5, "", _handleMessage);
            }
        }
        function setActionAssociation(association) {
            _association = association;
        }
        function _getFunction(functionName) {
            if (functionName) {
                var nameUpperCase = functionName.toUpperCase();
                var call = _association.mappings[nameUpperCase];
                if (!_isNullOrUndefined(call) && typeof (call) === "function") {
                    return call;
                }
            }
            throw OfficeExtension.Utility.createRuntimeError("invalidOperation", "sourceData", "ActionProxy._getFunction");
        }
        function _handleMessage(args) {
            try {
                OfficeExtension.Utility.log('ActionProxy._handleMessage');
                OfficeExtension.Utility.checkArgumentNull(args, "args");
                var entryArray = args.entries;
                var invocationArray = [];
                for (var i = 0; i < entryArray.length; i++) {
                    if (entryArray[i].messageCategory !== ActionMessageCategory) {
                        continue;
                    }
                    if (typeof (entryArray[i].message) === 'string') {
                        entryArray[i].message = JSON.parse(entryArray[i].message);
                    }
                    if (entryArray[i].messageType === ActionDispatchMessageType) {
                        var actionsArgs = null;
                        var actionName = entryArray[i].message[0];
                        var call = _getFunction(actionName);
                        if (entryArray[i].message.length >= 2) {
                            var actionArgsJson = entryArray[i].message[1];
                            if (actionArgsJson) {
                                if (_isJsonObjectString(actionArgsJson)) {
                                    actionsArgs = JSON.parse(actionArgsJson);
                                }
                                else {
                                    actionsArgs = actionArgsJson;
                                }
                            }
                        }
                        if (typeof (OSF) !== 'undefined' &&
                            OSF.AppTelemetry &&
                            OSF.AppTelemetry.CallOnAppActivatedIfPending) {
                            OSF.AppTelemetry.CallOnAppActivatedIfPending();
                        }
                        call.apply(null, [actionsArgs]);
                    }
                    else {
                        OfficeExtension.Utility.log('ActionProxy._handleMessage unknown message type ' + entryArray[i].messageType);
                    }
                }
            }
            catch (ex) {
                _tryLog(ex);
                throw ex;
            }
            return OfficeExtension.Utility._createPromiseFromResult(null);
        }
        function _isJsonObjectString(value) {
            if (typeof value === 'string' && value[0] === '{') {
                return true;
            }
            return false;
        }
        function toLogMessage(ex) {
            var ret = 'Unknown Error';
            if (ex) {
                try {
                    if (ex.toString) {
                        ret = ex.toString();
                    }
                    ret = ret + ' ' + JSON.stringify(ex);
                }
                catch (otherEx) {
                    ret = 'Unexpected Error';
                }
            }
            return ret;
        }
        function _tryLog(ex) {
            var message = toLogMessage(ex);
            OfficeExtension.Utility.log(message);
        }
        function notifyActionHandlerReady() {
            var context = new OfficeExtension.ClientRequestContext();
            var addinInternalService = OfficeCore.AddinInternalService.newObject(context);
            context._customData = 'WacPartition';
            addinInternalService.notifyActionHandlerReady();
            return context.sync();
        }
        function handlerOnReadyInternal() {
            try {
                Microsoft.Office.WebExtension.onReadyInternal()
                    .then(function () {
                    return init();
                })
                    .then(function () {
                    var hostInfo = OSF._OfficeAppFactory.getHostInfo();
                    if (hostInfo.isDialog === true || (hostInfo.hostPlatform === "web" && hostInfo.hostType !== "word" && hostInfo.hostType !== "excel")) {
                        return;
                    }
                    else {
                        return notifyActionHandlerReady();
                    }
                });
            }
            catch (ex) {
            }
        }
        function initFromHostBridge(hostBridge) {
            hostBridge.addHostMessageHandler(function (bridgeMessage) {
                if (bridgeMessage.type === 3) {
                    _handleMessage(bridgeMessage.message);
                }
            });
        }
        function initOnce() {
            OfficeExtension.Utility.log('ActionProxy.initOnce');
            if (typeof (Office.actions) != 'undefined') {
                setActionAssociation(Office.actions._association);
            }
            if (typeof (document) !== 'undefined') {
                if (document.readyState && document.readyState !== 'loading') {
                    OfficeExtension.Utility.log('ActionProxy.initOnce: document.readyState is not loading state');
                    handlerOnReadyInternal();
                }
                else if (document.addEventListener) {
                    document.addEventListener("DOMContentLoaded", function () {
                        OfficeExtension.Utility.log('ActionProxy.initOnce: DOMContentLoaded event triggered');
                        handlerOnReadyInternal();
                    });
                }
            }
            OfficeExtension.HostBridge.onInited(function (hostBridge) {
                initFromHostBridge(hostBridge);
            });
        }
        initOnce();
    })(actionProxy || (actionProxy = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var actions;
    (function (actions) {
        function _createRequestContext() {
            var context = new OfficeCore.RequestContext();
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function areShortcutsInUse(shortcuts) {
            return __awaiter(this, void 0, void 0, function () {
                var context, actionService, inUseArray, inUseInfoArray, i;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            actionService = OfficeCore.ActionService.newObject(context);
                            inUseArray = actionService.areShortcutsInUse(shortcuts);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            inUseInfoArray = [];
                            for (i = 0; i < shortcuts.length; i++) {
                                inUseInfoArray.push({
                                    shortcut: shortcuts[i],
                                    inUse: inUseArray.value[i]
                                });
                            }
                            return [2, inUseInfoArray];
                    }
                });
            });
        }
        actions.areShortcutsInUse = areShortcutsInUse;
        function replaceShortcuts(shortcuts) {
            return __awaiter(this, void 0, void 0, function () {
                var context, actionService;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            actionService = OfficeCore.ActionService.newObject(context);
                            actionService.replaceShortcuts(shortcuts);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        actions.replaceShortcuts = replaceShortcuts;
        function getShortcuts() {
            return __awaiter(this, void 0, void 0, function () {
                var context, actionService, shortcuts;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            actionService = OfficeCore.ActionService.newObject(context);
                            shortcuts = actionService.getShortcuts();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, shortcuts.value];
                    }
                });
            });
        }
        actions.getShortcuts = getShortcuts;
    })(actions = Office.actions || (Office.actions = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var dialogPage;
    (function (dialogPage_1) {
        function close() {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.close();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.close = close;
        function readyToShow() {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.readyToShow();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.readyToShow = readyToShow;
        function onShow(callback) {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage, removeListener;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.onOnShowEvent.add(callback);
                            removeListener = function () {
                                dialogPage.onOnShowEvent.remove(callback);
                                return null;
                            };
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, removeListener];
                    }
                });
            });
        }
        dialogPage_1.onShow = onShow;
        function sendMessageToHost(message) {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.sendMessageToHost(message);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.sendMessageToHost = sendMessageToHost;
    })(dialogPage = Office.dialogPage || (Office.dialogPage = {}));
})(Office || (Office = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b)
            if (b.hasOwnProperty(p))
                d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var Visio;
(function (Visio) {
    var _hostName = "Visio";
    var _defaultApiSetName = "";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _typeApplication = "Application";
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Application.prototype, "_className", {
            get: function () {
                return "Application";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["name", "userName", "isVisible"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Name", "UserName", "IsVisible"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [false, true, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["activeDocument", "activePage", "documents"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "activeDocument", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.Document, this, "ActiveDocument", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "activePage", {
            get: function () {
                if (!this._Ac) {
                    this._Ac = _createPropertyObject(Visio.Page, this, "ActivePage", false, 4);
                }
                return this._Ac;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "documents", {
            get: function () {
                if (!this._D) {
                    this._D = _createPropertyObject(Visio.DocumentCollection, this, "Documents", true, 4);
                }
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "isVisible", {
            get: function () {
                _throwIfNotLoaded("isVisible", this._I, _typeApplication, this._isNull);
                return this._I;
            },
            set: function (value) {
                this._I = value;
                _invokeSetProperty(this, "IsVisible", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typeApplication, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "userName", {
            get: function () {
                _throwIfNotLoaded("userName", this._U, _typeApplication, this._isNull);
                return this._U;
            },
            set: function (value) {
                this._U = value;
                _invokeSetProperty(this, "UserName", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Application.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["userName", "isVisible"], ["activeDocument", "activePage"], [
                "documents"
            ]);
        };
        Application.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Application.prototype.getStencilInfo = function (stencilName, includeHiddenMasters) {
            return _invokeMethod(this, "GetStencilInfo", 1, [stencilName, includeHiddenMasters], 4, 0);
        };
        Application.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["IsVisible"])) {
                this._I = obj["IsVisible"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            if (!_isUndefined(obj["UserName"])) {
                this._U = obj["UserName"];
            }
            _handleNavigationPropertyResults(this, obj, ["activeDocument", "ActiveDocument", "activePage", "ActivePage", "documents", "Documents"]);
        };
        Application.prototype.load = function (options) {
            return _load(this, options);
        };
        Application.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Application.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Application.prototype.toJSON = function () {
            return _toJson(this, {
                "isVisible": this._I,
                "name": this._N,
                "userName": this._U
            }, {
                "activeDocument": this._A,
                "activePage": this._Ac,
                "documents": this._D
            });
        };
        Application.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Application.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Application;
    }(OfficeExtension.ClientObject));
    Visio.Application = Application;
    var _typeDocument = "Document";
    var Document = (function (_super) {
        __extends(Document, _super);
        function Document() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Document.prototype, "_className", {
            get: function () {
                return "Document";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["description", "fullName", "id", "index", "name"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Description", "FullName", "ID", "Index", "Name"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true, false, false, false, false];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["application", "pages"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "application", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.Application, this, "Application", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "pages", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(Visio.PageCollection, this, "Pages", true, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "description", {
            get: function () {
                _throwIfNotLoaded("description", this._D, _typeDocument, this._isNull);
                return this._D;
            },
            set: function (value) {
                this._D = value;
                _invokeSetProperty(this, "Description", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "fullName", {
            get: function () {
                _throwIfNotLoaded("fullName", this._F, _typeDocument, this._isNull);
                return this._F;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeDocument, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this._In, _typeDocument, this._isNull);
                return this._In;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typeDocument, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Document.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["description"], ["application"], [
                "pages"
            ]);
        };
        Document.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Document.prototype.close = function () {
            _invokeMethod(this, "Close", 1, [], 4, 0);
        };
        Document.prototype.getActivePage = function () {
            return _createMethodObject(Visio.Page, this, "GetActivePage", 1, [], false, false, null, 4);
        };
        Document.prototype.setActivePage = function (PageName) {
            _invokeMethod(this, "SetActivePage", 1, [PageName], 4, 0);
        };
        Document.prototype.showTaskPane = function (taskPaneType, initialProps, show) {
            _invokeMethod(this, "ShowTaskPane", 1, [taskPaneType, initialProps, show], 4, 0);
        };
        Document.prototype._RegisterDataVisualizerDiagramOperationCompletedEvent = function () {
            _invokeMethod(this, "_RegisterDataVisualizerDiagramOperationCompletedEvent", 0, [], 0, 0);
        };
        Document.prototype._RegisterSelectionChangedEvent = function () {
            _invokeMethod(this, "_RegisterSelectionChangedEvent", 0, [], 0, 0);
        };
        Document.prototype._RegisterShapeAddedEvent = function () {
            _invokeMethod(this, "_RegisterShapeAddedEvent", 0, [], 0, 0);
        };
        Document.prototype._UnregisterDataVisualizerDiagramOperationCompletedEvent = function () {
            _invokeMethod(this, "_UnregisterDataVisualizerDiagramOperationCompletedEvent", 0, [], 0, 0);
        };
        Document.prototype._UnregisterSelectionChangedEvent = function () {
            _invokeMethod(this, "_UnregisterSelectionChangedEvent", 0, [], 0, 0);
        };
        Document.prototype._UnregisterShapeAddedEvent = function () {
            _invokeMethod(this, "_UnregisterShapeAddedEvent", 0, [], 0, 0);
        };
        Document.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Description"])) {
                this._D = obj["Description"];
            }
            if (!_isUndefined(obj["FullName"])) {
                this._F = obj["FullName"];
            }
            if (!_isUndefined(obj["ID"])) {
                this._I = obj["ID"];
            }
            if (!_isUndefined(obj["Index"])) {
                this._In = obj["Index"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "pages", "Pages"]);
        };
        Document.prototype.load = function (options) {
            return _load(this, options);
        };
        Document.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Document.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Object.defineProperty(Document.prototype, "onDataVisualizerDiagramOperationCompleted", {
            get: function () {
                var _this = this;
                if (!this.m_dataVisualizerDiagramOperationCompleted) {
                    this.m_dataVisualizerDiagramOperationCompleted = new OfficeExtension.GenericEventHandlers(this.context, this, "DataVisualizerDiagramOperationCompleted", {
                        eventType: 3,
                        registerFunc: function () { _this._RegisterDataVisualizerDiagramOperationCompletedEvent(); },
                        unregisterFunc: function () { _this._UnregisterDataVisualizerDiagramOperationCompletedEvent(); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var newArgs = {
                                dataValidationErrors: value.dataValidationErrors,
                                diagramID: value.diagramId,
                                operationType: value.operationType,
                                pageID: value.pageId,
                                resultMessage: value.resultMessage,
                                resultTitle: value.resultTitle,
                                resultType: value.resultType,
                                tabularData: value.tabularData
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(newArgs);
                        }
                    });
                }
                return this.m_dataVisualizerDiagramOperationCompleted;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "SelectionChanged", {
                        eventType: 2,
                        registerFunc: function () { _this._RegisterSelectionChangedEvent(); },
                        unregisterFunc: function () { _this._UnregisterSelectionChangedEvent(); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var newArgs = {
                                pageID: value.pageID,
                                shapeIDs: value.shapeIDs
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(newArgs);
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onShapeAdded", {
            get: function () {
                var _this = this;
                if (!this.m_shapeAdded) {
                    this.m_shapeAdded = new OfficeExtension.EventHandlers(this.context, this, "ShapeAdded", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(1, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(1, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_shapeAdded;
            },
            enumerable: true,
            configurable: true
        });
        Document.prototype.toJSON = function () {
            return _toJson(this, {
                "description": this._D,
                "fullName": this._F,
                "id": this._I,
                "index": this._In,
                "name": this._N
            }, {
                "application": this._A,
                "pages": this._P
            });
        };
        Document.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Document.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Document;
    }(OfficeExtension.ClientObject));
    Visio.Document = Document;
    var _typeDocumentCollection = "DocumentCollection";
    var DocumentCollection = (function (_super) {
        __extends(DocumentCollection, _super);
        function DocumentCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DocumentCollection.prototype, "_className", {
            get: function () {
                return "DocumentCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeDocumentCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        DocumentCollection.prototype.add = function (FileName) {
            return _createMethodObject(Visio.Document, this, "Add", 1, [FileName], false, true, null, 4);
        };
        DocumentCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        DocumentCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Document, this, [key]);
        };
        DocumentCollection.prototype.getItemOrNullObject = function (index) {
            return _createMethodObject(Visio.Document, this, "GetItemOrNullObject", 1, [index], false, false, null, 4);
        };
        DocumentCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Document, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        DocumentCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        DocumentCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DocumentCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Document, true, _this, childItemData, index); });
        };
        DocumentCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        DocumentCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Document, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return DocumentCollection;
    }(OfficeExtension.ClientObject));
    Visio.DocumentCollection = DocumentCollection;
    var _typePage = "Page";
    var Page = (function (_super) {
        __extends(Page, _super);
        function Page() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Page.prototype, "_className", {
            get: function () {
                return "Page";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["index", "name", "id"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Index", "Name", "ID"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["shapes", "application", "document", "dataVisualizerDiagrams"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "application", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.Application, this, "Application", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "dataVisualizerDiagrams", {
            get: function () {
                if (!this._D) {
                    this._D = _createPropertyObject(Visio.DataVisualizerDiagramCollection, this, "DataVisualizerDiagrams", true, 4);
                }
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "document", {
            get: function () {
                if (!this._Do) {
                    this._Do = _createPropertyObject(Visio.Document, this, "Document", false, 4);
                }
                return this._Do;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "shapes", {
            get: function () {
                if (!this._S) {
                    this._S = _createPropertyObject(Visio.ShapeCollection, this, "Shapes", true, 4);
                }
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typePage, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this._In, _typePage, this._isNull);
                return this._In;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typePage, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Page.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["application", "document"], [
                "dataVisualizerDiagrams",
                "shapes"
            ]);
        };
        Page.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Page.prototype.boundingBox = function (Flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top) {
            _invokeMethod(this, "BoundingBox", 1, [Flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top], 4, 0);
        };
        Page.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ID"])) {
                this._I = obj["ID"];
            }
            if (!_isUndefined(obj["Index"])) {
                this._In = obj["Index"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "dataVisualizerDiagrams", "DataVisualizerDiagrams", "document", "Document", "shapes", "Shapes"]);
        };
        Page.prototype.load = function (options) {
            return _load(this, options);
        };
        Page.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Page.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Page.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I,
                "index": this._In,
                "name": this._N
            }, {
                "application": this._A,
                "dataVisualizerDiagrams": this._D,
                "document": this._Do,
                "shapes": this._S
            });
        };
        Page.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Page.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Page;
    }(OfficeExtension.ClientObject));
    Visio.Page = Page;
    var _typePageCollection = "PageCollection";
    var PageCollection = (function (_super) {
        __extends(PageCollection, _super);
        function PageCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PageCollection.prototype, "_className", {
            get: function () {
                return "PageCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typePageCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PageCollection.prototype.add = function (FileName) {
            return _createMethodObject(Visio.Page, this, "Add", 1, [FileName], false, true, null, 4);
        };
        PageCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        PageCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Page, this, [key]);
        };
        PageCollection.prototype.getItemOrNullObject = function (index) {
            return _createMethodObject(Visio.Page, this, "GetItemOrNullObject", 1, [index], false, false, null, 4);
        };
        PageCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Page, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PageCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        PageCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PageCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Page, true, _this, childItemData, index); });
        };
        PageCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        PageCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Page, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return PageCollection;
    }(OfficeExtension.ClientObject));
    Visio.PageCollection = PageCollection;
    var _typeShape = "Shape";
    var Shape = (function (_super) {
        __extends(Shape, _super);
        function Shape() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Shape.prototype, "_className", {
            get: function () {
                return "Shape";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["name", "id", "text", "isCallout", "isDataGraphicCallout", "isOpenForTextEdit", "objType", "isBoundToData"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Name", "Id", "Text", "IsCallout", "IsDataGraphicCallout", "IsOpenForTextEdit", "ObjType", "IsBoundToData"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["application", "document"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "application", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.Application, this, "Application", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "document", {
            get: function () {
                if (!this._D) {
                    this._D = _createPropertyObject(Visio.Document, this, "Document", false, 4);
                }
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeShape, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "isBoundToData", {
            get: function () {
                _throwIfNotLoaded("isBoundToData", this._Is, _typeShape, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "isCallout", {
            get: function () {
                _throwIfNotLoaded("isCallout", this._IsC, _typeShape, this._isNull);
                return this._IsC;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "isDataGraphicCallout", {
            get: function () {
                _throwIfNotLoaded("isDataGraphicCallout", this._IsD, _typeShape, this._isNull);
                return this._IsD;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "isOpenForTextEdit", {
            get: function () {
                _throwIfNotLoaded("isOpenForTextEdit", this._IsO, _typeShape, this._isNull);
                return this._IsO;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typeShape, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "objType", {
            get: function () {
                _throwIfNotLoaded("objType", this._O, _typeShape, this._isNull);
                return this._O;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this._T, _typeShape, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Shape.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["application", "document"], []);
        };
        Shape.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Shape.prototype.boundingBox = function (Flags, left, bottom, right, top) {
            _invokeMethod(this, "BoundingBox", 1, [Flags, left, bottom, right, top], 4, 0);
        };
        Shape.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
            if (!_isUndefined(obj["IsBoundToData"])) {
                this._Is = obj["IsBoundToData"];
            }
            if (!_isUndefined(obj["IsCallout"])) {
                this._IsC = obj["IsCallout"];
            }
            if (!_isUndefined(obj["IsDataGraphicCallout"])) {
                this._IsD = obj["IsDataGraphicCallout"];
            }
            if (!_isUndefined(obj["IsOpenForTextEdit"])) {
                this._IsO = obj["IsOpenForTextEdit"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            if (!_isUndefined(obj["ObjType"])) {
                this._O = obj["ObjType"];
            }
            if (!_isUndefined(obj["Text"])) {
                this._T = obj["Text"];
            }
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "document", "Document"]);
        };
        Shape.prototype.load = function (options) {
            return _load(this, options);
        };
        Shape.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Shape.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        Shape.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Shape.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I,
                "isBoundToData": this._Is,
                "isCallout": this._IsC,
                "isDataGraphicCallout": this._IsD,
                "isOpenForTextEdit": this._IsO,
                "name": this._N,
                "objType": this._O,
                "text": this._T
            }, {
                "application": this._A,
                "document": this._D
            });
        };
        Shape.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Shape.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Shape;
    }(OfficeExtension.ClientObject));
    Visio.Shape = Shape;
    var _typeShapeCollection = "ShapeCollection";
    var ShapeCollection = (function (_super) {
        __extends(ShapeCollection, _super);
        function ShapeCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ShapeCollection.prototype, "_className", {
            get: function () {
                return "ShapeCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeShapeCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ShapeCollection.prototype.add = function (FileName) {
            return _createMethodObject(Visio.Shape, this, "Add", 1, [FileName], false, true, null, 4);
        };
        ShapeCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        ShapeCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Shape, this, [key]);
        };
        ShapeCollection.prototype.getItemOrNullObject = function (index) {
            return _createMethodObject(Visio.Shape, this, "GetItemOrNullObject", 1, [index], false, false, null, 4);
        };
        ShapeCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Shape, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ShapeCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        ShapeCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ShapeCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Shape, true, _this, childItemData, index); });
        };
        ShapeCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        ShapeCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Shape, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return ShapeCollection;
    }(OfficeExtension.ClientObject));
    Visio.ShapeCollection = ShapeCollection;
    var _typeDataVisualizerDiagramCollection = "DataVisualizerDiagramCollection";
    var DataVisualizerDiagramCollection = (function (_super) {
        __extends(DataVisualizerDiagramCollection, _super);
        function DataVisualizerDiagramCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "_className", {
            get: function () {
                return "DataVisualizerDiagramCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeDataVisualizerDiagramCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        DataVisualizerDiagramCollection.prototype.add = function (data, settings) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "Add", 1, [data, settings], false, true, null, 4);
        };
        DataVisualizerDiagramCollection.prototype.addPreferred = function (data, diagramType) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "AddPreferred", 0, [data, diagramType], false, false, null, 0);
        };
        DataVisualizerDiagramCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        DataVisualizerDiagramCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.DataVisualizerDiagram, this, [key]);
        };
        DataVisualizerDiagramCollection.prototype.getItemAt = function (index) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "GetItemAt", 1, [index], false, false, null, 4);
        };
        DataVisualizerDiagramCollection.prototype.getItemOrNullObject = function (key) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
        };
        DataVisualizerDiagramCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.DataVisualizerDiagram, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        DataVisualizerDiagramCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        DataVisualizerDiagramCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DataVisualizerDiagramCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.DataVisualizerDiagram, true, _this, childItemData, index); });
        };
        DataVisualizerDiagramCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        DataVisualizerDiagramCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.DataVisualizerDiagram, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return DataVisualizerDiagramCollection;
    }(OfficeExtension.ClientObject));
    Visio.DataVisualizerDiagramCollection = DataVisualizerDiagramCollection;
    var _typeDataVisualizerDiagram = "DataVisualizerDiagram";
    var DataVisualizerDiagram = (function (_super) {
        __extends(DataVisualizerDiagram, _super);
        function DataVisualizerDiagram() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DataVisualizerDiagram.prototype, "_className", {
            get: function () {
                return "DataVisualizerDiagram";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id", "mappings", "type", "dataHeaders"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["ID", "Mappings", "Type", "DataHeaders"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["page"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "page", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(Visio.Page, this, "Page", false, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "dataHeaders", {
            get: function () {
                _throwIfNotLoaded("dataHeaders", this._D, _typeDataVisualizerDiagram, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeDataVisualizerDiagram, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "mappings", {
            get: function () {
                _throwIfNotLoaded("mappings", this._M, _typeDataVisualizerDiagram, this._isNull);
                return this._M;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this._T, _typeDataVisualizerDiagram, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        DataVisualizerDiagram.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["page"], []);
        };
        DataVisualizerDiagram.prototype["delete"] = function () {
            _invokeMethod(this, "Delete", 1, [], 4, 0);
        };
        DataVisualizerDiagram.prototype.getDataColumnValuesAsString = function (columnName) {
            return _invokeMethod(this, "GetDataColumnValuesAsString", 1, [columnName], 4, 0);
        };
        DataVisualizerDiagram.prototype.setConnection = function (connectionInfo) {
            _invokeMethod(this, "SetConnection", 1, [connectionInfo], 4, 0);
        };
        DataVisualizerDiagram.prototype.update = function (data, mappings, ignoreConflicts) {
            _invokeMethod(this, "Update", 1, [data, mappings, ignoreConflicts], 4, 0);
        };
        DataVisualizerDiagram.prototype.updateMappings = function (mappings) {
            _invokeMethod(this, "UpdateMappings", 1, [mappings], 4, 0);
        };
        DataVisualizerDiagram.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["DataHeaders"])) {
                this._D = obj["DataHeaders"];
            }
            if (!_isUndefined(obj["ID"])) {
                this._I = obj["ID"];
            }
            if (!_isUndefined(obj["Mappings"])) {
                this._M = obj["Mappings"];
            }
            if (!_isUndefined(obj["Type"])) {
                this._T = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["page", "Page"]);
        };
        DataVisualizerDiagram.prototype.load = function (options) {
            return _load(this, options);
        };
        DataVisualizerDiagram.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DataVisualizerDiagram.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        DataVisualizerDiagram.prototype.toJSON = function () {
            return _toJson(this, {
                "dataHeaders": this._D,
                "id": this._I,
                "mappings": this._M,
                "type": this._T
            }, {
                "page": this._P
            });
        };
        DataVisualizerDiagram.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        DataVisualizerDiagram.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return DataVisualizerDiagram;
    }(OfficeExtension.ClientObject));
    Visio.DataVisualizerDiagram = DataVisualizerDiagram;
    var EventType;
    (function (EventType) {
        EventType["shapeAdded"] = "ShapeAdded";
        EventType["selectionChanged"] = "SelectionChanged";
        EventType["dataVisualizerDiagramOperationCompleted"] = "DataVisualizerDiagramOperationCompleted";
    })(EventType = Visio.EventType || (Visio.EventType = {}));
    var TaskPaneType;
    (function (TaskPaneType) {
        TaskPaneType["none"] = "None";
        TaskPaneType["dataVisualizerProcessMappings"] = "DataVisualizerProcessMappings";
        TaskPaneType["dataVisualizerOrgChartMappings"] = "DataVisualizerOrgChartMappings";
    })(TaskPaneType = Visio.TaskPaneType || (Visio.TaskPaneType = {}));
    var DataVisualizerDiagramResultType;
    (function (DataVisualizerDiagramResultType) {
        DataVisualizerDiagramResultType["success"] = "Success";
        DataVisualizerDiagramResultType["unexpected"] = "Unexpected";
        DataVisualizerDiagramResultType["validationError"] = "ValidationError";
        DataVisualizerDiagramResultType["conflictError"] = "ConflictError";
    })(DataVisualizerDiagramResultType = Visio.DataVisualizerDiagramResultType || (Visio.DataVisualizerDiagramResultType = {}));
    var DataVisualizerDiagramOperationType;
    (function (DataVisualizerDiagramOperationType) {
        DataVisualizerDiagramOperationType["unknown"] = "Unknown";
        DataVisualizerDiagramOperationType["create"] = "Create";
        DataVisualizerDiagramOperationType["updateMappings"] = "UpdateMappings";
        DataVisualizerDiagramOperationType["updateData"] = "UpdateData";
        DataVisualizerDiagramOperationType["update"] = "Update";
        DataVisualizerDiagramOperationType["delete"] = "Delete";
    })(DataVisualizerDiagramOperationType = Visio.DataVisualizerDiagramOperationType || (Visio.DataVisualizerDiagramOperationType = {}));
    var DataVisualizerDiagramType;
    (function (DataVisualizerDiagramType) {
        DataVisualizerDiagramType["unknown"] = "Unknown";
        DataVisualizerDiagramType["basicFlowchart"] = "BasicFlowchart";
        DataVisualizerDiagramType["crossFunctionalFlowchart_Horizontal"] = "CrossFunctionalFlowchart_Horizontal";
        DataVisualizerDiagramType["crossFunctionalFlowchart_Vertical"] = "CrossFunctionalFlowchart_Vertical";
        DataVisualizerDiagramType["audit"] = "Audit";
        DataVisualizerDiagramType["orgChart"] = "OrgChart";
        DataVisualizerDiagramType["network"] = "Network";
    })(DataVisualizerDiagramType = Visio.DataVisualizerDiagramType || (Visio.DataVisualizerDiagramType = {}));
    var ColumnType;
    (function (ColumnType) {
        ColumnType["unknown"] = "Unknown";
        ColumnType["string"] = "String";
        ColumnType["number"] = "Number";
        ColumnType["date"] = "Date";
        ColumnType["currency"] = "Currency";
    })(ColumnType = Visio.ColumnType || (Visio.ColumnType = {}));
    var DataSourceType;
    (function (DataSourceType) {
        DataSourceType["unknown"] = "Unknown";
        DataSourceType["excel"] = "Excel";
    })(DataSourceType = Visio.DataSourceType || (Visio.DataSourceType = {}));
    var CrossFunctionalFlowchartOrientation;
    (function (CrossFunctionalFlowchartOrientation) {
        CrossFunctionalFlowchartOrientation["horizontal"] = "Horizontal";
        CrossFunctionalFlowchartOrientation["vertical"] = "Vertical";
    })(CrossFunctionalFlowchartOrientation = Visio.CrossFunctionalFlowchartOrientation || (Visio.CrossFunctionalFlowchartOrientation = {}));
    var LayoutVariant;
    (function (LayoutVariant) {
        LayoutVariant["unknown"] = "Unknown";
        LayoutVariant["pageDefault"] = "PageDefault";
        LayoutVariant["flowchart_TopToBottom"] = "Flowchart_TopToBottom";
        LayoutVariant["flowchart_LeftToRight"] = "Flowchart_LeftToRight";
        LayoutVariant["radial"] = "Radial";
        LayoutVariant["flowchart_BottomToTop"] = "Flowchart_BottomToTop";
        LayoutVariant["flowchart_RightToLeft"] = "Flowchart_RightToLeft";
        LayoutVariant["circular"] = "Circular";
        LayoutVariant["wideTree_DownThenRight"] = "WideTree_DownThenRight";
        LayoutVariant["wideTree_RightThenDown"] = "WideTree_RightThenDown";
        LayoutVariant["wideTree_RightThenUp"] = "WideTree_RightThenUp";
        LayoutVariant["wideTree_UpThenRight"] = "WideTree_UpThenRight";
        LayoutVariant["wideTree_UpThenLeft"] = "WideTree_UpThenLeft";
        LayoutVariant["wideTree_LeftThenUp"] = "WideTree_LeftThenUp";
        LayoutVariant["wideTree_LeftThenDown"] = "WideTree_LeftThenDown";
        LayoutVariant["wideTree_DownThenLeft"] = "WideTree_DownThenLeft";
        LayoutVariant["parentDefault"] = "ParentDefault";
        LayoutVariant["hierarchy_TopToBottomLeft"] = "Hierarchy_TopToBottomLeft";
        LayoutVariant["hierarchy_TopToBottomCenter"] = "Hierarchy_TopToBottomCenter";
        LayoutVariant["hierarchy_TopToBottomRight"] = "Hierarchy_TopToBottomRight";
        LayoutVariant["hierarchy_BottomToTopLeft"] = "Hierarchy_BottomToTopLeft";
        LayoutVariant["hierarchy_BottomToTopCenter"] = "Hierarchy_BottomToTopCenter";
        LayoutVariant["hierarchy_BottomToTopRight"] = "Hierarchy_BottomToTopRight";
        LayoutVariant["hierarchy_LeftToRightTop"] = "Hierarchy_LeftToRightTop";
        LayoutVariant["hierarchy_LeftToRightMiddle"] = "Hierarchy_LeftToRightMiddle";
        LayoutVariant["hierarchy_LeftToRightBottom"] = "Hierarchy_LeftToRightBottom";
        LayoutVariant["hierarchy_RightToLeftTop"] = "Hierarchy_RightToLeftTop";
        LayoutVariant["hierarchy_RightToLeftMiddle"] = "Hierarchy_RightToLeftMiddle";
        LayoutVariant["hierarchy_RightToLeftBottom"] = "Hierarchy_RightToLeftBottom";
        LayoutVariant["orgChart_HorizontalCenter"] = "OrgChart_HorizontalCenter";
        LayoutVariant["orgChart_HorizontalCenter_LeftToRight"] = "OrgChart_HorizontalCenter_LeftToRight";
        LayoutVariant["orgChart_Hybrid_HorizontalCenter_VerticalRight"] = "OrgChart_Hybrid_HorizontalCenter_VerticalRight";
        LayoutVariant["orgChart_SideBySide"] = "OrgChart_SideBySide";
    })(LayoutVariant = Visio.LayoutVariant || (Visio.LayoutVariant = {}));
    var DataValidationErrorType;
    (function (DataValidationErrorType) {
        DataValidationErrorType["none"] = "None";
        DataValidationErrorType["columnNotMapped"] = "ColumnNotMapped";
        DataValidationErrorType["uniqueIdColumnError"] = "UniqueIdColumnError";
        DataValidationErrorType["swimlaneColumnError"] = "SwimlaneColumnError";
        DataValidationErrorType["delimiterError"] = "DelimiterError";
        DataValidationErrorType["connectorColumnError"] = "ConnectorColumnError";
        DataValidationErrorType["connectorColumnMappedElsewhere"] = "ConnectorColumnMappedElsewhere";
        DataValidationErrorType["connectorLabelColumnMappedElsewhere"] = "ConnectorLabelColumnMappedElsewhere";
        DataValidationErrorType["connectorColumnAndConnectorLabelMappedElsewhere"] = "ConnectorColumnAndConnectorLabelMappedElsewhere";
    })(DataValidationErrorType = Visio.DataValidationErrorType || (Visio.DataValidationErrorType = {}));
    var ConnectorDirection;
    (function (ConnectorDirection) {
        ConnectorDirection["fromTarget"] = "FromTarget";
        ConnectorDirection["toTarget"] = "ToTarget";
    })(ConnectorDirection = Visio.ConnectorDirection || (Visio.ConnectorDirection = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["generalException"] = "GeneralException";
    })(ErrorCodes = Visio.ErrorCodes || (Visio.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = Visio.Interfaces || (Visio.Interfaces = {}));
})(Visio || (Visio = {}));
var Visio;
(function (Visio) {
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this.m_document = new Visio.Document(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
            _this._rootObject = _this.m_document;
            return _this;
        }
        Object.defineProperty(RequestContext.prototype, "document", {
            get: function () {
                return this.m_document;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeCore.RequestContext));
    Visio.RequestContext = RequestContext;
    function run(arg1, arg2, arg3) {
        return OfficeExtension.ClientRequestContext._runBatch("Visio.run", arguments, function (requestInfo) {
            var ret = new Visio.RequestContext(requestInfo);
            return ret;
        });
    }
    Visio.run = run;
})(Visio || (Visio = {}));
OSFAriaLogger.AriaLogger.EnableSendingTelemetryWithOTel = true;
OSFAriaLogger.AriaLogger.EnableSendingTelemetryWithLegacyAria = false;
