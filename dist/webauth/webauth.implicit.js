!function(e,t){"object"==typeof exports&&"object"==typeof module?module.exports=t():"function"==typeof define&&define.amd?define("Implicit",[],t):"object"==typeof exports?exports.Implicit=t():e.Implicit=t()}(window,function(){return function(e){var t={};function r(o){if(t[o])return t[o].exports;var n=t[o]={i:o,l:!1,exports:{}};return e[o].call(n.exports,n,n.exports,r),n.l=!0,n.exports}return r.m=e,r.c=t,r.d=function(e,t,o){r.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:o})},r.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},r.t=function(e,t){if(1&t&&(e=r(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var o=Object.create(null);if(r.r(o),Object.defineProperty(o,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var n in e)r.d(o,n,function(t){return e[t]}.bind(null,n));return o},r.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return r.d(t,"a",t),t},r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},r.p="",r(r.s=0)}({"./packages/Microsoft.Office.WebAuth.Implicit/lib/api.js":function(e,t,r){"use strict";Object.defineProperty(t,"__esModule",{value:!0}),t.logUserAction=function(e,t,r,o,n,i,a){void 0===t&&(t=!0);void 0===r&&(r=c);void 0===o&&(o=c);void 0===n&&(n=c);void 0===i&&(i=0);void 0===a&&(a=[]);l({name:null,actionName:e,commandSurface:n,parentName:r,triggerMethod:o,durationMs:i,succeeded:t,dataFields:a})},t.logActivity=function(e,t,r,o){void 0===t&&(t=!0);void 0===r&&(r=0);void 0===o&&(o=[]);u({name:e,succeeded:t,durationMs:r,dataFields:o})},t.sendTelemetryEvent=function(e){h({kind:"event",event:e,timestamp:(new Date).getTime()})},t.sendActivityEvent=u,t.sendOtelEvent=function(e){h({kind:"otel",event:e})},t.sendUserActionEvent=l,t.addNamespaceMapping=function(e,t){h({kind:"addNamespaceMapping",namespace:e,ariaTenantToken:t})},t.setEnabledState=function(e){(n=e)||(i=[])},t.shutdown=function(){return h({kind:"shutdown"}),i.length+a},t.registerEventHandler=function(e){o=e,i.forEach(function(e){return h(e)}),i=[]};var o,n=!1,i=[],a=0,s=2e4,c="Unknown";function u(e){h({kind:"activity",event:e,timestamp:(new Date).getTime()})}function l(e){h({kind:"action",event:e,timestamp:(new Date).getTime()})}function h(e){n&&(o?o(e):i.length<=s?i.push(e):a+=1)}},"./packages/Microsoft.Office.WebAuth.Implicit/lib/msal.js":function(e,t,r){"use strict";
/*! msal v1.3.3 2020-07-14 */
/*! msal v1.3.3 2020-07-14 */
var o;window,o=function(){return function(e){var t={};function r(o){if(t[o])return t[o].exports;var n=t[o]={i:o,l:!1,exports:{}};return e[o].call(n.exports,n,n.exports,r),n.l=!0,n.exports}return r.m=e,r.c=t,r.d=function(e,t,o){r.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:o})},r.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},r.t=function(e,t){if(1&t&&(e=r(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var o=Object.create(null);if(r.r(o),Object.defineProperty(o,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var n in e)r.d(o,n,function(t){return e[t]}.bind(null,n));return o},r.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return r.d(t,"a",t),t},r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},r.p="",r(r.s=29)}([function(e,t,r){
/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
Object.defineProperty(t,"__esModule",{value:!0});var o=function(e,t){return(o=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var r in t)t.hasOwnProperty(r)&&(e[r]=t[r])})(e,t)};function n(e){var t="function"==typeof Symbol&&e[Symbol.iterator],r=0;return t?t.call(e):{next:function(){return e&&r>=e.length&&(e=void 0),{value:e&&e[r++],done:!e}}}}function i(e,t){var r="function"==typeof Symbol&&e[Symbol.iterator];if(!r)return e;var o,n,i=r.call(e),a=[];try{for(;(void 0===t||t-- >0)&&!(o=i.next()).done;)a.push(o.value)}catch(e){n={error:e}}finally{try{o&&!o.done&&(r=i.return)&&r.call(i)}finally{if(n)throw n.error}}return a}function a(e){return this instanceof a?(this.v=e,this):new a(e)}t.__extends=function(e,t){function r(){this.constructor=e}o(e,t),e.prototype=null===t?Object.create(t):(r.prototype=t.prototype,new r)},t.__assign=function(){return t.__assign=Object.assign||function(e){for(var t,r=1,o=arguments.length;r<o;r++)for(var n in t=arguments[r])Object.prototype.hasOwnProperty.call(t,n)&&(e[n]=t[n]);return e},t.__assign.apply(this,arguments)},t.__rest=function(e,t){var r={};for(var o in e)Object.prototype.hasOwnProperty.call(e,o)&&t.indexOf(o)<0&&(r[o]=e[o]);if(null!=e&&"function"==typeof Object.getOwnPropertySymbols){var n=0;for(o=Object.getOwnPropertySymbols(e);n<o.length;n++)t.indexOf(o[n])<0&&Object.prototype.propertyIsEnumerable.call(e,o[n])&&(r[o[n]]=e[o[n]])}return r},t.__decorate=function(e,t,r,o){var n,i=arguments.length,a=i<3?t:null===o?o=Object.getOwnPropertyDescriptor(t,r):o;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)a=Reflect.decorate(e,t,r,o);else for(var s=e.length-1;s>=0;s--)(n=e[s])&&(a=(i<3?n(a):i>3?n(t,r,a):n(t,r))||a);return i>3&&a&&Object.defineProperty(t,r,a),a},t.__param=function(e,t){return function(r,o){t(r,o,e)}},t.__metadata=function(e,t){if("object"==typeof Reflect&&"function"==typeof Reflect.metadata)return Reflect.metadata(e,t)},t.__awaiter=function(e,t,r,o){return new(r||(r=Promise))(function(n,i){function a(e){try{c(o.next(e))}catch(e){i(e)}}function s(e){try{c(o.throw(e))}catch(e){i(e)}}function c(e){e.done?n(e.value):new r(function(t){t(e.value)}).then(a,s)}c((o=o.apply(e,t||[])).next())})},t.__generator=function(e,t){var r,o,n,i,a={label:0,sent:function(){if(1&n[0])throw n[1];return n[1]},trys:[],ops:[]};return i={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function s(i){return function(s){return function(i){if(r)throw new TypeError("Generator is already executing.");for(;a;)try{if(r=1,o&&(n=2&i[0]?o.return:i[0]?o.throw||((n=o.return)&&n.call(o),0):o.next)&&!(n=n.call(o,i[1])).done)return n;switch(o=0,n&&(i=[2&i[0],n.value]),i[0]){case 0:case 1:n=i;break;case 4:return a.label++,{value:i[1],done:!1};case 5:a.label++,o=i[1],i=[0];continue;case 7:i=a.ops.pop(),a.trys.pop();continue;default:if(!(n=(n=a.trys).length>0&&n[n.length-1])&&(6===i[0]||2===i[0])){a=0;continue}if(3===i[0]&&(!n||i[1]>n[0]&&i[1]<n[3])){a.label=i[1];break}if(6===i[0]&&a.label<n[1]){a.label=n[1],n=i;break}if(n&&a.label<n[2]){a.label=n[2],a.ops.push(i);break}n[2]&&a.ops.pop(),a.trys.pop();continue}i=t.call(e,a)}catch(e){i=[6,e],o=0}finally{r=n=0}if(5&i[0])throw i[1];return{value:i[0]?i[1]:void 0,done:!0}}([i,s])}}},t.__exportStar=function(e,t){for(var r in e)t.hasOwnProperty(r)||(t[r]=e[r])},t.__values=n,t.__read=i,t.__spread=function(){for(var e=[],t=0;t<arguments.length;t++)e=e.concat(i(arguments[t]));return e},t.__spreadArrays=function(){for(var e=0,t=0,r=arguments.length;t<r;t++)e+=arguments[t].length;var o=Array(e),n=0;for(t=0;t<r;t++)for(var i=arguments[t],a=0,s=i.length;a<s;a++,n++)o[n]=i[a];return o},t.__await=a,t.__asyncGenerator=function(e,t,r){if(!Symbol.asyncIterator)throw new TypeError("Symbol.asyncIterator is not defined.");var o,n=r.apply(e,t||[]),i=[];return o={},s("next"),s("throw"),s("return"),o[Symbol.asyncIterator]=function(){return this},o;function s(e){n[e]&&(o[e]=function(t){return new Promise(function(r,o){i.push([e,t,r,o])>1||c(e,t)})})}function c(e,t){try{(r=n[e](t)).value instanceof a?Promise.resolve(r.value.v).then(u,l):h(i[0][2],r)}catch(e){h(i[0][3],e)}var r}function u(e){c("next",e)}function l(e){c("throw",e)}function h(e,t){e(t),i.shift(),i.length&&c(i[0][0],i[0][1])}},t.__asyncDelegator=function(e){var t,r;return t={},o("next"),o("throw",function(e){throw e}),o("return"),t[Symbol.iterator]=function(){return this},t;function o(o,n){t[o]=e[o]?function(t){return(r=!r)?{value:a(e[o](t)),done:"return"===o}:n?n(t):t}:n}},t.__asyncValues=function(e){if(!Symbol.asyncIterator)throw new TypeError("Symbol.asyncIterator is not defined.");var t,r=e[Symbol.asyncIterator];return r?r.call(e):(e=n(e),t={},o("next"),o("throw"),o("return"),t[Symbol.asyncIterator]=function(){return this},t);function o(r){t[r]=e[r]&&function(t){return new Promise(function(o,n){!function(e,t,r,o){Promise.resolve(o).then(function(t){e({value:t,done:r})},t)}(o,n,(t=e[r](t)).done,t.value)})}}},t.__makeTemplateObject=function(e,t){return Object.defineProperty?Object.defineProperty(e,"raw",{value:t}):e.raw=t,e},t.__importStar=function(e){if(e&&e.__esModule)return e;var t={};if(null!=e)for(var r in e)Object.hasOwnProperty.call(e,r)&&(t[r]=e[r]);return t.default=e,t},t.__importDefault=function(e){return e&&e.__esModule?e:{default:e}}},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o,n=function(){function e(){}return Object.defineProperty(e,"libraryName",{get:function(){return"Msal.js"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"claims",{get:function(){return"claims"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"clientId",{get:function(){return"clientId"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"adalIdToken",{get:function(){return"adal.idtoken"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"cachePrefix",{get:function(){return"msal"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"scopes",{get:function(){return"scopes"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"no_account",{get:function(){return"NO_ACCOUNT"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"upn",{get:function(){return"upn"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"domain_hint",{get:function(){return"domain_hint"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"prompt_select_account",{get:function(){return"&prompt=select_account"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"prompt_none",{get:function(){return"&prompt=none"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"prompt",{get:function(){return"prompt"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"response_mode_fragment",{get:function(){return"&response_mode=fragment"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"resourceDelimiter",{get:function(){return"|"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"cacheDelimiter",{get:function(){return"."},enumerable:!0,configurable:!0}),Object.defineProperty(e,"popUpWidth",{get:function(){return this._popUpWidth},set:function(e){this._popUpWidth=e},enumerable:!0,configurable:!0}),Object.defineProperty(e,"popUpHeight",{get:function(){return this._popUpHeight},set:function(e){this._popUpHeight=e},enumerable:!0,configurable:!0}),Object.defineProperty(e,"login",{get:function(){return"LOGIN"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"renewToken",{get:function(){return"RENEW_TOKEN"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"unknown",{get:function(){return"UNKNOWN"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"ADFS",{get:function(){return"adfs"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"homeAccountIdentifier",{get:function(){return"homeAccountIdentifier"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"common",{get:function(){return"common"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"openidScope",{get:function(){return"openid"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"profileScope",{get:function(){return"profile"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"interactionTypeRedirect",{get:function(){return"redirectInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"interactionTypePopup",{get:function(){return"popupInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"interactionTypeSilent",{get:function(){return"silentInteraction"},enumerable:!0,configurable:!0}),Object.defineProperty(e,"inProgress",{get:function(){return"inProgress"},enumerable:!0,configurable:!0}),e._popUpWidth=483,e._popUpHeight=600,e}();t.Constants=n,function(e){e.SCOPE="scope",e.STATE="state",e.ERROR="error",e.ERROR_DESCRIPTION="error_description",e.ACCESS_TOKEN="access_token",e.ID_TOKEN="id_token",e.EXPIRES_IN="expires_in",e.SESSION_STATE="session_state",e.CLIENT_INFO="client_info"}(t.ServerHashParamKeys||(t.ServerHashParamKeys={})),function(e){e.AUTHORITY="authority",e.ACQUIRE_TOKEN_ACCOUNT="acquireTokenAccount",e.SESSION_STATE="session.state",e.STATE_LOGIN="state.login",e.STATE_ACQ_TOKEN="state.acquireToken",e.STATE_RENEW="state.renew",e.NONCE_IDTOKEN="nonce.idtoken",e.LOGIN_REQUEST="login.request",e.RENEW_STATUS="token.renew.status",e.URL_HASH="urlHash",e.INTERACTION_STATUS="interaction_status",e.REDIRECT_REQUEST="redirect_request"}(t.TemporaryCacheKeys||(t.TemporaryCacheKeys={})),function(e){e.IDTOKEN="idtoken",e.CLIENT_INFO="client.info"}(t.PersistentCacheKeys||(t.PersistentCacheKeys={})),function(e){e.LOGIN_ERROR="login.error",e.ERROR="error",e.ERROR_DESC="error.description"}(t.ErrorCacheKeys||(t.ErrorCacheKeys={})),t.DEFAULT_AUTHORITY="https://login.microsoftonline.com/common/",t.AAD_INSTANCE_DISCOVERY_ENDPOINT=t.DEFAULT_AUTHORITY+"/discovery/instance?api-version=1.1&authorization_endpoint=",function(e){e.ACCOUNT="account",e.SID="sid",e.LOGIN_HINT="login_hint",e.ID_TOKEN="id_token",e.ACCOUNT_ID="accountIdentifier",e.HOMEACCOUNT_ID="homeAccountIdentifier"}(o=t.SSOTypes||(t.SSOTypes={})),t.BlacklistedEQParams=[o.SID,o.LOGIN_HINT],t.NetworkRequestType={GET:"GET",POST:"POST"},t.PromptState={LOGIN:"login",SELECT_ACCOUNT:"select_account",CONSENT:"consent",NONE:"none"},t.FramePrefix={ID_TOKEN_FRAME:"msalIdTokenFrame",TOKEN_FRAME:"msalRenewFrame"},t.libraryVersion=function(){return"1.3.3"}},function(e,t,r){Object.defineProperty(t,"__esModule",{value:!0});var o=function(){function e(){}return e.createNewGuid=function(){var t=window.crypto;if(t&&t.getRandomValues){var r=new Uint8Array(16);return t.getRandomValues(r),r[6]|=64,r[6]&=79,r[8]|=128,r[8]&=191,e.decimalToHex(r[0])+e.decimalToHex(r[1])+e.decimalToHex(r[2])+e.decimalToHex(r[3])+"-"+e.decimalToHex(r[4])+e.decimalToHex(r[5])+"-"+e.decimalToHex(r[6])+e.decimalToHex(r[7])+"-"+e.decimalToHex(r[8])+e.decimalToHex(r[9])+"-"+e.decimalToHex(r[10])+e.decimalToHex(r[11])+e.decimalToHex(r[12])+e.decimalToHex(r[13])+e.decimalToHex(r[14])+e.decimalToHex(r[15])}for(var o="xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx",n="0123456789abcdef",i=0,a="",s=0;s<36;s++)"-"!==o[s]&&"4"!==o[s]&&(i=16*Math.random()|0),"x"===o[s]?a+=n[i]:"y"===o[s]?(i&=3,a+=n[i|=8]):a+=o[s];return a},e.isGuid=function(e){return/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(e)},e.decimalToHex=function(e){for(var t=e.toString(16);t.length<2;)t="0"+t;return t},e.base64Encode=function(e){return btoa(encodeURIComponent(e).replace(/%([0-9A-F]{2})/g,function(e,t){return String.fromCharCode(Number("0x"+t))}))},e.base64Decode=function(e){var t=e.replace(/-/g,"+").replace(/_/g,"/");switch(t.length%4){case 0:break;case 2:t+="==";break;case 3:t+="=";br