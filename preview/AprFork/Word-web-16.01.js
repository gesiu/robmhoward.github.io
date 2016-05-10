/* Word web application specific API library */
/* Version: 16.0.6929.3000 */
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

var __extends=this&&this.__extends||function(n,t){function r(){this.constructor=n}for(var i in t)t.hasOwnProperty(i)&&(n[i]=t[i]);n.prototype=t===null?Object.create(t):(r.prototype=t.prototype,new r)},OsfMsAjaxFactory,OSF,OSFLog,Logger,OSFAppTelemetry,OSF_DDA_Marshaling_FilePropertiesKeys,OSF_DDA_Marshaling_File_FilePropertiesKeys,OSF_DDA_Marshaling_File_SlicePropertiesKeys,OSF_DDA_Marshaling_File_FileType,OSF_DDA_Marshaling_File_ParameterKeys,OSF_DDA_Marshaling_GoToType,OSF_DDA_Marshaling_SelectionMode,OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys,OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys,OfficeExt,OSF_DDA_Marshaling_Dialog_DialogMessageReceivedEventKeys,OSFWordWAC,OfficeExtension,Word;(function(n){var t=function(){function t(){}var i=null,n=!0;return t.prototype.isMsAjaxLoaded=function(){var t="function",i="undefined";return typeof Sys!==i&&typeof Type!==i&&Sys.StringBuilder&&typeof Sys.StringBuilder===t&&Type.registerNamespace&&typeof Type.registerNamespace===t&&Type.registerClass&&typeof Type.registerClass===t&&typeof Function._validateParams===t?n:!1},t.prototype.loadMsAjaxFull=function(n){var t=(window.location.protocol.toLowerCase()==="https:"?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(t,n)},Object.defineProperty(t.prototype,"msAjaxError",{get:function(){var n=this;return n._msAjaxError==i&&n.isMsAjaxLoaded()&&(n._msAjaxError=Error),n._msAjaxError},set:function(n){this._msAjaxError=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxSerializer",{get:function(){var n=this;return n._msAjaxSerializer==i&&n.isMsAjaxLoaded()&&(n._msAjaxSerializer=Sys.Serialization.JavaScriptSerializer),n._msAjaxSerializer},set:function(n){this._msAjaxSerializer=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxString",{get:function(){var n=this;return n._msAjaxString==i&&n.isMsAjaxLoaded()&&(n._msAjaxSerializer=String),n._msAjaxString},set:function(n){this._msAjaxString=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxDebug",{get:function(){var n=this;return n._msAjaxDebug==i&&n.isMsAjaxLoaded()&&(n._msAjaxDebug=Sys.Debug),n._msAjaxDebug},set:function(n){this._msAjaxDebug=n},enumerable:n,configurable:n}),t}();n.MicrosoftAjaxFactory=t})(OfficeExt||(OfficeExt={}));OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory;OSF=OSF||{},function(n){var t=function(){function n(n){this._internalStorage=n}return n.prototype.getItem=function(n){try{return this._internalStorage&&this._internalStorage.getItem(n)}catch(t){return null}},n.prototype.setItem=function(n,t){try{this._internalStorage&&this._internalStorage.setItem(n,t)}catch(i){}},n.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(n){}},n.prototype.removeItem=function(n){try{this._internalStorage&&this._internalStorage.removeItem(n)}catch(t){}},n.prototype.getKeysWithPrefix=function(n){var r=[],u,t,i;try{for(u=this._internalStorage&&this._internalStorage.length||0,t=0;t<u;t++)i=this._internalStorage.key(t),i.indexOf(n)===0&&r.push(i)}catch(f){}return r},n}();n.SafeStorage=t}(OfficeExt||(OfficeExt={}));OSF.XdmFieldName={ConversationUrl:"ConversationUrl",AppId:"AppId"};OSF.OUtil=function(){function d(){var n=s*Math.random();return n^=r^(new Date).getMilliseconds()<<Math.floor(Math.random()*21),n.toString(16)}function g(){if(!v){try{var t=window.sessionStorage}catch(i){t=n}v=new OfficeExt.SafeStorage(t)}return v}var u="on",p="configurable",w="writable",f="enumerable",o="undefined",i=!0,t=!1,s=2147483647,n=null,h=-1,b="&_xdm_Info=",c="&_serializer_version=",nt="&_app_context=",k="_xdm_",tt="_serializer_version=",l="#",e="class",a={},it=3e4,v=n,y=n,r=(new Date).getTime();return{set_entropy:function(n){var t,u,i;if(typeof n=="string")for(t=0;t<n.length;t+=4){for(u=0,i=0;i<4&&t+i<n.length;i++)u=(u<<8)+n.charCodeAt(t+i);r^=u}else r^=typeof n=="number"?n:s*Math.random();r&=s},extend:function(n,t){var i=function(){};i.prototype=t.prototype;n.prototype=new i;n.prototype.constructor=n;n.uber=t.prototype;t.prototype.constructor===Object.prototype.constructor&&(t.prototype.constructor=t)},setNamespace:function(n,t){t&&n&&!t[n]&&(t[n]={})},unsetNamespace:function(n,t){t&&n&&t[n]&&delete t[n]},loadScript:function(r,u,f){var s,e,o,h,c;r&&u&&(s=window.document,e=a[r],e?e.loaded?u():e.pendingCallbacks.push(u):(o=s.createElement("script"),o.type="text/javascript",e={loaded:t,pendingCallbacks:[u],timer:n},a[r]=e,h=function(){var r,t,u;for(e.timer!=n&&(clearTimeout(e.timer),delete e.timer),e.loaded=i,r=e.pendingCallbacks.length,t=0;t<r;t++)u=e.pendingCallbacks.shift(),u()},c=function(){var i,t,u;for(delete a[r],e.timer!=n&&(clearTimeout(e.timer),delete e.timer),i=e.pendingCallbacks.length,t=0;t<i;t++)u=e.pendingCallbacks.shift(),u()},o.readyState?o.onreadystatechange=function(){(o.readyState=="loaded"||o.readyState=="complete")&&(o.onreadystatechange=n,h())}:o.onload=h,o.onerror=c,f=f||it,e.timer=setTimeout(c,f),o.src=r,s.getElementsByTagName("head")[0].appendChild(o)))},loadCSS:function(n){if(n){var i=window.document,t=i.createElement("link");t.type="text/css";t.rel="stylesheet";t.href=n;i.getElementsByTagName("head")[0].appendChild(t)}},parseEnum:function(n,t){var i=t[n.trim()];if(typeof i==o){OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+n);throw OsfMsAjaxFactory.msAjaxError.argument("str");}return i},delayExecutionAndCache:function(){var n={calc:arguments[0]};return function(){return n.calc&&(n.val=n.calc.apply(this,arguments),delete n.calc),n.val}},getUniqueId:function(){return h=h+1,h.toString()},formatString:function(){var n=arguments,t=n[0];return t.replace(/{(\d+)}/gm,function(t,i){var r=parseInt(i,10)+1;return n[r]===undefined?"{"+i+"}":n[r]})},generateConversationId:function(){return[d(),d(),(new Date).getTime().toString()].join("_")},getFrameNameAndConversationId:function(n,t){var i=k+n+this.generateConversationId();return t.setAttribute("name",i),this.generateConversationId()},addXdmInfoAsHash:function(n,t){return OSF.OUtil.addInfoAsHash(n,b,t)},addSerializerVersionAsHash:function(n,t){return OSF.OUtil.addInfoAsHash(n,c,t)},addInfoAsHash:function(n,t,i){n=n.trim()||"";var r=n.split(l),u=r.shift(),f=r.join(l);return[u,l,f,t,i].join("")},parseXdmInfo:function(n){return OSF.OUtil.parseXdmInfoWithGivenFragment(n,window.location.hash)},parseXdmInfoWithGivenFragment:function(n,t){return OSF.OUtil.parseInfoWithGivenFragment(b,k,c,n,t)},parseSerializerVersion:function(n){return OSF.OUtil.parseSerializerVersionWithGivenFragment(n,window.location.hash)},parseSerializerVersionWithGivenFragment:function(n,t){return parseInt(OSF.OUtil.parseInfoWithGivenFragment(c,tt,nt,n,t))},parseInfoWithGivenFragment:function(t,i,r,u,f){var h=f.split(t),a=h.length>1?h[h.length-1]:n,e=a!=n?a.split(r)[0]:n,c=g(),o,s,l;return!u&&c&&(o=window.name.indexOf(i),o>-1&&(s=window.name.indexOf(";",o),s==-1&&(s=window.name.length),l=window.name.substring(o,s),e?c.setItem(l,e):e=c.getItem(l))),e},getConversationId:function(){var i=window.location.search,t=n,r;return i&&(r=i.indexOf("&"),t=r>0?i.substring(1,r):i.substr(1),t&&t.charAt(t.length-1)==="="&&(t=t.substring(0,t.length-1),t&&(t=decodeURIComponent(t)))),t},getInfoItems:function(n){var t=n.split("$");return typeof t[1]==o&&(t=n.split("|")),t},getXdmFieldValue:function(n,t){var r="",u=OSF.OUtil.parseXdmInfo(t),i;if(u&&(i=OSF.OUtil.getInfoItems(u),i!=undefined&&i.length>=3))switch(n){case OSF.XdmFieldName.ConversationUrl:r=i[2];case OSF.XdmFieldName.AppId:r=i[1]}return r},validateParamObject:function(n,r){var u=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:t},{name:"expectedProperties",type:Object,mayBeNull:t},{name:"callback",type:Function,mayBeNull:i}]),f;if(u)throw u;for(f in r)if(u=Function._validateParameter(n[f],r[f],f),u)throw u;},writeProfilerMark:function(n){window.msWriteProfilerMark&&(window.msWriteProfilerMark(n),OsfMsAjaxFactory.msAjaxDebug.trace(n))},outputDebug:function(n){typeof OsfMsAjaxFactory!==o&&OsfMsAjaxFactory.msAjaxDebug&&OsfMsAjaxFactory.msAjaxDebug.trace&&OsfMsAjaxFactory.msAjaxDebug.trace(n)},defineNondefaultProperty:function(n,t,r,u){var e,f;r=r||{};for(e in u)f=u[e],r[f]==undefined&&(r[f]=i);return Object.defineProperty(n,t,r),n},defineNondefaultProperties:function(n,t,i){t=t||{};for(var r in t)OSF.OUtil.defineNondefaultProperty(n,r,t[r],i);return n},defineEnumerableProperty:function(n,t,i){return OSF.OUtil.defineNondefaultProperty(n,t,i,[f])},defineEnumerableProperties:function(n,t){return OSF.OUtil.defineNondefaultProperties(n,t,[f])},defineMutableProperty:function(n,t,i){return OSF.OUtil.defineNondefaultProperty(n,t,i,[w,f,p])},defineMutableProperties:function(n,t){return OSF.OUtil.defineNondefaultProperties(n,t,[w,f,p])},finalizeProperties:function(n,r){var e,u;r=r||{};for(var o=Object.getOwnPropertyNames(n),s=o.length,f=0;f<s;f++)e=o[f],u=Object.getOwnPropertyDescriptor(n,e),u.get||u.set||(u.writable=r.writable||t),u.configurable=r.configurable||t,u.enumerable=r.enumerable||i,Object.defineProperty(n,e,u);return n},mapList:function(n,t){var i=[],r;if(n)for(r in n)i.push(t(n[r]));return i},listContainsKey:function(n,r){for(var u in n)if(r==u)return i;return t},listContainsValue:function(n,r){for(var u in n)if(r==n[u])return i;return t},augmentList:function(n,t){var r=n.push?function(t,i){n.push(i)}:function(t,i){n[t]=i};for(var i in t)r(i,t[i])},redefineList:function(n,t){var r,i;for(r in n)delete n[r];for(i in t)n[i]=t[i]},isArray:function(n){return Object.prototype.toString.apply(n)==="[object Array]"},isFunction:function(n){return Object.prototype.toString.apply(n)==="[object Function]"},isDate:function(n){return Object.prototype.toString.apply(n)==="[object Date]"},addEventListener:function(n,i,r){n.addEventListener?n.addEventListener(i,r,t):Sys.Browser.agent===Sys.Browser.InternetExplorer&&n.attachEvent?n.attachEvent(u+i,r):n[u+i]=r},removeEventListener:function(i,r,f){i.removeEventListener?i.removeEventListener(r,f,t):Sys.Browser.agent===Sys.Browser.InternetExplorer&&i.detachEvent?i.detachEvent(u+r,f):i[u+r]=n},getCookieValue:function(n){var t=RegExp(n+"[^;]+").exec(document.cookie);return t.toString().replace(/^[^=]+./,"")},xhrGet:function(n,t,r){var u;try{u=new XMLHttpRequest;u.onreadystatechange=function(){u.readyState==4&&(u.status==200?t(u.responseText):r(u.status))};u.open("GET",n,i);u.send()}catch(f){r(f)}},xhrGetFull:function(n,t,r,u){var f,e=t;try{f=new XMLHttpRequest;f.onreadystatechange=function(){f.readyState==4&&(f.status==200?r(f,e):u(f.status))};f.open("GET",n,i);f.send()}catch(o){u(o)}},encodeBase64:function(n){var h;if(!n)return n;var a="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",l=[],i=[],o=0,c,e,s,r,u,f,t,v=n.length;do for(c=n.charCodeAt(o++),e=n.charCodeAt(o++),s=n.charCodeAt(o++),t=0,r=c&255,u=c>>8,f=e&255,i[t++]=r>>2,i[t++]=(r&3)<<4|u>>4,i[t++]=(u&15)<<2|f>>6,i[t++]=f&63,isNaN(e)||(r=e>>8,u=s&255,f=s>>8,i[t++]=r>>2,i[t++]=(r&3)<<4|u>>4,i[t++]=(u&15)<<2|f>>6,i[t++]=f&63),isNaN(e)?i[t-1]=64:isNaN(s)&&(i[t-2]=64,i[t-1]=64),h=0;h<t;h++)l.push(a.charAt(i[h]));while(o<v);return l.join("")},getSessionStorage:function(){return g()},getLocalStorage:function(){if(!y){try{var t=window.localStorage}catch(i){t=n}y=new OfficeExt.SafeStorage(t)}return y},convertIntToCssHexColor:function(n){return"#"+(Number(n)+16777216).toString(16).slice(-6)},attachClickHandler:function(n,t){n.onclick=function(){t()};n.ontouchend=function(n){t();n.preventDefault()}},getQueryStringParamValue:function(n,i){var u=Function._validateParams(arguments,[{name:"queryString",type:String,mayBeNull:t},{name:"paramName",type:String,mayBeNull:t}]),r;return u?(OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null."),""):(r=new RegExp("[\\?&]"+i+"=([^&#]*)","i"),!r.test(n))?(OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found."),""):r.exec(n)[1]},isiOS:function(){return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g)?i:t},shallowCopy:function(n){var i=n.constructor();for(var t in n)n.hasOwnProperty(t)&&(i[t]=n[t]);return i},createObject:function(t){var r=n,u,i;if(t)for(r={},u=t.length,i=0;i<u;i++)r[t[i].name]=t[i].value;return r},addClass:function(n,t){if(!OSF.OUtil.hasClass(n,t)){var i=n.getAttribute(e);i?n.setAttribute(e,i+" "+t):n.setAttribute(e,t)}},hasClass:function(n,t){var i=n.getAttribute(e);return i&&i.match(new RegExp("(\\s|^)"+t+"(\\s|$)"))}}}();OSF.OUtil.Guid=function(){var n=["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f"];return{generateNewGuid:function(){for(var i="",r=(new Date).getTime(),t=0;t<32&&r>0;t++)(t==8||t==12||t==16||t==20)&&(i+="-"),i+=n[r%16],r=Math.floor(r/16);for(;t<32;t++)(t==8||t==12||t==16||t==20)&&(i+="-"),i+=n[Math.floor(Math.random()*16)];return i}}}();window.OSF=OSF;OSF.OUtil.setNamespace("OSF",window);OSF.AppName={Unsupported:0,Excel:1,Word:2,PowerPoint:4,Outlook:8,ExcelWebApp:16,WordWebApp:32,OutlookWebApp:64,Project:128,AccessWebApp:256,PowerpointWebApp:512,ExcelIOS:1024,Sway:2048,WordIOS:4096,PowerPointIOS:8192,Access:16384,Lync:32768,OutlookIOS:65536,OneNoteWebApp:131072,OneNote:262144};OSF.InternalPerfMarker={DataCoercionBegin:"Agave.HostCall.CoerceDataStart",DataCoercionEnd:"Agave.HostCall.CoerceDataEnd"};OSF.HostCallPerfMarker={IssueCall:"Agave.HostCall.IssueCall",ReceiveResponse:"Agave.HostCall.ReceiveResponse",RuntimeExceptionRaised:"Agave.HostCall.RuntimeExecptionRaised"};OSF.AgaveHostAction={Select:0,UnSelect:1,CancelDialog:2,InsertAgave:3,CtrlF6In:4,CtrlF6Exit:5,CtrlF6ExitShift:6,SelectWithError:7,NotifyHostError:8,RefreshAddinCommands:9,PageIsReady:10};OSF.SharedConstants={NotificationConversationIdSuffix:"_ntf"};OSF.DialogMessageType={DialogMessageReceived:0,DialogClosed:1,NavigationFailed:2,InvalidSchema:3};OSF.OfficeAppContext=function(n,t,i,r,u,f,e,o,s,h,c,l,a,v,y,p,w){var b=this;b._id=n;b._appName=t;b._appVersion=i;b._appUILocale=r;b._dataLocale=u;b._docUrl=f;b._clientMode=e;b._settings=o;b._reason=s;b._osfControlType=h;b._eToken=c;b._correlationId=l;b._appInstanceId=a;b._touchEnabled=v;b._commerceAllowed=y;b._appMinorVersion=p;b._requirementMatrix=w;b._isDialog=!1;b.get_id=function(){return this._id};b.get_appName=function(){return this._appName};b.get_appVersion=function(){return this._appVersion};b.get_appUILocale=function(){return this._appUILocale};b.get_dataLocale=function(){return this._dataLocale};b.get_docUrl=function(){return this._docUrl};b.get_clientMode=function(){return this._clientMode};b.get_bindings=function(){return this._bindings};b.get_settings=function(){return this._settings};b.get_reason=function(){return this._reason};b.get_osfControlType=function(){return this._osfControlType};b.get_eToken=function(){return this._eToken};b.get_correlationId=function(){return this._correlationId};b.get_appInstanceId=function(){return this._appInstanceId};b.get_touchEnabled=function(){return this._touchEnabled};b.get_commerceAllowed=function(){return this._commerceAllowed};b.get_appMinorVersion=function(){return this._appMinorVersion};b.get_requirementMatrix=function(){return this._requirementMatrix};b.get_isDialog=function(){return this._isDialog}};OSF.OsfControlType={DocumentLevel:0,ContainerLevel:1};OSF.ClientMode={ReadOnly:0,ReadWrite:1};OSF.OUtil.setNamespace("Microsoft",window);OSF.OUtil.setNamespace("Office",Microsoft);OSF.OUtil.setNamespace("Client",Microsoft.Office);OSF.OUtil.setNamespace("WebExtension",Microsoft.Office);Microsoft.Office.WebExtension.InitializationReason={Inserted:"inserted",DocumentOpened:"documentOpened"};Microsoft.Office.WebExtension.ValueFormat={Unformatted:"unformatted",Formatted:"formatted"};Microsoft.Office.WebExtension.FilterType={All:"all"};Microsoft.Office.WebExtension.Parameters={BindingType:"bindingType",CoercionType:"coercionType",ValueFormat:"valueFormat",FilterType:"filterType",Columns:"columns",SampleData:"sampleData",GoToType:"goToType",SelectionMode:"selectionMode",Id:"id",PromptText:"promptText",ItemName:"itemName",FailOnCollision:"failOnCollision",StartRow:"startRow",StartColumn:"startColumn",RowCount:"rowCount",ColumnCount:"columnCount",Callback:"callback",AsyncContext:"asyncContext",Data:"data",Rows:"rows",OverwriteIfStale:"overwriteIfStale",FileType:"fileType",EventType:"eventType",Handler:"handler",SliceSize:"sliceSize",SliceIndex:"sliceIndex",ActiveView:"activeView",Status:"status",Xml:"xml",Namespace:"namespace",Prefix:"prefix",XPath:"xPath",Text:"text",ImageLeft:"imageLeft",ImageTop:"imageTop",ImageWidth:"imageWidth",ImageHeight:"imageHeight",TaskId:"taskId",FieldId:"fieldId",FieldValue:"fieldValue",ServerUrl:"serverUrl",ListName:"listName",ResourceId:"resourceId",ViewType:"viewType",ViewName:"viewName",GetRawValue:"getRawValue",CellFormat:"cellFormat",TableOptions:"tableOptions",TaskIndex:"taskIndex",ResourceIndex:"resourceIndex",Url:"url",MessageHandler:"messageHandler",Width:"width",Height:"height",RequireHTTPs:"requireHTTPS",MessageToParent:"messageToParent",XFrameDenySafe:"xFrameDenySafe"};OSF.OUtil.setNamespace("DDA",OSF);OSF.DDA.DocumentMode={ReadOnly:1,ReadWrite:0};OSF.DDA.PropertyDescriptors={AsyncResultStatus:"AsyncResultStatus"};OSF.DDA.EventDescriptors={};OSF.DDA.ListDescriptors={};OSF.DDA.UI={};OSF.DDA.getXdmEventName=function(n,t){return t==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged||t==Microsoft.Office.WebExtension.EventType.BindingDataChanged?n+"_"+t:t};OSF.DDA.MethodDispId={dispidMethodMin:64,dispidGetSelectedDataMethod:64,dispidSetSelectedDataMethod:65,dispidAddBindingFromSelectionMethod:66,dispidAddBindingFromPromptMethod:67,dispidGetBindingMethod:68,dispidReleaseBindingMethod:69,dispidGetBindingDataMethod:70,dispidSetBindingDataMethod:71,dispidAddRowsMethod:72,dispidClearAllRowsMethod:73,dispidGetAllBindingsMethod:74,dispidLoadSettingsMethod:75,dispidSaveSettingsMethod:76,dispidGetDocumentCopyMethod:77,dispidAddBindingFromNamedItemMethod:78,dispidAddColumnsMethod:79,dispidGetDocumentCopyChunkMethod:80,dispidReleaseDocumentCopyMethod:81,dispidNavigateToMethod:82,dispidGetActiveViewMethod:83,dispidGetDocumentThemeMethod:84,dispidGetOfficeThemeMethod:85,dispidGetFilePropertiesMethod:86,dispidClearFormatsMethod:87,dispidSetTableOptionsMethod:88,dispidSetFormatsMethod:89,dispidExecuteRichApiRequestMethod:93,dispidAppCommandInvocationCompletedMethod:94,dispidAddDataPartMethod:128,dispidGetDataPartByIdMethod:129,dispidGetDataPartsByNamespaceMethod:130,dispidGetDataPartXmlMethod:131,dispidGetDataPartNodesMethod:132,dispidDeleteDataPartMethod:133,dispidGetDataNodeValueMethod:134,dispidGetDataNodeXmlMethod:135,dispidGetDataNodesMethod:136,dispidSetDataNodeValueMethod:137,dispidSetDataNodeXmlMethod:138,dispidAddDataNamespaceMethod:139,dispidGetDataUriByPrefixMethod:140,dispidGetDataPrefixByUriMethod:141,dispidGetDataNodeTextMethod:142,dispidSetDataNodeTextMethod:143,dispidMessageParentMethod:144,dispidMethodMax:144,dispidGetSelectedTaskMethod:110,dispidGetSelectedResourceMethod:111,dispidGetTaskMethod:112,dispidGetResourceFieldMethod:113,dispidGetWSSUrlMethod:114,dispidGetTaskFieldMethod:115,dispidGetProjectFieldMethod:116,dispidGetSelectedViewMethod:117,dispidGetTaskByIndexMethod:118,dispidGetResourceByIndexMethod:119,dispidSetTaskFieldMethod:120,dispidSetResourceFieldMethod:121,dispidGetMaxTaskIndexMethod:122,dispidGetMaxResourceIndexMethod:123};OSF.DDA.EventDispId={dispidEventMin:0,dispidInitializeEvent:0,dispidSettingsChangedEvent:1,dispidDocumentSelectionChangedEvent:2,dispidBindingSelectionChangedEvent:3,dispidBindingDataChangedEvent:4,dispidDocumentOpenEvent:5,dispidDocumentCloseEvent:6,dispidActiveViewChangedEvent:7,dispidDocumentThemeChangedEvent:8,dispidOfficeThemeChangedEvent:9,dispidDialogMessageReceivedEvent:10,dispidActivationStatusChangedEvent:32,dispidAppCommandInvokedEvent:39,dispidTaskSelectionChangedEvent:56,dispidResourceSelectionChangedEvent:57,dispidViewSelectionChangedEvent:58,dispidDataNodeAddedEvent:60,dispidDataNodeReplacedEvent:61,dispidDataNodeDeletedEvent:62,dispidEventMax:63};OSF.DDA.ErrorCodeManager=function(){var n={};return{getErrorArgs:function(t){var i=n[t];return i?(i.name||(i.name=n[this.errorCodes.ooeInternalError].name),i.message||(i.message=n[this.errorCodes.ooeInternalError].message)):i=n[this.errorCodes.ooeInternalError],i},addErrorMessage:function(t,i){n[t]=i},errorCodes:{ooeSuccess:0,ooeChunkResult:1,ooeCoercionTypeNotSupported:1e3,ooeGetSelectionNotMatchDataType:1001,ooeCoercionTypeNotMatchBinding:1002,ooeInvalidGetRowColumnCounts:1003,ooeSelectionNotSupportCoercionType:1004,ooeInvalidGetStartRowColumn:1005,ooeNonUniformPartialGetNotSupported:1006,ooeGetDataIsTooLarge:1008,ooeFileTypeNotSupported:1009,ooeGetDataParametersConflict:1010,ooeInvalidGetColumns:1011,ooeInvalidGetRows:1012,ooeInvalidReadForBlankRow:1013,ooeUnsupportedDataObject:2e3,ooeCannotWriteToSelection:2001,ooeDataNotMatchSelection:2002,ooeOverwriteWorksheetData:2003,ooeDataNotMatchBindingSize:2004,ooeInvalidSetStartRowColumn:2005,ooeInvalidDataFormat:2006,ooeDataNotMatchCoercionType:2007,ooeDataNotMatchBindingType:2008,ooeSetDataIsTooLarge:2009,ooeNonUniformPartialSetNotSupported:2010,ooeInvalidSetColumns:2011,ooeInvalidSetRows:2012,ooeSetDataParametersConflict:2013,ooeCellDataAmountBeyondLimits:2014,ooeSelectionCannotBound:3e3,ooeBindingNotExist:3002,ooeBindingToMultipleSelection:3003,ooeInvalidSelectionForBindingType:3004,ooeOperationNotSupportedOnThisBindingType:3005,ooeNamedItemNotFound:3006,ooeMultipleNamedItemFound:3007,ooeInvalidNamedItemForBindingType:3008,ooeUnknownBindingType:3009,ooeOperationNotSupportedOnMatrixData:3010,ooeInvalidColumnsForBinding:3011,ooeSettingNameNotExist:4e3,ooeSettingsCannotSave:4001,ooeSettingsAreStale:4002,ooeOperationNotSupported:5e3,ooeInternalError:5001,ooeDocumentReadOnly:5002,ooeEventHandlerNotExist:5003,ooeInvalidApiCallInContext:5004,ooeShuttingDown:5005,ooeUnsupportedEnumeration:5007,ooeIndexOutOfRange:5008,ooeBrowserAPINotSupported:5009,ooeInvalidParam:5010,ooeRequestTimeout:5011,ooeTooManyIncompleteRequests:5100,ooeRequestTokenUnavailable:5101,ooeActivityLimitReached:5102,ooeCustomXmlNodeNotFound:6e3,ooeCustomXmlError:6100,ooeCustomXmlExceedQuota:6101,ooeCustomXmlOutOfDate:6102,ooeNoCapability:7e3,ooeCannotNavTo:7001,ooeSpecifiedIdNotExist:7002,ooeNavOutOfBound:7004,ooeElementMissing:8e3,ooeProtectedError:8001,ooeInvalidCellsValue:8010,ooeInvalidTableOptionValue:8011,ooeInvalidFormatValue:8012,ooeRowIndexOutOfRange:8020,ooeColIndexOutOfRange:8021,ooeFormatValueOutOfRange:8022,ooeCellFormatAmountBeyondLimits:8023,ooeMemoryFileLimit:11e3,ooeNetworkProblemRetrieveFile:11001,ooeInvalidSliceSize:11002,ooeInvalidCallback:11101,ooeInvalidWidth:12e3,ooeInvalidHeight:12001,ooeNavigationError:12002,ooeInvalidScheme:12003,ooeAppDomains:12004,ooeRequireHTTPS:12005,ooeWebDialogClosed:12006,ooeDialogAlreadyOpened:12007,ooeEndUserAllow:12008,ooeEndUserIgnore:12009},initializeErrorMessages:function(t){n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported]={name:t.L_InvalidCoercion,message:t.L_CoercionTypeNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType]={name:t.L_DataReadError,message:t.L_GetSelectionNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding]={name:t.L_InvalidCoercion,message:t.L_CoercionTypeNotMatchBinding};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts]={name:t.L_DataReadError,message:t.L_InvalidGetRowColumnCounts};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType]={name:t.L_DataReadError,message:t.L_SelectionNotSupportCoercionType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn]={name:t.L_DataReadError,message:t.L_InvalidGetStartRowColumn};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported]={name:t.L_DataReadError,message:t.L_NonUniformPartialGetNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge]={name:t.L_DataReadError,message:t.L_GetDataIsTooLarge};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported]={name:t.L_DataReadError,message:t.L_FileTypeNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict]={name:t.L_DataReadError,message:t.L_GetDataParametersConflict};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns]={name:t.L_DataReadError,message:t.L_InvalidGetColumns};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows]={name:t.L_DataReadError,message:t.L_InvalidGetRows};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow]={name:t.L_DataReadError,message:t.L_InvalidReadForBlankRow};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject]={name:t.L_DataWriteError,message:t.L_UnsupportedDataObject};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection]={name:t.L_DataWriteError,message:t.L_CannotWriteToSelection};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection]={name:t.L_DataWriteError,message:t.L_DataNotMatchSelection};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData]={name:t.L_DataWriteError,message:t.L_OverwriteWorksheetData};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize]={name:t.L_DataWriteError,message:t.L_DataNotMatchBindingSize};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn]={name:t.L_DataWriteError,message:t.L_InvalidSetStartRowColumn};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat]={name:t.L_InvalidFormat,message:t.L_InvalidDataFormat};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType]={name:t.L_InvalidDataObject,message:t.L_DataNotMatchCoercionType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType]={name:t.L_InvalidDataObject,message:t.L_DataNotMatchBindingType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge]={name:t.L_DataWriteError,message:t.L_SetDataIsTooLarge};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported]={name:t.L_DataWriteError,message:t.L_NonUniformPartialSetNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns]={name:t.L_DataWriteError,message:t.L_InvalidSetColumns};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows]={name:t.L_DataWriteError,message:t.L_InvalidSetRows};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict]={name:t.L_DataWriteError,message:t.L_SetDataParametersConflict};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound]={name:t.L_BindingCreationError,message:t.L_SelectionCannotBound};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist]={name:t.L_InvalidBindingError,message:t.L_BindingNotExist};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection]={name:t.L_BindingCreationError,message:t.L_BindingToMultipleSelection};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType]={name:t.L_BindingCreationError,message:t.L_InvalidSelectionForBindingType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType]={name:t.L_InvalidBindingOperation,message:t.L_OperationNotSupportedOnThisBindingType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound]={name:t.L_BindingCreationError,message:t.L_NamedItemNotFound};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound]={name:t.L_BindingCreationError,message:t.L_MultipleNamedItemFound};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType]={name:t.L_BindingCreationError,message:t.L_InvalidNamedItemForBindingType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType]={name:t.L_InvalidBinding,message:t.L_UnknownBindingType};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData]={name:t.L_InvalidBindingOperation,message:t.L_OperationNotSupportedOnMatrixData};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding]={name:t.L_InvalidBinding,message:t.L_InvalidColumnsForBinding};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist]={name:t.L_ReadSettingsError,message:t.L_SettingNameNotExist};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave]={name:t.L_SaveSettingsError,message:t.L_SettingsCannotSave};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale]={name:t.L_SettingsStaleError,message:t.L_SettingsAreStale};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported]={name:t.L_HostError,message:t.L_OperationNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError]={name:t.L_InternalError,message:t.L_InternalErrorDescription};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly]={name:t.L_PermissionDenied,message:t.L_DocumentReadOnly};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist]={name:t.L_EventRegistrationError,message:t.L_EventHandlerNotExist};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext]={name:t.L_InvalidAPICall,message:t.L_InvalidApiCallInContext};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown]={name:t.L_ShuttingDown,message:t.L_ShuttingDown};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration]={name:t.L_UnsupportedEnumeration,message:t.L_UnsupportedEnumerationMessage};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange]={name:t.L_IndexOutOfRange,message:t.L_IndexOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported]={name:t.L_APINotSupported,message:t.L_BrowserAPINotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout]={name:t.L_APICallFailed,message:t.L_RequestTimeout};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests]={name:t.L_APICallFailed,message:t.L_TooManyIncompleteRequests};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable]={name:t.L_APICallFailed,message:t.L_RequestTokenUnavailable};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached]={name:t.L_APICallFailed,message:t.L_ActivityLimitReached};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound]={name:t.L_InvalidNode,message:t.L_CustomXmlNodeNotFound};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError]={name:t.L_CustomXmlError,message:t.L_CustomXmlError};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota]={name:t.L_CustomXmlError,message:t.L_CustomXmlError};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate]={name:t.L_CustomXmlError,message:t.L_CustomXmlError};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability]={name:t.L_PermissionDenied,message:t.L_NoCapability};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo]={name:t.L_CannotNavigateTo,message:t.L_CannotNavigateTo};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist]={name:t.L_SpecifiedIdNotExist,message:t.L_SpecifiedIdNotExist};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound]={name:t.L_NavOutOfBound,message:t.L_NavOutOfBound};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits]={name:t.L_DataWriteReminder,message:t.L_CellDataAmountBeyondLimits};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing]={name:t.L_MissingParameter,message:t.L_ElementMissing};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError]={name:t.L_PermissionDenied,message:t.L_NoCapability};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue]={name:t.L_InvalidValue,message:t.L_InvalidCellsValue};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue]={name:t.L_InvalidValue,message:t.L_InvalidTableOptionValue};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue]={name:t.L_InvalidValue,message:t.L_InvalidFormatValue};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange]={name:t.L_OutOfRange,message:t.L_RowIndexOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange]={name:t.L_OutOfRange,message:t.L_ColIndexOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange]={name:t.L_OutOfRange,message:t.L_FormatValueOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits]={name:t.L_FormattingReminder,message:t.L_CellFormatAmountBeyondLimits};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit]={name:t.L_MemoryLimit,message:t.L_CloseFileBeforeRetrieve};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile]={name:t.L_NetworkProblem,message:t.L_NetworkProblemRetrieveFile};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize]={name:t.L_InvalidValue,message:t.L_SliceSizeNotSupported};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened]={name:t.L_DisplayDialogError,message:t.L_DialogAlreadyOpened};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth]={name:t.L_IndexOutOfRange,message:t.L_IndexOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight]={name:t.L_IndexOutOfRange,message:t.L_IndexOutOfRange};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError]={name:t.L_DisplayDialogError,message:t.L_NetworkProblem};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme]={name:t.L_DialogNavigateError,message:t.L_DialogAddressNotTrusted};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains]={name:t.L_DisplayDialogError,message:t.L_DialogAddressNotTrusted};n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS]={name:t.L_DisplayDialogError,message:t.L_DialogAddressNotTrusted}}}}(),function(n){var t;(function(n){var t=1.1,r=function(){function n(n){this.isSetSupported=function(n,t){var u,i,r;return n==undefined?!1:(t==undefined&&(t=0),u=this._setMap,i=u._sets,i.hasOwnProperty(n.toLowerCase())?(r=i[n.toLowerCase()],r>0&&r>=t):!1)};this._setMap=n}return n}(),i,u,h,c,f,l,e,a,v,y,p,w,o,b,k,d,s,g,nt,tt,it;n.RequirementMatrix=r;i=function(){function n(n){this._addSetMap=function(n){for(var t in n)this._sets[t]=n[t]};this._sets=n}return n}();n.DefaultSetRequirement=i;u=function(n){function i(){n.call(this,{bindingevents:t,documentevents:t,excelapi:t,matrixbindings:t,matrixcoercion:t,selection:t,settings:t,tablebindings:t,tablecoercion:t,textbindings:t,textcoercion:t})}return __extends(i,n),i}(i);n.ExcelClientDefaultSetRequirement=u;h=function(n){function i(){n.call(this);this._addSetMap({imagecoercion:t})}return __extends(i,n),i}(u);n.ExcelClientV1DefaultSetRequirement=h;c=function(n){function t(){n.call(this,{mailbox:1.3})}return __extends(t,n),t}(i);n.OutlookClientDefaultSetRequirement=c;f=function(n){function i(){n.call(this,{bindingevents:t,compressedfile:t,customxmlparts:t,documentevents:t,file:t,htmlcoercion:t,matrixbindings:t,matrixcoercion:t,ooxmlcoercion:t,pdffile:t,selection:t,settings:t,tablebindings:t,tablecoercion:t,textbindings:t,textcoercion:t,textfile:t,wordapi:t})}return __extends(i,n),i}(i);n.WordClientDefaultSetRequirement=f;l=function(n){function i(){n.call(this);this._addSetMap({customxmlparts:1.2,wordapi:1.2,imagecoercion:t})}return __extends(i,n),i}(f);n.WordClientV1DefaultSetRequirement=l;e=function(n){function i(){n.call(this,{activeview:t,compressedfile:t,documentevents:t,file:t,pdffile:t,selection:t,settings:t,textcoercion:t})}return __extends(i,n),i}(i);n.PowerpointClientDefaultSetRequirement=e;a=function(n){function i(){n.call(this);this._addSetMap({imagecoercion:t})}return __extends(i,n),i}(e);n.PowerpointClientV1DefaultSetRequirement=a;v=function(n){function i(){n.call(this,{selection:t,textcoercion:t})}return __extends(i,n),i}(i);n.ProjectClientDefaultSetRequirement=v;y=function(n){function i(){n.call(this,{bindingevents:t,documentevents:t,matrixbindings:t,matrixcoercion:t,selection:t,settings:t,tablebindings:t,tablecoercion:t,textbindings:t,textcoercion:t,file:t})}return __extends(i,n),i}(i);n.ExcelWebDefaultSetRequirement=y;p=function(n){function i(){n.call(this,{customxmlparts:t,documentevents:t,file:t,ooxmlcoercion:t,selection:t,settings:t,textcoercion:t})}return __extends(i,n),i}(i);n.WordWebDefaultSetRequirement=p;w=function(n){function i(){n.call(this,{activeview:t,settings:t})}return __extends(i,n),i}(i);n.PowerpointWebDefaultSetRequirement=w;o=function(n){function t(){n.call(this,{mailbox:1.3})}return __extends(t,n),t}(i);n.OutlookWebDefaultSetRequirement=o;b=function(n){function i(){n.call(this,{activeview:t,documentevents:t,selection:t,settings:t,textcoercion:t})}return __extends(i,n),i}(i);n.SwayWebDefaultSetRequirement=b;k=function(n){function i(){n.call(this,{bindingevents:t,partialtablebindings:t,settings:t,tablebindings:t,tablecoercion:t})}return __extends(i,n),i}(i);n.AccessWebDefaultSetRequirement=k;d=function(n){function i(){n.call(this,{bindingevents:t,documentevents:t,matrixbindings:t,matrixcoercion:t,selection:t,settings:t,tablebindings:t,tablecoercion:t,textbindings:t,textcoercion:t})}return __extends(i,n),i}(i);n.ExcelIOSDefaultSetRequirement=d;s=function(n){function i(){n.call(this,{bindingevents:t,compressedfile:t,customxmlparts:t,documentevents:t,file:t,htmlcoercion:t,matrixbindings:t,matrixcoercion:t,ooxmlcoercion:t,pdffile:t,selection:t,settings:t,tablebindings:t,tablecoercion:t,textbindings:t,textcoercion:t,textfile:t})}return __extends(i,n),i}(i);n.WordIOSDefaultSetRequirement=s;g=function(n){function t(){n.call(this);this._addSetMap({customxmlparts:1.2,wordapi:1.2})}return __extends(t,n),t}(s);n.WordIOSV1DefaultSetRequirement=g;nt=function(n){function i(){n.call(this,{activeview:t,compressedfile:t,documentevents:t,file:t,pdffile:t,selection:t,settings:t,textcoercion:t})}return __extends(i,n),i}(i);n.PowerpointIOSDefaultSetRequirement=nt;tt=function(n){function i(){n.call(this,{mailbox:t})}return __extends(i,n),i}(i);n.OutlookIOSDefaultSetRequirement=tt;it=function(){function n(){}return n.initializeOsfDda=function(){OSF.OUtil.setNamespace("Requirement",OSF.DDA)},n.getDefaultRequirementMatrix=function(t){var u,f,o,e;return this.initializeDefaultSetMatrix(),u=undefined,f=t.get_requirementMatrix(),f!=undefined&&f.length>0&&typeof JSON!="undefined"?(o=JSON.parse(t.get_requirementMatrix().toLowerCase()),u=new r(new i(o))):(e=n.getClientFullVersionString(t),u=n.DefaultSetArrayMatrix!=undefined&&n.DefaultSetArrayMatrix[e]!=undefined?new r(n.DefaultSetArrayMatrix[e]):new r(new i({}))),u},n.getClientFullVersionString=function(n){var i=n.get_appMinorVersion(),u="",r="",t=n.get_appName(),f=t==1024||t==4096||t==8192||t==65536;return f&&n.get_appVersion()==1?r=t==4096&&i>=15?"16.00.01":"16.00":n.get_appName()==64?r=n.get_appVersion():(u=i<10?"0"+i:""+i,r=n.get_appVersion()+"."+u),n.get_appName()+"-"+r},n.initializeDefaultSetMatrix=function(){n.DefaultSetArrayMatrix[n.Excel_RCLIENT_1600]=new u;n.DefaultSetArrayMatrix[n.Word_RCLIENT_1600]=new f;n.DefaultSetArrayMatrix[n.PowerPoint_RCLIENT_1600]=new e;n.DefaultSetArrayMatrix[n.Excel_RCLIENT_1601]=new h;n.DefaultSetArrayMatrix[n.Word_RCLIENT_1601]=new l;n.DefaultSetArrayMatrix[n.PowerPoint_RCLIENT_1601]=new a;n.DefaultSetArrayMatrix[n.Outlook_RCLIENT_1600]=new c;n.DefaultSetArrayMatrix[n.Excel_WAC_1600]=new y;n.DefaultSetArrayMatrix[n.Word_WAC_1600]=new p;n.DefaultSetArrayMatrix[n.Outlook_WAC_1600]=new o;n.DefaultSetArrayMatrix[n.Outlook_WAC_1601]=new o;n.DefaultSetArrayMatrix[n.Project_RCLIENT_1600]=new v;n.DefaultSetArrayMatrix[n.Access_WAC_1600]=new k;n.DefaultSetArrayMatrix[n.PowerPoint_WAC_1600]=new w;n.DefaultSetArrayMatrix[n.Excel_IOS_1600]=new d;n.DefaultSetArrayMatrix[n.SWAY_WAC_1600]=new b;n.DefaultSetArrayMatrix[n.Word_IOS_1600]=new s;n.DefaultSetArrayMatrix[n.Word_IOS_16001]=new g;n.DefaultSetArrayMatrix[n.PowerPoint_IOS_1600]=new nt;n.DefaultSetArrayMatrix[n.Outlook_IOS_1600]=new tt},n.Excel_RCLIENT_1600="1-16.00",n.Excel_RCLIENT_1601="1-16.01",n.Word_RCLIENT_1600="2-16.00",n.Word_RCLIENT_1601="2-16.01",n.PowerPoint_RCLIENT_1600="4-16.00",n.PowerPoint_RCLIENT_1601="4-16.01",n.Outlook_RCLIENT_1600="8-16.00",n.Excel_WAC_1600="16-16.00",n.Word_WAC_1600="32-16.00",n.Outlook_WAC_1600="64-16.00",n.Outlook_WAC_1601="64-16.01",n.Project_RCLIENT_1600="128-16.00",n.Access_WAC_1600="256-16.00",n.PowerPoint_WAC_1600="512-16.00",n.Excel_IOS_1600="1024-16.00",n.SWAY_WAC_1600="2048-16.00",n.Word_IOS_1600="4096-16.00",n.Word_IOS_16001="4096-16.00.01",n.PowerPoint_IOS_1600="8192-16.00",n.Outlook_IOS_1600="65536-16.00",n.DefaultSetArrayMatrix={},n}();n.RequirementsMatrixFactory=it})(t=n.Requirement||(n.Requirement={}))}(OfficeExt||(OfficeExt={}));OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();Microsoft.Office.WebExtension.ApplicationMode={WebEditor:"webEditor",WebViewer:"webViewer",Client:"client"};Microsoft.Office.WebExtension.DocumentMode={ReadOnly:"readOnly",ReadWrite:"readWrite"};OSF.NamespaceManager=function(){var t,n=!1;return{enableShortcut:function(){n||(window.Office?t=window.Office:OSF.OUtil.setNamespace("Office",window),window.Office=Microsoft.Office.WebExtension,n=!0)},disableShortcut:function(){n&&(t?window.Office=t:OSF.OUtil.unsetNamespace("Office",window),n=!1)}}}();OSF.NamespaceManager.enableShortcut();Microsoft.Office.WebExtension.useShortNamespace=function(n){n?OSF.NamespaceManager.enableShortcut():OSF.NamespaceManager.disableShortcut()};Microsoft.Office.WebExtension.select=function(n,t){var i,r,o,u,f,e;if(n&&typeof n=="string"&&(r=n.indexOf("#"),r!=-1)){o=n.substring(0,r);u=n.substring(r+1);switch(o){case"binding":case"bindings":u&&(i=new OSF.DDA.BindingPromise(u))}}if(i)return i.onFail=t,i;else if(t)if(f=typeof t,f=="function")e={},e[Microsoft.Office.WebExtension.Parameters.Callback]=t,OSF.DDA.issueAsyncResult(e,OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext,OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));else throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction,f);};OSF.DDA.Context=function(n,t,i,r,u){var f=this,e,o;OSF.OUtil.defineEnumerableProperties(f,{contentLanguage:{value:n.get_dataLocale()},displayLanguage:{value:n.get_appUILocale()},touchEnabled:{value:n.get_touchEnabled()},commerceAllowed:{value:n.get_commerceAllowed()}});i&&OSF.OUtil.defineEnumerableProperty(f,"license",{value:i});n.ui&&OSF.OUtil.defineEnumerableProperty(f,"ui",{value:n.ui});n.get_isDialog()||(t&&OSF.OUtil.defineEnumerableProperty(f,"document",{value:t}),r&&(e=r.displayName||"appOM",delete r.displayName,OSF.OUtil.defineEnumerableProperty(f,e,{value:r})),u&&OSF.OUtil.defineEnumerableProperty(f,"officeTheme",{get:function(){return u()}}),o=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(n),OSF.OUtil.defineEnumerableProperty(f,"requirements",{value:o}))};OSF.DDA.OutlookContext=function(n,t,i,r,u){OSF.DDA.OutlookContext.uber.constructor.call(this,n,null,i,r,u);t&&OSF.OUtil.defineEnumerableProperty(this,"roamingSettings",{value:t})};OSF.OUtil.extend(OSF.DDA.OutlookContext,OSF.DDA.Context);OSF.DDA.OutlookAppOm=function(){};OSF.DDA.Document=function(n,t){var i;switch(n.get_clientMode()){case OSF.ClientMode.ReadOnly:i=Microsoft.Office.WebExtension.DocumentMode.ReadOnly;break;case OSF.ClientMode.ReadWrite:i=Microsoft.Office.WebExtension.DocumentMode.ReadWrite}t&&OSF.OUtil.defineEnumerableProperty(this,"settings",{value:t});OSF.OUtil.defineMutableProperties(this,{mode:{value:i},url:{value:n.get_docUrl()}})};OSF.DDA.JsomDocument=function(n,t,i){var r=this,u;OSF.DDA.JsomDocument.uber.constructor.call(r,n,i);t&&OSF.OUtil.defineEnumerableProperty(r,"bindings",{get:function(){return t}});u=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(r,[u.GetSelectedDataAsync,u.SetSelectedDataAsync]);OSF.DDA.DispIdHost.addEventSupport(r,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))};OSF.OUtil.extend(OSF.DDA.JsomDocument,OSF.DDA.Document);OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension,"context",{get:function(){var n;return OSF&&OSF._OfficeAppFactory&&(n=OSF._OfficeAppFactory.getContext()),n}});OSF.DDA.License=function(n){OSF.OUtil.defineEnumerableProperty(this,"value",{value:n})};OSF.DDA.ApiMethodCall=function(n,t,i,r,u){var f=this,e=n.length,o=OSF.OUtil.delayExecutionAndCache(function(){return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters,u)});f.verifyArguments=function(n,t){var u,i,r;for(u in n){if(i=n[u],r=t[u],i["enum"])switch(typeof r){case"string":if(OSF.OUtil.listContainsValue(i["enum"],r))break;case"undefined":throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;default:throw o();}if(i.types&&!OSF.OUtil.listContainsValue(i.types,typeof r))throw o();}};f.extractRequiredArguments=function(t,i,r){var f,u,h,s,c,l;if(t.length<e)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);for(f=[],u=0;u<e;u++)f.push(t[u]);for(this.verifyArguments(n,f),h={},u=0;u<e;u++){if(s=n[u],c=f[u],s.verify&&(l=s.verify(c,i,r),!l))throw o();h[s.name]=c}return h};f.fillOptions=function(n,i,r,u){var o,f,e;n=n||{};for(o in t)OSF.OUtil.listContainsKey(n,o)||(f=undefined,e=t[o],e.calculate&&i&&(f=e.calculate(i,r,u)),f||e.defaultValue===undefined||(f=e.defaultValue),n[o]=f);return n};f.constructCallArgs=function(n,t,u,f){var e={},o,s,h;for(o in n)e[o]=n[o];for(s in t)e[s]=t[s];for(h in i)e[h]=i[h](u,f);return r&&(e=r(e,u,f)),e}};OSF.OUtil.setNamespace("AsyncResultEnum",OSF.DDA);OSF.DDA.AsyncResultEnum.Properties={Context:"Context",Value:"Value",Status:"Status",Error:"Error"};Microsoft.Office.WebExtension.AsyncResultStatus={Succeeded:"succeeded",Failed:"failed"};OSF.DDA.AsyncResultEnum.ErrorCode={Success:0,Failed:1};OSF.DDA.AsyncResultEnum.ErrorProperties={Name:"Name",Message:"Message",Code:"Code"};OSF.DDA.AsyncMethodNames={};OSF.DDA.AsyncMethodNames.addNames=function(n){var t,i;for(t in n)i={},OSF.OUtil.defineEnumerableProperties(i,{id:{value:t},displayName:{value:n[t]}}),OSF.DDA.AsyncMethodNames[t]=i};OSF.DDA.AsyncMethodCall=function(n,t,i,r,u,f,e){function c(n,i,r,u){var f,e,c,l;if(n.length>s+2)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);for(c=n.length-1;c>=s;c--){l=n[c];switch(typeof l){case"object":if(f)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);else f=l;break;case h:if(e)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);else e=l;break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);}}if(f=o.fillOptions(f,i,r,u),e)if(f[Microsoft.Office.WebExtension.Parameters.Callback])throw Strings.OfficeOM.L_RedundantCallbackSpecification;else f[Microsoft.Office.WebExtension.Parameters.Callback]=e;return o.verifyArguments(t,f),f}var h="function",s=n.length,o=new OSF.DDA.ApiMethodCall(n,t,i,f,e);this.verifyAndExtractCall=function(n,t,i){var r=o.extractRequiredArguments(n,t,i),u=c(n,r,t,i);return o.constructCallArgs(r,u,t,i)};this.processResponse=function(n,t,i,f){return n==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess?r?r(t,i,f):t:u?u(n,t):OSF.DDA.ErrorCodeManager.getErrorArgs(n)};this.getCallArgs=function(n){for(var t,u,r,i=n.length-1;i>=s;i--){r=n[i];switch(typeof r){case"object":t=r;break;case h:u=r}}return t=t||{},u&&(t[Microsoft.Office.WebExtension.Parameters.Callback]=u),t}};OSF.DDA.AsyncMethodCallFactory=function(){return{manufacture:function(n){var t=n.supportedOptions?OSF.OUtil.createObject(n.supportedOptions):[],i=n.privateStateCallbacks?OSF.OUtil.createObject(n.privateStateCallbacks):[];return new OSF.DDA.AsyncMethodCall(n.requiredArguments||[],t,i,n.onSucceeded,n.onFailed,n.checkCallArgs,n.method.displayName)}}}();OSF.DDA.AsyncMethodCalls={};OSF.DDA.AsyncMethodCalls.define=function(n){OSF.DDA.AsyncMethodCalls[n.method.id]=OSF.DDA.AsyncMethodCallFactory.manufacture(n)};OSF.DDA.Error=function(n,t,i){OSF.OUtil.defineEnumerableProperties(this,{name:{value:n},message:{value:t},code:{value:i}})};OSF.DDA.AsyncResult=function(n,t){OSF.OUtil.defineEnumerableProperties(this,{value:{value:n[OSF.DDA.AsyncResultEnum.Properties.Value]},status:{value:t?Microsoft.Office.WebExtension.AsyncResultStatus.Failed:Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded}});n[OSF.DDA.AsyncResultEnum.Properties.Context]&&OSF.OUtil.defineEnumerableProperty(this,"asyncContext",{value:n[OSF.DDA.AsyncResultEnum.Properties.Context]});t&&OSF.OUtil.defineEnumerableProperty(this,"error",{value:new OSF.DDA.Error(t[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],t[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],t[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])})};OSF.DDA.issueAsyncResult=function(n,t,i){var f=n[Microsoft.Office.WebExtension.Parameters.Callback],u,r;f&&(u={},u[OSF.DDA.AsyncResultEnum.Properties.Context]=n[Microsoft.Office.WebExtension.Parameters.AsyncContext],t==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess?u[OSF.DDA.AsyncResultEnum.Properties.Value]=i:(r={},i=i||OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError),r[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=t||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,r[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=i.name||i,r[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=i.message||i),f(new OSF.DDA.AsyncResult(u,r)))};OSF.DDA.SyncMethodNames={};OSF.DDA.SyncMethodNames.addNames=function(n){var t,i;for(t in n)i={},OSF.OUtil.defineEnumerableProperties(i,{id:{value:t},displayName:{value:n[t]}}),OSF.DDA.SyncMethodNames[t]=i};OSF.DDA.SyncMethodCall=function(n,t,i,r,u){function o(n,i,r,u){var o,c,s,h;if(n.length>e+1)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);for(s=n.length-1;s>=e;s--){h=n[s];switch(typeof h){case"object":if(o)throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);else o=h;break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);}}return o=f.fillOptions(o,i,r,u),f.verifyArguments(t,o),o}var e=n.length,f=new OSF.DDA.ApiMethodCall(n,t,i,r,u);this.verifyAndExtractCall=function(n,t,i){var r=f.extractRequiredArguments(n,t,i),u=o(n,r,t,i);return f.constructCallArgs(r,u,t,i)}};OSF.DDA.SyncMethodCallFactory=function(){return{manufacture:function(n){var t=n.supportedOptions?OSF.OUtil.createObject(n.supportedOptions):[];return new OSF.DDA.SyncMethodCall(n.requiredArguments||[],t,n.privateStateCallbacks,n.checkCallArgs,n.method.displayName)}}}();OSF.DDA.SyncMethodCalls={};OSF.DDA.SyncMethodCalls.define=function(n){OSF.DDA.SyncMethodCalls[n.method.id]=OSF.DDA.SyncMethodCallFactory.manufacture(n)};OSF.DDA.ListType=function(){var n={};return{setListType:function(t,i){n[t]=i},isListType:function(t){return OSF.OUtil.listContainsKey(n,t)},getDescriptor:function(t){return n[t]}}}();OSF.DDA.HostParameterMap=function(n,t){function e(i,f){var a=i?{}:undefined,s,h,o,v,c,l;for(s in i){if(h=i[s],OSF.DDA.ListType.isListType(s)){o=[];for(v in h)o.push(e(h[v],f))}else OSF.OUtil.listContainsKey(r,s)?o=r[s][f](h):f==u&&n.preserveNesting(s)?o=e(h,f):(c=t[s],c?(l=c[f],l&&(o=l[h],o===undefined&&(o=h))):o=h);a[s]=o}return a}function c(i,r){var e,u,h,s;for(u in r)h=n.isComplexType(u)?c(i,t[u][o]):i[u],h!=undefined&&(e||(e={}),s=r[u],s==f&&(s=u),e[s]=n.pack(u,h));return e}function s(i,r,e){var o,l,h,c,y,a,v;e||(e={});for(o in r)if(l=r[o],h=l==f?i:i[l],h===null||h===undefined)e[o]=undefined;else if(h=n.unpack(o,h),n.isComplexType(o))c=t[o][u],n.preserveNesting(o)?e[o]=s(h,c):s(h,c,e);else if(OSF.DDA.ListType.isListType(o)){c={};y=OSF.DDA.ListType.getDescriptor(o);c[y]=f;a=new Array(h.length);for(v in h)a[v]=s(h[v],c);e[o]=a}else e[o]=h;return e}function l(n,i,r){var f=t[n][r],u,o,l;return r=="toHost"?(o=e(i,r),u=c(o,f)):r==h&&(l=s(i,f),u=e(l,r)),u}var h="fromHost",i=this,o="toHost",u=h,f="self",r={};r[Microsoft.Office.WebExtension.Parameters.Data]={toHost:function(n){if(n!=null&&n.rows!==undefined){var t={};t[OSF.DDA.TableDataProperties.TableRows]=n.rows;t[OSF.DDA.TableDataProperties.TableHeaders]=n.headers;n=t}return n},fromHost:function(n){return n}};r[Microsoft.Office.WebExtension.Parameters.SampleData]=r[Microsoft.Office.WebExtension.Parameters.Data];t||(t={});i.addMapping=function(n,i){var e,h,c,l,r,s,a,v;if(i.map){e=i.map;h={};for(c in e)l=e[c],l==f&&(l=c),h[l]=c}else e=i.toHost,h=i.fromHost;if(r=t[n],r){s=r[o];for(a in s)e[a]=s[a];s=r[u];for(v in s)h[v]=s[v]}else r=t[n]={};r[o]=e;r[u]=h};i.toHost=function(n,t){return l(n,t,o)};i.fromHost=function(n,t){return l(n,t,u)};i.self=f;i.addComplexType=function(t){n.addComplexType(t)};i.getDynamicType=function(t){return n.getDynamicType(t)};i.setDynamicType=function(t,i){n.setDynamicType(t,i)};i.dynamicTypes=r;i.doMapValues=function(n,t){return e(n,t)}};OSF.DDA.SpecialProcessor=function(n,t){var i=this;i.addComplexType=function(t){n.push(t)};i.getDynamicType=function(n){return t[n]};i.setDynamicType=function(n,i){t[n]=i};i.isComplexType=function(t){return OSF.OUtil.listContainsValue(n,t)};i.isDynamicType=function(n){return OSF.OUtil.listContainsKey(t,n)};i.preserveNesting=function(n){var t=[];return OSF.DDA.PropertyDescriptors&&t.push(OSF.DDA.PropertyDescriptors.Subset),OSF.DDA.DataNodeEventProperties&&(t=t.concat([OSF.DDA.DataNodeEventProperties.OldNode,OSF.DDA.DataNodeEventProperties.NewNode,OSF.DDA.DataNodeEventProperties.NextSiblingNode])),OSF.OUtil.listContainsValue(t,n)};i.pack=function(n,i){return this.isDynamicType(n)?t[n].toHost(i):i};i.unpack=function(n,i){return this.isDynamicType(n)?t[n].fromHost(i):i}};OSF.DDA.getDecoratedParameterMap=function(n,t){function r(n){var i=null,r,t;if(n)for(i={},r=n.length,t=0;t<r;t++)i[n[t].name]=n[t].value;return i}var i=new OSF.DDA.HostParameterMap(n),f=i.self,u;i.define=function(n){var t={},u=r(n.toHost);n.invertible?t.map=u:n.canonical?t.toHost=t.fromHost=u:(t.toHost=u,t.fromHost=r(n.fromHost));i.addMapping(n.type,t);n.isComplexType&&i.addComplexType(n.type)};for(u in t)i.define(t[u]);return i};OSF.OUtil.setNamespace("DispIdHost",OSF.DDA);OSF.DDA.DispIdHost.Methods={InvokeMethod:"invokeMethod",AddEventHandler:"addEventHandler",RemoveEventHandler:"removeEventHandler",OpenDialog:"openDialog",CloseDialog:"closeDialog",MessageParent:"messageParent"};OSF.DDA.DispIdHost.Delegates={ExecuteAsync:"executeAsync",RegisterEventAsync:"registerEventAsync",UnregisterEventAsync:"unregisterEventAsync",ParameterMap:"parameterMap",OpenDialog:"openDialog",CloseDialog:"closeDialog",MessageParent:"messageParent"};OSF.DDA.DispIdHost.Facade=function(n,t){function o(n,t,i,r){if(typeof n=="number")r||(r=t.getCallArgs(i)),OSF.DDA.issueAsyncResult(r,n,OSF.DDA.ErrorCodeManager.getErrorArgs(n));else throw n;}var s=null,e=this,r={},u=OSF.DDA.AsyncMethodNames,i=OSF.DDA.MethodDispId,a={GoToByIdAsync:i.dispidNavigateToMethod,GetSelectedDataAsync:i.dispidGetSelectedDataMethod,SetSelectedDataAsync:i.dispidSetSelectedDataMethod,GetDocumentCopyChunkAsync:i.dispidGetDocumentCopyChunkMethod,ReleaseDocumentCopyAsync:i.dispidReleaseDocumentCopyMethod,GetDocumentCopyAsync:i.dispidGetDocumentCopyMethod,AddFromSelectionAsync:i.dispidAddBindingFromSelectionMethod,AddFromPromptAsync:i.dispidAddBindingFromPromptMethod,AddFromNamedItemAsync:i.dispidAddBindingFromNamedItemMethod,GetAllAsync:i.dispidGetAllBindingsMethod,GetByIdAsync:i.dispidGetBindingMethod,ReleaseByIdAsync:i.dispidReleaseBindingMethod,GetDataAsync:i.dispidGetBindingDataMethod,SetDataAsync:i.dispidSetBindingDataMethod,AddRowsAsync:i.dispidAddRowsMethod,AddColumnsAsync:i.dispidAddColumnsMethod,DeleteAllDataValuesAsync:i.dispidClearAllRowsMethod,RefreshAsync:i.dispidLoadSettingsMethod,SaveAsync:i.dispidSaveSettingsMethod,GetActiveViewAsync:i.dispidGetActiveViewMethod,GetFilePropertiesAsync:i.dispidGetFilePropertiesMethod,GetOfficeThemeAsync:i.dispidGetOfficeThemeMethod,GetDocumentThemeAsync:i.dispidGetDocumentThemeMethod,ClearFormatsAsync:i.dispidClearFormatsMethod,SetTableOptionsAsync:i.dispidSetTableOptionsMethod,SetFormatsAsync:i.dispidSetFormatsMethod,ExecuteRichApiRequestAsync:i.dispidExecuteRichApiRequestMethod,AppCommandInvocationCompletedAsync:i.dispidAppCommandInvocationCompletedMethod,AddDataPartAsync:i.dispidAddDataPartMethod,GetDataPartByIdAsync:i.dispidGetDataPartByIdMethod,GetDataPartsByNameSpaceAsync:i.dispidGetDataPartsByNamespaceMethod,GetPartXmlAsync:i.dispidGetDataPartXmlMethod,GetPartNodesAsync:i.dispidGetDataPartNodesMethod,DeleteDataPartAsync:i.dispidDeleteDataPartMethod,GetNodeValueAsync:i.dispidGetDataNodeValueMethod,GetNodeXmlAsync:i.dispidGetDataNodeXmlMethod,GetRelativeNodesAsync:i.dispidGetDataNodesMethod,SetNodeValueAsync:i.dispidSetDataNodeValueMethod,SetNodeXmlAsync:i.dispidSetDataNodeXmlMethod,AddDataPartNamespaceAsync:i.dispidAddDataNamespaceMethod,GetDataPartNamespaceAsync:i.dispidGetDataUriByPrefixMethod,GetDataPartPrefixAsync:i.dispidGetDataPrefixByUriMethod,GetNodeTextAsync:i.dispidGetDataNodeTextMethod,SetNodeTextAsync:i.dispidSetDataNodeTextMethod,GetSelectedTask:i.dispidGetSelectedTaskMethod,GetTask:i.dispidGetTaskMethod,GetWSSUrl:i.dispidGetWSSUrlMethod,GetTaskField:i.dispidGetTaskFieldMethod,GetSelectedResource:i.dispidGetSelectedResourceMethod,GetResourceField:i.dispidGetResourceFieldMethod,GetProjectField:i.dispidGetProjectFieldMethod,GetSelectedView:i.dispidGetSelectedViewMethod,GetTaskByIndex:i.dispidGetTaskByIndexMethod,GetResourceByIndex:i.dispidGetResourceByIndexMethod,SetTaskField:i.dispidSetTaskFieldMethod,SetResourceField:i.dispidSetResourceFieldMethod,GetMaxTaskIndex:i.dispidGetMaxTaskIndexMethod,GetMaxResourceIndex:i.dispidGetMaxResourceIndexMethod},c,f,l,h;for(f in a)u[f]&&(r[u[f].id]=a[f]);u=OSF.DDA.SyncMethodNames;i=OSF.DDA.MethodDispId;c={MessageParent:i.dispidMessageParentMethod};for(f in c)u[f]&&(r[u[f].id]=c[f]);u=Microsoft.Office.WebExtension.EventType;i=OSF.DDA.EventDispId;l={SettingsChanged:i.dispidSettingsChangedEvent,DocumentSelectionChanged:i.dispidDocumentSelectionChangedEvent,BindingSelectionChanged:i.dispidBindingSelectionChangedEvent,BindingDataChanged:i.dispidBindingDataChangedEvent,ActiveViewChanged:i.dispidActiveViewChangedEvent,OfficeThemeChanged:i.dispidOfficeThemeChangedEvent,DocumentThemeChanged:i.dispidDocumentThemeChangedEvent,AppCommandInvoked:i.dispidAppCommandInvokedEvent,DialogMessageReceived:i.dispidDialogMessageReceivedEvent,TaskSelectionChanged:i.dispidTaskSelectionChangedEvent,ResourceSelectionChanged:i.dispidResourceSelectionChangedEvent,ViewSelectionChanged:i.dispidViewSelectionChangedEvent,DataNodeInserted:i.dispidDataNodeAddedEvent,DataNodeReplaced:i.dispidDataNodeReplacedEvent,DataNodeDeleted:i.dispidDataNodeDeletedEvent};for(h in l)u[h]&&(r[u[h]]=l[h]);e[OSF.DDA.DispIdHost.Methods.InvokeMethod]=function(i,u,f,e){var h,l,a,y,p,w;try{l=i.id;a=OSF.DDA.AsyncMethodCalls[l];h=a.verifyAndExtractCall(u,f,e);var v=r[l],b=n(l),c=s;window.Excel&&window.Office.context.requirements.isSetSupported("RedirectV1Api")&&(window.Excel._RedirectV1APIs=!0);window.Excel&&window.Excel._RedirectV1APIs&&(c=window.Excel._V1APIMap[l])?(c.preprocess&&(h=c.preprocess(h)),y=new window.Excel.RequestContext,p=c.call(y,h),y.sync().then(function(){var n=p.value,t=n.status;delete n.status;delete n["@odata.type"];c.postprocess&&(n=c.postprocess(n,h));t!=0&&(n=OSF.DDA.ErrorCodeManager.getErrorArgs(t));OSF.DDA.issueAsyncResult(h,t,n)})["catch"](function(){OSF.DDA.issueAsyncResult(h,OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure,s)})):(w=t.toHost?t.toHost(v,h):h,b[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({dispId:v,hostCallArgs:w,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:function(n,i){var r,u;r=n==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess?t.fromHost?t.fromHost(v,i):i:i;u=a.processResponse(n,r,f,h);OSF.DDA.issueAsyncResult(h,n,u)}}))}catch(k){o(k,a,u,h)}};e[OSF.DDA.DispIdHost.Methods.AddEventHandler]=function(i,u,f){function a(n){var t,i;n==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess&&(t=u.addEventHandler(e,l),t||(n=OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed));n!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess&&(i=OSF.DDA.ErrorCodeManager.getErrorArgs(n));OSF.DDA.issueAsyncResult(s,n,i)}var s,e,l,h,c,v;try{h=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];s=h.verifyAndExtractCall(i,f,u);e=s[Microsoft.Office.WebExtension.Parameters.EventType];l=s[Microsoft.Office.WebExtension.Parameters.Handler];u.getEventHandlerCount(e)==0?(c=r[e],v=n(e)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync],v({eventType:e,dispId:c,targetId:f.id||"",onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:a,onEvent:function(n){var i=t.fromHost(c,n);u.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(e,f,i))}})):a(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)}catch(y){o(y,h,i,s)}};e[OSF.DDA.DispIdHost.Methods.RemoveEventHandler]=function(t,i,u){function v(n){var t;n!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess&&(t=OSF.DDA.ErrorCodeManager.getErrorArgs(n));OSF.DDA.issueAsyncResult(e,n,t)}var e,f,c,l,a,h,y,p;try{l=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];e=l.verifyAndExtractCall(t,u,i);f=e[Microsoft.Office.WebExtension.Parameters.EventType];c=e[Microsoft.Office.WebExtension.Parameters.Handler];c===s?(h=i.clearEventHandlers(f),a=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess):(h=i.removeEventHandler(f,c),a=h?OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess:OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist);h&&i.getEventHandlerCount(f)==0?(y=r[f],p=n(f)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync],p({eventType:f,dispId:y,targetId:u.id||"",onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:v})):v(a)}catch(w){o(w,l,t,e)}};e[OSF.DDA.DispIdHost.Methods.OpenDialog]=function(i,u,f){function v(n){var t,i;n!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess?i=OSF.DDA.ErrorCodeManager.getErrorArgs(n):(t={},t[Microsoft.Office.WebExtension.Parameters.Id]=a,t[Microsoft.Office.WebExtension.Parameters.Data]=u,i=l.processResponse(n,t,f,h),OSF.DialogShownStatus.hasDialogShown=!0,u.clearEventHandlers(e),u.clearEventHandlers(c));OSF.DDA.issueAsyncResult(h,n,i)}var h,a,e=Microsoft.Office.WebExtension.EventType.DialogMessageReceived,c=Microsoft.Office.WebExtension.EventType.DialogEventReceived,l;try{if((e==undefined||c==undefined)&&v(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported),OSF.DDA.AsyncMethodNames.DisplayDialogAsync==s){v(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);return}l=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];h=l.verifyAndExtractCall(i,f,u);var p=r[e],y=n(e),w=y[OSF.DDA.DispIdHost.Delegates.OpenDialog]!=undefined?y[OSF.DDA.DispIdHost.Delegates.OpenDialog]:y[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];a=JSON.stringify(h);w({eventType:e,dispId:p,targetId:a,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:v,onEvent:function(n){var s=t.fromHost(p,n),o=OSF.DDA.OMFactory.manufactureEventArgs(e,f,s),r,i;o.type==c&&(r=OSF.DDA.ErrorCodeManager.getErrorArgs(o.error),i={},i[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status||OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,i[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=r.name||r,i[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=r.message||r,o.error=new OSF.DDA.Error(i[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],i[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],i[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]));u.fireOrQueueEvent(o);s[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogClosed&&(u.clearEventHandlers(e),u.clearEventHandlers(c),OSF.DialogShownStatus.hasDialogShown=!1)}})}catch(b){o(b,l,i,h)}};e[OSF.DDA.DispIdHost.Methods.CloseDialog]=function(t,i,u,f){function v(n){s=n;OSF.DialogShownStatus.hasDialogShown=!1}var l,e,a,s=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess,h;try{h=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];l=h.verifyAndExtractCall(t,f,u);e=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;a=Microsoft.Office.WebExtension.EventType.DialogEventReceived;u.clearEventHandlers(e);u.clearEventHandlers(a);var y=r[e],c=n(e),p=c[OSF.DDA.DispIdHost.Delegates.CloseDialog]!=undefined?c[OSF.DDA.DispIdHost.Delegates.CloseDialog]:c[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];p({eventType:e,dispId:y,targetId:i,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)},onComplete:v})}catch(w){o(w,h,t,l)}if(s!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed,OSF.DDA.AsyncMethodNames.CloseAsync.displayName,s);};e[OSF.DDA.DispIdHost.Methods.MessageParent]=function(t,i){var u={},f=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id],e=f.verifyAndExtractCall(t,i,u),o=n(OSF.DDA.SyncMethodNames.MessageParent.id),s=o[OSF.DDA.DispIdHost.Delegates.MessageParent],h=r[OSF.DDA.SyncMethodNames.MessageParent.id];return s({dispId:h,hostCallArgs:e,onCalling:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)},onReceiving:function(){OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)}})}};OSF.DDA.DispIdHost.addAsyncMethods=function(n,t,i){var f,r,u;for(f in t)r=t[f],u=r.displayName,n[u]||OSF.OUtil.defineEnumerableProperty(n,u,{value:function(t){return function(){var r=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];r(t,arguments,n,i)}}(r)})};OSF.DDA.DispIdHost.addEventSupport=function(n,t){var i=OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName,r=OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;n[i]||OSF.OUtil.defineEnumerableProperty(n,i,{value:function(){var i=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];i(arguments,t,n)}});n[r]||OSF.OUtil.defineEnumerableProperty(n,r,{value:function(){var i=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];i(arguments,t,n)}})};OSF.ShowWindowDialogParameterKeys={WindowUrl:"windowUrl",WindowSpecs:"windowSpecs",WindowName:"windowName",MarketPlaceId:"marketPlaceId"},function(n){var t;(function(n){function r(n,i){return OSF.OUtil.addInfoAsHash(n,t,i)}function u(n,r){return OSF.OUtil.parseInfoWithGivenFragment(t,i,null,n,r)}function f(n){if(typeof JSON!="undefined")try{return JSON.stringify(n)}catch(t){}return""}var t="&_app_context=",i="_app_context=";n.addAppContextAsHash=r;n.parseAppContextWithGivenFragment=u;n.serializeObjectToString=f})(t=n.WACUtils||(n.WACUtils={}))}(OfficeExt||(OfficeExt={})),function(n){var u="\n",f=!0,t=null,i="undefined",s=function(){function n(){}return n.isInstanceOfType=function(n,r){if(typeof r===i||r===t)return!1;if(r instanceof n)return f;var u=r.constructor;return u&&typeof u=="function"&&u.__typeName!=="Object"||(u=Object),!!(u===n)||u.inheritsFrom&&u.inheritsFrom(n)||u.implementsInterface&&u.implementsInterface(n)},n}(),e,r,o;n.MsAjaxTypeHelper=s;e=function(){function n(){}var f="Parameter name: {0}";return n.create=function(n,t){var i=new Error(n),r;if(i.message=n,t)for(r in t)i[r]=t[r];return i.popStackFrame(),i},n.parameterCount=function(t){var r="Sys.ParameterCountException: "+(t?t:"Parameter count mismatch."),i=n.create(r,{name:"Sys.ParameterCountException"});return i.popStackFrame(),i},n.argument=function(t,i){var o="Sys.ArgumentException: "+(i?i:"Value does not fall within the expected range."),e;return t&&(o+=u+r.format(f,t)),e=n.create(o,{name:"Sys.ArgumentException",paramName:t}),e.popStackFrame(),e},n.argumentNull=function(t,i){var o="Sys.ArgumentNullException: "+(i?i:"Value cannot be null."),e;return t&&(o+=u+r.format(f,t)),e=n.create(o,{name:"Sys.ArgumentNullException",paramName:t}),e.popStackFrame(),e},n.argumentOutOfRange=function(e,o,s){var h="Sys.ArgumentOutOfRangeException: "+(s?s:"Specified argument was out of the range of valid values."),c;return e&&(h+=u+r.format(f,e)),typeof o!==i&&o!==t&&(h+=u+r.format("Actual value was {0}.",o)),c=n.create(h,{name:"Sys.ArgumentOutOfRangeException",paramName:e,actualValue:o}),c.popStackFrame(),c},n.argumentType=function(t,i,e,o){var s="Sys.ArgumentTypeException: ",h;return s+=o?o:i&&e?r.format("Object of type '{0}' cannot be converted to type '{1}'.",i.getName?i.getName():i,e.getName?e.getName():e):"Object cannot be converted to the required type.",t&&(s+=u+r.format(f,t)),h=n.create(s,{name:"Sys.ArgumentTypeException",paramName:t,actualType:i,expectedType:e}),h.popStackFrame(),h},n.argumentUndefined=function(t,i){var o="Sys.ArgumentUndefinedException: "+(i?i:"Value cannot be undefined."),e;return t&&(o+=u+r.format(f,t)),e=n.create(o,{name:"Sys.ArgumentUndefinedException",paramName:t}),e.popStackFrame(),e},n.invalidOperation=function(t){var r="Sys.InvalidOperationException: "+(t?t:"Operation is not valid due to the current state of the object."),i=n.create(r,{name:"Sys.InvalidOperationException"});return i.popStackFrame(),i},n}();n.MsAjaxError=e;r=function(){function n(){}return n.format=function(n){for(var r,i=[],t=1;t<arguments.length;t++)i[t-1]=arguments[t];return r=n,r.replace(/{(\d+)}/gm,function(n,t){var r=parseInt(t,10);return i[r]===undefined?"{"+t+"}":i[r]})},n.startsWith=function(n,t){return n.substr(0,t.length)===t},n}();n.MsAjaxString=r;o=function(){function n(){}return n.trace=function(){},n}();n.MsAjaxDebug=o;OsfMsAjaxFactory.isMsAjaxLoaded()||(Function.createCallback||(Function.createCallback=function(n,t){var i=Function._validateParams(arguments,[{name:"method",type:Function},{name:"context",mayBeNull:f}]);if(i)throw i;return function(){var u=arguments.length,r,i;if(u>0){for(r=[],i=0;i<u;i++)r[i]=arguments[i];return r[u]=t,n.apply(this,r)}return n.call(this,t)}}),Function.createDelegate||(Function.createDelegate=function(n,t){var i=Function._validateParams(arguments,[{name:"instance",mayBeNull:f},{name:"method",type:Function}]);if(i)throw i;return function(){return t.apply(n,arguments)}}),Function._validateParams||(Function._validateParams=function(n,r,u){var f,s=r.length,e,c,o,h;if(u=u||typeof u===i,f=Function._validateParameterCount(n,r,u),f)return f.popStackFrame(),f;for(e=0,c=n.length;e<c;e++){if(o=r[Math.min(e,s-1)],h=o.name,o.parameterArray)h+="["+(e-s+1)+"]";else if(!u&&e>=s)break;if(f=Function._validateParameter(n[e],o,h),f)return f.popStackFrame(),f}return t}),Function._validateParameterCount||(Function._validateParameterCount=function(n,i,r){var u,s,o=i.length,h=n.length,c,l,a;if(h<o){for(c=o,u=0;u<o;u++)l=i[u],(l.optional||l.parameterArray)&&c--;h<c&&(s=f)}else if(r&&h>o)for(s=f,u=0;u<o;u++)if(i[u].parameterArray){s=!1;break}return s?(a=e.parameterCount(),a.popStackFrame(),a):t}),Function._validateParameter||(Function._validateParameter=function(n,r,u){var f,h=r.type,l=!!r.integer,a=!!r.domElement,v=!!r.mayBeNull,o,s,c;if(f=Function._validateParameterType(n,h,l,a,v,u),f)return f.popStackFrame(),f;if(o=r.elementType,s=!!r.elementMayBeNull,h===Array&&typeof n!==i&&n!==t&&(o||!s))for(var y=!!r.elementInteger,p=!!r.elementDomElement,e=0;e<n.length;e++)if(c=n[e],f=Function._validateParameterType(c,o,y,p,s,u+"["+e+"]"),f)return f.popStackFrame(),f;return t}),Function._validateParameterType||(Function._validateParameterType=function(r,u,f,e,o,s){var h,c;return typeof r===i?o?t:(h=n.MsAjaxError.argumentUndefined(s),h.popStackFrame(),h):r===t?o?t:(h=n.MsAjaxError.argumentNull(s),h.popStackFrame(),h):u&&!n.MsAjaxTypeHelper.isInstanceOfType(u,r)?(h=n.MsAjaxError.argumentType(s,typeof r,u),h.popStackFrame(),h):t}),window.Type||(window.Type=Function),Type.registerNamespace||(Type.registerNamespace=function(n){for(var i=n.split("."),r=window,t=0;t<i.length;t++)r[i[t]]=r[i[t]]||{},r=r[i[t]]}),Type.prototype.registerClass||(Type.prototype.registerClass=function(n){n={}}),typeof Sys===i&&Type.registerNamespace("Sys"),Error.prototype.popStackFrame||(Error.prototype.popStackFrame=function(){var n=this,s,f;if(arguments.length!==0)throw e.parameterCount();if(typeof n.stack!==i&&n.stack!==t&&typeof n.fileName!==i&&n.fileName!==t&&typeof n.lineNumber!==i&&n.lineNumber!==t){for(var r=n.stack.split(u),o=r[0],h=n.fileName+":"+n.lineNumber;typeof o!==i&&o!==t&&o.indexOf(h)===-1;)r.shift(),o=r[0];(s=r[1],typeof s!==i&&s!==t)&&(f=s.match(/@(.*):(\d+)$/),typeof f!==i&&f!==t)&&(n.fileName=f[1],n.lineNumber=parseInt(f[2]),r.shift(),n.stack=r.join(u))}}),OsfMsAjaxFactory.msAjaxError=e,OsfMsAjaxFactory.msAjaxString=r,OsfMsAjaxFactory.msAjaxDebug=o)}(OfficeExt||(OfficeExt={})),function(n){var t="undefined",i=null,u=function(){function u(){}var o='["\\\\\\x00-\\x1F]',e='"',f="g";return u._init=function(){var i=["\\u0000","\\u0001","\\u0002","\\u0003","\\u0004","\\u0005","\\u0006","\\u0007","\\b","\\t","\\n","\\u000b","\\f","\\r","\\u000e","\\u000f","\\u0010","\\u0011","\\u0012","\\u0013","\\u0014","\\u0015","\\u0016","\\u0017","\\u0018","\\u0019","\\u001a","\\u001b","\\u001c","\\u001d","\\u001e","\\u001f"],n,t;for(u._charsToEscape[0]="\\",u._charsToEscapeRegExs["\\"]=new RegExp("\\\\",f),u._escapeChars["\\"]="\\\\",u._charsToEscape[1]=e,u._charsToEscapeRegExs[e]=new RegExp(e,f),u._escapeChars[e]='\\"',n=0;n<32;n++)t=String.fromCharCode(n),u._charsToEscape[n+2]=t,u._charsToEscapeRegExs[t]=new RegExp(t,f),u._escapeChars[t]=i[n]},u.serialize=function(n){var t=new r;return u.serializeWithBuilder(n,t,!1),t.toString()},u.deserialize=function(t,r){if(t.length===0)throw n.MsAjaxError.argument("data","Cannot deserialize empty string.");try{var f=t.replace(u._dateRegEx,"$1new Date($2)");if(r&&u._jsonRegEx.test(f.replace(u._jsonStringRegEx,"")))throw i;return eval("("+f+")")}catch(e){throw n.MsAjaxError.argument("data","Cannot deserialize. The data does not correspond to valid JSON.");}},u.serializeBooleanWithBuilder=function(n,t){t.append(n.toString())},u.serializeNumberWithBuilder=function(t,i){if(isFinite(t))i.append(String(t));else throw n.MsAjaxError.invalidOperation("Cannot serialize non finite numbers.");},u.serializeStringWithBuilder=function(n,t){var r,i;if(t.append(e),u._escapeRegEx.test(n))if(u._charsToEscape.length===0&&u._init(),n.length<128)n=n.replace(u._escapeRegExGlobal,function(n){return u._escapeChars[n]});else for(r=0;r<34;r++)i=u._charsToEscape[r],n.indexOf(i)!==-1&&(n=navigator.userAgent.indexOf("OPR/")>-1||navigator.userAgent.indexOf("Firefox")>-1?n.split(i).join(u._escapeChars[i]):n.replace(u._charsToEscapeRegExs[i],u._escapeChars[i]));t.append(n);t.append(e)},u.serializeWithBuilder=function(i,r,f,e){var o,l,s,h,c,v,a;switch(typeof i){case"object":if(i){if(e){for(l=0;l<e.length;l++)if(e[l]===i)throw n.MsAjaxError.invalidOperation("Cannot serialize object with cyclic reference within child properties.");}else e=[];try{if(n.MsAjaxArray.add(e,i),n.MsAjaxTypeHelper.isInstanceOfType(Number,i))u.serializeNumberWithBuilder(i,r);else if(n.MsAjaxTypeHelper.isInstanceOfType(Boolean,i))u.serializeBooleanWithBuilder(i,r);else if(n.MsAjaxTypeHelper.isInstanceOfType(String,i))u.serializeStringWithBuilder(i,r);else if(n.MsAjaxTypeHelper.isInstanceOfType(Array,i)){for(r.append("["),o=0;o<i.length;++o)o>0&&r.append(","),u.serializeWithBuilder(i[o],r,!1,e);r.append("]")}else{if(n.MsAjaxTypeHelper.isInstanceOfType(Date,i)){r.append('"\\/Date(');r.append(i.getTime());r.append(')\\/"');break}s=[];h=0;for(c in i)n.MsAjaxString.startsWith(c,"$")||(c===u._serverTypeFieldName&&h!==0?(s[h++]=s[0],s[0]=c):s[h++]=c);for(f&&s.sort(),r.append("{"),v=!1,o=0;o<h;o++)a=i[s[o]],typeof a!==t&&typeof a!="function"&&(v?r.append(","):v=!0,u.serializeWithBuilder(s[o],r,f,e),r.append(":"),u.serializeWithBuilder(a,r,f,e));r.append("}")}}finally{n.MsAjaxArray.removeAt(e,e.length-1)}}else r.append("null");break;case"number":u.serializeNumberWithBuilder(i,r);break;case"string":u.serializeStringWithBuilder(i,r);break;case"boolean":u.serializeBooleanWithBuilder(i,r);break;default:r.append("null")}},u.__patchVersion=0,u._charsToEscapeRegExs=[],u._charsToEscape=[],u._dateRegEx=new RegExp('(^|[^\\\\])\\"\\\\/Date\\((-?[0-9]+)(?:[a-zA-Z]|(?:\\+|-)[0-9]{4})?\\)\\\\/\\"',f),u._escapeChars={},u._escapeRegEx=new RegExp(o,"i"),u._escapeRegExGlobal=new RegExp(o,f),u._jsonRegEx=new RegExp("[^,:{}\\[\\]0-9.\\-+Eaeflnr-u \\n\\r\\t]",f),u._jsonStringRegEx=new RegExp('"(\\\\.|[^"\\\\])*"',f),u._serverTypeFieldName="__type",u}(),f,r;n.MsAjaxJavaScriptSerializer=u;f=function(){function n(){}return n.add=function(n,t){n[n.length]=t},n.removeAt=function(n,t){n.splice(t,1)},n.clone=function(n){return n.length===1?[n[0]]:Array.apply(i,n)},n.remove=function(t,i){var r=n.indexOf(t,i);return r>=0&&t.splice(r,1),r>=0},n.indexOf=function(n,i,r){var f,u;if(typeof i===t)return-1;if(f=n.length,f!==0)for(r=+r,isNaN(r)?r=0:(isFinite(r)&&(r=r-r%1),r<0&&(r=Math.max(0,f+r))),u=r;u<f;u++)if(typeof n[u]!==t&&n[u]===i)return u;return-1},n}();n.MsAjaxArray=f;r=function(){function n(n){this._parts=typeof n!==t&&n!==i&&n!==""?[n.toString()]:[];this._value={};this._len=0}return n.prototype.append=function(n){this._parts[this._parts.length]=n},n.prototype.toString=function(n){var f=this,r,e,u;if(n=n||"",r=f._parts,f._len!==r.length&&(f._value={},f._len=r.length),e=f._value,typeof e[n]===t){if(n!=="")for(u=0;u<r.length;)typeof r[u]===t||r[u]===""||r[u]===i?r.splice(u,1):u++;e[n]=f._parts.join(n)}return e[n]},n}();n.MsAjaxStringBuilder=r;OsfMsAjaxFactory.isMsAjaxLoaded()||(OsfMsAjaxFactory.msAjaxSerializer=u)}(OfficeExt||(OfficeExt={}));OSF.OUtil.setNamespace("Microsoft",window);OSF.OUtil.setNamespace("Office",Microsoft);OSF.OUtil.setNamespace("Common",Microsoft.Office);OSF.SerializerVersion={MsAjax:0,Browser:1},function(n){function t(){return!0}n.appSpecificCheckOrigin=t}(OfficeExt||(OfficeExt={})),function(n){"use strict";function s(n,t,i){n.addEventListener?n.addEventListener(t,i,!1):n.attachEvent&&n.attachEvent("on"+t,i)}function h(){return OsfMsAjaxFactory.msAjaxSerializer?OsfMsAjaxFactory.msAjaxSerializer:null}function c(i,s,h){var c;if(!s)return h(i);if(n.JSON&&n.JSON.parse)return n.JSON.parse(i);if(c=i.replace(r,"[]"),c=c.replace(u,"[]"),c=c.replace(f,"[]"),e.test(c))throw t;if(o.test(c))throw t;try{eval("("+i+")")}catch(l){throw t;}}function i(){var n=h(),t;return n===null||typeof n.deserialize!="function"?!1:n.__patchVersion>=1?!0:(t=n.deserialize,n.deserialize=function(n){return c(n,!0,t)},n.__patchVersion=1,!0)}var r=new RegExp('"(\\\\.|[^"\\\\])*"',"g"),u=new RegExp("\\b(true|false|null)\\b","g"),f=new RegExp("-?(0|([1-9]\\d*))(\\.\\d+)?([eE][+-]?\\d+)?","g"),e=new RegExp("[^{:,\\[\\s](?=\\s*\\[)"),o=new RegExp("[^\\s\\[\\]{}:,]"),t="Cannot deserialize. The data does not correspond to valid JSON.";i()||s(n,"load",function(){i()})}(window);Microsoft.Office.Common.InvokeType={async:0,sync:1,asyncRegisterEvent:2,asyncUnregisterEvent:3,syncRegisterEvent:4,syncUnregisterEvent:5};Microsoft.Office.Common.InvokeResultCode={noError:0,errorInRequest:-1,errorHandlingRequest:-2,errorInResponse:-3,errorHandlingResponse:-4,errorHandlingRequestAccessDenied:-5,errorHandlingMethodCallTimedout:-6};Microsoft.Office.Common.MessageType={request:0,response:1};Microsoft.Office.Common.ActionType={invoke:0,registerEvent:1,unregisterEvent:2};Microsoft.Office.Common.ResponseType={forCalling:0,forEventing:1};Microsoft.Office.Common.MethodObject=function(n,t,i){this._method=n;this._invokeType=t;this._blockingOthers=i};Microsoft.Office.Common.MethodObject.prototype={getMethod:function(){return this._method},getInvokeType:function(){return this._invokeType},getBlockingFlag:function(){return this._blockingOthers}};Microsoft.Office.Common.EventMethodObject=function(n,t){this._registerMethodObject=n;this._unregisterMethodObject=t};Microsoft.Office.Common.EventMethodObject.prototype={getRegisterMethodObject:function(){return this._registerMethodObject},getUnregisterMethodObject:function(){return this._unregisterMethodObject}};Microsoft.Office.Common.ServiceEndPoint=function(n){var t=this,i=Function._validateParams(arguments,[{name:"serviceEndPointId",type:String,mayBeNull:!1}]);if(i)throw i;t._methodObjectList={};t._eventHandlerProxyList={};t._Id=n;t._conversations={};t._policyManager=null;t._appDomains={};t._onHandleRequestError=null};Microsoft.Office.Common.ServiceEndPoint.prototype={registerMethod:function(n,t,i,r){var f="invokeType",u=!1,e=Function._validateParams(arguments,[{name:"methodName",type:String,mayBeNull:u},{name:"method",type:Function,mayBeNull:u},{name:f,type:Number,mayBeNull:u},{name:"blockingOthers",type:Boolean,mayBeNull:u}]),o;if(e)throw e;if(i!==Microsoft.Office.Common.InvokeType.async&&i!==Microsoft.Office.Common.InvokeType.sync)throw OsfMsAjaxFactory.msAjaxError.argument(f);o=new Microsoft.Office.Common.MethodObject(t,i,r);this._methodObjectList[n]=o},unregisterMethod:function(n){var t=Function._validateParams(arguments,[{name:"methodName",type:String,mayBeNull:!1}]);if(t)throw t;delete this._methodObjectList[n]},registerEvent:function(n,t,i){var r=!1,u=Function._validateParams(arguments,[{name:"eventName",type:String,mayBeNull:r},{name:"registerMethod",type:Function,mayBeNull:r},{name:"unregisterMethod",type:Function,mayBeNull:r}]),f;if(u)throw u;f=new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(t,Microsoft.Office.Common.InvokeType.syncRegisterEvent,r),new Microsoft.Office.Common.MethodObject(i,Microsoft.Office.Common.InvokeType.syncUnregisterEvent,r));this._methodObjectList[n]=f},registerEventEx:function(n,t,i,r,u){var f=!1,e=Function._validateParams(arguments,[{name:"eventName",type:String,mayBeNull:f},{name:"registerMethod",type:Function,mayBeNull:f},{name:"registerMethodInvokeType",type:Number,mayBeNull:f},{name:"unregisterMethod",type:Function,mayBeNull:f},{name:"unregisterMethodInvokeType",type:Number,mayBeNull:f}]),o;if(e)throw e;o=new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(t,i,f),new Microsoft.Office.Common.MethodObject(r,u,f));this._methodObjectList[n]=o},unregisterEvent:function(n){var t=Function._validateParams(arguments,[{name:"eventName",type:String,mayBeNull:!1}]);if(t)throw t;this.unregisterMethod(n)},registerConversation:function(n,t,i,r){var f="appDomains",u=!0,e=Function._validateParams(arguments,[{name:"conversationId",type:String,mayBeNull:!1},{name:"conversationUrl",type:String,mayBeNull:!1,optional:u},{name:f,type:Object,mayBeNull:u,optional:u},{name:"serializerVersion",type:Number,mayBeNull:u,optional:u}]);if(e)throw e;if(i){if(!(i instanceof Array))throw OsfMsAjaxFactory.msAjaxError.argument(f);this._appDomains[n]=i}this._conversations[n]={url:t,serializerVersion:r}},unregisterConversation:function(n){var t=Function._validateParams(arguments,[{name:"conversationId",type:String,mayBeNull:!1}]);if(t)throw t;delete this._conversations[n]},setPolicyManager:function(n){var t="policyManager",i=Function._validateParams(arguments,[{name:t,type:Object,mayBeNull:!1}]);if(i)throw i;if(!n.checkPermission)throw OsfMsAjaxFactory.msAjaxError.argument(t);this._policyManager=n},getPolicyManager:function(){return this._policyManager}};Microsoft.Office.Common.ClientEndPoint=function(n,t,i,r){var f="targetWindow",u=this,e=Function._validateParams(arguments,[{name:"conversationId",type:String,mayBeNull:!1},{name:f,mayBeNull:!1},{name:"targetUrl",type:String,mayBeNull:!1},{name:"serializerVersion",type:Number,mayBeNull:!0,optional:!0}]);if(e)throw e;if(!t.postMessage)throw OsfMsAjaxFactory.msAjaxError.argument(f);u._conversationId=n;u._targetWindow=t;u._targetUrl=i;u._callingIndex=0;u._callbackList={};u._eventHandlerList={};u._serializerVersion=r!=null?r:OSF.SerializerVersion.MsAjax};Microsoft.Office.Common.ClientEndPoint.prototype={invoke:function(n,t,i){var r=this,f=Function._validateParams(arguments,[{name:"targetMethodName",type:String,mayBeNull:!1},{name:"callback",type:Function,mayBeNull:!0},{name:"param",mayBeNull:!0}]),o,s;if(f)throw f;var u=r._callingIndex++,h=new Date,e={callback:t,createdOn:h.getTime()};i&&typeof i=="object"&&typeof i.__timeout__=="number"&&(e.timeout=i.__timeout__,delete i.__timeout__);r._callbackList[u]=e;try{o=new Microsoft.Office.Common.Request(n,Microsoft.Office.Common.ActionType.invoke,r._conversationId,u,i);s=Microsoft.Office.Common.MessagePackager.envelope(o,r._serializerVersion);r._targetWindow.postMessage(s,r._targetUrl);Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()}catch(c){try{t!==null&&t(Microsoft.Office.Common.InvokeResultCode.errorInRequest,c)}finally{delete r._callbackList[u]}}},registerForEvent:function(n,t,i,r){var u=this,e=Function._validateParams(arguments,[{name:"targetEventName",type:String,mayBeNull:!1},{name:"eventHandler",type:Function,mayBeNull:!1},{name:"callback",type:Function,mayBeNull:!0},{name:"data",mayBeNull:!0,optional:!0}]),f,o,s,h;if(e)throw e;f=u._callingIndex++;o=new Date;u._callbackList[f]={callback:i,createdOn:o.getTime()};try{s=new Microsoft.Office.Common.Request(n,Microsoft.Office.Common.ActionType.registerEvent,u._conversationId,f,r);h=Microsoft.Office.Common.MessagePackager.envelope(s,u._serializerVersion);u._targetWindow.postMessage(h,u._targetUrl);Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();u._eventHandlerList[n]=t}catch(c){try{i!==null&&i(Microsoft.Office.Common.InvokeResultCode.errorInRequest,c)}finally{delete u._callbackList[f]}}},unregisterForEvent:function(n,t,i){var r=this,f=Function._validateParams(arguments,[{name:"targetEventName",type:String,mayBeNull:!1},{name:"callback",type:Function,mayBeNull:!0},{name:"data",mayBeNull:!0,optional:!0}]),u,e,o,s;if(f)throw f;u=r._callingIndex++;e=new Date;r._callbackList[u]={callback:t,createdOn:e.getTime()};try{o=new Microsoft.Office.Common.Request(n,Microsoft.Office.Common.ActionType.unregisterEvent,r._conversationId,u,i);s=Microsoft.Office.Common.MessagePackager.envelope(o,r._serializerVersion);r._targetWindow.postMessage(s,r._targetUrl);Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()}catch(h){try{t!==null&&t(Microsoft.Office.Common.InvokeResultCode.errorInRequest,h)}finally{delete r._callbackList[u]}}finally{delete r._eventHandlerList[n]}}};Microsoft.Office.Common.XdmCommunicationManager=function(){function tt(n){for(var t in f)if(f[t]._conversations[n])return f[t];OsfMsAjaxFactory.msAjaxDebug.trace(v);throw OsfMsAjaxFactory.msAjaxError.argument(o);}function it(n){var t=i[n];if(!t){OsfMsAjaxFactory.msAjaxDebug.trace(v);throw OsfMsAjaxFactory.msAjaxError.argument(o);}return t}function rt(t,i){var r=t._methodObjectList[i._actionName],u;if(!r){OsfMsAjaxFactory.msAjaxDebug.trace("The specified method is not registered on service endpoint:"+i._actionName);throw OsfMsAjaxFactory.msAjaxError.argument("messageObject");}return u=n,i._actionType===Microsoft.Office.Common.ActionType.invoke?r:i._actionType===Microsoft.Office.Common.ActionType.registerEvent?r.getRegisterMethodObject():r.getUnregisterMethodObject()}function ut(n){h.push(n)}function ft(){if(r!==n){if(!s)if(h.length>0){var t=h.shift();k(t)}else clearInterval(r),r=n}else OsfMsAjaxFactory.msAjaxDebug.trace(a)}function k(n){s=n.getInvokeBlockingFlag();n.invoke();c=(new Date).getTime()}function et(){var r,f,s,e,h,o,t;if(u){f=0;s=new Date;for(h in i){r=i[h];for(o in r._callbackList)if(t=r._callbackList[o],e=t.timeout?t.timeout:w,e>=0&&Math.abs(s.getTime()-t.createdOn)>=e)try{t.callback&&t.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout,n)}finally{delete r._callbackList[o]}else f++}f===0&&(clearInterval(u),u=n)}else OsfMsAjaxFactory.msAjaxDebug.trace(a)}function ot(){s=t}function st(n){if(window.addEventListener)window.addEventListener("message",n,t);else if(navigator.userAgent.indexOf("MSIE")>-1&&window.attachEvent)window.attachEvent("onmessage",n);else{OsfMsAjaxFactory.msAjaxDebug.trace("Browser doesn't support the required API.");throw OsfMsAjaxFactory.msAjaxError.argument("Browser");}}function l(n,i){var f=t,r,u;return n===e?e:!n||!i||!n.length||!i.length?f:(r=document.createElement("a"),u=document.createElement("a"),r.href=n,u.href=i,f=d(r,u),delete r,u,f)}function ht(n,i){var u=t,f,e,r;if(!i||!i.length||!n||!(n instanceof Array)||!n.length)return u;for(f=document.createElement("a"),e=document.createElement("a"),f.href=i,r=0;r<n.length&&!u;r++)n[r].indexOf("://")!==-1&&(e.href=n[r],u=d(f,e));return delete f,e,u}function d(n,t){return n.hostname==t.hostname&&n.protocol==t.protocol&&n.port==t.port}function ct(i){var et="Access Denied",u,f,w,d,o,a,g,nt,lt,at,h,v,p;if(i.data!=""){f=OSF.SerializerVersion.MsAjax;w=i.data;try{u=Microsoft.Office.Common.MessagePackager.unenvelope(w,OSF.SerializerVersion.Browser);f=u._serializerVersion!=n?u._serializerVersion:f}catch(b){}if(f!=OSF.SerializerVersion.Browser)try{u=Microsoft.Office.Common.MessagePackager.unenvelope(w,f)}catch(b){return}if(typeof u._messageType=="undefined")return;if(u._messageType===Microsoft.Office.Common.MessageType.request){d=i.origin==n||i.origin=="null"?u._origin:i.origin;try{if(o=tt(u._conversationId),a=o._conversations[u._conversationId],f=a.serializerVersion!=n?a.serializerVersion:f,!l(a.url,i.origin)&&!ht(o._appDomains[u._conversationId],i.origin)&&!OfficeExt.appSpecificCheckOrigin(a.url,i,u,l))throw"Failed origin check";if(g=o.getPolicyManager(),g&&!g.checkPermission(u._conversationId,u._actionName,u._data))throw et;var vt=rt(o,u),yt=new Microsoft.Office.Common.InvokeCompleteCallback(i.source,d,u._actionName,u._conversationId,u._correlationId,ot,f),st=new Microsoft.Office.Common.Invoker(vt,u._data,yt,o._eventHandlerProxyList,u._conversationId,u._actionName,f),ct=e;r==n&&((c==n||(new Date).getTime()-c>y)&&!s?(k(st),ct=t):r=setInterval(ft,y));ct&&ut(st)}catch(b){o&&o._onHandleRequestError&&o._onHandleRequestError(u,b);nt=Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;b==et&&(nt=Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied);lt=new Microsoft.Office.Common.Response(u._actionName,u._conversationId,u._correlationId,nt,Microsoft.Office.Common.ResponseType.forCalling,b);at=Microsoft.Office.Common.MessagePackager.envelope(lt,f);i.source&&i.source.postMessage&&i.source.postMessage(at,d)}}else if(u._messageType===Microsoft.Office.Common.MessageType.response){if(h=it(u._conversationId),h._serializerVersion=f,!l(h._targetUrl,i.origin))throw"Failed orgin check";if(u._responseType===Microsoft.Office.Common.ResponseType.forCalling){if(v=h._callbackList[u._correlationId],v)try{v.callback&&v.callback(u._errorCode,u._data)}finally{delete h._callbackList[u._correlationId]}}else p=h._eventHandlerList[u._actionName],p!==undefined&&p!==n&&p(u._data)}else return}}function g(){b||(st(ct),b=e)}var e=!0,a="channel is not ready.",o="conversationId",v="Unknown conversation Id.",t=!1,n=null,h=[],c=n,r=n,y=10,s=t,u=n,nt=2e3,p=65e3,w=p,f={},i={},b=t;return{connect:function(n,t,r,u){var f=i[n];return f||(g(),f=new Microsoft.Office.Common.ClientEndPoint(n,t,r,u),i[n]=f),f},getClientEndPoint:function(n){var r=Function._validateParams(arguments,[{name:o,type:String,mayBeNull:t}]);if(r)throw r;return i[n]},createServiceEndPoint:function(n){g();var t=new Microsoft.Office.Common.ServiceEndPoint(n);return f[n]=t,t},getServiceEndPoint:function(n){var i=Function._validateParams(arguments,[{name:"serviceEndPointId",type:String,mayBeNull:t}]);if(i)throw i;return f[n]},deleteClientEndPoint:function(n){var r=Function._validateParams(arguments,[{name:o,type:String,mayBeNull:t}]);if(r)throw r;delete i[n]},_setMethodTimeout:function(n){var i=Function._validateParams(arguments,[{name:"methodTimeout",type:Number,mayBeNull:t}]);if(i)throw i;w=n<=0?p:n},_startMethodTimeoutTimer:function(){u||(u=setInterval(et,nt))}}}();Microsoft.Office.Common.Message=function(n,t,i,r,u){var e=!1,f=this,o=Function._validateParams(arguments,[{name:"messageType",type:Number,mayBeNull:e},{name:"actionName",type:String,mayBeNull:e},{name:"conversationId",type:String,mayBeNull:e},{name:"correlationId",mayBeNull:e},{name:"data",mayBeNull:!0,optional:!0}]);if(o)throw o;f._messageType=n;f._actionName=t;f._conversationId=i;f._correlationId=r;f._origin=window.location.href;f._data=typeof u=="undefined"?null:u};Microsoft.Office.Common.Message.prototype={getActionName:function(){return this._actionName},getConversationId:function(){return this._conversationId},getCorrelationId:function(){return this._correlationId},getOrigin:function(){return this._origin},getData:function(){return this._data},getMessageType:function(){return this._messageType}};Microsoft.Office.Common.Request=function(n,t,i,r,u){Microsoft.Office.Common.Request.uber.constructor.call(this,Microsoft.Office.Common.MessageType.request,n,i,r,u);this._actionType=t};OSF.OUtil.extend(Microsoft.Office.Common.Request,Microsoft.Office.Common.Message);Microsoft.Office.Common.Request.prototype.getActionType=function(){return this._actionType};Microsoft.Office.Common.Response=function(n,t,i,r,u,f){Microsoft.Office.Common.Response.uber.constructor.call(this,Microsoft.Office.Common.MessageType.response,n,t,i,f);this._errorCode=r;this._responseType=u};OSF.OUtil.extend(Microsoft.Office.Common.Response,Microsoft.Office.Common.Message);Microsoft.Office.Common.Response.prototype.getErrorCode=function(){return this._errorCode};Microsoft.Office.Common.Response.prototype.getResponseType=function(){return this._responseType};Microsoft.Office.Common.MessagePackager={envelope:function(n,t){return t==OSF.SerializerVersion.Browser&&typeof JSON!="undefined"?(typeof n=="object"&&(n._serializerVersion=t),JSON.stringify(n)):(typeof n=="object"&&(n._serializerVersion=OSF.SerializerVersion.MsAjax),OsfMsAjaxFactory.msAjaxSerializer.serialize(n))},unenvelope:function(n,t){return t==OSF.SerializerVersion.Browser&&typeof JSON!="undefined"?JSON.parse(n):OsfMsAjaxFactory.msAjaxSerializer.deserialize(n,!0)}};Microsoft.Office.Common.ResponseSender=function(n,t,i,r,u,f,e){var h=!1,o=this,c=Function._validateParams(arguments,[{name:"requesterWindow",mayBeNull:h},{name:"requesterUrl",type:String,mayBeNull:h},{name:"actionName",type:String,mayBeNull:h},{name:"conversationId",type:String,mayBeNull:h},{name:"correlationId",mayBeNull:h},{name:"responsetype",type:Number,maybeNull:h},{name:"serializerVersion",type:Number,maybeNull:!0,optional:!0}]),s;if(c)throw c;o._requesterWindow=n;o._requesterUrl=t;o._actionName=i;o._conversationId=r;o._correlationId=u;o._invokeResultCode=Microsoft.Office.Common.InvokeResultCode.noError;o._responseType=f;s=o;o._send=function(n){try{var t=new Microsoft.Office.Common.Response(s._actionName,s._conversationId,s._correlationId,s._invokeResultCode,s._responseType,n),i=Microsoft.Office.Common.MessagePackager.envelope(t,e);s._requesterWindow.postMessage(i,s._requesterUrl)}catch(r){OsfMsAjaxFactory.msAjaxDebug.trace("ResponseSender._send error:"+r.message)}}};Microsoft.Office.Common.ResponseSender.prototype={getRequesterWindow:function(){return this._requesterWindow},getRequesterUrl:function(){return this._requesterUrl},getActionName:function(){return this._actionName},getConversationId:function(){return this._conversationId},getCorrelationId:function(){return this._correlationId},getSend:function(){return this._send},setResultCode:function(n){this._invokeResultCode=n}};Microsoft.Office.Common.InvokeCompleteCallback=function(n,t,i,r,u,f,e){var s=this,o;Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(s,n,t,i,r,u,Microsoft.Office.Common.ResponseType.forCalling,e);s._postCallbackHandler=f;o=s;s._send=function(n){try{var t=new Microsoft.Office.Common.Response(o._actionName,o._conversationId,o._correlationId,o._invokeResultCode,o._responseType,n),i=Microsoft.Office.Common.MessagePackager.envelope(t,e);o._requesterWindow.postMessage(i,o._requesterUrl);o._postCallbackHandler()}catch(r){OsfMsAjaxFactory.msAjaxDebug.trace("InvokeCompleteCallback._send error:"+r.message)}}};OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback,Microsoft.Office.Common.ResponseSender);Microsoft.Office.Common.Invoker=function(n,t,i,r,u,f,e){var s=!0,h=!1,o=this,c=Function._validateParams(arguments,[{name:"methodObject",mayBeNull:h},{name:"paramValue",mayBeNull:s},{name:"invokeCompleteCallback",mayBeNull:h},{name:"eventHandlerProxyList",mayBeNull:s},{name:"conversationId",type:String,mayBeNull:h},{name:"eventName",type:String,mayBeNull:h},{name:"serializerVersion",type:Number,mayBeNull:s,optional:s}]);if(c)throw c;o._methodObject=n;o._param=t;o._invokeCompleteCallback=i;o._eventHandlerProxyList=r;o._conversationId=u;o._eventName=f;o._serializerVersion=e};Microsoft.Office.Common.Invoker.prototype={invoke:function(){var n=this,t,i,u,r,f;try{switch(n._methodObject.getInvokeType()){case Microsoft.Office.Common.InvokeType.async:n._methodObject.getMethod()(n._param,n._invokeCompleteCallback.getSend());break;case Microsoft.Office.Common.InvokeType.sync:t=n._methodObject.getMethod()(n._param);n._invokeCompleteCallback.getSend()(t);break;case Microsoft.Office.Common.InvokeType.syncRegisterEvent:i=n._createEventHandlerProxyObject(n._invokeCompleteCallback);t=n._methodObject.getMethod()(i.getSend(),n._param);n._eventHandlerProxyList[n._conversationId+n._eventName]=i.getSend();n._invokeCompleteCallback.getSend()(t);break;case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:u=n._eventHandlerProxyList[n._conversationId+n._eventName];t=n._methodObject.getMethod()(u,n._param);delete n._eventHandlerProxyList[n._conversationId+n._eventName];n._invokeCompleteCallback.getSend()(t);break;case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:r=n._createEventHandlerProxyObject(n._invokeCompleteCallback);n._methodObject.getMethod()(r.getSend(),n._invokeCompleteCallback.getSend(),n._param);n._eventHandlerProxyList[n._callerId+n._eventName]=r.getSend();break;case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:f=n._eventHandlerProxyList[n._callerId+n._eventName];n._methodObject.getMethod()(f,n._invokeCompleteCallback.getSend(),n._param);delete n._eventHandlerProxyList[n._callerId+n._eventName]}}catch(e){n._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);n._invokeCompleteCallback.getSend()(e)}},getInvokeBlockingFlag:function(){return this._methodObject.getBlockingFlag()},_createEventHandlerProxyObject:function(n){return new Microsoft.Office.Common.ResponseSender(n.getRequesterWindow(),n.getRequesterUrl(),n.getActionName(),n.getConversationId(),n.getCorrelationId(),Microsoft.Office.Common.ResponseType.forEventing,this._serializerVersion)}};OSF.OUtil.setNamespace("WAC",OSF.DDA);OSF.DDA.WAC.UniqueArguments={Data:"Data",Properties:"Properties",BindingRequest:"DdaBindingsMethod",BindingResponse:"Bindings",SingleBindingResponse:"singleBindingResponse",GetData:"DdaGetBindingData",AddRowsColumns:"DdaAddRowsColumns",SetData:"DdaSetBindingData",ClearFormats:"DdaClearBindingFormats",SetFormats:"DdaSetBindingFormats",SettingsRequest:"DdaSettingsMethod",BindingEventSource:"ddaBinding",ArrayData:"ArrayData"};OSF.OUtil.setNamespace("Delegate",OSF.DDA.WAC);OSF.DDA.WAC.Delegate.SpecialProcessor=function(){var n=[OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,OSF.DDA.WAC.UniqueArguments.BindingRequest,OSF.DDA.WAC.UniqueArguments.BindingResponse,OSF.DDA.WAC.UniqueArguments.GetData,OSF.DDA.WAC.UniqueArguments.AddRowsColumns,OSF.DDA.WAC.UniqueArguments.SetData,OSF.DDA.WAC.UniqueArguments.ClearFormats,OSF.DDA.WAC.UniqueArguments.SetFormats,OSF.DDA.WAC.UniqueArguments.SettingsRequest,OSF.DDA.WAC.UniqueArguments.BindingEventSource],t={};OSF.DDA.WAC.Delegate.SpecialProcessor.uber.constructor.call(this,n,t)};OSF.OUtil.extend(OSF.DDA.WAC.Delegate.SpecialProcessor,OSF.DDA.SpecialProcessor);OSF.DDA.WAC.Delegate.ParameterMap=OSF.DDA.getDecoratedParameterMap(new OSF.DDA.WAC.Delegate.SpecialProcessor,[]);OSF.OUtil.setNamespace("WAC",OSF.DDA);OSF.OUtil.setNamespace("Delegate",OSF.DDA.WAC);OSF.DDA.WAC.getDelegateMethods=function(){var n={};return n[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.WAC.Delegate.executeAsync,n[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.WAC.Delegate.registerEventAsync,n[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.WAC.Delegate.unregisterEventAsync,n[OSF.DDA.DispIdHost.Delegates.OpenDialog]=OSF.DDA.WAC.Delegate.openDialog,n[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.WAC.Delegate.messageParent,n[OSF.DDA.DispIdHost.Delegates.CloseDialog]=OSF.DDA.WAC.Delegate.closeDialog,n};OSF.DDA.WAC.Delegate.version=1;OSF.DDA.WAC.Delegate.executeAsync=function(n){n.hostCallArgs||(n.hostCallArgs={});n.hostCallArgs.DdaMethod={ControlId:OSF._OfficeAppFactory.getId(),Version:OSF.DDA.WAC.Delegate.version,DispatchId:n.dispId};n.hostCallArgs.__timeout__=-1;n.onCalling&&n.onCalling();var t=(new Date).getTime();OSF.getClientEndPoint().invoke("executeMethod",function(i,r){n.onReceiving&&n.onReceiving();var u;if(i==Microsoft.Office.Common.InvokeResultCode.noError)OSF.DDA.WAC.Delegate.version=r.Version,u=r.Error;else switch(i){case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:u=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;break;default:u=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError}n.onComplete&&n.onComplete(u,r);OSF.AppTelemetry&&OSF.AppTelemetry.onMethodDone(n.dispId,n.hostCallArgs,Math.abs((new Date).getTime()-t),u)},n.hostCallArgs)};OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent=function(n,t){var i=(new Date).getTime();return function(r,u){t.onReceiving&&t.onReceiving();var f;if(r!=Microsoft.Office.Common.InvokeResultCode.noError)switch(r){case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:f=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;break;default:f=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError}else f=u?u.Error?u.Error:OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess:OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;t.onComplete&&t.onComplete(f);OSF.AppTelemetry&&OSF.AppTelemetry.onRegisterDone(n,t.dispId,Math.abs((new Date).getTime()-i),f)}};OSF.DDA.WAC.Delegate.registerEventAsync=function(n){n.onCalling&&n.onCalling();OSF.getClientEndPoint().registerForEvent(OSF.DDA.getXdmEventName(n.targetId,n.eventType),function(t){n.onEvent&&n.onEvent(t);OSF.AppTelemetry&&OSF.AppTelemetry.onEventDone(n.dispId)},OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(!0,n),{controlId:OSF._OfficeAppFactory.getId(),eventDispId:n.dispId,targetId:n.targetId})};OSF.DDA.WAC.Delegate.unregisterEventAsync=function(n){n.onCalling&&n.onCalling();OSF.getClientEndPoint().unregisterForEvent(OSF.DDA.getXdmEventName(n.targetId,n.eventType),OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(!1,n),{controlId:OSF._OfficeAppFactory.getId(),eventDispId:n.dispId,targetId:n.targetId})},function(n){var t;(function(t){var i;(function(t){function h(t){function c(n){if(n.source==r)try{u([OSF.DialogMessageType.DialogMessageReceived,n.data])}catch(t){OsfMsAjaxFactory.msAjaxDebug.trace("Error happened during message handler. Exception: "+t)}}function l(){try{(r==i||r.closed)&&(window.clearInterval(e),u([OSF.DialogMessageType.DialogClosed]))}catch(n){OsfMsAjaxFactory.msAjaxDebug.trace("Error happened during check or handle window close. Exception: "+n)}}var o=t.input,h=OSF._OfficeAppFactory.getInitializationHelper()._appContext,s;h._id=o[OSF.ShowWindowDialogParameterKeys.MarketPlaceId];s=o[OSF.ShowWindowDialogParameterKeys.WindowUrl];s=n.WACUtils.addAppContextAsHash(s,n.WACUtils.serializeObjectToString(h));r=window.open(s,o[OSF.ShowWindowDialogParameterKeys.WindowName],o[OSF.ShowWindowDialogParameterKeys.WindowSpecs]);window.addEventListener("message",c);e=window.setInterval(l,1e3);f!=i?f(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess):OsfMsAjaxFactory.msAjaxDebug.trace("showDialogCallback can not be null.");t.completed()}function c(n){r!=i?(r.postMessage(o,s()),window.clearInterval(e),r=i,n(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)):n(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)}function l(n){var t=n.hostCallArgs[Microsoft.Office.WebExtension.Parameters.MessageToParent];window.opener.postMessage(t,s())}function a(){function n(n){n.source==window.opener&&n.data.indexOf(o)>-1&&window.close()}window.addEventListener("message",n)}function v(n,t){u=n;f=t}function s(){return window.location.origin?window.location.origin:window.location.protocol+"//"+window.location.hostname+(window.location.port?":"+window.location.port:"")}var i=null,r=i,u=i,o="action=closeDialog",f=i,e=-1;t.showDialog=h;t.closeDialog=c;t.messageParent=l;t.registerMessageReceivedEvent=a;t.setHandlerAndShowDialogCallback=v})(i=t.Dialog||(t.Dialog={}))})(t=n.AddinNativeAction||(n.AddinNativeAction={}))}(OfficeExt||(OfficeExt={}));OSF.OUtil.setNamespace("WebApp",OSF);OSF.WebApp.AddHostInfoAndXdmInfo=function(n){return OSF._OfficeAppFactory.getWindowLocationSearch&&OSF._OfficeAppFactory.getWindowLocationHash?n+OSF._OfficeAppFactory.getWindowLocationSearch()+OSF._OfficeAppFactory.getWindowLocationHash():n};OSF.WebApp._UpdateLinksForHostAndXdmInfo=function(){for(var r,i,t=document.querySelectorAll("a[data-officejs-navigate]"),n=0;n<t.length;n++)OSF.WebApp._isGoodUrl(t[n].href)&&(t[n].href=OSF.WebApp.AddHostInfoAndXdmInfo(t[n].href));for(r=document.querySelectorAll("form[data-officejs-navigate]"),n=0;n<r.length;n++)i=r[n],OSF.WebApp._isGoodUrl(i.action)&&(i.action=OSF.WebApp.AddHostInfoAndXdmInfo(i.action))};OSF.WebApp._isGoodUrl=function(n){if(typeof n=="undefined")return!1;n=n.trim();var i=n.indexOf(":"),t=i>0?n.substr(0,i):null,r=t!==null?t.toLowerCase()==="http"||t.toLowerCase()==="https":!0;return r&&n!="#"&&n!="/"&&n!=""&&n!=OSF._OfficeAppFactory.getWebAppState().webAppUrl};OSF.InitializationHelper=function(n,t,i,r,u){var f=this,e;f._hostInfo=n;f._webAppState=t;f._context=i;f._settings=r;f._hostFacade=u;f._appContext={};f._initializeSettings=function(n,t){var e="undefined",r=n.get_settings(),u=OSF.OUtil.getSessionStorage(),i,f;return u&&(i=u.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey()),i?r=typeof JSON!==e?JSON.parse(i):OsfMsAjaxFactory.msAjaxSerializer.deserialize(i,!0):(i=typeof JSON!==e?JSON.stringify(r):OsfMsAjaxFactory.msAjaxSerializer.serialize(r),u.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(),i))),f=OSF.DDA.SettingsManager.deserializeSettings(r),t?new OSF.DDA.RefreshableSettings(f):new OSF.DDA.Settings(f)};e=function(n){var t=window.open;n.open=function(n,i,r){var u=null,f;try{u=t(n,i,r)}catch(e){OSF.AppTelemetry&&OSF.AppTelemetry.logAppCommonMessage("Exception happens at windowOpen."+e)}return u||(f={strUrl:n,strWindowName:i,strWindowFeatures:r},OSF._OfficeAppFactory.getClientEndPoint().invoke("ContextActivationManager_openWindowInHost",null,f)),u}};e(window)};OSF.InitializationHelper.prototype.saveAndSetDialogInfo=function(n){var r="IsDialog",u=function(){var n=OSF.OUtil.parseXdmInfo(!0),t;return n?(t=n.split("|"),t[1]):null},t=OSF.OUtil.getSessionStorage(),i;t&&(n&&n.indexOf("isDialog")>-1&&(i=u(),i!=null&&t.setItem(i+r,"true")),this._hostInfo.isDialog=t.getItem(OSF.OUtil.getXdmFieldValue(OSF.XdmFieldName.AppId,!1)+r)!=null?!0:!1)};OSF.InitializationHelper.prototype.getAppContext=function(n,t){var u=null,r=this,l=r,s=function(n,i){var r,u,h,f,e,o,s;if(i._appName===OSF.AppName.ExcelWebApp){u=i._settings;r={};for(h in u)f=u[h],r[f[0]]=f[1]}else r=i._settings;if(n===0&&i._id!=undefined&&i._appName!=undefined&&i._appVersion!=undefined&&i._appUILocale!=undefined&&i._dataLocale!=undefined&&i._docUrl!=undefined&&i._clientMode!=undefined&&i._settings!=undefined&&i._reason!=undefined){l._appContext=i;var a=i._appInstanceId?i._appInstanceId:i._id,v=!1,y=!0,c=0;i._appMinorVersion!=undefined&&(c=i._appMinorVersion);e=undefined;i._requirementMatrix!=undefined&&(e=i._requirementMatrix);o=new OSF.OfficeAppContext(i._id,i._appName,i._appVersion,i._appUILocale,i._dataLocale,i._docUrl,i._clientMode,r,i._reason,i._osfControlType,i._eToken,i._correlationId,a,v,y,c,e);OSF.AppTelemetry&&OSF.AppTelemetry.initialize(o);t(o)}else{s="Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is "+n+", exception: "+i;OSF.AppTelemetry&&OSF.AppTelemetry.logAppException(s);throw s;}},i,e;try{if(r._hostInfo.isDialog&&window.opener!=u){i=OfficeExt.WACUtils.parseAppContextWithGivenFragment(!0,OSF._OfficeAppFactory.getWindowLocationHash());i=i!=u?decodeURIComponent(i):u;var h=OSF.OUtil.getXdmFieldValue(OSF.XdmFieldName.AppId,!1),o=h+"AppContext",f=window.sessionStorage;!i&&f&&f.getItem(o)&&(i=f.getItem(o));i!=u&&(f.setItem(o,i),e=JSON.parse(i),e._correlationId=r._hostInfo.osfControlAppCorrelationId,e._appInstanceId=h,s(0,e))}else r._webAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync",s,r._webAppState.id)}catch(c){OSF.AppTelemetry&&OSF.AppTelemetry.logAppException("Exception thrown when trying to invoke getAppContextAsync. Exception:["+c+"]");throw c;}};OSF.InitializationHelper.prototype.setAgaveHostCommunication=function(){var u="ContextActivationManager_notifyHost",r=null,i=!1,n,f,t,e,o,s;try{n=this;f=OSF.OUtil.parseXdmInfoWithGivenFragment(i,OSF._OfficeAppFactory.getWindowLocationHash());f&&(t=OSF.OUtil.getInfoItems(f),t!=undefined&&t.length>=3&&(n._webAppState.conversationID=t[0],n._webAppState.id=t[1],n._webAppState.webAppUrl=t[2].indexOf(":")>=0?t[2]:decodeURIComponent(t[2])));n._webAppState.wnd=window.opener!=r?window.opener:window.parent;n._webAppState.serializerVersion=OSF.OUtil.parseSerializerVersionWithGivenFragment(i,OSF._OfficeAppFactory.getWindowLocationHash());n._webAppState.clientEndPoint=Microsoft.Office.Common.XdmCommunicationManager.connect(n._webAppState.conversationID,n._webAppState.wnd,n._webAppState.webAppUrl,n._webAppState.serializerVersion);n._webAppState.serviceEndPoint=Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(n._webAppState.id);e=n._webAppState.conversationID+OSF.SharedConstants.NotificationConversationIdSuffix;n._webAppState.serviceEndPoint.registerConversation(e,n._webAppState.webAppUrl);o=function(){var i,t,r,u;if(!n._webAppState.focused)for(n._webAppState.focused=!0,i=document.querySelectorAll("input,a,button"),t=0;t<i.length;t++)if(r=i[t],r instanceof HTMLElement){u=r;u.focus();break}};s=function(t){switch(t){case OSF.AgaveHostAction.Select:n._webAppState.focused=!0;break;case OSF.AgaveHostAction.UnSelect:n._webAppState.focused=i;break;case OSF.AgaveHostAction.CtrlF6In:o();default:OsfMsAjaxFactory.msAjaxDebug.trace("actionId "+t+" notifyAgave is wrong.")}};n._webAppState.serviceEndPoint.registerMethod("Office_notifyAgave",s,Microsoft.Office.Common.InvokeType.async,i);OSF.OUtil.addEventListener(window,"focus",function(){n._webAppState.focused||(n._webAppState.focused=!0);n._webAppState.clientEndPoint.invoke(u,r,[n._webAppState.id,OSF.AgaveHostAction.Select])});OSF.OUtil.addEventListener(window,"blur",function(){n._webAppState.focused&&(n._webAppState.focused=i);n._webAppState.clientEndPoint.invoke(u,r,[n._webAppState.id,OSF.AgaveHostAction.UnSelect])});OSF.OUtil.addEventListener(window,"keydown",function(t){if(t.keyCode==117&&t.ctrlKey){t.preventDefault?t.preventDefault():t.returnValue=i;var f=OSF.AgaveHostAction.CtrlF6Exit;t.shiftKey&&(f=OSF.AgaveHostAction.CtrlF6ExitShift);n._webAppState.clientEndPoint.invoke(u,r,[n._webAppState.id,f])}});OSF.OUtil.addEventListener(window,"keypress",function(n){n.keyCode==117&&n.ctrlKey&&(n.preventDefault?n.preventDefault():n.returnValue=i)})}catch(h){OSF.AppTelemetry&&OSF.AppTelemetry.logAppException("Exception thrown in setAgaveHostCommunication. Exception:["+h+"]");throw h;}};OSF.InitializationHelper.prototype.initWebDialog=function(n){n.get_isDialog()?OSF.DDA.UI.ChildUI&&(n.ui=new OSF.DDA.UI.ChildUI,window.opener!=null&&OfficeExt.AddinNativeAction.Dialog.registerMessageReceivedEvent()):OSF.DDA.UI.ParentUI&&(n.ui=new OSF.DDA.UI.ParentUI)};OSF.getClientEndPoint=function(){var n=OSF._OfficeAppFactory.getInitializationHelper();return n._webAppState.clientEndPoint},function(n){var u="ResponseTime",i="SessionId",r="CorrelationId",t=!0,f=function(){function n(n){this._table=n;this._fields={}}return Object.defineProperty(n.prototype,"Fields",{get:function(){return this._fields},enumerable:t,configurable:t}),Object.defineProperty(n.prototype,"Table",{get:function(){return this._table},enumerable:t,configurable:t}),n.prototype.SerializeFields=function(){},n.prototype.SetSerializedField=function(n,t){typeof t!="undefined"&&t!==null&&(this._serializedFields[n]=t.toString())},n.prototype.SerializeRow=function(){var n=this;return n._serializedFields={},n.SetSerializedField("Table",n._table),n.SerializeFields(),JSON.stringify(n._serializedFields)},n}(),e,o,s,h,c;n.BaseUsageData=f;e=function(n){function u(){n.call(this,"AppActivated")}var f="AppSizeHeight",e="AppSizeWidth",o="ClientId",s="HostVersion",h="Host",c="UserId",l="Browser",a="AssetId",v="AppURL",y="AppInstanceId",p="AppId";return __extends(u,n),Object.defineProperty(u.prototype,r,{get:function(){return this.Fields[r]},set:function(n){this.Fields[r]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,i,{get:function(){return this.Fields[i]},set:function(n){this.Fields[i]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,p,{get:function(){return this.Fields[p]},set:function(n){this.Fields[p]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,y,{get:function(){return this.Fields[y]},set:function(n){this.Fields[y]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,v,{get:function(){return this.Fields[v]},set:function(n){this.Fields[v]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,a,{get:function(){return this.Fields[a]},set:function(n){this.Fields[a]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,l,{get:function(){return this.Fields[l]},set:function(n){this.Fields[l]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,c,{get:function(){return this.Fields[c]},set:function(n){this.Fields[c]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,h,{get:function(){return this.Fields[h]},set:function(n){this.Fields[h]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,s,{get:function(){return this.Fields[s]},set:function(n){this.Fields[s]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,o,{get:function(){return this.Fields[o]},set:function(n){this.Fields[o]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,e,{get:function(){return this.Fields[e]},set:function(n){this.Fields[e]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,f,{get:function(){return this.Fields[f]},set:function(n){this.Fields[f]=n},enumerable:t,configurable:t}),u.prototype.SerializeFields=function(){var n=this;n.SetSerializedField(r,n.CorrelationId);n.SetSerializedField(i,n.SessionId);n.SetSerializedField(p,n.AppId);n.SetSerializedField(y,n.AppInstanceId);n.SetSerializedField(v,n.AppURL);n.SetSerializedField(a,n.AssetId);n.SetSerializedField(l,n.Browser);n.SetSerializedField(c,n.UserId);n.SetSerializedField(h,n.Host);n.SetSerializedField(s,n.HostVersion);n.SetSerializedField(o,n.ClientId);n.SetSerializedField(e,n.AppSizeWidth);n.SetSerializedField(f,n.AppSizeHeight)},u}(f);n.AppActivatedUsageData=e;o=function(n){function f(){n.call(this,"ScriptLoad")}var e="StartTime",o="ScriptId";return __extends(f,n),Object.defineProperty(f.prototype,r,{get:function(){return this.Fields[r]},set:function(n){this.Fields[r]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,i,{get:function(){return this.Fields[i]},set:function(n){this.Fields[i]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,o,{get:function(){return this.Fields[o]},set:function(n){this.Fields[o]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,e,{get:function(){return this.Fields[e]},set:function(n){this.Fields[e]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,u,{get:function(){return this.Fields[u]},set:function(n){this.Fields[u]=n},enumerable:t,configurable:t}),f.prototype.SerializeFields=function(){var n=this;n.SetSerializedField(r,n.CorrelationId);n.SetSerializedField(i,n.SessionId);n.SetSerializedField(o,n.ScriptId);n.SetSerializedField(e,n.StartTime);n.SetSerializedField(u,n.ResponseTime)},f}(f);n.ScriptLoadUsageData=o;s=function(n){function u(){n.call(this,"AppClosed")}var f="CloseMethod",e="OpenTime",o="AppSizeFinalHeight",s="AppSizeFinalWidth",h="FocusTime";return __extends(u,n),Object.defineProperty(u.prototype,r,{get:function(){return this.Fields[r]},set:function(n){this.Fields[r]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,i,{get:function(){return this.Fields[i]},set:function(n){this.Fields[i]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,h,{get:function(){return this.Fields[h]},set:function(n){this.Fields[h]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,s,{get:function(){return this.Fields[s]},set:function(n){this.Fields[s]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,o,{get:function(){return this.Fields[o]},set:function(n){this.Fields[o]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,e,{get:function(){return this.Fields[e]},set:function(n){this.Fields[e]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,f,{get:function(){return this.Fields[f]},set:function(n){this.Fields[f]=n},enumerable:t,configurable:t}),u.prototype.SerializeFields=function(){var n=this;n.SetSerializedField(r,n.CorrelationId);n.SetSerializedField(i,n.SessionId);n.SetSerializedField(h,n.FocusTime);n.SetSerializedField(s,n.AppSizeFinalWidth);n.SetSerializedField(o,n.AppSizeFinalHeight);n.SetSerializedField(e,n.OpenTime);n.SetSerializedField(f,n.CloseMethod)},u}(f);n.AppClosedUsageData=s;h=function(n){function f(){n.call(this,"APIUsage")}var e="ErrorType",o="Parameters",s="APIID",h="APIType";return __extends(f,n),Object.defineProperty(f.prototype,r,{get:function(){return this.Fields[r]},set:function(n){this.Fields[r]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,i,{get:function(){return this.Fields[i]},set:function(n){this.Fields[i]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,h,{get:function(){return this.Fields[h]},set:function(n){this.Fields[h]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,s,{get:function(){return this.Fields[s]},set:function(n){this.Fields[s]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,o,{get:function(){return this.Fields[o]},set:function(n){this.Fields[o]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,u,{get:function(){return this.Fields[u]},set:function(n){this.Fields[u]=n},enumerable:t,configurable:t}),Object.defineProperty(f.prototype,e,{get:function(){return this.Fields[e]},set:function(n){this.Fields[e]=n},enumerable:t,configurable:t}),f.prototype.SerializeFields=function(){var n=this;n.SetSerializedField(r,n.CorrelationId);n.SetSerializedField(i,n.SessionId);n.SetSerializedField(h,n.APIType);n.SetSerializedField(s,n.APIID);n.SetSerializedField(o,n.Parameters);n.SetSerializedField(u,n.ResponseTime);n.SetSerializedField(e,n.ErrorType)},f}(f);n.APIUsageUsageData=h;c=function(n){function u(){n.call(this,"AppInitialization")}var f="Message",e="SuccessCode";return __extends(u,n),Object.defineProperty(u.prototype,r,{get:function(){return this.Fields[r]},set:function(n){this.Fields[r]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,i,{get:function(){return this.Fields[i]},set:function(n){this.Fields[i]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,e,{get:function(){return this.Fields[e]},set:function(n){this.Fields[e]=n},enumerable:t,configurable:t}),Object.defineProperty(u.prototype,f,{get:function(){return this.Fields[f]},set:function(n){this.Fields[f]=n},enumerable:t,configurable:t}),u.prototype.SerializeFields=function(){var n=this;n.SetSerializedField(r,n.CorrelationId);n.SetSerializedField(i,n.SessionId);n.SetSerializedField(e,n.SuccessCode);n.SetSerializedField(f,n.Message)},u}(f);n.AppInitializationUsageData=c}(OSFLog||(OSFLog={})),function(n){"use strict";function u(){OSF.Logger&&OSF.Logger.ulsEndpoint&&OSF.Logger.ulsEndpoint.loadProxyFrame()}function f(n,t,i){if(OSF.Logger&&OSF.Logger.ulsEndpoint){var r={traceLevel:n,message:t,flag:i,internalLog:!0},u=JSON.stringify(r);OSF.Logger.ulsEndpoint.writeLog(u)}}function e(){try{return new t}catch(n){return null}}var i,r,t;(function(n){n[n.info=0]="info";n[n.warning=1]="warning";n[n.error=2]="error"})(n.TraceLevel||(n.TraceLevel={}));i=n.TraceLevel,function(n){n[n.none=0]="none";n[n.flush=1]="flush"}(n.SendFlag||(n.SendFlag={}));r=n.SendFlag;n.allowUploadingData=u;n.sendLog=f;t=function(){function n(){var n=this,t=n;n.proxyFrame=null;n.telemetryEndPoint="https://telemetryservice.firstpartyapps.oaspapps.com/telemetryservice/telemetryproxy.html";n.buffer=[];n.proxyFrameReady=!1;OSF.OUtil.addEventListener(window,"message",function(n){return t.tellProxyFrameReady(n)});setTimeout(function(){t.loadProxyFrame()},3e3)}return n.prototype.writeLog=function(t){var i=this;i.proxyFrameReady===!0?i.proxyFrame.contentWindow.postMessage(t,n.telemetryOrigin):i.buffer.length<128&&i.buffer.push(t)},n.prototype.loadProxyFrame=function(){var n=this;n.proxyFrame==null&&(n.proxyFrame=document.createElement("iframe"),n.proxyFrame.setAttribute("style","display:none"),n.proxyFrame.setAttribute("src",n.telemetryEndPoint),document.head.appendChild(n.proxyFrame))},n.prototype.tellProxyFrameReady=function(t){var i=this,e=i,r,u,f;if(t.data==="ProxyFrameReadyToLog"){for(i.proxyFrameReady=!0,r=0;r<i.buffer.length;r++)i.writeLog(i.buffer[r]);i.buffer.length=0;OSF.OUtil.removeEventListener(window,"message",function(n){return e.tellProxyFrameReady(n)})}else t.data==="ProxyFrameReadyToInit"&&(u={appName:"Office APPs",sessionId:OSF.OUtil.Guid.generateNewGuid()},f=JSON.stringify(u),i.proxyFrame.contentWindow.postMessage(f,n.telemetryOrigin))},n.telemetryOrigin="https://telemetryservice.firstpartyapps.oaspapps.com",n}();OSF.Logger||(OSF.Logger=n);n.ulsEndpoint=e()}(Logger||(Logger={})),function(n){function c(r){if(OSF.Logger&&!t){t=new h;t.hostVersion=r.get_appVersion();t.appId=r.get_id();t.host=r.get_appName();t.browser=window.navigator.userAgent;t.correlationId=r.get_correlationId();t.clientId=(new o).getClientId();t.appInstanceId=r.get_appInstanceId();t.appInstanceId&&(t.appInstanceId=t.appInstanceId.replace(/[{}]/g,"").toLowerCase());var f=location.href.indexOf("?");t.appURL=f==-1?location.href:location.href.substring(0,f),function(n,t){var u,f,r;t.assetId="";t.userId="";try{u=decodeURIComponent(n);f=new DOMParser;r=f.parseFromString(u,"text/xml");t.userId=r.getElementsByTagName("t")[0].attributes.getNamedItem("cid").nodeValue;t.assetId=r.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue}catch(e){}finally{u=i;r=i;f=i}}(r.get_eToken(),t),function(){var c=new Date,r=i,o=0,h=!1,f=function(){document.hasFocus()?r==i&&(r=new Date):r&&(o+=Math.abs((new Date).getTime()-r.getTime()),r=i)},t=[],s,e;for(t.push(new u("focus",f)),t.push(new u("blur",f)),t.push(new u("focusout",f)),t.push(new u("focusin",f)),s=function(){for(var u=0;u<t.length;u++)OSF.OUtil.removeEventListener(window,t[u].name,t[u].handler);if(t.length=0,!h){document.hasFocus()&&r&&(o+=Math.abs((new Date).getTime()-r.getTime()),r=i);n.onAppClosed(Math.abs((new Date).getTime()-c.getTime()),o);h=!0}},t.push(new u("beforeunload",s)),t.push(new u("unload",s)),e=0;e<t.length;e++)OSF.OUtil.addEventListener(window,t[e].name,t[e].handler);f()}();n.onAppActivated()}}function l(){if(t){(new o).enumerateLog(function(n,t){return(new f).LogRawData(t)},!0);var n=new OSFLog.AppActivatedUsageData;n.SessionId=r;n.AppId=t.appId;n.AssetId=t.assetId;n.AppURL=t.appURL;n.UserId=t.userId;n.ClientId=t.clientId;n.Browser=t.browser;n.Host=t.host;n.HostVersion=t.hostVersion;n.CorrelationId=t.correlationId;n.AppSizeWidth=window.innerWidth;n.AppSizeHeight=window.innerHeight;n.AppInstanceId=t.appInstanceId;(new f).LogData(n);setTimeout(function(){OSF.Logger&&OSF.Logger.allowUploadingData()},100)}}function a(n,t,i,u){var e=new OSFLog.ScriptLoadUsageData;e.CorrelationId=u;e.SessionId=r;e.ScriptId=n;e.StartTime=t;e.ResponseTime=i;(new f).LogData(e)}function v(n,i,u,o,s){if(t){var h=new OSFLog.APIUsageUsageData;h.CorrelationId=e;h.SessionId=r;h.APIType=n;h.APIID=i;h.Parameters=u;h.ResponseTime=o;h.ErrorType=s;(new f).LogData(h)}}function y(n,t,r,u){var f=i,e;if(t)if(typeof t=="number")f=String(t);else if(typeof t=="object")for(e in t)f!==i?f+=",":f="",typeof t[e]=="number"&&(f+=String(t[e]));else f="";OSF.AppTelemetry.onCallDone("method",n,f,r,u)}function p(n,t){OSF.AppTelemetry.onCallDone("property",-1,n,t)}function w(n,t){OSF.AppTelemetry.onCallDone("event",n,i,0,t)}function b(n,t,r,u){OSF.AppTelemetry.onCallDone(n?"registerevent":"unregisterevent",t,i,r,u)}function k(n,i){if(t){var u=new OSFLog.AppClosedUsageData;u.CorrelationId=e;u.SessionId=r;u.FocusTime=i;u.OpenTime=n;u.AppSizeFinalWidth=window.innerWidth;u.AppSizeFinalHeight=window.innerHeight;(new o).saveLog(r,u.SerializeRow())}}function d(n){e=n}function s(n,t){var i=new OSFLog.AppInitializationUsageData;i.CorrelationId=e;i.SessionId=r;i.SuccessCode=n?1:0;i.Message=t;(new f).LogData(i)}function g(n){s(!1,n)}function nt(n){s(!0,n)}var i=null;"use strict";var t,r=OSF.OUtil.Guid.generateNewGuid(),e="",h=function(){function n(){}return n}(),u=function(){function n(n,t){this.name=n;this.handler=t}return n}(),o=function(){function n(){this.clientIDKey="Office API client";this.logIdSetKey="Office App Log Id Set"}return n.prototype.getClientId=function(){var t=this,n=t.getValue(t.clientIDKey);return(!n||n.length<=0||n.length>40)&&(n=OSF.OUtil.Guid.generateNewGuid(),t.setValue(t.clientIDKey,n)),n},n.prototype.saveLog=function(n,t){var i=this,r=i.getValue(i.logIdSetKey);r=(r&&r.length>0?r+";":"")+n;i.setValue(i.logIdSetKey,r);i.setValue(n,t)},n.prototype.enumerateLog=function(n,t){var i=this,e=i.getValue(i.logIdSetKey),u,o,r,f;if(e){u=e.split(";");for(o in u)r=u[o],f=i.getValue(r),f&&(n&&n(r,f),t&&i.remove(r));t&&i.remove(i.logIdSetKey)}},n.prototype.getValue=function(n){var t=OSF.OUtil.getLocalStorage(),i="";return t&&(i=t.getItem(n)),i},n.prototype.setValue=function(n,t){var i=OSF.OUtil.getLocalStorage();i&&i.setItem(n,t)},n.prototype.remove=function(n){var t=OSF.OUtil.getLocalStorage();if(t)try{t.removeItem(n)}catch(i){}},n}(),f=function(){function n(){}return n.prototype.LogData=function(n){OSF.Logger&&OSF.Logger.sendLog(OSF.Logger.TraceLevel.info,n.SerializeRow(),OSF.Logger.SendFlag.none)},n.prototype.LogRawData=function(n){OSF.Logger&&OSF.Logger.sendLog(OSF.Logger.TraceLevel.info,n,OSF.Logger.SendFlag.none)},n}();n.initialize=c;n.onAppActivated=l;n.onScriptDone=a;n.onCallDone=v;n.onMethodDone=y;n.onPropertyDone=p;n.onEventDone=w;n.onRegisterDone=b;n.onAppClosed=k;n.setOsfControlAppCorrelationId=d;n.doAppInitializationLogging=s;n.logAppCommonMessage=g;n.logAppException=nt;OSF.AppTelemetry=n}(OSFAppTelemetry||(OSFAppTelemetry={}));Microsoft.Office.WebExtension.TableData=function(n,t){function i(n){if(n==null||n==undefined)return null;try{for(var t=OSF.DDA.DataCoercion.findArrayDimensionality(n,2);t<2;t++)n=[n];return n}catch(i){}}OSF.OUtil.defineEnumerableProperties(this,{headers:{get:function(){return t},set:function(n){t=i(n)}},rows:{get:function(){return n},set:function(t){n=t==null||OSF.OUtil.isArray(t)&&t.length==0?[]:i(t)}}});this.headers=t;this.rows=n};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureTableData=function(n){return new Microsoft.Office.WebExtension.TableData(n[OSF.DDA.TableDataProperties.TableRows],n[OSF.DDA.TableDataProperties.TableHeaders])};Microsoft.Office.WebExtension.CoercionType={Text:"text",Matrix:"matrix",Table:"table"};OSF.DDA.DataCoercion=function(){var n=null;return{findArrayDimensionality:function(n){if(OSF.OUtil.isArray(n)){for(var t=0,i=0;i<n.length;i++)t=Math.max(t,OSF.DDA.DataCoercion.findArrayDimensionality(n[i]));return t+1}else return 0},getCoercionDefaultForBinding:function(n){switch(n){case Microsoft.Office.WebExtension.BindingType.Matrix:return Microsoft.Office.WebExtension.CoercionType.Matrix;case Microsoft.Office.WebExtension.BindingType.Table:return Microsoft.Office.WebExtension.CoercionType.Table;case Microsoft.Office.WebExtension.BindingType.Text:default:return Microsoft.Office.WebExtension.CoercionType.Text}},getBindingDefaultForCoercion:function(n){switch(n){case Microsoft.Office.WebExtension.CoercionType.Matrix:return Microsoft.Office.WebExtension.BindingType.Matrix;case Microsoft.Office.WebExtension.CoercionType.Table:return Microsoft.Office.WebExtension.BindingType.Table;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:return Microsoft.Office.WebExtension.BindingType.Text}},determineCoercionType:function(t){if(t==n||t==undefined)return n;var i=n,r=typeof t;if(t.rows!==undefined)i=Microsoft.Office.WebExtension.CoercionType.Table;else if(OSF.OUtil.isArray(t))i=Microsoft.Office.WebExtension.CoercionType.Matrix;else if(r=="string"||r=="number"||r=="boolean"||OSF.OUtil.isDate(t))i=Microsoft.Office.WebExtension.CoercionType.Text;else throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;return i},coerceData:function(n,t,i){return i=i||OSF.DDA.DataCoercion.determineCoercionType(n),i&&i!=t&&(OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin),n=OSF.DDA.DataCoercion._coerceDataFromTable(t,OSF.DDA.DataCoercion._coerceDataToTable(n,i)),OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd)),n},_matrixToText:function(n){if(n.length==1&&n[0].length==1)return""+n[0][0];for(var t="",i=0;i<n.length;i++)t+=n[i].join("\t")+"\n";return t.substring(0,t.length-1)},_textToMatrix:function(n){for(var t=n.split("\n"),i=0;i<t.length;i++)t[i]=t[i].split("\t");return t},_tableToText:function(t){var i="",r;return t.headers!=n&&(i=OSF.DDA.DataCoercion._matrixToText([t.headers])+"\n"),r=OSF.DDA.DataCoercion._matrixToText(t.rows),r==""&&(i=i.substring(0,i.length-1)),i+r},_tableToMatrix:function(t){var i=t.rows;return t.headers!=n&&i.unshift(t.headers),i},_coerceDataFromTable:function(t,i){var r;switch(t){case Microsoft.Office.WebExtension.CoercionType.Table:r=i;break;case Microsoft.Office.WebExtension.CoercionType.Matrix:r=OSF.DDA.DataCoercion._tableToMatrix(i);break;case Microsoft.Office.WebExtension.CoercionType.SlideRange:r=n;OSF.DDA.OMFactory.manufactureSlideRange&&(r=OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(i)));r==n&&(r=OSF.DDA.DataCoercion._tableToText(i));break;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:r=OSF.DDA.DataCoercion._tableToText(i)}return r},_coerceDataToTable:function(n,t){t==undefined&&(t=OSF.DDA.DataCoercion.determineCoercionType(n));var i;switch(t){case Microsoft.Office.WebExtension.CoercionType.Table:i=n;break;case Microsoft.Office.WebExtension.CoercionType.Matrix:i=new Microsoft.Office.WebExtension.TableData(n);break;case Microsoft.Office.WebExtension.CoercionType.Text:case Microsoft.Office.WebExtension.CoercionType.Html:case Microsoft.Office.WebExtension.CoercionType.Ooxml:default:i=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(n))}return i}}}();OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{Image:"image"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType,{Ooxml:"ooxml"});Microsoft.Office.WebExtension.EventType={};OSF.EventDispatch=function(n){var t=this,r,i;t._eventHandlers={};t._queuedEventsArgs={};for(r in n)i=n[r],t._eventHandlers[i]=[],t._queuedEventsArgs[i]=[]};OSF.EventDispatch.prototype={getSupportedEvents:function(){var n=[];for(var t in this._eventHandlers)n.push(t);return n},supportsEvent:function(n){var t=!1;for(var i in this._eventHandlers)if(n==i){t=!0;break}return t},hasEventHandler:function(n,t){var i=this._eventHandlers[n],r;if(i&&i.length>0)for(r in i)if(i[r]===t)return!0;return!1},addEventHandler:function(n,t){if(typeof t!="function")return!1;var i=this._eventHandlers[n];return i&&!this.hasEventHandler(n,t)?(i.push(t),!0):!1},addEventHandlerAndFireQueuedEvent:function(n,t){var r=this._eventHandlers[n],u=r.length==0,i=this.addEventHandler(n,t);return u&&i&&this.fireQueuedEvent(n),i},removeEventHandler:function(n,t){var i=this._eventHandlers[n],r;if(i&&i.length>0)for(r=0;r<i.length;r++)if(i[r]===t)return i.splice(r,1),!0;return!1},clearEventHandlers:function(n){return typeof this._eventHandlers[n]!="undefined"&&this._eventHandlers[n].length>0?(this._eventHandlers[n]=[],!0):!1},getEventHandlerCount:function(n){return this._eventHandlers[n]!=undefined?this._eventHandlers[n].length:-1},fireEvent:function(n){var t,i,r;if(n.type==undefined)return!1;if(t=n.type,t&&this._eventHandlers[t]){i=this._eventHandlers[t];for(r in i)i[r](n);return!0}else return!1},fireOrQueueEvent:function(n){var t=this,i=n.type,r,u;return i&&t._eventHandlers[i]?(r=t._eventHandlers[i],u=t._queuedEventsArgs[i],r.length==0?u.push(n):t.fireEvent(n),!0):!1},fireQueuedEvent:function(n){var t,i,r,u;if(n&&this._eventHandlers[n]&&(t=this._eventHandlers[n],i=this._queuedEventsArgs[n],t.length>0)){for(r=t[0];i.length>0;)u=i.shift(),r(u);return!0}return!1}};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureEventArgs=function(n,t,i){var u=this,r;switch(n){case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:r=new OSF.DDA.DocumentSelectionChangedEventArgs(t);break;case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:r=new OSF.DDA.BindingSelectionChangedEventArgs(u.manufactureBinding(i,t.document),i[OSF.DDA.PropertyDescriptors.Subset]);break;case Microsoft.Office.WebExtension.EventType.BindingDataChanged:r=new OSF.DDA.BindingDataChangedEventArgs(u.manufactureBinding(i,t.document));break;case Microsoft.Office.WebExtension.EventType.SettingsChanged:r=new OSF.DDA.SettingsChangedEventArgs(t);break;case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:r=new OSF.DDA.ActiveViewChangedEventArgs(i);break;case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:r=new OSF.DDA.Theming.OfficeThemeChangedEventArgs(i);break;case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:r=new OSF.DDA.Theming.DocumentThemeChangedEventArgs(i);break;case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:r=OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(i);break;case Microsoft.Office.WebExtension.EventType.DataNodeInserted:r=new OSF.DDA.NodeInsertedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NewNode]),i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:r=new OSF.DDA.NodeReplacedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.OldNode]),u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NewNode]),i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:r=new OSF.DDA.NodeDeletedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.OldNode]),u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NextSiblingNode]),i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);break;case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:r=new OSF.DDA.TaskSelectionChangedEventArgs(t);break;case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:r=new OSF.DDA.ResourceSelectionChangedEventArgs(t);break;case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:r=new OSF.DDA.ViewSelectionChangedEventArgs(t);break;case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:r=new OSF.DDA.DialogEventArgs(i);break;default:throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,n));}return r};OSF.DDA.AsyncMethodNames.addNames({AddHandlerAsync:"addHandlerAsync",RemoveHandlerAsync:"removeHandlerAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddHandlerAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(n,t,i){return i.supportsEvent(n)}},{name:Microsoft.Office.WebExtension.Parameters.Handler,types:["function"]}],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(n,t,i){return i.supportsEvent(n)}}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Handler,value:{types:["function","object"],defaultValue:null}}],privateStateCallbacks:[]});OSF.DDA.DataPartProperties={Id:Microsoft.Office.WebExtension.Parameters.Id,BuiltIn:"DataPartBuiltIn"};OSF.DDA.DataNodeProperties={Handle:"DataNodeHandle",BaseName:"DataNodeBaseName",NamespaceUri:"DataNodeNamespaceUri",NodeType:"DataNodeType"};OSF.DDA.DataNodeEventProperties={OldNode:"OldNode",NewNode:"NewNode",NextSiblingNode:"NextSiblingNode",InUndoRedo:"InUndoRedo"};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{DataPartProperties:"DataPartProperties",DataNodeProperties:"DataNodeProperties"});OSF.OUtil.augmentList(OSF.DDA.ListDescriptors,{DataPartList:"DataPartList",DataNodeList:"DataNodeList"});OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataPartList,OSF.DDA.PropertyDescriptors.DataPartProperties);OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataNodeList,OSF.DDA.PropertyDescriptors.DataNodeProperties);OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{DataNodeInsertedEvent:"DataNodeInsertedEvent",DataNodeReplacedEvent:"DataNodeReplacedEvent",DataNodeDeletedEvent:"DataNodeDeletedEvent"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{DataNodeDeleted:"nodeDeleted",DataNodeInserted:"nodeInserted",DataNodeReplaced:"nodeReplaced"});OSF.DDA.CustomXmlParts=function(){this._eventDispatches=[];var n=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[n.AddDataPartAsync,n.GetDataPartByIdAsync,n.GetDataPartsByNameSpaceAsync])};OSF.DDA.CustomXmlPart=function(n,t,i){var u,e,r,f;OSF.OUtil.defineEnumerableProperties(this,{builtIn:{value:i},id:{value:t},namespaceManager:{value:new OSF.DDA.CustomXmlPrefixMappings(t)}});u=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[u.DeleteDataPartAsync,u.GetPartNodesAsync,u.GetPartXmlAsync]);e=n._eventDispatches;r=e[t];r||(f=Microsoft.Office.WebExtension.EventType,r=new OSF.EventDispatch([f.DataNodeDeleted,f.DataNodeInserted,f.DataNodeReplaced]),e[t]=r);OSF.DDA.DispIdHost.addEventSupport(this,r)};OSF.DDA.CustomXmlPrefixMappings=function(n){var t=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[t.AddDataPartNamespaceAsync,t.GetDataPartNamespaceAsync,t.GetDataPartPrefixAsync],n)};OSF.DDA.CustomXmlNode=function(n,t,i,r){OSF.OUtil.defineEnumerableProperties(this,{baseName:{value:r},namespaceUri:{value:i},nodeType:{value:t}});var u=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[u.GetRelativeNodesAsync,u.GetNodeValueAsync,u.GetNodeXmlAsync,u.SetNodeValueAsync,u.SetNodeXmlAsync,u.GetNodeTextAsync,u.SetNodeTextAsync],n)};OSF.DDA.NodeInsertedEventArgs=function(n,t){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeInserted},newNode:{value:n},inUndoRedo:{value:t}})};OSF.DDA.NodeReplacedEventArgs=function(n,t,i){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeReplaced},oldNode:{value:n},newNode:{value:t},inUndoRedo:{value:i}})};OSF.DDA.NodeDeletedEventArgs=function(n,t,i){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeDeleted},oldNode:{value:n},oldNextSibling:{value:t},inUndoRedo:{value:i}})};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureDataNode=function(n){if(n)return new OSF.DDA.CustomXmlNode(n[OSF.DDA.DataNodeProperties.Handle],n[OSF.DDA.DataNodeProperties.NodeType],n[OSF.DDA.DataNodeProperties.NamespaceUri],n[OSF.DDA.DataNodeProperties.BaseName])};OSF.DDA.OMFactory.manufactureDataPart=function(n,t){return new OSF.DDA.CustomXmlPart(t,n[OSF.DDA.DataPartProperties.Id],n[OSF.DDA.DataPartProperties.BuiltIn])};OSF.DDA.AsyncMethodNames.addNames({AddDataPartAsync:"addAsync",GetDataPartByIdAsync:"getByIdAsync",GetDataPartsByNameSpaceAsync:"getByNamespaceAsync",DeleteDataPartAsync:"deleteAsync",GetPartNodesAsync:"getNodesAsync",GetPartXmlAsync:"getXmlAsync",AddDataPartNamespaceAsync:"addNamespaceAsync",GetDataPartNamespaceAsync:"getNamespaceAsync",GetDataPartPrefixAsync:"getPrefixAsync",GetRelativeNodesAsync:"getNodesAsync",GetNodeValueAsync:"getNodeValueAsync",GetNodeXmlAsync:"getXmlAsync",SetNodeValueAsync:"setNodeValueAsync",SetNodeXmlAsync:"setXmlAsync",GetNodeTextAsync:"getTextAsync",SetNodeTextAsync:"setTextAsync"}),function(){function r(n){return OSF.DDA.OMFactory.manufactureDataPart(n,Microsoft.Office.WebExtension.context.document.customXmlParts)}function e(n){return OSF.DDA.OMFactory.manufactureDataNode(n)}function i(n){var t=n[Microsoft.Office.WebExtension.Parameters.Data];return t==undefined?null:t}function u(n){return n.id}function f(n,t){return t}function t(n,t){return t}var n="string";OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddDataPartAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Xml,types:[n]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:r});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataPartByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[n]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:r});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataPartsByNameSpaceAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Namespace,types:[n]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(n){return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataPartList],r)}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.DeleteDataPartAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:u}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetPartNodesAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.XPath,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:u}],onSucceeded:function(n){return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataNodeList],e)}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetPartXmlAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:u}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddDataPartNamespaceAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Prefix,types:[n]},{name:Microsoft.Office.WebExtension.Parameters.Namespace,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:f}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataPartNamespaceAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Prefix,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:f}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataPartPrefixAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Namespace,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataPartProperties.Id,value:f}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetRelativeNodesAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.XPath,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}],onSucceeded:function(n){return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataNodeList],e)}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetNodeValueAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetNodeXmlAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetNodeValueAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetNodeXmlAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Xml,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetNodeTextAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}],onSucceeded:i});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetNodeTextAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Text,types:[n]}],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.DataNodeProperties.Handle,value:t}]})}();OSF.OUtil.setNamespace("Marshaling",OSF.DDA);OSF.DDA.Marshaling.CustomXmlPartsKeys={Id:"id",Namespace:"namespace",Xml:"xml",XPath:"xpath",Prefix:"prefix"};OSF.DDA.Marshaling.DataPartProperties={Id:"id",BuiltIn:"DataPartBuiltIn"};OSF.DDA.Marshaling.PropertyDescriptors={DataPartProperties:"dataPartProperties",DataNodeProperties:"dataNodeProperties"};OSF.DDA.Marshaling.DataNodeProperties={Handle:"DataNodeHandle",BaseName:"DataNodeBaseName",NamespaceUri:"DataNodeNamespaceUri",NodeType:"DataNodeType"};OSF.DDA.Marshaling.ListDescriptors={DataPartList:"DataPartList",DataNodeList:"DataNodeList"};OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.DataPartProperties,fromHost:[{name:OSF.DDA.DataPartProperties.Id,value:OSF.DDA.Marshaling.DataPartProperties.Id},{name:OSF.DDA.DataPartProperties.BuiltIn,value:OSF.DDA.Marshaling.DataPartProperties.BuiltIn}],isComplexType:!0});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.DataNodeProperties,fromHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle},{name:OSF.DDA.DataNodeProperties.BaseName,value:OSF.DDA.Marshaling.DataNodeProperties.BaseName},{name:OSF.DDA.DataNodeProperties.NamespaceUri,value:OSF.DDA.Marshaling.DataNodeProperties.NamespaceUri},{name:OSF.DDA.DataNodeProperties.NodeType,value:OSF.DDA.Marshaling.DataNodeProperties.NodeType}],isComplexType:!0});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddDataPartMethod,fromHost:[{name:OSF.DDA.PropertyDescriptors.DataPartProperties,value:OSF.DDA.Marshaling.PropertyDescriptors.DataPartProperties}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Xml,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Xml}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataPartByIdMethod,fromHost:[{name:OSF.DDA.PropertyDescriptors.DataPartProperties,value:OSF.DDA.Marshaling.PropertyDescriptors.DataPartProperties}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Id}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataPartsByNamespaceMethod,fromHost:[{name:OSF.DDA.ListDescriptors.DataPartList,value:OSF.DDA.Marshaling.ListDescriptors.DataPartList}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Namespace,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Namespace}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataPartXmlMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Id}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataPartNodesMethod,fromHost:[{name:OSF.DDA.ListDescriptors.DataNodeList,value:OSF.DDA.Marshaling.ListDescriptors.DataNodeList}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Id},{name:Microsoft.Office.WebExtension.Parameters.XPath,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.XPath}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidDeleteDataPartMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Id}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataNodeValueMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataNodeXmlMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataNodesMethod,fromHost:[{name:OSF.DDA.ListDescriptors.DataNodeList,value:OSF.DDA.Marshaling.ListDescriptors.DataNodeList}],toHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle},{name:Microsoft.Office.WebExtension.Parameters.XPath,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.XPath}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSetDataNodeValueMethod,toHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle},{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSetDataNodeXmlMethod,toHost:[{name:OSF.DDA.DataNodeProperties.Handle,value:OSF.DDA.Marshaling.DataNodeProperties.Handle},{name:Microsoft.Office.WebExtension.Parameters.Xml,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Xml}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddDataNamespaceMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.DataPartProperties.Id},{name:Microsoft.Office.WebExtension.Parameters.Prefix,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Prefix},{name:Microsoft.Office.WebExtension.Parameters.Namespace,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Namespace}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataUriByPrefixMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.DataPartProperties.Id},{name:Microsoft.Office.WebExtension.Parameters.Prefix,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Prefix}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDataPrefixByUriMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.DataPartProperties.Id},{name:Microsoft.Office.WebExtension.Parameters.Namespace,value:OSF.DDA.Marshaling.CustomXmlPartsKeys.Namespace}]});OSF.DDA.AsyncMethodNames.addNames({GetSelectedDataAsync:"getSelectedDataAsync",SetSelectedDataAsync:"setSelectedDataAsync"}),function(){function r(n,t,i){var r=n[Microsoft.Office.WebExtension.Parameters.Data];return OSF.DDA.TableDataProperties&&r&&(r[OSF.DDA.TableDataProperties.TableRows]!=undefined||r[OSF.DDA.TableDataProperties.TableHeaders]!=undefined)&&(r=OSF.DDA.OMFactory.manufactureTableData(r)),r=OSF.DDA.DataCoercion.coerceData(r,i[Microsoft.Office.WebExtension.Parameters.CoercionType]),r==undefined?null:r}var i=!1,n="boolean",t="number";OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,"enum":Microsoft.Office.WebExtension.CoercionType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:{"enum":Microsoft.Office.WebExtension.ValueFormat,defaultValue:Microsoft.Office.WebExtension.ValueFormat.Unformatted}},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:{"enum":Microsoft.Office.WebExtension.FilterType,defaultValue:Microsoft.Office.WebExtension.FilterType.All}}],privateStateCallbacks:[],onSucceeded:r});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["string","object",t,n]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(n){return OSF.DDA.DataCoercion.determineCoercionType(n[Microsoft.Office.WebExtension.Parameters.Data])}}},{name:Microsoft.Office.WebExtension.Parameters.ImageLeft,value:{types:[t,n],defaultValue:i}},{name:Microsoft.Office.WebExtension.Parameters.ImageTop,value:{types:[t,n],defaultValue:i}},{name:Microsoft.Office.WebExtension.Parameters.ImageWidth,value:{types:[t,n],defaultValue:i}},{name:Microsoft.Office.WebExtension.Parameters.ImageHeight,value:{types:[t,n],defaultValue:i}}],privateStateCallbacks:[]})}();OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.GetData,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:"BindingId"},{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:"CoerceType"},{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:"ValueFormat"},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:"FilterType"},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:"Rows"},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:"Columns"},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:"StartRow"},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:"StartCol"},{name:Microsoft.Office.WebExtension.Parameters.RowCount,value:"RowCount"},{name:Microsoft.Office.WebExtension.Parameters.ColumnCount,value:"ColCount"}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.SetData,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:"BindingId"},{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:"CoerceType"},{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:"Rows"},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:"Columns"},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:"StartRow"},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:"StartCol"},{name:Microsoft.Office.WebExtension.Parameters.ImageLeft,value:"ImageLeft"},{name:Microsoft.Office.WebExtension.Parameters.ImageTop,value:"ImageTop"},{name:Microsoft.Office.WebExtension.Parameters.ImageWidth,value:"ImageWidth"},{name:Microsoft.Office.WebExtension.Parameters.ImageHeight,value:"ImageHeight"}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:OSF.DDA.WAC.UniqueArguments.GetData,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.SetData,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.SettingsManager={SerializedSettings:"serializedSettings",RefreshingSettings:"refreshingSettings",DateJSONPrefix:"Date(",DataJSONSuffix:")",serializeSettings:function(n){var r={},i,t;for(i in n){t=n[i];try{t=JSON?JSON.stringify(t,function(n,t){return OSF.OUtil.isDate(this[n])?OSF.DDA.SettingsManager.DateJSONPrefix+this[n].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix:t}):Sys.Serialization.JavaScriptSerializer.serialize(t);r[i]=t}catch(u){}}return r},deserializeSettings:function(n){var r={},i,t;n=n||{};for(i in n){t=n[i];try{t=JSON?JSON.parse(t,function(n,t){var i;return typeof t=="string"&&t&&t.length>6&&t.slice(0,5)===OSF.DDA.SettingsManager.DateJSONPrefix&&t.slice(-1)===OSF.DDA.SettingsManager.DataJSONSuffix&&(i=new Date(parseInt(t.slice(5,-1))),i)?i:t}):Sys.Serialization.JavaScriptSerializer.deserialize(t,!0);r[i]=t}catch(u){}}return r}};OSF.DDA.Settings=function(n){var t="name",i;n=n||{};i=function(n){var i=OSF.OUtil.getSessionStorage(),t,r;i&&(t=OSF.DDA.SettingsManager.serializeSettings(n),r=JSON?JSON.stringify(t):Sys.Serialization.JavaScriptSerializer.serialize(t),i.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(),r))};OSF.OUtil.defineEnumerableProperties(this,{get:{value:function(i){var u=Function._validateParams(arguments,[{name:t,type:String,mayBeNull:!1}]),r;if(u)throw u;return r=n[i],typeof r=="undefined"?null:r}},set:{value:function(r,u){var f=Function._validateParams(arguments,[{name:t,type:String,mayBeNull:!1},{name:"value",mayBeNull:!0}]);if(f)throw f;n[r]=u;i(n)}},remove:{value:function(r){var u=Function._validateParams(arguments,[{name:t,type:String,mayBeNull:!1}]);if(u)throw u;delete n[r];i(n)}}});OSF.DDA.DispIdHost.addAsyncMethods(this,[OSF.DDA.AsyncMethodNames.SaveAsync],n)};OSF.DDA.RefreshableSettings=function(n){OSF.DDA.RefreshableSettings.uber.constructor.call(this,n);OSF.DDA.DispIdHost.addAsyncMethods(this,[OSF.DDA.AsyncMethodNames.RefreshAsync],n);OSF.DDA.DispIdHost.addEventSupport(this,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]))};OSF.OUtil.extend(OSF.DDA.RefreshableSettings,OSF.DDA.Settings);OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{SettingsChanged:"settingsChanged"});OSF.DDA.SettingsChangedEventArgs=function(n){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.SettingsChanged},settings:{value:n}})};OSF.DDA.AsyncMethodNames.addNames({RefreshAsync:"refreshAsync",SaveAsync:"saveAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.RefreshAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:OSF.DDA.SettingsManager.RefreshingSettings,value:function(n,t){return t}}],onSucceeded:function(n,t,i){var f=n[OSF.DDA.SettingsManager.SerializedSettings],u=OSF.DDA.SettingsManager.deserializeSettings(f),e=i[OSF.DDA.SettingsManager.RefreshingSettings];for(var r in e)t.remove(r);for(r in u)t.set(r,u[r]);return t}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SaveAsync,requiredArguments:[],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,value:{types:["boolean"],defaultValue:!0}}],privateStateCallbacks:[{name:OSF.DDA.SettingsManager.SerializedSettings,value:function(n,t){return OSF.DDA.SettingsManager.serializeSettings(t)}}]});OSF.DDA.WAC.SettingsTranslator=function(){var n=0,t=1;return{read:function(i){var u={},f=i.Settings,e,r;for(e in f)r=f[e],u[r[n]]=r[t];return u},write:function(i){var f=[],u,r;for(u in i)r=[],r[n]=u,r[t]=i[u],f.push(r);return f}}}();OSF.DDA.WAC.Delegate.ParameterMap.setDynamicType(OSF.DDA.SettingsManager.SerializedSettings,{toHost:OSF.DDA.WAC.SettingsTranslator.write,fromHost:OSF.DDA.WAC.SettingsTranslator.read});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.SettingsRequest,toHost:[{name:Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,value:"OverwriteIfStale"},{name:OSF.DDA.SettingsManager.SerializedSettings,value:OSF.DDA.WAC.UniqueArguments.Properties}],invertible:!0});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidLoadSettingsMethod,fromHost:[{name:OSF.DDA.SettingsManager.SerializedSettings,value:OSF.DDA.WAC.UniqueArguments.Properties}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSaveSettingsMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.SettingsRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidSettingsChangedEvent});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{DocumentSelectionChanged:"documentSelectionChanged"});OSF.DDA.DocumentSelectionChangedEventArgs=function(n){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged},document:{value:n}})};OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent});OSF.DDA.FilePropertiesDescriptor={Url:"Url"};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{FilePropertiesDescriptor:"FilePropertiesDescriptor"});Microsoft.Office.WebExtension.FileProperties=function(n){OSF.OUtil.defineEnumerableProperties(this,{url:{value:n[OSF.DDA.FilePropertiesDescriptor.Url]}})};OSF.DDA.AsyncMethodNames.addNames({GetFilePropertiesAsync:"getFilePropertiesAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,fromHost:[{name:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,value:0}],requiredArguments:[],supportedOptions:[],onSucceeded:function(n){return new Microsoft.Office.WebExtension.FileProperties(n)}});OSF.OUtil.setNamespace("Marshaling",OSF.DDA),function(n){var t="Properties";n[n[t]=0]=t;n[n.Url=1]="Url"}(OSF_DDA_Marshaling_FilePropertiesKeys||(OSF_DDA_Marshaling_FilePropertiesKeys={}));OSF.DDA.Marshaling.FilePropertiesKeys=OSF_DDA_Marshaling_FilePropertiesKeys;OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,fromHost:[{name:OSF.DDA.FilePropertiesDescriptor.Url,value:OSF.DDA.Marshaling.FilePropertiesKeys.Url}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,fromHost:[{name:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,value:OSF.DDA.Marshaling.FilePropertiesKeys.Properties}]});Microsoft.Office.WebExtension.FileType={Text:"text",Compressed:"compressed",Pdf:"pdf"};OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{FileProperties:"FileProperties",FileSliceProperties:"FileSliceProperties"});OSF.DDA.FileProperties={Handle:"FileHandle",FileSize:"FileSize",SliceSize:Microsoft.Office.WebExtension.Parameters.SliceSize};OSF.DDA.File=function(n,t,i){var r,u;OSF.OUtil.defineEnumerableProperties(this,{size:{value:t},sliceCount:{value:Math.ceil(t/i)}});r={};r[OSF.DDA.FileProperties.Handle]=n;r[OSF.DDA.FileProperties.SliceSize]=i;u=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[u.GetDocumentCopyChunkAsync,u.ReleaseDocumentCopyAsync],r)};OSF.DDA.FileSliceOffset="fileSliceoffset";OSF.DDA.AsyncMethodNames.addNames({GetDocumentCopyAsync:"getFileAsync",GetDocumentCopyChunkAsync:"getSliceAsync",ReleaseDocumentCopyAsync:"closeAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.FileType,"enum":Microsoft.Office.WebExtension.FileType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.SliceSize,value:{types:["number"],defaultValue:4194304}}],checkCallArgs:function(n){var t=n[Microsoft.Office.WebExtension.Parameters.SliceSize];if(t<=0||t>4194304)throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;return n},onSucceeded:function(n,t,i){return new OSF.DDA.File(n[OSF.DDA.FileProperties.Handle],n[OSF.DDA.FileProperties.FileSize],i[Microsoft.Office.WebExtension.Parameters.SliceSize])}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.SliceIndex,types:["number"]}],privateStateCallbacks:[{name:OSF.DDA.FileProperties.Handle,value:function(n,t){return t[OSF.DDA.FileProperties.Handle]}},{name:OSF.DDA.FileProperties.SliceSize,value:function(n,t){return t[OSF.DDA.FileProperties.SliceSize]}}],checkCallArgs:function(n,t,i){var r=n[Microsoft.Office.WebExtension.Parameters.SliceIndex];if(r<0||r>=t.sliceCount)throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;return n[OSF.DDA.FileSliceOffset]=parseInt((r*i[OSF.DDA.FileProperties.SliceSize]).toString()),n},onSucceeded:function(n,t,i){var r={};return OSF.OUtil.defineEnumerableProperties(r,{data:{value:n[Microsoft.Office.WebExtension.Parameters.Data]},index:{value:i[Microsoft.Office.WebExtension.Parameters.SliceIndex]},size:{value:n[OSF.DDA.FileProperties.SliceSize]}}),r}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,privateStateCallbacks:[{name:OSF.DDA.FileProperties.Handle,value:function(n,t){return t[OSF.DDA.FileProperties.Handle]}}]});OSF.OUtil.setNamespace("Marshaling",OSF.DDA);OSF.OUtil.setNamespace("File",OSF.DDA.Marshaling),function(n){var t="FileSize";n[n.Handle=0]="Handle";n[n[t]=1]=t}(OSF_DDA_Marshaling_File_FilePropertiesKeys||(OSF_DDA_Marshaling_File_FilePropertiesKeys={}));OSF.DDA.Marshaling.File.FilePropertiesKeys=OSF_DDA_Marshaling_File_FilePropertiesKeys,function(n){var t="SliceSize";n[n.Data=0]="Data";n[n[t]=1]=t}(OSF_DDA_Marshaling_File_SlicePropertiesKeys||(OSF_DDA_Marshaling_File_SlicePropertiesKeys={}));OSF.DDA.Marshaling.File.SlicePropertiesKeys=OSF_DDA_Marshaling_File_SlicePropertiesKeys,function(n){var t="Compressed";n[n.Text=0]="Text";n[n[t]=1]=t;n[n.Pdf=2]="Pdf"}(OSF_DDA_Marshaling_File_FileType||(OSF_DDA_Marshaling_File_FileType={}));OSF.DDA.Marshaling.File.FileType=OSF_DDA_Marshaling_File_FileType,function(n){var t="SliceIndex",i="SliceSize",r="FileType";n[n[r]=0]=r;n[n[i]=1]=i;n[n.Handle=2]="Handle";n[n[t]=3]=t}(OSF_DDA_Marshaling_File_ParameterKeys||(OSF_DDA_Marshaling_File_ParameterKeys={}));OSF.DDA.Marshaling.File.ParameterKeys=OSF_DDA_Marshaling_File_ParameterKeys;OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.PropertyDescriptors.FileProperties);OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.PropertyDescriptors.FileSliceProperties);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FileProperties,fromHost:[{name:OSF.DDA.FileProperties.Handle,value:OSF.DDA.Marshaling.File.FilePropertiesKeys.Handle},{name:OSF.DDA.FileProperties.FileSize,value:OSF.DDA.Marshaling.File.FilePropertiesKeys.FileSize}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.FileSliceProperties,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.Marshaling.File.SlicePropertiesKeys.Data},{name:OSF.DDA.FileProperties.SliceSize,value:OSF.DDA.Marshaling.File.SlicePropertiesKeys.SliceSize}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.FileType,toHost:[{name:Microsoft.Office.WebExtension.FileType.Text,value:OSF.DDA.Marshaling.File.FileType.Text},{name:Microsoft.Office.WebExtension.FileType.Compressed,value:OSF.DDA.Marshaling.File.FileType.Compressed},{name:Microsoft.Office.WebExtension.FileType.Pdf,value:OSF.DDA.Marshaling.File.FileType.Pdf}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.FileType,value:OSF.DDA.Marshaling.File.ParameterKeys.FileType},{name:Microsoft.Office.WebExtension.Parameters.SliceSize,value:OSF.DDA.Marshaling.File.ParameterKeys.SliceSize}],fromHost:[{name:OSF.DDA.PropertyDescriptors.FileProperties,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,toHost:[{name:OSF.DDA.FileProperties.Handle,value:OSF.DDA.Marshaling.File.ParameterKeys.Handle},{name:Microsoft.Office.WebExtension.Parameters.SliceIndex,value:OSF.DDA.Marshaling.File.ParameterKeys.SliceIndex}],fromHost:[{name:OSF.DDA.PropertyDescriptors.FileSliceProperties,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,toHost:[{name:OSF.DDA.FileProperties.Handle,value:OSF.DDA.Marshaling.File.ParameterKeys.Handle}]});OSF.DDA.AsyncMethodNames.addNames({ExecuteRichApiRequestAsync:"executeRichApiRequestAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[]});OSF.OUtil.setNamespace("RichApi",OSF.DDA);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.ArrayData}],fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}]});Microsoft.Office.WebExtension.BindingType={Table:"table",Text:"text",Matrix:"matrix"};OSF.DDA.BindingProperties={Id:"BindingId",Type:Microsoft.Office.WebExtension.Parameters.BindingType};OSF.OUtil.augmentList(OSF.DDA.ListDescriptors,{BindingList:"BindingList"});OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{Subset:"subset",BindingProperties:"BindingProperties"});OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList,OSF.DDA.PropertyDescriptors.BindingProperties);OSF.DDA.BindingPromise=function(n,t){this._id=n;OSF.OUtil.defineEnumerableProperty(this,"onFail",{get:function(){return t},set:function(n){var i=typeof n;if(i!="undefined"&&i!="function")throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction,i);t=n}})};OSF.DDA.BindingPromise.prototype={_fetch:function(n){var t=this,i;return t.binding?n&&n(t.binding):t._binding||(i=t,Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(t._id,function(t){t.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded?(OSF.OUtil.defineEnumerableProperty(i,"binding",{value:t.value}),n&&n(i.binding)):i.onFail&&i.onFail(t)})),t},getDataAsync:function(){var n=arguments;return this._fetch(function(t){t.getDataAsync.apply(t,n)}),this},setDataAsync:function(){var n=arguments;return this._fetch(function(t){t.setDataAsync.apply(t,n)}),this},addHandlerAsync:function(){var n=arguments;return this._fetch(function(t){t.addHandlerAsync.apply(t,n)}),this},removeHandlerAsync:function(){var n=arguments;return this._fetch(function(t){t.removeHandlerAsync.apply(t,n)}),this}};OSF.DDA.BindingFacade=function(n){this._eventDispatches=[];OSF.OUtil.defineEnumerableProperty(this,"document",{value:n});var t=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[t.AddFromSelectionAsync,t.AddFromNamedItemAsync,t.GetAllAsync,t.GetByIdAsync,t.ReleaseByIdAsync])};OSF.DDA.UnknownBinding=function(n,t){OSF.OUtil.defineEnumerableProperties(this,{document:{value:t},id:{value:n}})};OSF.DDA.Binding=function(n,t){var r,u,i,f;OSF.OUtil.defineEnumerableProperties(this,{document:{value:t},id:{value:n}});r=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[r.GetDataAsync,r.SetDataAsync]);u=Microsoft.Office.WebExtension.EventType;i=t.bindings._eventDispatches;i[n]||(i[n]=new OSF.EventDispatch([u.BindingSelectionChanged,u.BindingDataChanged]));f=i[n];OSF.DDA.DispIdHost.addEventSupport(this,f)};OSF.DDA.generateBindingId=function(){return"UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+(new Date).getTime()};OSF.DDA.OMFactory=OSF.DDA.OMFactory||{};OSF.DDA.OMFactory.manufactureBinding=function(n,t){var r=n[OSF.DDA.BindingProperties.Id],u=n[OSF.DDA.BindingProperties.RowCount],f=n[OSF.DDA.BindingProperties.ColumnCount],s=n[OSF.DDA.BindingProperties.HasHeaders],i,e,o;switch(n[OSF.DDA.BindingProperties.Type]){case Microsoft.Office.WebExtension.BindingType.Text:i=new OSF.DDA.TextBinding(r,t);break;case Microsoft.Office.WebExtension.BindingType.Matrix:i=new OSF.DDA.MatrixBinding(r,t,u,f);break;case Microsoft.Office.WebExtension.BindingType.Table:e=function(){return OSF.DDA.ExcelDocument&&Microsoft.Office.WebExtension.context.document&&Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument};o=e()&&OSF.DDA.ExcelTableBinding?OSF.DDA.ExcelTableBinding:OSF.DDA.TableBinding;i=new o(r,t,u,f,s);break;default:i=new OSF.DDA.UnknownBinding(r,t)}return i};OSF.DDA.AsyncMethodNames.addNames({AddFromSelectionAsync:"addFromSelectionAsync",AddFromNamedItemAsync:"addFromNamedItemAsync",GetAllAsync:"getAllAsync",GetByIdAsync:"getByIdAsync",ReleaseByIdAsync:"releaseByIdAsync",GetDataAsync:"getDataAsync",SetDataAsync:"setDataAsync"}),function(){function u(n){return OSF.DDA.OMFactory.manufactureBinding(n,Microsoft.Office.WebExtension.context.document)}function f(n){return n.id}function e(n,t,i){var u=n[Microsoft.Office.WebExtension.Parameters.Data];return OSF.DDA.TableDataProperties&&u&&(u[OSF.DDA.TableDataProperties.TableRows]!=undefined||u[OSF.DDA.TableDataProperties.TableHeaders]!=undefined)&&(u=OSF.DDA.OMFactory.manufactureTableData(u)),u=OSF.DDA.DataCoercion.coerceData(u,i[Microsoft.Office.WebExtension.Parameters.CoercionType]),u==undefined?r:u}var t="number",i="object",n="string",r=null;OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.BindingType,"enum":Microsoft.Office.WebExtension.BindingType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:{types:[n],calculate:OSF.DDA.generateBindingId}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[i],defaultValue:r}}],privateStateCallbacks:[],onSucceeded:u});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.ItemName,types:[n]},{name:Microsoft.Office.WebExtension.Parameters.BindingType,"enum":Microsoft.Office.WebExtension.BindingType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:{types:[n],calculate:OSF.DDA.generateBindingId}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[i],defaultValue:r}}],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.FailOnCollision,value:function(){return!0}}],onSucceeded:u});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetAllAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(n){return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.BindingList],u)}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[n]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:u});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[n]}],supportedOptions:[],privateStateCallbacks:[],onSucceeded:function(n,t,i){var r=i[Microsoft.Office.WebExtension.Parameters.Id];delete t._eventDispatches[r]}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GetDataAsync,requiredArguments:[],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(n,t){return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type)}}},{name:Microsoft.Office.WebExtension.Parameters.ValueFormat,value:{"enum":Microsoft.Office.WebExtension.ValueFormat,defaultValue:Microsoft.Office.WebExtension.ValueFormat.Unformatted}},{name:Microsoft.Office.WebExtension.Parameters.FilterType,value:{"enum":Microsoft.Office.WebExtension.FilterType,defaultValue:Microsoft.Office.WebExtension.FilterType.All}},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:{types:[i,n],defaultValue:r}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[i],defaultValue:r}},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:{types:[t],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:{types:[t],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.RowCount,value:{types:[t],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.ColumnCount,value:{types:[t],defaultValue:0}}],checkCallArgs:function(n,t){if(n[Microsoft.Office.WebExtension.Parameters.StartRow]==0&&n[Microsoft.Office.WebExtension.Parameters.StartColumn]==0&&n[Microsoft.Office.WebExtension.Parameters.RowCount]==0&&n[Microsoft.Office.WebExtension.Parameters.ColumnCount]==0&&(delete n[Microsoft.Office.WebExtension.Parameters.StartRow],delete n[Microsoft.Office.WebExtension.Parameters.StartColumn],delete n[Microsoft.Office.WebExtension.Parameters.RowCount],delete n[Microsoft.Office.WebExtension.Parameters.ColumnCount]),n[Microsoft.Office.WebExtension.Parameters.CoercionType]!=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type)&&(n[Microsoft.Office.WebExtension.Parameters.StartRow]||n[Microsoft.Office.WebExtension.Parameters.StartColumn]||n[Microsoft.Office.WebExtension.Parameters.RowCount]||n[Microsoft.Office.WebExtension.Parameters.ColumnCount]))throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;return n},privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:f}],onSucceeded:e});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.SetDataAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:[n,i,t,"boolean"]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.CoercionType,value:{"enum":Microsoft.Office.WebExtension.CoercionType,calculate:function(n){return OSF.DDA.DataCoercion.determineCoercionType(n[Microsoft.Office.WebExtension.Parameters.Data])}}},{name:Microsoft.Office.WebExtension.Parameters.Rows,value:{types:[i,n],defaultValue:r}},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:{types:[i],defaultValue:r}},{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:{types:[t],defaultValue:0}},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:{types:[t],defaultValue:0}}],checkCallArgs:function(n,t){if(n[Microsoft.Office.WebExtension.Parameters.StartRow]==0&&n[Microsoft.Office.WebExtension.Parameters.StartColumn]==0&&(delete n[Microsoft.Office.WebExtension.Parameters.StartRow],delete n[Microsoft.Office.WebExtension.Parameters.StartColumn]),n[Microsoft.Office.WebExtension.Parameters.CoercionType]!=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type)&&(n[Microsoft.Office.WebExtension.Parameters.StartRow]||n[Microsoft.Office.WebExtension.Parameters.StartColumn]))throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;return n},privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:f}]})}();OSF.DDA.TextBinding=function(n,t){OSF.DDA.TextBinding.uber.constructor.call(this,n,t);OSF.OUtil.defineEnumerableProperty(this,"type",{value:Microsoft.Office.WebExtension.BindingType.Text})};OSF.OUtil.extend(OSF.DDA.TextBinding,OSF.DDA.Binding);OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{TableDataProperties:"TableDataProperties"});OSF.OUtil.augmentList(OSF.DDA.BindingProperties,{RowCount:"BindingRowCount",ColumnCount:"BindingColumnCount",HasHeaders:"HasHeaders"});OSF.DDA.TableDataProperties={TableRows:"TableRows",TableHeaders:"TableHeaders"};OSF.DDA.TableBinding=function(n,t,i,r,u){OSF.DDA.TableBinding.uber.constructor.call(this,n,t);OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.BindingType.Table},rowCount:{value:i?i:0},columnCount:{value:r?r:0},hasHeaders:{value:u?u:!1}});var f=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[f.AddRowsAsync,f.AddColumnsAsync,f.DeleteAllDataValuesAsync])};OSF.OUtil.extend(OSF.DDA.TableBinding,OSF.DDA.Binding);OSF.DDA.AsyncMethodNames.addNames({AddRowsAsync:"addRowsAsync",AddColumnsAsync:"addColumnsAsync",DeleteAllDataValuesAsync:"deleteAllDataValuesAsync"}),function(){function n(n){return n.id}OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddRowsAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:n}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AddColumnsAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Data,types:["object"]}],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:n}]});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:n}]})}();OSF.OUtil.augmentList(OSF.DDA.BindingProperties,{RowCount:"BindingRowCount",ColumnCount:"BindingColumnCount",HasHeaders:"HasHeaders"});OSF.DDA.MatrixBinding=function(n,t,i,r){OSF.DDA.MatrixBinding.uber.constructor.call(this,n,t);OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.BindingType.Matrix},rowCount:{value:i?i:0},columnCount:{value:r?r:0}})};OSF.OUtil.extend(OSF.DDA.MatrixBinding,OSF.DDA.Binding);OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.PropertyDescriptors.BindingProperties);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.BindingRequest,toHost:[{name:Microsoft.Office.WebExtension.Parameters.ItemName,value:"ItemName"},{name:Microsoft.Office.WebExtension.Parameters.Id,value:"BindingId"},{name:Microsoft.Office.WebExtension.Parameters.BindingType,value:"BindingType"},{name:Microsoft.Office.WebExtension.Parameters.PromptText,value:"PromptText"},{name:Microsoft.Office.WebExtension.Parameters.Columns,value:"Columns"},{name:Microsoft.Office.WebExtension.Parameters.SampleData,value:"SampleData"},{name:Microsoft.Office.WebExtension.Parameters.FailOnCollision,value:"FailOnCollision"}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.BindingType,toHost:[{name:Microsoft.Office.WebExtension.BindingType.Text,value:2},{name:Microsoft.Office.WebExtension.BindingType.Matrix,value:3},{name:Microsoft.Office.WebExtension.BindingType.Table,value:1}],invertible:!0});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.BindingProperties,fromHost:[{name:OSF.DDA.BindingProperties.Id,value:"Name"},{name:OSF.DDA.BindingProperties.Type,value:"BindingType"},{name:OSF.DDA.BindingProperties.RowCount,value:"RowCount"},{name:OSF.DDA.BindingProperties.ColumnCount,value:"ColCount"},{name:OSF.DDA.BindingProperties.HasHeaders,value:"HasHeaders"}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,fromHost:[{name:OSF.DDA.PropertyDescriptors.BindingProperties,value:0}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,fromHost:[{name:OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,value:OSF.DDA.WAC.UniqueArguments.BindingResponse}],toHost:[{name:OSF.DDA.WAC.UniqueArguments.BindingRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,fromHost:[{name:OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,value:OSF.DDA.WAC.UniqueArguments.BindingResponse}],toHost:[{name:OSF.DDA.WAC.UniqueArguments.BindingRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidReleaseBindingMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.BindingRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetBindingMethod,fromHost:[{name:OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,value:OSF.DDA.WAC.UniqueArguments.BindingResponse}],toHost:[{name:OSF.DDA.WAC.UniqueArguments.BindingRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,fromHost:[{name:OSF.DDA.ListDescriptors.BindingList,value:OSF.DDA.WAC.UniqueArguments.BindingResponse}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidGetBindingDataMethod,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}],toHost:[{name:OSF.DDA.WAC.UniqueArguments.GetData,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidSetBindingDataMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.SetData,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{BindingSelectionChanged:"bindingSelectionChanged",BindingDataChanged:"bindingDataChanged"});OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{BindingSelectionChangedEvent:"BindingSelectionChangedEvent"});OSF.DDA.BindingSelectionChangedEventArgs=function(n,t){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.BindingSelectionChanged},binding:{value:n}});for(var i in t)OSF.OUtil.defineEnumerableProperty(this,i,{value:t[i]})};OSF.DDA.BindingDataChangedEventArgs=function(n){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.BindingDataChanged},binding:{value:n}})};OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.BindingSelectionChangedEvent);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,fromHost:[{name:OSF.DDA.PropertyDescriptors.BindingProperties,value:OSF.DDA.WAC.UniqueArguments.BindingEventSource},{name:OSF.DDA.PropertyDescriptors.Subset,value:OSF.DDA.PropertyDescriptors.Subset}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidBindingDataChangedEvent,fromHost:[{name:OSF.DDA.PropertyDescriptors.BindingProperties,value:OSF.DDA.WAC.UniqueArguments.BindingEventSource}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddRowsMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.AddRowsColumns,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidAddColumnsMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.AddRowsColumns,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidClearAllRowsMethod,toHost:[{name:OSF.DDA.WAC.UniqueArguments.BindingRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.WAC.UniqueArguments.AddRowsColumns,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:"BindingId"},{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.WAC.UniqueArguments.Data}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.PropertyDescriptors.Subset,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.StartRow,value:"StartRow"},{name:Microsoft.Office.WebExtension.Parameters.StartColumn,value:"StartCol"},{name:Microsoft.Office.WebExtension.Parameters.RowCount,value:"RowCount"},{name:Microsoft.Office.WebExtension.Parameters.ColumnCount,value:"ColCount"}]});Microsoft.Office.WebExtension.GoToType={Binding:"binding",NamedItem:"namedItem",Slide:"slide",Index:"index"};Microsoft.Office.WebExtension.SelectionMode={Default:"default",Selected:"selected",None:"none"};Microsoft.Office.WebExtension.Index={First:"first",Last:"last",Next:"next",Previous:"previous"};OSF.DDA.AsyncMethodNames.addNames({GoToByIdAsync:"goToByIdAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.GoToByIdAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:["string","number"]},{name:Microsoft.Office.WebExtension.Parameters.GoToType,"enum":Microsoft.Office.WebExtension.GoToType}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.SelectionMode,value:{"enum":Microsoft.Office.WebExtension.SelectionMode,defaultValue:Microsoft.Office.WebExtension.SelectionMode.Default}}]});OSF.OUtil.setNamespace("Marshaling",OSF.DDA);OSF.DDA.Marshaling.NavigationKeys={NavigationRequest:"DdaGoToByIdMethod",Id:"Id",GoToType:"GoToType",SelectionMode:"SelectionMode"},function(n){var t="NamedItem";n[n.Binding=0]="Binding";n[n[t]=1]=t;n[n.Slide=2]="Slide";n[n.Index=3]="Index"}(OSF_DDA_Marshaling_GoToType||(OSF_DDA_Marshaling_GoToType={}));OSF.DDA.Marshaling.GoToType=OSF_DDA_Marshaling_GoToType,function(n){var t="Selected";n[n.Default=0]="Default";n[n[t]=1]=t;n[n.None=2]="None"}(OSF_DDA_Marshaling_SelectionMode||(OSF_DDA_Marshaling_SelectionMode={}));OSF.DDA.Marshaling.SelectionMode=OSF_DDA_Marshaling_SelectionMode;OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.Marshaling.NavigationKeys.NavigationRequest);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.Marshaling.NavigationKeys.NavigationRequest,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.NavigationKeys.Id},{name:Microsoft.Office.WebExtension.Parameters.GoToType,value:OSF.DDA.Marshaling.NavigationKeys.GoToType},{name:Microsoft.Office.WebExtension.Parameters.SelectionMode,value:OSF.DDA.Marshaling.NavigationKeys.SelectionMode}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.GoToType,toHost:[{name:Microsoft.Office.WebExtension.GoToType.Binding,value:OSF.DDA.Marshaling.GoToType.Binding},{name:Microsoft.Office.WebExtension.GoToType.NamedItem,value:OSF.DDA.Marshaling.GoToType.NamedItem},{name:Microsoft.Office.WebExtension.GoToType.Slide,value:OSF.DDA.Marshaling.GoToType.Slide},{name:Microsoft.Office.WebExtension.GoToType.Index,value:OSF.DDA.Marshaling.GoToType.Index}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:Microsoft.Office.WebExtension.Parameters.SelectionMode,toHost:[{name:Microsoft.Office.WebExtension.SelectionMode.Default,value:OSF.DDA.Marshaling.SelectionMode.Default},{name:Microsoft.Office.WebExtension.SelectionMode.Selected,value:OSF.DDA.Marshaling.SelectionMode.Selected},{name:Microsoft.Office.WebExtension.SelectionMode.None,value:OSF.DDA.Marshaling.SelectionMode.None}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidNavigateToMethod,toHost:[{name:OSF.DDA.Marshaling.NavigationKeys.NavigationRequest,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]}),function(n){var t;(function(t){var u=function(){function r(){var n=this,t=n;n._pseudoDocument=u;n._eventDispatch=u;n._processAppCommandInvocation=function(n){var i=t._verifyManifestCallback(n.callbackName),r;if(i.errorCode!=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess){t._invokeAppCommandCompletedMethod(n.appCommandId,i.errorCode,"");return}r=t._constructEventObjectForCallback(n);r?window.setTimeout(function(){i.callback(r)},0):t._invokeAppCommandCompletedMethod(n.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,"")}}var f="object",e="string",u=null;return r.initializeOsfDda=function(){OSF.DDA.AsyncMethodNames.addNames({AppCommandInvocationCompletedAsync:"appCommandInvocationCompletedAsync"});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Id,types:[e]},{name:Microsoft.Office.WebExtension.Parameters.Status,types:["number"]},{name:Microsoft.Office.WebExtension.Parameters.Data,types:[e]}]});OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{AppCommandInvokedEvent:"AppCommandInvokedEvent"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AppCommandInvoked:"appCommandInvoked"});OSF.OUtil.setNamespace("AppCommand",OSF.DDA);OSF.DDA.AppCommand.AppCommandInvokedEventArgs=n.AppCommand.AppCommandInvokedEventArgs},r.prototype.initializeAndChangeOnce=function(n){var i=this,r;t.registerDdaFacade();i._pseudoDocument={};OSF.DDA.DispIdHost.addAsyncMethods(i._pseudoDocument,[OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync]);i._eventDispatch=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.AppCommandInvoked]);r=function(t){n&&(t.status=="succeeded"?n(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess):n(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError))};OSF.DDA.DispIdHost.addEventSupport(i._pseudoDocument,i._eventDispatch);i._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked,i._processAppCommandInvocation,r)},r.prototype._verifyManifestCallback=function(n){var e={callback:u,errorCode:OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback},o;n=n.trim();try{for(var t=n.split("."),i=window,r=0;r<t.length-1;r++)if(i[t[r]]&&typeof i[t[r]]==f)i=i[t[r]];else return e;if(o=i[t[t.length-1]],typeof o!="function")return e}catch(s){return e}return{callback:o,errorCode:OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess}},r.prototype._invokeAppCommandCompletedMethod=function(n,t,i){this._pseudoDocument.appCommandInvocationCompletedAsync(n,t,i)},r.prototype._constructEventObjectForCallback=function(n){var f=this,t=new i,r;try{r=JSON.parse(n.eventObjStr);this._translateEventObjectInternal(r,t);Object.defineProperty(t,"completed",{value:function(){var i=JSON.stringify(t);f._invokeAppCommandCompletedMethod(n.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess,i)},enumerable:!0})}catch(e){t=u}return t},r.prototype._translateEventObjectInternal=function(n,t){var i,r;for(i in n)n.hasOwnProperty(i)&&(r=n[i],typeof r==f&&r!=u?(OSF.OUtil.defineEnumerableProperty(t,i,{value:{}}),this._translateEventObjectInternal(r,t[i])):Object.defineProperty(t,i,{value:r,enumerable:!0,writable:!0}))},r.prototype._constructObjectByTemplate=function(n,t){var r={},i;if(!n||!t)return r;for(i in n)if(n.hasOwnProperty(i)&&(r[i]=u,t[i]!=u)){var o=n[i],s=t[i],h=typeof s;typeof o==f&&o!=u?r[i]=this._constructObjectByTemplate(o,s):(h=="number"||h==e||h=="boolean")&&(r[i]=s)}return r},r.instance=function(){return r._instance==u&&(r._instance=new r),r._instance},r._instance=u,r}(),r,i;t.AppCommandManager=u;r=function(){function n(n,t,i){var r=this;r.type=Microsoft.Office.WebExtension.EventType.AppCommandInvoked;r.appCommandId=n;r.callbackName=t;r.eventObjStr=i}return n.create=function(i){return new n(i[t.AppCommandInvokedEventEnums.AppCommandId],i[t.AppCommandInvokedEventEnums.CallbackName],i[t.AppCommandInvokedEventEnums.EventObjStr])},n}();t.AppCommandInvokedEventArgs=r;i=function(){function n(){}return n}();t.AppCommandCallbackEventArgs=i;t.AppCommandInvokedEventEnums={AppCommandId:"appCommandId",CallbackName:"callbackName",EventObjStr:"eventObjStr"}})(t=n.AppCommand||(n.AppCommand={}))}(OfficeExt||(OfficeExt={}));OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();OSF.OUtil.setNamespace("Marshaling",OSF.DDA);OSF.OUtil.setNamespace("AppCommand",OSF.DDA.Marshaling),function(n){var t="EventObjStr",i="CallbackName",r="AppCommandId";n[n[r]=0]=r;n[n[i]=1]=i;n[n[t]=2]=t}(OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys||(OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys={}));OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys=OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys,function(n){n[n.Id=0]="Id";n[n.Status=1]="Status";n[n.Data=2]="Data"}(OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys||(OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys={}));OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys=OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys,function(n){var t;(function(t){function i(){if(OSF.DDA.WAC){var t=OSF.DDA.WAC.Delegate.ParameterMap;t.define({type:OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.Id,value:OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Id},{name:Microsoft.Office.WebExtension.Parameters.Status,value:OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Status},{name:Microsoft.Office.WebExtension.Parameters.Data,value:OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Data}]});t.define({type:OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.AppCommandInvokedEvent,value:t.self}]});t.addComplexType(OSF.DDA.EventDescriptors.AppCommandInvokedEvent);t.define({type:OSF.DDA.EventDescriptors.AppCommandInvokedEvent,fromHost:[{name:n.AppCommand.AppCommandInvokedEventEnums.AppCommandId,value:OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.AppCommandId},{name:n.AppCommand.AppCommandInvokedEventEnums.CallbackName,value:OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.CallbackName},{name:n.AppCommand.AppCommandInvokedEventEnums.EventObjStr,value:OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.EventObjStr}]})}}t.registerDdaFacade=i})(t=n.AppCommand||(n.AppCommand={}))}(OfficeExt||(OfficeExt={}));OSF.DialogShownStatus={hasDialogShown:!1,isWindowDialog:!1};OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{DialogMessageReceivedEvent:"DialogMessageReceivedEvent"});OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{DialogMessageReceived:"dialogMessageReceived",DialogEventReceived:"dialogEventReceived"});OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{MessageType:"messageType",MessageContent:"messageContent"});OSF.DDA.DialogEventType={};OSF.OUtil.augmentList(OSF.DDA.DialogEventType,{DialogClosed:"dialogClosed",NavigationFailed:"naviationFailed"});OSF.DDA.AsyncMethodNames.addNames({DisplayDialogAsync:"displayDialogAsync",CloseAsync:"close"});OSF.DDA.SyncMethodNames.addNames({MessageParent:"messageParent",AddMessageHandler:"addEventHandler"});OSF.DDA.UI.ParentUI=function(){var i=new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived,Microsoft.Office.WebExtension.EventType.DialogEventReceived]),t=OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName,n=this;n[t]||OSF.OUtil.defineEnumerableProperty(n,t,{value:function(){var t=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];t(arguments,i,n)}});OSF.OUtil.finalizeProperties(this)};OSF.DDA.UI.ChildUI=function(){var t=OSF.DDA.SyncMethodNames.MessageParent.displayName,n=this;n[t]||OSF.OUtil.defineEnumerableProperty(n,t,{value:function(){var t=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];return t(arguments,n)}});OSF.OUtil.finalizeProperties(this)};OSF.DialogHandler=function(){};OSF.DDA.DialogEventArgs=function(n){n[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogMessageReceived?OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DialogMessageReceived},message:{value:n[OSF.DDA.PropertyDescriptors.MessageContent]}}):OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DialogEventReceived},error:{value:n[OSF.DDA.PropertyDescriptors.MessageType]}})};OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.DisplayDialogAsync,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.Url,types:["string"]}],supportedOptions:[{name:Microsoft.Office.WebExtension.Parameters.Width,value:{types:["number"],defaultValue:99}},{name:Microsoft.Office.WebExtension.Parameters.Height,value:{types:["number"],defaultValue:99}},{name:Microsoft.Office.WebExtension.Parameters.RequireHTTPs,value:{types:["boolean"],defaultValue:!0}},{name:Microsoft.Office.WebExtension.Parameters.XFrameDenySafe,value:{types:["boolean"],defaultValue:!0}}],privateStateCallbacks:[],onSucceeded:function(n){var u=n[Microsoft.Office.WebExtension.Parameters.Id],i=n[Microsoft.Office.WebExtension.Parameters.Data],t=new OSF.DialogHandler,f=OSF.DDA.AsyncMethodNames.CloseAsync.displayName,r;return OSF.OUtil.defineEnumerableProperty(t,f,{value:function(){var n=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];n(arguments,u,i,t)}}),r=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName,OSF.OUtil.defineEnumerableProperty(t,r,{value:function(){var r=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id],n=r.verifyAndExtractCall(arguments,t,i),u=n[Microsoft.Office.WebExtension.Parameters.EventType],f=n[Microsoft.Office.WebExtension.Parameters.Handler];return i.addEventHandlerAndFireQueuedEvent(u,f)}}),t},checkCallArgs:function(n){return n[Microsoft.Office.WebExtension.Parameters.Width]<=0&&(n[Microsoft.Office.WebExtension.Parameters.Width]=1),n[Microsoft.Office.WebExtension.Parameters.Width]>100&&(n[Microsoft.Office.WebExtension.Parameters.Width]=99),n[Microsoft.Office.WebExtension.Parameters.Height]<=0&&(n[Microsoft.Office.WebExtension.Parameters.Height]=1),n[Microsoft.Office.WebExtension.Parameters.Height]>100&&(n[Microsoft.Office.WebExtension.Parameters.Height]=99),n}});OSF.DDA.AsyncMethodCalls.define({method:OSF.DDA.AsyncMethodNames.CloseAsync,requiredArguments:[],supportedOptions:[],privateStateCallbacks:[]});OSF.DDA.SyncMethodCalls.define({method:OSF.DDA.SyncMethodNames.MessageParent,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.MessageToParent,types:["string","number","boolean"]}],supportedOptions:[]});OSF.DDA.SyncMethodCalls.define({method:OSF.DDA.SyncMethodNames.AddMessageHandler,requiredArguments:[{name:Microsoft.Office.WebExtension.Parameters.EventType,"enum":Microsoft.Office.WebExtension.EventType,verify:function(n,t,i){return i.supportsEvent(n)}},{name:Microsoft.Office.WebExtension.Parameters.Handler,types:["function"]}],supportedOptions:[]});OSF.OUtil.setNamespace("Marshaling",OSF.DDA);OSF.OUtil.setNamespace("Dialog",OSF.DDA.Marshaling),function(n){var t="MessageContent",i="MessageType";n[n[i]=0]=i;n[n[t]=1]=t}(OSF_DDA_Marshaling_Dialog_DialogMessageReceivedEventKeys||(OSF_DDA_Marshaling_Dialog_DialogMessageReceivedEventKeys={}));OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys=OSF_DDA_Marshaling_Dialog_DialogMessageReceivedEventKeys;OSF.DDA.Marshaling.MessageParentKeys={MessageToParent:"messageToParent"};OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,fromHost:[{name:OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,value:OSF.DDA.WAC.Delegate.ParameterMap.self}]});OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.DialogMessageReceivedEvent);OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,fromHost:[{name:OSF.DDA.PropertyDescriptors.MessageType,value:OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType},{name:OSF.DDA.PropertyDescriptors.MessageContent,value:OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent}]});OSF.DDA.WAC.Delegate.ParameterMap.define({type:OSF.DDA.MethodDispId.dispidMessageParentMethod,toHost:[{name:Microsoft.Office.WebExtension.Parameters.MessageToParent,value:OSF.DDA.Marshaling.MessageParentKeys.MessageToParent}]});OSF.DDA.WAC.Delegate.openDialog=function(n){function i(n){var t={Error:n};u(Microsoft.Office.Common.InvokeResultCode.noError,t)}var r=JSON.parse(n.targetId),u=OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(!0,n),t;if(OSF.DialogShownStatus.hasDialogShown)i(OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened);else{if(r.xFrameDenySafe){t=n.onComplete;function f(n){n!=OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserAllow&&t!=null&&t(n)}n.onComplete=f;OSF.DialogShownStatus.isWindowDialog=!0;OfficeExt.AddinNativeAction.Dialog.setHandlerAndShowDialogCallback(function(t){n.onEvent&&n.onEvent(t);OSF.AppTelemetry&&OSF.AppTelemetry.onEventDone(n.dispId)},i)}else OSF.DialogShownStatus.isWindowDialog=!1;OSF.DDA.WAC.Delegate.registerEventAsync(n)}};OSF.DDA.WAC.Delegate.messageParent=function(n){window.opener!=null?OfficeExt.AddinNativeAction.Dialog.messageParent(n):OSF.DDA.WAC.Delegate.executeAsync(n)};OSF.DDA.WAC.Delegate.closeDialog=function(n){function t(n){var t={Error:n};i(Microsoft.Office.Common.InvokeResultCode.noError,t)}var i=OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(!1,n);OSF.DialogShownStatus.hasDialogShown?OSF.DialogShownStatus.isWindowDialog?(n.onCalling&&n.onCalling(),OfficeExt.AddinNativeAction.Dialog.closeDialog(t)):OSF.DDA.WAC.Delegate.unregisterEventAsync(n):t(OSF.DDA.ErrorCodeManager.errorCodes.ooeWebDialogClosed)},function(n){var t=function(){function n(n,t){var i=this;OSF.DDA.WordDocument.uber.constructor.call(i,n,t);OSF.DDA.DispIdHost.addAsyncMethods(i,[OSF.DDA.AsyncMethodNames.GoToByIdAsync,OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,OSF.DDA.AsyncMethodNames.SaveAsync,OSF.DDA.AsyncMethodNames.RefreshAsync,OSF.DDA.SyncMethodNames.MessageParent]);OSF.DDA.DispIdHost.addEventSupport(i,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));OSF.DDA.DispIdHost.addEventSupport(i,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.BindingSelectionChanged]));OSF.OUtil.defineEnumerableProperty(i,"customXmlParts",{value:new OSF.DDA.CustomXmlParts});OSF.OUtil.defineEnumerableProperty(i,"bindings",{value:new OSF.DDA.BindingFacade(i)});OSF.OUtil.finalizeProperties(i)}return n}();n.WordDocument=t}(OSFWordWAC||(OSFWordWAC={}));OSF.DDA.WordDocument=OSFWordWAC.WordDocument;OSF.OUtil.extend(OSF.DDA.WordDocument,OSF.DDA.Document);OSF.OUtil.redefineList(Microsoft.Office.WebExtension.CoercionType,{Html:"html",Text:"text",Ooxml:"ooxml",Table:"table",Matrix:"matrix",Image:"image"});OSF.DDA.TableDataProperties={TableRows:"TableRows",TableHeaders:"TableHeaders"};OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM=function(n,t){OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);n.doc=new OSF.DDA.WordDocument(n,this._initializeSettings(n,!0));OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi,[OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);t()};OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize=function(n){var t,i,r;OSF.WebApp._UpdateLinksForHostAndXdmInfo();t=new OSF.DDA.License(n.get_eToken());this.initWebDialog(n);OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(n,n.doc,t));OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.WAC.getDelegateMethods,OSF.DDA.WAC.Delegate.ParameterMap));i=n.get_reason();Microsoft.Office.WebExtension.initialize(i);r=OfficeExt.AppCommand.AppCommandManager.instance();r.initializeAndChangeOnce()},function(n){var t=function(){function n(n,t){this.m_actionInfo=n;this.m_isWriteOperation=t}return Object.defineProperty(n.prototype,"actionInfo",{get:function(){return this.m_actionInfo},enumerable:!0,configurable:!0}),Object.defineProperty(n.prototype,"isWriteOperation",{get:function(){return this.m_isWriteOperation},enumerable:!0,configurable:!0}),n}();n.Action=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.createSetPropertyAction=function(t,i,r,u){var f;n.Utility.validateObjectPath(i);var e={Id:t._nextId(),ActionType:4,Name:r,ObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}},s=[u],o=n.Utility.setMethodArguments(t,e.ArgumentInfo,s);return n.Utility.validateReferencedObjectPaths(o),f=new n.Action(e,!0),t._pendingRequest.addAction(f),t._pendingRequest.addReferencedObjectPath(i._objectPath),t._pendingRequest.addReferencedObjectPaths(o),f},t.createMethodAction=function(t,i,r,u,f){var e,o,h,s;return n.Utility.validateObjectPath(i),e={Id:t._nextId(),ActionType:3,Name:r,ObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}},o=n.Utility.setMethodArguments(t,e.ArgumentInfo,f),n.Utility.validateReferencedObjectPaths(o),h=u!=1,s=new n.Action(e,h),t._pendingRequest.addAction(s),t._pendingRequest.addReferencedObjectPath(i._objectPath),t._pendingRequest.addReferencedObjectPaths(o),s},t.createQueryAction=function(t,i,r){var u,f;return n.Utility.validateObjectPath(i),u={Id:t._nextId(),ActionType:2,Name:"",ObjectPathId:i._objectPath.objectPathInfo.Id},u.QueryInfo=r,f=new n.Action(u,!1),t._pendingRequest.addAction(f),t._pendingRequest.addReferencedObjectPath(i._objectPath),f},t.createInstantiateAction=function(t,i){n.Utility.validateObjectPath(i);var u={Id:t._nextId(),ActionType:1,Name:"",ObjectPathId:i._objectPath.objectPathInfo.Id},r=new n.Action(u,!1);return t._pendingRequest.addAction(r),t._pendingRequest.addReferencedObjectPath(i._objectPath),t._pendingRequest.addActionResultHandler(r,new n.InstantiateActionResultHandler(i)),r},t.createTraceAction=function(t,i){var r={Id:t._nextId(),ActionType:5,Name:"Trace",ObjectPathId:0},u=new n.Action(r,!1);return t._pendingRequest.addAction(u),t._pendingRequest.addTrace(r.Id,i),u},t}();n.ActionFactory=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(t,i){n.Utility.checkArgumentNull(t,"context");this.m_context=t;this.m_objectPath=i;this.m_objectPath&&(t._processingResult||(n.ActionFactory.createInstantiateAction(t,this),t._autoCleanup&&this._KeepReference&&t.trackedObjects._autoAdd(this)))}return Object.defineProperty(t.prototype,"context",{get:function(){return this.m_context},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_objectPath",{get:function(){return this.m_objectPath},set:function(n){this.m_objectPath=n},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isNull",{get:function(){return n.Utility.throwIfNotLoaded("isNull",this._isNull,null,this._isNull),this._isNull},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_isNull",{get:function(){return this.m_isNull},set:function(n){this.m_isNull=n;n&&this.m_objectPath&&this.m_objectPath._updateAsNullObject()},enumerable:!0,configurable:!0}),t.prototype._handleResult=function(t){this.m_isNull=n.Utility.isNullOrUndefined(t);this.m_isNull&&this.m_objectPath&&this.m_objectPath._updateAsNullObject()},t}();n.ClientObject=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(n){this.m_context=n;this.m_actions=[];this.m_actionResultHandler={};this.m_referencedObjectPaths={};this.m_flags=0;this.m_traceInfos={}}return Object.defineProperty(t.prototype,"flags",{get:function(){return this.m_flags},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"traceInfos",{get:function(){return this.m_traceInfos},enumerable:!0,configurable:!0}),t.prototype.addAction=function(n){n.isWriteOperation&&(this.m_flags=this.m_flags|1);this.m_actions.push(n)},Object.defineProperty(t.prototype,"hasActions",{get:function(){return this.m_actions.length>0},enumerable:!0,configurable:!0}),t.prototype.addTrace=function(n,t){this.m_traceInfos[n]=t},t.prototype.addReferencedObjectPath=function(t){if(!this.m_referencedObjectPaths[t.objectPathInfo.Id])for(t.isValid||n.Utility.throwError(n.ResourceStrings.invalidObjectPath,n.Utility.getObjectPathExpression(t));t;)t.isWriteOperation&&(this.m_flags=this.m_flags|1),this.m_referencedObjectPaths[t.objectPathInfo.Id]=t,t.objectPathInfo.ObjectPathType==3&&this.addReferencedObjectPaths(t.argumentObjectPaths),t=t.parentObjectPath},t.prototype.addReferencedObjectPaths=function(n){if(n)for(var t=0;t<n.length;t++)this.addReferencedObjectPath(n[t])},t.prototype.addActionResultHandler=function(n,t){this.m_actionResultHandler[n.actionInfo.Id]=t},t.prototype.buildRequestMessageBody=function(){var r={},t,i,n;for(t in this.m_referencedObjectPaths)r[t]=this.m_referencedObjectPaths[t].objectPathInfo;for(i=[],n=0;n<this.m_actions.length;n++)i.push(this.m_actions[n].actionInfo);return{Actions:i,ObjectPaths:r}},t.prototype.processResponse=function(n){var t,i,r;if(n&&n.Results)for(t=0;t<n.Results.length;t++)i=n.Results[t],r=this.m_actionResultHandler[i.ActionId],r&&r._handleResult(i.Value)},t.prototype.invalidatePendingInvalidObjectPaths=function(){for(var n in this.m_referencedObjectPaths)this.m_referencedObjectPaths[n].isInvalidAfterRequest&&(this.m_referencedObjectPaths[n].isValid=!1)},t}();n.ClientRequest=t}(OfficeExtension||(OfficeExtension={})),function(n){function r(n){t=n}var t=function(){return new n.OfficeJsRequestExecutor},i;n._setRequestExecutorFactory=r;i=function(){function i(i){this.m_nextId=0;this.m_url=i;n.Utility.isNullOrEmptyString(this.m_url)&&(this.m_url=n.Constants.localDocument);this._processingResult=!1;this._customData=n.Constants.iterativeExecutor;this._requestExecutor=t();this.sync=this.sync.bind(this)}return Object.defineProperty(i.prototype,"_pendingRequest",{get:function(){return this.m_pendingRequest==null&&(this.m_pendingRequest=new n.ClientRequest(this)),this.m_pendingRequest},enumerable:!0,configurable:!0}),Object.defineProperty(i.prototype,"trackedObjects",{get:function(){return this.m_trackedObjects||(this.m_trackedObjects=new n.TrackedObjects(this)),this.m_trackedObjects},enumerable:!0,configurable:!0}),i.prototype.load=function(t,i){var u,f,r,e;n.Utility.validateContext(this,t);u={};typeof i=="string"?(f=i,u.Select=this.parseSelectExpand(f)):Array.isArray(i)?u.Select=i:typeof i=="object"?(r=i,typeof r.select=="string"?u.Select=this.parseSelectExpand(r.select):Array.isArray(r.select)?u.Select=r.select:n.Utility.isNullOrUndefined(r.select)||n.Utility.throwError(n.ResourceStrings.invalidArgument,"option.select"),typeof r.expand=="string"?u.Expand=this.parseSelectExpand(r.expand):Array.isArray(r.expand)?u.Expand=r.expand:n.Utility.isNullOrUndefined(r.expand)||n.Utility.throwError(n.ResourceStrings.invalidArgument,"option.expand"),typeof r.top=="number"?u.Top=r.top:n.Utility.isNullOrUndefined(r.top)||n.Utility.throwError(n.ResourceStrings.invalidArgument,"option.top"),typeof r.skip=="number"?u.Skip=r.skip:n.Utility.isNullOrUndefined(r.skip)||n.Utility.throwError(n.ResourceStrings.invalidArgument,"option.skip")):n.Utility.isNullOrUndefined(i)||n.Utility.throwError(n.ResourceStrings.invalidArgument,"option");e=n.ActionFactory.createQueryAction(this,t,u);this._pendingRequest.addActionResultHandler(e,t)},i.prototype.trace=function(t){n.ActionFactory.createTraceAction(this,t)},i.prototype.parseSelectExpand=function(t){var f=[],u,i,r;if(!n.Utility.isNullOrEmptyString(t))for(u=t.split(","),i=0;i<u.length;i++)r=u[i],r=r.trim(),f.push(r);return f},i.prototype.syncPrivate=function(t,i){var r=this._pendingRequest,e,f;if(!r.hasActions){t();return}this.m_pendingRequest=null;var o=r.buildRequestMessageBody(),s=r.flags,u=this._requestExecutor;u||(u=new n.OfficeJsRequestExecutor);e={Url:this.m_url,Headers:null,Body:o};r.invalidatePendingInvalidObjectPaths();f=this;u.executeAsync(this._customData,s,e,function(u){var e,s=[],h,o,c,l;if(n.Utility.isNullOrEmptyString(u.ErrorCode)?u.Body&&u.Body.Error&&(e=new n._Internal.RuntimeError(u.Body.Error.Code,u.Body.Error.Message,s,{errorLocation:u.Body.Error.Location})):e=new n._Internal.RuntimeError(u.ErrorCode,u.ErrorMessage,s,{}),u.Body&&u.Body.TraceIds)for(h=r.traceInfos,o=0;o<u.Body.TraceIds.length;o++)c=u.Body.TraceIds[o],l=h[c],s.push(l);if(e){i(e);return}else{f._processingResult=!0;try{r.processResponse(u.Body)}finally{f._processingResult=!1}t();return}})},i.prototype.sync=function(t){var i=this;return new n.Promise(function(n,r){i.syncPrivate(function(){n(t)},function(n){r(n)})})},i._run=function(t,i,r,u,f,e){r===void 0&&(r=3);u===void 0&&(u=5e3);var c=new n.Promise(function(n,t){n()}),o,h=!1,s;return c.then(function(){o=t();o._autoCleanup=!0;var r=i(o);return(n.Utility.isNullOrUndefined(r)||typeof r.then!="function")&&n.Utility.throwError(n.ResourceStrings.runMustReturnPromise),r}).then(function(n){return o.sync(n)}).then(function(n){h=!0;s=n}).catch(function(n){s=n}).then(function(){function s(){n++;for(var i in t)o.trackedObjects.remove(t[i]);o.sync().then(function(){f&&f(n)}).catch(function(){e&&e(n);n<r&&setTimeout(function(){s()},u)})}var t=o.trackedObjects._retrieveAndClearAutoCleanupList(),i,n;o._autoCleanup=!1;for(i in t)t[i]._objectPath.isValid=!1;n=0;s()}).then(function(){if(h)return s;else throw s;})},i.prototype._nextId=function(){return++this.m_nextId},i}();n.ClientRequestContext=i}(OfficeExtension||(OfficeExtension={})),function(n){(function(n){n[n.None=0]="None";n[n.WriteOperation=1]="WriteOperation"})(n.ClientRequestFlags||(n.ClientRequestFlags={}));var t=n.ClientRequestFlags}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function n(){}return Object.defineProperty(n.prototype,"value",{get:function(){return this.m_value},enumerable:!0,configurable:!0}),n.prototype._handleResult=function(n){typeof n=="object"&&n&&n._IsNull||(this.m_value=n)},n}();n.ClientResult=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function n(){}return n.getItemAt="GetItemAt",n.id="Id",n.idPrivate="_Id",n.index="_Index",n.items="_Items",n.iterativeExecutor="IterativeExecutor",n.localDocument="http://document.localhost/",n.localDocumentApiPrefix="http://document.localhost/_api/",n.referenceId="_ReferenceId",n.sourceLibHeader="X-OfficeExtension-Source",n}();n.Constants=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.prototype.executeAsync=function(i,r,u,f){var o=n.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(i,r,u,t.SourceLibHeaderValue),e=n.Embedded&&n.Embedded._getEndpoint();if(!e){n.RichApiMessageUtility.sendResponseOnError(n.Embedded&&n.Embedded.EmbeddedApiStatus.InternalError,"",f);return}e.invoke("executeMethod",function(t,i){n.Utility.log("Response:");n.Utility.log(JSON.stringify(i));t==n.Embedded.EmbeddedApiStatus.Success?n.RichApiMessageUtility.sendResponseOnSuccess(n.RichApiMessageUtility.getResponseBodyFromSafeArray(i.Data),n.RichApiMessageUtility.getResponseHeadersFromSafeArray(i.Data),f):n.RichApiMessageUtility.sendResponseOnError(i.error.Code,i.error.Message,f)},t._transformMessageArrayIntoParams(o))},t._transformMessageArrayIntoParams=function(n){return{ArrayData:n,DdaMethod:{DispatchId:t.DispidExecuteRichApiRequestMethod}}},t.DispidExecuteRichApiRequestMethod=93,t.SourceLibHeaderValue="Embedded",t}();n.EmbedRequestExecutor=t}(OfficeExtension||(OfficeExtension={}));__extends=this.__extends||function(n,t){function r(){this.constructor=n}for(var i in t)t.hasOwnProperty(i)&&(n[i]=t[i]);r.prototype=t.prototype;n.prototype=new r},function(n){var t;(function(n){var t=function(n){function t(t,i,r,u){n.call(this,i);this.name="OfficeExtension.Error";this.code=t;this.message=i;this.traceMessages=r;this.debugInfo=u}return __extends(t,n),t.prototype.toString=function(){return this.code+": "+this.message},t}(Error);n.RuntimeError=t})(t=n._Internal||(n._Internal={}));n.Error=n._Internal.RuntimeError}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function n(){}return n.accessDenied="AccessDenied",n.generalException="GeneralException",n.activityLimitReached="ActivityLimitReached",n}();n.ErrorCodes=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(n){this.m_clientObject=n}return t.prototype._handleResult=function(t){this.m_clientObject._isNull=n.Utility.isNullOrUndefined(t);n.Utility.fixObjectPathIfNecessary(this.m_clientObject,t);t&&!n.Utility.isNullOrUndefined(t[n.Constants.referenceId])&&this.m_clientObject._initReferenceId&&this.m_clientObject._initReferenceId(t[n.Constants.referenceId])},t}();n.InstantiateActionResultHandler=t}(OfficeExtension||(OfficeExtension={})),function(n){}(OfficeExtension||(OfficeExtension={})),function(n){var t,i,r,u;(function(n){n[n.CustomData=0]="CustomData";n[n.Method=1]="Method";n[n.PathAndQuery=2]="PathAndQuery";n[n.Headers=3]="Headers";n[n.Body=4]="Body";n[n.AppPermission=5]="AppPermission";n[n.RequestFlags=6]="RequestFlags"})(n.RichApiRequestMessageIndex||(n.RichApiRequestMessageIndex={}));t=n.RichApiRequestMessageIndex,function(n){n[n.StatusCode=0]="StatusCode";n[n.Headers=1]="Headers";n[n.Body=2]="Body"}(n.RichApiResponseMessageIndex||(n.RichApiResponseMessageIndex={}));i=n.RichApiResponseMessageIndex,function(n){n[n.Instantiate=1]="Instantiate";n[n.Query=2]="Query";n[n.Method=3]="Method";n[n.SetProperty=4]="SetProperty";n[n.Trace=5]="Trace"}(n.ActionType||(n.ActionType={}));r=n.ActionType,function(n){n[n.GlobalObject=1]="GlobalObject";n[n.NewObject=2]="NewObject";n[n.Method=3]="Method";n[n.Property=4]="Property";n[n.Indexer=5]="Indexer";n[n.ReferenceId=6]="ReferenceId";n[n.NullObject=7]="NullObject"}(n.ObjectPathType||(n.ObjectPathType={}));u=n.ObjectPathType}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(n,t,i,r){this.m_objectPathInfo=n;this.m_parentObjectPath=t;this.m_isWriteOperation=!1;this.m_isCollection=i;this.m_isInvalidAfterRequest=r;this.m_isValid=!0}return Object.defineProperty(t.prototype,"objectPathInfo",{get:function(){return this.m_objectPathInfo},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isWriteOperation",{get:function(){return this.m_isWriteOperation},set:function(n){this.m_isWriteOperation=n},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isCollection",{get:function(){return this.m_isCollection},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isInvalidAfterRequest",{get:function(){return this.m_isInvalidAfterRequest},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"parentObjectPath",{get:function(){return this.m_parentObjectPath},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"argumentObjectPaths",{get:function(){return this.m_argumentObjectPaths},set:function(n){this.m_argumentObjectPaths=n},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isValid",{get:function(){return this.m_isValid},set:function(n){this.m_isValid=n},enumerable:!0,configurable:!0}),t.prototype._updateAsNullObject=function(){this.m_isInvalidAfterRequest=!1;this.m_isValid=!0;this.m_objectPathInfo.ObjectPathType=7;this.m_objectPathInfo.Name="";this.m_objectPathInfo.ArgumentInfo={};this.m_parentObjectPath=null;this.m_argumentObjectPaths=null},t.prototype.updateUsingObjectData=function(t){var r=t[n.Constants.referenceId],i;if(!n.Utility.isNullOrEmptyString(r)){this.m_isInvalidAfterRequest=!1;this.m_isValid=!0;this.m_objectPathInfo.ObjectPathType=6;this.m_objectPathInfo.Name=r;this.m_objectPathInfo.ArgumentInfo={};this.m_parentObjectPath=null;this.m_argumentObjectPaths=null;return}if(this.parentObjectPath&&this.parentObjectPath.isCollection&&(i=t[n.Constants.id],n.Utility.isNullOrUndefined(i)&&(i=t[n.Constants.idPrivate]),!n.Utility.isNullOrUndefined(i))){this.m_isInvalidAfterRequest=!1;this.m_isValid=!0;this.m_objectPathInfo.ObjectPathType=5;this.m_objectPathInfo.Name="";this.m_objectPathInfo.ArgumentInfo={};this.m_objectPathInfo.ArgumentInfo.Arguments=[i];this.m_argumentObjectPaths=null;return}},t}();n.ObjectPath=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.createGlobalObjectObjectPath=function(t){var i={Id:t._nextId(),ObjectPathType:1,Name:""};return new n.ObjectPath(i,null,!1,!1)},t.createNewObjectObjectPath=function(t,i,r){var u={Id:t._nextId(),ObjectPathType:2,Name:i};return new n.ObjectPath(u,null,r,!1)},t.createPropertyObjectPath=function(t,i,r,u,f){var e={Id:t._nextId(),ObjectPathType:4,Name:r,ParentObjectPathId:i._objectPath.objectPathInfo.Id};return new n.ObjectPath(e,i._objectPath,u,f)},t.createIndexerObjectPath=function(t,i,r){var u={Id:t._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}};return u.ArgumentInfo.Arguments=r,new n.ObjectPath(u,i._objectPath,!1,!1)},t.createIndexerObjectPathUsingParentPath=function(t,i,r){var u={Id:t._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:i.objectPathInfo.Id,ArgumentInfo:{}};return u.ArgumentInfo.Arguments=r,new n.ObjectPath(u,i,!1,!1)},t.createMethodObjectPath=function(t,i,r,u,f,e,o){var h={Id:t._nextId(),ObjectPathType:3,Name:r,ParentObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}},c=n.Utility.setMethodArguments(t,h.ArgumentInfo,f),s=new n.ObjectPath(h,i._objectPath,e,o);return s.argumentObjectPaths=c,s.isWriteOperation=u!=1,s},t.createChildItemObjectPathUsingIndexerOrGetItemAt=function(i,r,u,f,e){var o=f[n.Constants.id];return n.Utility.isNullOrUndefined(o)&&(o=f[n.Constants.idPrivate]),i&&!n.Utility.isNullOrUndefined(o)?t.createChildItemObjectPathUsingIndexer(r,u,f):t.createChildItemObjectPathUsingGetItemAt(r,u,f,e)},t.createChildItemObjectPathUsingIndexer=function(t,i,r){var f=r[n.Constants.id],u;return n.Utility.isNullOrUndefined(f)&&(f=r[n.Constants.idPrivate]),u=u={Id:t._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}},u.ArgumentInfo.Arguments=[f],new n.ObjectPath(u,i._objectPath,!1,!1)},t.createChildItemObjectPathUsingGetItemAt=function(t,i,r,u){var e=r[n.Constants.index],f;return e&&(u=e),f={Id:t._nextId(),ObjectPathType:3,Name:n.Constants.getItemAt,ParentObjectPathId:i._objectPath.objectPathInfo.Id,ArgumentInfo:{}},f.ArgumentInfo.Arguments=[u],new n.ObjectPath(f,i._objectPath,!1,!1)},t}();n.ObjectPathFactory=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.prototype.executeAsync=function(i,r,u,f){var e=n.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(i,r,u,t.SourceLibHeaderValue);OSF.DDA.RichApi.executeRichApiRequestAsync(e,function(t){n.Utility.log("Response:");n.Utility.log(JSON.stringify(t));t.status=="succeeded"?n.RichApiMessageUtility.sendResponseOnSuccess(n.RichApiMessageUtility.getResponseBody(t),n.RichApiMessageUtility.getResponseHeaders(t),f):n.RichApiMessageUtility.sendResponseOnError(t.error.code,t.error.message,f)})},t.SourceLibHeaderValue="OfficeJs",t}();n.OfficeJsRequestExecutor=t}(OfficeExtension||(OfficeExtension={})),function(n){function u(n){return t.settings.oldxhr=n,r}function r(){return new t}var i=function(){function n(){}return n}(),t;n.OfficeXHRSettings=i;n.resetXHRFactory=u;n.officeXHRFactory=r;t=function(){function t(){}return t.prototype.open=function(i,r){if(this.m_method=i,this.m_url=r,this.m_url.toLowerCase().indexOf(n.Constants.localDocumentApiPrefix)==0)this.m_url=this.m_url.substr(n.Constants.localDocumentApiPrefix.length);else{this.m_innerXhr=t.settings.oldxhr();var u=this;this.m_innerXhr.onreadystatechange=function(){u.innerXhrOnreadystatechage()};this.m_innerXhr.open(i,this.m_url)}},t.prototype.abort=function(){this.m_innerXhr&&this.m_innerXhr.abort()},t.prototype.send=function(i){var f,u,r;this.m_innerXhr?this.m_innerXhr.send(i):(f=this,u=0,n.Utility.isReadonlyRestRequest(this.m_method)||(u=1),r=t.settings.executeRichApiRequestAsync,r||(r=OSF.DDA.RichApi.executeRichApiRequestAsync),r(n.RichApiMessageUtility.buildRequestMessageSafeArray("",u,this.m_method,this.m_url,this.m_requestHeaders,i),function(n){f.officeContextRequestCallback(n)}))},t.prototype.setRequestHeader=function(n,t){this.m_innerXhr?this.m_innerXhr.setRequestHeader(n,t):(this.m_requestHeaders||(this.m_requestHeaders={}),this.m_requestHeaders[n]=t)},t.prototype.getResponseHeader=function(n){return this.m_responseHeaders?this.m_responseHeaders[n.toUpperCase()]:null},t.prototype.getAllResponseHeaders=function(){return this.m_allResponseHeaders},t.prototype.overrideMimeType=function(n){this.m_innerXhr&&this.m_innerXhr.overrideMimeType(n)},t.prototype.innerXhrOnreadystatechage=function(){this.readyState=this.m_innerXhr.readyState;this.readyState==t.DONE&&(this.status=this.m_innerXhr.status,this.statusText=this.m_innerXhr.statusText,this.responseText=this.m_innerXhr.responseText,this.response=this.m_innerXhr.response,this.responseType=this.m_innerXhr.responseType,this.setAllResponseHeaders(this.m_innerXhr.getAllResponseHeaders()));this.onreadystatechange&&this.onreadystatechange()},t.prototype.officeContextRequestCallback=function(i){this.readyState=t.DONE;i.status=="succeeded"?(this.status=n.RichApiMessageUtility.getResponseStatusCode(i),this.m_responseHeaders=n.RichApiMessageUtility.getResponseHeaders(i),console.debug("ResponseHeaders="+JSON.stringify(this.m_responseHeaders)),this.responseText=n.RichApiMessageUtility.getResponseBody(i),console.debug("ResponseText="+this.responseText),this.response=this.responseText):(this.status=500,this.statusText="Internal Error");this.onreadystatechange&&this.onreadystatechange()},t.prototype.setAllResponseHeaders=function(t){var s,o,r,i,u,f,e;if(this.m_allResponseHeaders=t,this.m_responseHeaders={},this.m_allResponseHeaders!=null)for(s=new RegExp("\r?\n"),o=this.m_allResponseHeaders.split(s),r=0;r<o.length;r++)i=o[r],i!=null&&(u=i.indexOf(":"),u>0&&(f=i.substr(0,u),e=i.substr(u+1),f=n.Utility.trim(f),e=n.Utility.trim(e),this.m_responseHeaders[f.toUpperCase()]=e))},t.UNSENT=0,t.OPENED=1,t.DONE=4,t.settings=new i,t}();n.OfficeXHR=t}(OfficeExtension||(OfficeExtension={})),function(n){var t;(function(t){var i;(function(i){function r(){(function(){"use strict";function bt(n){return typeof n=="function"||typeof n=="object"&&n!==null}function k(n){return typeof n=="function"}function kt(n){return typeof n=="object"&&n!==null}function dt(n){d=n}function gt(n){e=n}function ii(){var t=process.nextTick,n=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);return Array.isArray(n)&&n[1]==="0"&&n[2]==="10"&&(t=setImmediate),function(){t(l)}}function ri(){return function(){it(l)}}function ui(){var n=0,i=new ft(l),t=document.createTextNode("");return i.observe(t,{characterData:!0}),function(){t.data=n=++n%2}}function fi(){var n=new MessageChannel;return n.port1.onmessage=l,function(){n.port2.postMessage(0)}}function et(){return function(){setTimeout(l,1)}}function l(){for(var t,i,n=0;n<c;n+=2)t=o[n],i=o[n+1],t(i),o[n]=undefined,o[n+1]=undefined;c=0}function ei(){try{var t=require,n=t("vertx");return it=n.runOnLoop||n.runOnContext,ri()}catch(i){return et()}}function a(){}function oi(){return new TypeError("You cannot resolve a promise with itself")}function si(){return new TypeError("A promises callback cannot return that same promise.")}function hi(n){try{return n.then}catch(t){return p.error=t,p}}function ci(n,t,i,r){try{n.call(t,i,r)}catch(u){return u}}function li(n,t,r){e(function(n){var f=!1,e=ci(r,t,function(i){f||(f=!0,t!==i?y(n,i):u(n,i))},function(t){f||(f=!0,i(n,t))},"Settle: "+(n._label||" unknown promise"));!f&&e&&(f=!0,i(n,e))},n)}function ai(n,t){t._state===v?u(n,t._result):t._state===h?i(n,t._result):w(t,undefined,function(t){y(n,t)},function(t){i(n,t)})}function vi(n,t){if(t.constructor===n.constructor)ai(n,t);else{var r=hi(t);r===p?i(n,p.error):r===undefined?u(n,t):k(r)?li(n,t,r):u(n,t)}}function y(n,t){n===t?i(n,oi()):bt(t)?vi(n,t):u(n,t)}function yi(n){n._onerror&&n._onerror(n._result);g(n)}function u(n,t){n._state===s&&(n._result=t,n._state=v,n._subscribers.length!==0&&e(g,n))}function i(n,t){n._state===s&&(n._state=h,n._result=t,e(yi,n))}function w(n,t,i,r){var u=n._subscribers,f=u.length;n._onerror=null;u[f]=t;u[f+v]=i;u[f+h]=r;f===0&&n._state&&e(g,n)}function g(n){var i=n._subscribers,e=n._state,r,u,f,t;if(i.length!==0){for(f=n._result,t=0;t<i.length;t+=3)r=i[t],u=i[t+e],r?ht(e,r,u,f):u(f);n._subscribers.length=0}}function st(){this.error=null}function pi(n,t){try{return n(t)}catch(i){return b.error=i,b}}function ht(n,t,r,f){var c=k(r),e,l,o,a;if(c){if(e=pi(r,f),e===b?(a=!0,l=e.error,e=null):o=!0,t===e){i(t,si());return}}else e=f,o=!0;t._state!==s||(c&&o?y(t,e):a?i(t,l):n===v?u(t,e):n===h&&i(t,e))}function wi(n,t){try{t(function(t){y(n,t)},function(t){i(n,t)})}catch(r){i(n,r)}}function f(n,t){var r=this;r._instanceConstructor=n;r.promise=new n(a);r._validateInput(t)?(r._input=t,r.length=t.length,r._remaining=t.length,r._init(),r.length===0?u(r.promise,r._result):(r.length=r.length||0,r._enumerate(),r._remaining===0&&u(r.promise,r._result))):i(r.promise,r._validationError())}function bi(n){return new ct(this,n).promise}function ki(n){function e(n){y(t,n)}function o(n){i(t,n)}var u=this,t=new u(a),f,r;if(!tt(n))return i(t,new TypeError("You must pass an array to race.")),t;for(f=n.length,r=0;t._state===s&&r<f;r++)w(u.resolve(n[r]),undefined,e,o);return t}function di(n){var i=this,t;return n&&typeof n=="object"&&n.constructor===i?n:(t=new i(a),y(t,n),t)}function gi(n){var r=this,t=new r(a);return i(t,n),t}function nr(){throw new TypeError("You must pass a resolver function as the first argument to the promise constructor");}function tr(){throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");}function r(n){this._id=pt++;this._state=undefined;this._result=undefined;this._subscribers=[];a!==n&&(k(n)||nr(),this instanceof r||tr(),wi(this,n))}var nt,o,ot,b,ct,lt,at,vt,yt,pt,wt;nt=Array.isArray?Array.isArray:function(n){return Object.prototype.toString.call(n)==="[object Array]"};var tt=nt,c=0,ir={}.toString,it,d,e=function(n,t){o[c]=n;o[c+1]=t;c+=2;c===2&&(d?d(l):ot())};var rt=typeof window!="undefined"?window:undefined,ut=rt||{},ft=ut.MutationObserver||ut.WebKitMutationObserver,ni=typeof process!="undefined"&&{}.toString.call(process)==="[object process]",ti=typeof Uint8ClampedArray!="undefined"&&typeof importScripts!="undefined"&&typeof MessageChannel!="undefined";o=new Array(1e3);ot=ni?ii():ft?ui():ti?fi():rt===undefined&&typeof require=="function"?ei():et();var s=void 0,v=1,h=2,p=new st;b=new st;f.prototype._validateInput=function(n){return tt(n)};f.prototype._validationError=function(){return new t.Error("Array Methods must be provided an Array")};f.prototype._init=function(){this._result=new Array(this.length)};ct=f;f.prototype._enumerate=function(){for(var n=this,i=n.length,r=n.promise,u=n._input,t=0;r._state===s&&t<i;t++)n._eachEntry(u[t],t)};f.prototype._eachEntry=function(n,t){var i=this,r=i._instanceConstructor;kt(n)?n.constructor===r&&n._state!==s?(n._onerror=null,i._settledAt(n._state,t,n._result)):i._willSettleAt(r.resolve(n),t):(i._remaining--,i._result[t]=n)};f.prototype._settledAt=function(n,t,r){var f=this,e=f.promise;e._state===s&&(f._remaining--,n===h?i(e,r):f._result[t]=r);f._remaining===0&&u(e,f._result)};f.prototype._willSettleAt=function(n,t){var i=this;w(n,undefined,function(n){i._settledAt(v,t,n)},function(n){i._settledAt(h,t,n)})};lt=bi;at=ki;vt=di;yt=gi;pt=0;wt=r;r.all=lt;r.race=at;r.resolve=vt;r.reject=yt;r._setScheduler=dt;r._setAsap=gt;r._asap=e;r.prototype={constructor:r,then:function(n,t){var u=this,i=u._state,r,f,o;return i===v&&!n||i===h&&!t?this:(r=new this.constructor(a),f=u._result,i?(o=arguments[i-1],e(function(){ht(i,r,o,f)})):w(u,r,n,t),r)},"catch":function(n){return this.then(null,n)}};n.Promise=wt}).call(this)}i.Init=r})(i=t.PromiseImpl||(t.PromiseImpl={}))})(t=n._Internal||(n._Internal={}));n.Promise||(window.Promise?n.Promise=window.Promise:t.PromiseImpl.Init())}(OfficeExtension||(OfficeExtension={})),function(n){(function(n){n[n.Default=0]="Default";n[n.Read=1]="Read"})(n.OperationType||(n.OperationType={}));var t=n.OperationType}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(n){this._autoCleanupList={};this.m_context=n}return t.prototype.add=function(n){var t=this;Array.isArray(n)?n.forEach(function(n){return t._addCommon(n,!0)}):this._addCommon(n,!0)},t.prototype._autoAdd=function(n){this._addCommon(n,!1);this._autoCleanupList[n._objectPath.objectPathInfo.Id]=n},t.prototype._addCommon=function(t,i){var r=t[n.Constants.referenceId];n.Utility.isNullOrEmptyString(r)&&t._KeepReference&&(t._KeepReference(),n.ActionFactory.createInstantiateAction(this.m_context,t),i&&this.m_context._autoCleanup&&delete this._autoCleanupList[t._objectPath.objectPathInfo.Id])},t.prototype.remove=function(n){var t=this;Array.isArray(n)?n.forEach(function(n){return t._removeCommon(n)}):this._removeCommon(n)},t.prototype._removeCommon=function(t){var r=t[n.Constants.referenceId],i;n.Utility.isNullOrEmptyString(r)||(i=this.m_context._rootObject,i._RemoveReference&&i._RemoveReference(r))},t.prototype._retrieveAndClearAutoCleanupList=function(){var n=this._autoCleanupList;return this._autoCleanupList={},n},t}();n.TrackedObjects=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function n(){}return n.invalidObjectPath="InvalidObjectPath",n.propertyNotLoaded="PropertyNotLoaded",n.invalidRequestContext="InvalidRequestContext",n.invalidArgument="InvalidArgument",n.runMustReturnPromise="RunMustReturnPromise",n}();n.ResourceStrings=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.buildMessageArrayForIRequestExecutor=function(i,r,u,f){var o=JSON.stringify(u.Body),e;return n.Utility.log("Request:"),n.Utility.log(o),e={},e[n.Constants.sourceLibHeader]=f,t.buildRequestMessageSafeArray(i,r,"POST","ProcessQuery",e,o)},t.sendResponseOnSuccess=function(n,t,i){var r={ErrorCode:"",ErrorMessage:"",Headers:null,Body:null};r.Body=JSON.parse(n);r.Headers=t;i(r)},t.sendResponseOnError=function(i,r,u){var f={ErrorCode:"",ErrorMessage:"",Headers:null,Body:null};f.ErrorCode=n.ErrorCodes.generalException;i==t.OfficeJsErrorCode_ooeNoCapability?f.ErrorCode=n.ErrorCodes.accessDenied:i==t.OfficeJsErrorCode_ooeActivityLimitReached&&(f.ErrorCode=n.ErrorCodes.activityLimitReached);f.ErrorMessage=r;u(f)},t.buildRequestMessageSafeArray=function(n,t,i,r,u,f){var e=[],o;if(u)for(o in u)e.push(o),e.push(u[o]);var s=0,h="",c="",l="";return[n,i,r,e,f,s,t,h,c,l]},t.getResponseBody=function(n){return t.getResponseBodyFromSafeArray(n.value.data)},t.getResponseHeaders=function(n){return t.getResponseHeadersFromSafeArray(n.value.data)},t.getResponseBodyFromSafeArray=function(n){var t=n[2],i;return typeof t=="string"?t:(i=t,i.join(""))},t.getResponseHeadersFromSafeArray=function(n){var i=n[1],r,t;if(!i)return null;for(r={},t=0;t<i.length-1;t+=2)r[i[t]]=i[t+1];return r},t.getResponseStatusCode=function(n){return t.getResponseStatusCodeFromSafeArray(n.value.data)},t.getResponseStatusCodeFromSafeArray=function(n){return n[0]},t.OfficeJsErrorCode_ooeNoCapability=7e3,t.OfficeJsErrorCode_ooeActivityLimitReached=5102,t}();n.RichApiMessageUtility=t}(OfficeExtension||(OfficeExtension={})),function(n){var t=function(){function t(){}return t.checkArgumentNull=function(i,r){t.isNullOrUndefined(i)&&t.throwError(n.ResourceStrings.invalidArgument,r)},t.isNullOrUndefined=function(n){return n===null?!0:typeof n=="undefined"?!0:!1},t.isUndefined=function(n){return typeof n=="undefined"?!0:!1},t.isNullOrEmptyString=function(n){return n===null?!0:typeof n=="undefined"?!0:n.length==0?!0:!1},t.trim=function(n){return n.replace(new RegExp("^\\s+|\\s+$","g"),"")},t.caseInsensitiveCompareString=function(n,i){return t.isNullOrUndefined(n)?t.isNullOrUndefined(i):t.isNullOrUndefined(i)?!1:n.toUpperCase()==i.toUpperCase()},t.isReadonlyRestRequest=function(n){return t.caseInsensitiveCompareString(n,"GET")},t.setMethodArguments=function(n,i,r){if(t.isNullOrUndefined(r))return null;var u=[],f=[],e=t.collectObjectPathInfos(n,r,u,f);return(i.Arguments=r,e)?(i.ReferencedObjectPathIds=f,u):null},t.collectObjectPathInfos=function(i,r,u,f){for(var o,h,c,s=!1,e=0;e<r.length;e++)r[e]instanceof n.ClientObject?(o=r[e],t.validateContext(i,o),r[e]=o._objectPath.objectPathInfo.Id,f.push(o._objectPath.objectPathInfo.Id),u.push(o._objectPath),s=!0):Array.isArray(r[e])?(h=[],c=t.collectObjectPathInfos(i,r[e],u,h),c?(f.push(h),s=!0):f.push(0)):f.push(0);return s},t.fixObjectPathIfNecessary=function(n,t){n&&n._objectPath&&t&&n._objectPath.updateUsingObjectData(t)},t.validateObjectPath=function(i){for(var r=i._objectPath,u;r;)r.isValid||(u=t.getObjectPathExpression(r),t.throwError(n.ResourceStrings.invalidObjectPath,u)),r=r.parentObjectPath},t.validateReferencedObjectPaths=function(i){var u,r,f;if(i)for(u=0;u<i.length;u++)for(r=i[u];r;)r.isValid||(f=t.getObjectPathExpression(r),t.throwError(n.ResourceStrings.invalidObjectPath,f)),r=r.parentObjectPath},t.validateContext=function(i,r){r&&r.context!==i&&t.throwError(n.ResourceStrings.invalidRequestContext)},t.log=function(n){t._logEnabled&&window.console&&window.console.log&&window.console.log(n)},t.load=function(n,t){n.context.load(n,t)},t.throwError=function(i,r,u){throw new n._Internal.RuntimeError(i,t._getResourceString(i,r),[],u?{errorLocation:u}:{});},t.createRuntimeError=function(t,i,r){return new n._Internal.RuntimeError(t,i,[],{errorLocation:r})},t._getResourceString=function(n,i){var r=n,f,u;return window.Strings&&window.Strings.OfficeOM&&(f="L_"+n,u=window.Strings.OfficeOM[f],u&&(r=u)),t.isNullOrUndefined(i)||(r=r.replace("{0}",i)),r},t.throwIfNotLoaded=function(i,r,u,f){!f&&t.isUndefined(r)&&i.charCodeAt(0)!=t.s_underscoreCharCode&&t.throwError(n.ResourceStrings.propertyNotLoaded,i,u?u+"."+i:null)},t.getObjectPathExpression=function(n){for(var i="";n;){switch(n.objectPathInfo.ObjectPathType){case 1:i=i;break;case 2:i="new()"+(i.length>0?".":"")+i;break;case 3:i=t.normalizeName(n.objectPathInfo.Name)+"()"+(i.length>0?".":"")+i;break;case 4:i=t.normalizeName(n.objectPathInfo.Name)+(i.length>0?".":"")+i;break;case 5:i="getItem()"+(i.length>0?".":"")+i;break;case 6:i="_reference()"+(i.length>0?".":"")+i;break}n=n.parentObjectPath}return i},t._createPromiseFromResult=function(t){return new n.Promise(function(n,i){n(t)})},t._addActionResultHandler=function(n,t,i){n.context._pendingRequest.addActionResultHandler(t,i)},t._handleNavigationPropertyResults=function(n,i,r){for(var u=0;u<r.length-1;u+=2)t.isUndefined(i[r[u+1]])||n[r[u]]._handleResult(i[r[u+1]])},t.normalizeName=function(n){return n.substr(0,1).toLowerCase()+n.substr(1)},t._logEnabled=!1,t.s_underscoreCharCode="_".charCodeAt(0),t}();n.Utility=t}(OfficeExtension||(OfficeExtension={}));__extends=this.__extends||function(n,t){function r(){this.constructor=n}for(var i in t)t.hasOwnProperty(i)&&(n[i]=t[i]);r.prototype=t.prototype;n.prototype=new r},function(n){var e=OfficeExtension.ObjectPathFactory.createPropertyObjectPath,u=OfficeExtension.ObjectPathFactory.createMethodObjectPath,a=OfficeExtension.ObjectPathFactory.createIndexerObjectPath,et=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath,bt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer,kt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt,v=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt,f=OfficeExtension.ActionFactory.createMethodAction,r=OfficeExtension.ActionFactory.createSetPropertyAction,o=OfficeExtension.Utility.isNullOrUndefined,i=OfficeExtension.Utility.isUndefined,t=OfficeExtension.Utility.throwIfNotLoaded,s=OfficeExtension.Utility.load,h=OfficeExtension.Utility.fixObjectPathIfNecessary,c=OfficeExtension.Utility._addActionResultHandler,l=OfficeExtension.Utility._handleNavigationPropertyResults,ot=function(a){function v(){a.apply(this,arguments)}return __extends(v,a),Object.defineProperty(v.prototype,"contentControls",{get:function(){return this.m_contentControls||(this.m_contentControls=new n.ContentControlCollection(this.context,e(this.context,this,"ContentControls",!0,!1))),this.m_contentControls},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"font",{get:function(){return this.m_font||(this.m_font=new n.Font(this.context,e(this.context,this,"Font",!1,!1))),this.m_font},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"inlinePictures",{get:function(){return this.m_inlinePictures||(this.m_inlinePictures=new n.InlinePictureCollection(this.context,e(this.context,this,"InlinePictures",!0,!1))),this.m_inlinePictures},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"paragraphs",{get:function(){return this.m_paragraphs||(this.m_paragraphs=new n.ParagraphCollection(this.context,e(this.context,this,"Paragraphs",!0,!1))),this.m_paragraphs},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"parentContentControl",{get:function(){return this.m_parentContentControl||(this.m_parentContentControl=new n.ContentControl(this.context,e(this.context,this,"ParentContentControl",!1,!1))),this.m_parentContentControl},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"style",{get:function(){return t("style",this.m_style,"Body",this._isNull),this.m_style},set:function(n){this.m_style=n;r(this.context,this,"Style",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"text",{get:function(){return t("text",this.m_text,"Body",this._isNull),this.m_text},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"Body",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),v.prototype.clear=function(){f(this.context,this,"Clear",0,[])},v.prototype.getHtml=function(){var t=f(this.context,this,"GetHtml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.getOoxml=function(){var t=f(this.context,this,"GetOoxml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.insertBreak=function(n,t){f(this.context,this,"InsertBreak",0,[n,t])},v.prototype.insertContentControl=function(){return new n.ContentControl(this.context,u(this.context,this,"InsertContentControl",0,[],!1,!0))},v.prototype.insertFileFromBase64=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertFileFromBase64",0,[t,i],!1,!0))},v.prototype.insertHtml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertHtml",0,[t,i],!1,!0))},v.prototype.insertInlinePictureFromBase64=function(t,i){return new n.InlinePicture(this.context,u(this.context,this,"InsertInlinePictureFromBase64",0,[t,i],!1,!0))},v.prototype.insertOoxml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertOoxml",0,[t,i],!1,!0))},v.prototype.insertParagraph=function(t,i){return new n.Paragraph(this.context,u(this.context,this,"InsertParagraph",0,[t,i],!1,!0))},v.prototype.insertText=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertText",0,[t,i],!1,!0))},v.prototype.search=function(t,i){return new n.SearchResultCollection(this.context,u(this.context,this,"Search",1,[t,i],!0,!0))},v.prototype.select=function(n){f(this.context,this,"Select",1,[n])},v.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},v.prototype._handleResult=function(n){if(a.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.Style)||(this.m_style=t.Style);i(t.Text)||(this.m_text=t.Text);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["contentControls","ContentControls","font","Font","inlinePictures","InlinePictures","paragraphs","Paragraphs","parentContentControl","ParentContentControl"])}},v.prototype.load=function(n){return s(this,n),this},v.prototype._initReferenceId=function(n){this.m__ReferenceId=n},v}(OfficeExtension.ClientObject),y,p,w,b,k,d,g,nt,tt,it,rt,ut,ft,st,ht,ct,lt,at,vt,yt,pt,wt;n.Body=ot;y=function(a){function v(){a.apply(this,arguments)}return __extends(v,a),Object.defineProperty(v.prototype,"contentControls",{get:function(){return this.m_contentControls||(this.m_contentControls=new n.ContentControlCollection(this.context,e(this.context,this,"ContentControls",!0,!1))),this.m_contentControls},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"font",{get:function(){return this.m_font||(this.m_font=new n.Font(this.context,e(this.context,this,"Font",!1,!1))),this.m_font},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"inlinePictures",{get:function(){return this.m_inlinePictures||(this.m_inlinePictures=new n.InlinePictureCollection(this.context,e(this.context,this,"InlinePictures",!0,!1))),this.m_inlinePictures},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"paragraphs",{get:function(){return this.m_paragraphs||(this.m_paragraphs=new n.ParagraphCollection(this.context,e(this.context,this,"Paragraphs",!0,!1))),this.m_paragraphs},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"parentContentControl",{get:function(){return this.m_parentContentControl||(this.m_parentContentControl=new n.ContentControl(this.context,e(this.context,this,"ParentContentControl",!1,!1))),this.m_parentContentControl},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"appearance",{get:function(){return t("appearance",this.m_appearance,"ContentControl",this._isNull),this.m_appearance},set:function(n){this.m_appearance=n;r(this.context,this,"Appearance",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"cannotDelete",{get:function(){return t("cannotDelete",this.m_cannotDelete,"ContentControl",this._isNull),this.m_cannotDelete},set:function(n){this.m_cannotDelete=n;r(this.context,this,"CannotDelete",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"cannotEdit",{get:function(){return t("cannotEdit",this.m_cannotEdit,"ContentControl",this._isNull),this.m_cannotEdit},set:function(n){this.m_cannotEdit=n;r(this.context,this,"CannotEdit",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"color",{get:function(){return t("color",this.m_color,"ContentControl",this._isNull),this.m_color},set:function(n){this.m_color=n;r(this.context,this,"Color",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"id",{get:function(){return t("id",this.m_id,"ContentControl",this._isNull),this.m_id},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"placeholderText",{get:function(){return t("placeholderText",this.m_placeholderText,"ContentControl",this._isNull),this.m_placeholderText},set:function(n){this.m_placeholderText=n;r(this.context,this,"PlaceholderText",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"removeWhenEdited",{get:function(){return t("removeWhenEdited",this.m_removeWhenEdited,"ContentControl",this._isNull),this.m_removeWhenEdited},set:function(n){this.m_removeWhenEdited=n;r(this.context,this,"RemoveWhenEdited",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"style",{get:function(){return t("style",this.m_style,"ContentControl",this._isNull),this.m_style},set:function(n){this.m_style=n;r(this.context,this,"Style",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"tag",{get:function(){return t("tag",this.m_tag,"ContentControl",this._isNull),this.m_tag},set:function(n){this.m_tag=n;r(this.context,this,"Tag",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"text",{get:function(){return t("text",this.m_text,"ContentControl",this._isNull),this.m_text},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"title",{get:function(){return t("title",this.m_title,"ContentControl",this._isNull),this.m_title},set:function(n){this.m_title=n;r(this.context,this,"Title",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"type",{get:function(){return t("type",this.m_type,"ContentControl",this._isNull),this.m_type},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"ContentControl",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),v.prototype.clear=function(){f(this.context,this,"Clear",0,[])},v.prototype.delete=function(n){f(this.context,this,"Delete",0,[n])},v.prototype.getHtml=function(){var t=f(this.context,this,"GetHtml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.getOoxml=function(){var t=f(this.context,this,"GetOoxml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.insertBreak=function(n,t){f(this.context,this,"InsertBreak",0,[n,t])},v.prototype.insertFileFromBase64=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertFileFromBase64",0,[t,i],!1,!0))},v.prototype.insertHtml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertHtml",0,[t,i],!1,!0))},v.prototype.insertInlinePictureFromBase64=function(t,i){return new n.InlinePicture(this.context,u(this.context,this,"InsertInlinePictureFromBase64",0,[t,i],!1,!0))},v.prototype.insertOoxml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertOoxml",0,[t,i],!1,!0))},v.prototype.insertParagraph=function(t,i){return new n.Paragraph(this.context,u(this.context,this,"InsertParagraph",0,[t,i],!1,!0))},v.prototype.insertText=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertText",0,[t,i],!1,!0))},v.prototype.search=function(t,i){return new n.SearchResultCollection(this.context,u(this.context,this,"Search",1,[t,i],!0,!0))},v.prototype.select=function(n){f(this.context,this,"Select",1,[n])},v.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},v.prototype._handleResult=function(n){if(a.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.Appearance)||(this.m_appearance=t.Appearance);i(t.CannotDelete)||(this.m_cannotDelete=t.CannotDelete);i(t.CannotEdit)||(this.m_cannotEdit=t.CannotEdit);i(t.Color)||(this.m_color=t.Color);i(t.Id)||(this.m_id=t.Id);i(t.PlaceholderText)||(this.m_placeholderText=t.PlaceholderText);i(t.RemoveWhenEdited)||(this.m_removeWhenEdited=t.RemoveWhenEdited);i(t.Style)||(this.m_style=t.Style);i(t.Tag)||(this.m_tag=t.Tag);i(t.Text)||(this.m_text=t.Text);i(t.Title)||(this.m_title=t.Title);i(t.Type)||(this.m_type=t.Type);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["contentControls","ContentControls","font","Font","inlinePictures","InlinePictures","paragraphs","Paragraphs","parentContentControl","ParentContentControl"])}},v.prototype.load=function(n){return s(this,n),this},v.prototype._initReferenceId=function(n){this.m__ReferenceId=n},v}(OfficeExtension.ClientObject);n.ContentControl=y;p=function(r){function e(){r.apply(this,arguments)}return __extends(e,r),Object.defineProperty(e.prototype,"items",{get:function(){return t("items",this.m__items,"ContentControlCollection",this._isNull),this.m__items},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"ContentControlCollection",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),e.prototype.getById=function(t){return new n.ContentControl(this.context,u(this.context,this,"GetById",1,[t],!1,!1))},e.prototype.getByTag=function(t){return new n.ContentControlCollection(this.context,u(this.context,this,"GetByTag",1,[t],!0,!1))},e.prototype.getByTitle=function(t){return new n.ContentControlCollection(this.context,u(this.context,this,"GetByTitle",1,[t],!0,!1))},e.prototype.getItem=function(t){return new n.ContentControl(this.context,a(this.context,this,[t]))},e.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},e.prototype._handleResult=function(t){var u,e,f,s;if((r.prototype._handleResult.call(this,t),!o(t))&&(u=t,h(this,u),i(u._ReferenceId)||(this.m__ReferenceId=u._ReferenceId),!o(u[OfficeExtension.Constants.items])))for(this.m__items=[],e=u[OfficeExtension.Constants.items],f=0;f<e.length;f++)s=new n.ContentControl(this.context,v(!0,this.context,this,e[f],f)),s._handleResult(e[f]),this.m__items.push(s)},e.prototype.load=function(n){return s(this,n),this},e.prototype._initReferenceId=function(n){this.m__ReferenceId=n},e}(OfficeExtension.ClientObject);n.ContentControlCollection=p;w=function(r){function a(){r.apply(this,arguments)}return __extends(a,r),Object.defineProperty(a.prototype,"body",{get:function(){return this.m_body||(this.m_body=new n.Body(this.context,e(this.context,this,"Body",!1,!1))),this.m_body},enumerable:!0,configurable:!0}),Object.defineProperty(a.prototype,"contentControls",{get:function(){return this.m_contentControls||(this.m_contentControls=new n.ContentControlCollection(this.context,e(this.context,this,"ContentControls",!0,!1))),this.m_contentControls},enumerable:!0,configurable:!0}),Object.defineProperty(a.prototype,"sections",{get:function(){return this.m_sections||(this.m_sections=new n.SectionCollection(this.context,e(this.context,this,"Sections",!0,!1))),this.m_sections},enumerable:!0,configurable:!0}),Object.defineProperty(a.prototype,"saved",{get:function(){return t("saved",this.m_saved,"Document",this._isNull),this.m_saved},enumerable:!0,configurable:!0}),a.prototype.getSelection=function(){return new n.Range(this.context,u(this.context,this,"GetSelection",1,[],!1,!0))},a.prototype.save=function(){f(this.context,this,"Save",0,[])},a.prototype._GetObjectByReferenceId=function(n){var i=f(this.context,this,"_GetObjectByReferenceId",1,[n]),t=new OfficeExtension.ClientResult;return c(this,i,t),t},a.prototype._GetObjectTypeNameByReferenceId=function(n){var i=f(this.context,this,"_GetObjectTypeNameByReferenceId",1,[n]),t=new OfficeExtension.ClientResult;return c(this,i,t),t},a.prototype._RemoveAllReferences=function(){f(this.context,this,"_RemoveAllReferences",1,[])},a.prototype._RemoveReference=function(n){f(this.context,this,"_RemoveReference",1,[n])},a.prototype._handleResult=function(n){if(r.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.Saved)||(this.m_saved=t.Saved);l(this,t,["body","Body","contentControls","ContentControls","sections","Sections"])}},a.prototype.load=function(n){return s(this,n),this},a}(OfficeExtension.ClientObject);n.Document=w;b=function(n){function u(){n.apply(this,arguments)}return __extends(u,n),Object.defineProperty(u.prototype,"bold",{get:function(){return t("bold",this.m_bold,"Font",this._isNull),this.m_bold},set:function(n){this.m_bold=n;r(this.context,this,"Bold",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"color",{get:function(){return t("color",this.m_color,"Font",this._isNull),this.m_color},set:function(n){this.m_color=n;r(this.context,this,"Color",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"doubleStrikeThrough",{get:function(){return t("doubleStrikeThrough",this.m_doubleStrikeThrough,"Font",this._isNull),this.m_doubleStrikeThrough},set:function(n){this.m_doubleStrikeThrough=n;r(this.context,this,"DoubleStrikeThrough",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"highlightColor",{get:function(){return t("highlightColor",this.m_highlightColor,"Font",this._isNull),this.m_highlightColor},set:function(n){this.m_highlightColor=n;r(this.context,this,"HighlightColor",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"italic",{get:function(){return t("italic",this.m_italic,"Font",this._isNull),this.m_italic},set:function(n){this.m_italic=n;r(this.context,this,"Italic",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"name",{get:function(){return t("name",this.m_name,"Font",this._isNull),this.m_name},set:function(n){this.m_name=n;r(this.context,this,"Name",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"size",{get:function(){return t("size",this.m_size,"Font",this._isNull),this.m_size},set:function(n){this.m_size=n;r(this.context,this,"Size",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"strikeThrough",{get:function(){return t("strikeThrough",this.m_strikeThrough,"Font",this._isNull),this.m_strikeThrough},set:function(n){this.m_strikeThrough=n;r(this.context,this,"StrikeThrough",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"subscript",{get:function(){return t("subscript",this.m_subscript,"Font",this._isNull),this.m_subscript},set:function(n){this.m_subscript=n;r(this.context,this,"Subscript",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"superscript",{get:function(){return t("superscript",this.m_superscript,"Font",this._isNull),this.m_superscript},set:function(n){this.m_superscript=n;r(this.context,this,"Superscript",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"underline",{get:function(){return t("underline",this.m_underline,"Font",this._isNull),this.m_underline},set:function(n){this.m_underline=n;r(this.context,this,"Underline",n)},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"Font",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),u.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},u.prototype._handleResult=function(t){if(n.prototype._handleResult.call(this,t),!o(t)){var r=t;h(this,r);i(r.Bold)||(this.m_bold=r.Bold);i(r.Color)||(this.m_color=r.Color);i(r.DoubleStrikeThrough)||(this.m_doubleStrikeThrough=r.DoubleStrikeThrough);i(r.HighlightColor)||(this.m_highlightColor=r.HighlightColor);i(r.Italic)||(this.m_italic=r.Italic);i(r.Name)||(this.m_name=r.Name);i(r.Size)||(this.m_size=r.Size);i(r.StrikeThrough)||(this.m_strikeThrough=r.StrikeThrough);i(r.Subscript)||(this.m_subscript=r.Subscript);i(r.Superscript)||(this.m_superscript=r.Superscript);i(r.Underline)||(this.m_underline=r.Underline);i(r._ReferenceId)||(this.m__ReferenceId=r._ReferenceId)}},u.prototype.load=function(n){return s(this,n),this},u.prototype._initReferenceId=function(n){this.m__ReferenceId=n},u}(OfficeExtension.ClientObject);n.Font=b;k=function(a){function v(){a.apply(this,arguments)}return __extends(v,a),Object.defineProperty(v.prototype,"paragraph",{get:function(){return this.m_paragraph||(this.m_paragraph=new n.Paragraph(this.context,e(this.context,this,"Paragraph",!1,!1))),this.m_paragraph},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"parentContentControl",{get:function(){return this.m_parentContentControl||(this.m_parentContentControl=new n.ContentControl(this.context,e(this.context,this,"ParentContentControl",!1,!1))),this.m_parentContentControl},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"altTextDescription",{get:function(){return t("altTextDescription",this.m_altTextDescription,"InlinePicture",this._isNull),this.m_altTextDescription},set:function(n){this.m_altTextDescription=n;r(this.context,this,"AltTextDescription",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"altTextTitle",{get:function(){return t("altTextTitle",this.m_altTextTitle,"InlinePicture",this._isNull),this.m_altTextTitle},set:function(n){this.m_altTextTitle=n;r(this.context,this,"AltTextTitle",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"height",{get:function(){return t("height",this.m_height,"InlinePicture",this._isNull),this.m_height},set:function(n){this.m_height=n;r(this.context,this,"Height",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"hyperlink",{get:function(){return t("hyperlink",this.m_hyperlink,"InlinePicture",this._isNull),this.m_hyperlink},set:function(n){this.m_hyperlink=n;r(this.context,this,"Hyperlink",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"lockAspectRatio",{get:function(){return t("lockAspectRatio",this.m_lockAspectRatio,"InlinePicture",this._isNull),this.m_lockAspectRatio},set:function(n){this.m_lockAspectRatio=n;r(this.context,this,"LockAspectRatio",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"width",{get:function(){return t("width",this.m_width,"InlinePicture",this._isNull),this.m_width},set:function(n){this.m_width=n;r(this.context,this,"Width",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_Id",{get:function(){return t("_Id",this.m__Id,"InlinePicture",this._isNull),this.m__Id},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"InlinePicture",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),v.prototype.delete=function(){f(this.context,this,"Delete",0,[])},v.prototype.getBase64ImageSrc=function(){var t=f(this.context,this,"GetBase64ImageSrc",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.insertBreak=function(n,t){f(this.context,this,"InsertBreak",0,[n,t])},v.prototype.insertContentControl=function(){return new n.ContentControl(this.context,u(this.context,this,"InsertContentControl",0,[],!1,!0))},v.prototype.insertFileFromBase64=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertFileFromBase64",0,[t,i],!1,!0))},v.prototype.insertHtml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertHtml",0,[t,i],!1,!0))},v.prototype.insertInlinePictureFromBase64=function(t,i){return new n.InlinePicture(this.context,u(this.context,this,"InsertInlinePictureFromBase64",0,[t,i],!1,!0))},v.prototype.insertOoxml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertOoxml",0,[t,i],!1,!0))},v.prototype.insertParagraph=function(t,i){return new n.Paragraph(this.context,u(this.context,this,"InsertParagraph",0,[t,i],!1,!0))},v.prototype.insertText=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertText",0,[t,i],!1,!0))},v.prototype.select=function(n){f(this.context,this,"Select",1,[n])},v.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},v.prototype._handleResult=function(n){if(a.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.AltTextDescription)||(this.m_altTextDescription=t.AltTextDescription);i(t.AltTextTitle)||(this.m_altTextTitle=t.AltTextTitle);i(t.Height)||(this.m_height=t.Height);i(t.Hyperlink)||(this.m_hyperlink=t.Hyperlink);i(t.LockAspectRatio)||(this.m_lockAspectRatio=t.LockAspectRatio);i(t.Width)||(this.m_width=t.Width);i(t._Id)||(this.m__Id=t._Id);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["paragraph","Paragraph","parentContentControl","ParentContentControl"])}},v.prototype.load=function(n){return s(this,n),this},v.prototype._initReferenceId=function(n){this.m__ReferenceId=n},v}(OfficeExtension.ClientObject);n.InlinePicture=k;d=function(r){function u(){r.apply(this,arguments)}return __extends(u,r),Object.defineProperty(u.prototype,"items",{get:function(){return t("items",this.m__items,"InlinePictureCollection",this._isNull),this.m__items},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"InlinePictureCollection",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),u.prototype._GetItem=function(t){return new n.InlinePicture(this.context,a(this.context,this,[t]))},u.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},u.prototype._handleResult=function(t){var u,e,f,s;if((r.prototype._handleResult.call(this,t),!o(t))&&(u=t,h(this,u),i(u._ReferenceId)||(this.m__ReferenceId=u._ReferenceId),!o(u[OfficeExtension.Constants.items])))for(this.m__items=[],e=u[OfficeExtension.Constants.items],f=0;f<e.length;f++)s=new n.InlinePicture(this.context,v(!0,this.context,this,e[f],f)),s._handleResult(e[f]),this.m__items.push(s)},u.prototype.load=function(n){return s(this,n),this},u.prototype._initReferenceId=function(n){this.m__ReferenceId=n},u}(OfficeExtension.ClientObject);n.InlinePictureCollection=d;g=function(a){function v(){a.apply(this,arguments)}return __extends(v,a),Object.defineProperty(v.prototype,"contentControls",{get:function(){return this.m_contentControls||(this.m_contentControls=new n.ContentControlCollection(this.context,e(this.context,this,"ContentControls",!0,!1))),this.m_contentControls},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"font",{get:function(){return this.m_font||(this.m_font=new n.Font(this.context,e(this.context,this,"Font",!1,!1))),this.m_font},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"inlinePictures",{get:function(){return this.m_inlinePictures||(this.m_inlinePictures=new n.InlinePictureCollection(this.context,e(this.context,this,"InlinePictures",!0,!1))),this.m_inlinePictures},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"parentContentControl",{get:function(){return this.m_parentContentControl||(this.m_parentContentControl=new n.ContentControl(this.context,e(this.context,this,"ParentContentControl",!1,!1))),this.m_parentContentControl},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"alignment",{get:function(){return t("alignment",this.m_alignment,"Paragraph",this._isNull),this.m_alignment},set:function(n){this.m_alignment=n;r(this.context,this,"Alignment",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"firstLineIndent",{get:function(){return t("firstLineIndent",this.m_firstLineIndent,"Paragraph",this._isNull),this.m_firstLineIndent},set:function(n){this.m_firstLineIndent=n;r(this.context,this,"FirstLineIndent",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"leftIndent",{get:function(){return t("leftIndent",this.m_leftIndent,"Paragraph",this._isNull),this.m_leftIndent},set:function(n){this.m_leftIndent=n;r(this.context,this,"LeftIndent",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"lineSpacing",{get:function(){return t("lineSpacing",this.m_lineSpacing,"Paragraph",this._isNull),this.m_lineSpacing},set:function(n){this.m_lineSpacing=n;r(this.context,this,"LineSpacing",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"lineUnitAfter",{get:function(){return t("lineUnitAfter",this.m_lineUnitAfter,"Paragraph",this._isNull),this.m_lineUnitAfter},set:function(n){this.m_lineUnitAfter=n;r(this.context,this,"LineUnitAfter",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"lineUnitBefore",{get:function(){return t("lineUnitBefore",this.m_lineUnitBefore,"Paragraph",this._isNull),this.m_lineUnitBefore},set:function(n){this.m_lineUnitBefore=n;r(this.context,this,"LineUnitBefore",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"outlineLevel",{get:function(){return t("outlineLevel",this.m_outlineLevel,"Paragraph",this._isNull),this.m_outlineLevel},set:function(n){this.m_outlineLevel=n;r(this.context,this,"OutlineLevel",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"rightIndent",{get:function(){return t("rightIndent",this.m_rightIndent,"Paragraph",this._isNull),this.m_rightIndent},set:function(n){this.m_rightIndent=n;r(this.context,this,"RightIndent",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"spaceAfter",{get:function(){return t("spaceAfter",this.m_spaceAfter,"Paragraph",this._isNull),this.m_spaceAfter},set:function(n){this.m_spaceAfter=n;r(this.context,this,"SpaceAfter",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"spaceBefore",{get:function(){return t("spaceBefore",this.m_spaceBefore,"Paragraph",this._isNull),this.m_spaceBefore},set:function(n){this.m_spaceBefore=n;r(this.context,this,"SpaceBefore",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"style",{get:function(){return t("style",this.m_style,"Paragraph",this._isNull),this.m_style},set:function(n){this.m_style=n;r(this.context,this,"Style",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"text",{get:function(){return t("text",this.m_text,"Paragraph",this._isNull),this.m_text},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_Id",{get:function(){return t("_Id",this.m__Id,"Paragraph",this._isNull),this.m__Id},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"Paragraph",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),v.prototype.clear=function(){f(this.context,this,"Clear",0,[])},v.prototype.delete=function(){f(this.context,this,"Delete",0,[])},v.prototype.getHtml=function(){var t=f(this.context,this,"GetHtml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.getOoxml=function(){var t=f(this.context,this,"GetOoxml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.insertBreak=function(n,t){f(this.context,this,"InsertBreak",0,[n,t])},v.prototype.insertContentControl=function(){return new n.ContentControl(this.context,u(this.context,this,"InsertContentControl",0,[],!1,!0))},v.prototype.insertFileFromBase64=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertFileFromBase64",0,[t,i],!1,!0))},v.prototype.insertHtml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertHtml",0,[t,i],!1,!0))},v.prototype.insertInlinePictureFromBase64=function(t,i){return new n.InlinePicture(this.context,u(this.context,this,"InsertInlinePictureFromBase64",0,[t,i],!1,!0))},v.prototype.insertOoxml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertOoxml",0,[t,i],!1,!0))},v.prototype.insertParagraph=function(t,i){return new n.Paragraph(this.context,u(this.context,this,"InsertParagraph",0,[t,i],!1,!0))},v.prototype.insertText=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertText",0,[t,i],!1,!0))},v.prototype.search=function(t,i){return new n.SearchResultCollection(this.context,u(this.context,this,"Search",1,[t,i],!0,!0))},v.prototype.select=function(n){f(this.context,this,"Select",1,[n])},v.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},v.prototype._handleResult=function(n){if(a.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.Alignment)||(this.m_alignment=t.Alignment);i(t.FirstLineIndent)||(this.m_firstLineIndent=t.FirstLineIndent);i(t.LeftIndent)||(this.m_leftIndent=t.LeftIndent);i(t.LineSpacing)||(this.m_lineSpacing=t.LineSpacing);i(t.LineUnitAfter)||(this.m_lineUnitAfter=t.LineUnitAfter);i(t.LineUnitBefore)||(this.m_lineUnitBefore=t.LineUnitBefore);i(t.OutlineLevel)||(this.m_outlineLevel=t.OutlineLevel);i(t.RightIndent)||(this.m_rightIndent=t.RightIndent);i(t.SpaceAfter)||(this.m_spaceAfter=t.SpaceAfter);i(t.SpaceBefore)||(this.m_spaceBefore=t.SpaceBefore);i(t.Style)||(this.m_style=t.Style);i(t.Text)||(this.m_text=t.Text);i(t._Id)||(this.m__Id=t._Id);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["contentControls","ContentControls","font","Font","inlinePictures","InlinePictures","parentContentControl","ParentContentControl"])}},v.prototype.load=function(n){return s(this,n),this},v.prototype._initReferenceId=function(n){this.m__ReferenceId=n},v}(OfficeExtension.ClientObject);n.Paragraph=g;nt=function(r){function u(){r.apply(this,arguments)}return __extends(u,r),Object.defineProperty(u.prototype,"items",{get:function(){return t("items",this.m__items,"ParagraphCollection",this._isNull),this.m__items},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"ParagraphCollection",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),u.prototype._GetItem=function(t){return new n.Paragraph(this.context,a(this.context,this,[t]))},u.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},u.prototype._handleResult=function(t){var u,e,f,s;if((r.prototype._handleResult.call(this,t),!o(t))&&(u=t,h(this,u),i(u._ReferenceId)||(this.m__ReferenceId=u._ReferenceId),!o(u[OfficeExtension.Constants.items])))for(this.m__items=[],e=u[OfficeExtension.Constants.items],f=0;f<e.length;f++)s=new n.Paragraph(this.context,v(!0,this.context,this,e[f],f)),s._handleResult(e[f]),this.m__items.push(s)},u.prototype.load=function(n){return s(this,n),this},u.prototype._initReferenceId=function(n){this.m__ReferenceId=n},u}(OfficeExtension.ClientObject);n.ParagraphCollection=nt;tt=function(a){function v(){a.apply(this,arguments)}return __extends(v,a),Object.defineProperty(v.prototype,"contentControls",{get:function(){return this.m_contentControls||(this.m_contentControls=new n.ContentControlCollection(this.context,e(this.context,this,"ContentControls",!0,!1))),this.m_contentControls},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"font",{get:function(){return this.m_font||(this.m_font=new n.Font(this.context,e(this.context,this,"Font",!1,!1))),this.m_font},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"inlinePictures",{get:function(){return this.m_inlinePictures||(this.m_inlinePictures=new n.InlinePictureCollection(this.context,e(this.context,this,"InlinePictures",!0,!1))),this.m_inlinePictures},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"paragraphs",{get:function(){return this.m_paragraphs||(this.m_paragraphs=new n.ParagraphCollection(this.context,e(this.context,this,"Paragraphs",!0,!1))),this.m_paragraphs},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"parentContentControl",{get:function(){return this.m_parentContentControl||(this.m_parentContentControl=new n.ContentControl(this.context,e(this.context,this,"ParentContentControl",!1,!1))),this.m_parentContentControl},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"style",{get:function(){return t("style",this.m_style,"Range",this._isNull),this.m_style},set:function(n){this.m_style=n;r(this.context,this,"Style",n)},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"text",{get:function(){return t("text",this.m_text,"Range",this._isNull),this.m_text},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_Id",{get:function(){return t("_Id",this.m__Id,"Range",this._isNull),this.m__Id},enumerable:!0,configurable:!0}),Object.defineProperty(v.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"Range",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),v.prototype.clear=function(){f(this.context,this,"Clear",0,[])},v.prototype.delete=function(){f(this.context,this,"Delete",0,[])},v.prototype.getHtml=function(){var t=f(this.context,this,"GetHtml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.getOoxml=function(){var t=f(this.context,this,"GetOoxml",1,[]),n=new OfficeExtension.ClientResult;return c(this,t,n),n},v.prototype.insertBreak=function(n,t){f(this.context,this,"InsertBreak",0,[n,t])},v.prototype.insertContentControl=function(){return new n.ContentControl(this.context,u(this.context,this,"InsertContentControl",0,[],!1,!0))},v.prototype.insertFileFromBase64=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertFileFromBase64",0,[t,i],!1,!0))},v.prototype.insertHtml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertHtml",0,[t,i],!1,!0))},v.prototype.insertInlinePictureFromBase64=function(t,i){return new n.InlinePicture(this.context,u(this.context,this,"InsertInlinePictureFromBase64",0,[t,i],!1,!0))},v.prototype.insertOoxml=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertOoxml",0,[t,i],!1,!0))},v.prototype.insertParagraph=function(t,i){return new n.Paragraph(this.context,u(this.context,this,"InsertParagraph",0,[t,i],!1,!0))},v.prototype.insertText=function(t,i){return new n.Range(this.context,u(this.context,this,"InsertText",0,[t,i],!1,!0))},v.prototype.search=function(t,i){return new n.SearchResultCollection(this.context,u(this.context,this,"Search",1,[t,i],!0,!0))},v.prototype.select=function(n){f(this.context,this,"Select",1,[n])},v.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},v.prototype._handleResult=function(n){if(a.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.Style)||(this.m_style=t.Style);i(t.Text)||(this.m_text=t.Text);i(t._Id)||(this.m__Id=t._Id);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["contentControls","ContentControls","font","Font","inlinePictures","InlinePictures","paragraphs","Paragraphs","parentContentControl","ParentContentControl"])}},v.prototype.load=function(n){return s(this,n),this},v.prototype._initReferenceId=function(n){this.m__ReferenceId=n},v}(OfficeExtension.ClientObject);n.Range=tt;it=function(u){function f(){u.apply(this,arguments)}return __extends(f,u),Object.defineProperty(f.prototype,"matchWildCards",{get:function(){return t("matchWildCards",this.m_matchWildcards),this.m_matchWildcards},set:function(n){this.m_matchWildcards=n;r(this.context,this,"MatchWildCards",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"ignorePunct",{get:function(){return t("ignorePunct",this.m_ignorePunct,"SearchOptions",this._isNull),this.m_ignorePunct},set:function(n){this.m_ignorePunct=n;r(this.context,this,"IgnorePunct",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"ignoreSpace",{get:function(){return t("ignoreSpace",this.m_ignoreSpace,"SearchOptions",this._isNull),this.m_ignoreSpace},set:function(n){this.m_ignoreSpace=n;r(this.context,this,"IgnoreSpace",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchCase",{get:function(){return t("matchCase",this.m_matchCase,"SearchOptions",this._isNull),this.m_matchCase},set:function(n){this.m_matchCase=n;r(this.context,this,"MatchCase",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchPrefix",{get:function(){return t("matchPrefix",this.m_matchPrefix,"SearchOptions",this._isNull),this.m_matchPrefix},set:function(n){this.m_matchPrefix=n;r(this.context,this,"MatchPrefix",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchSoundsLike",{get:function(){return t("matchSoundsLike",this.m_matchSoundsLike,"SearchOptions",this._isNull),this.m_matchSoundsLike},set:function(n){this.m_matchSoundsLike=n;r(this.context,this,"MatchSoundsLike",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchSuffix",{get:function(){return t("matchSuffix",this.m_matchSuffix,"SearchOptions",this._isNull),this.m_matchSuffix},set:function(n){this.m_matchSuffix=n;r(this.context,this,"MatchSuffix",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchWholeWord",{get:function(){return t("matchWholeWord",this.m_matchWholeWord,"SearchOptions",this._isNull),this.m_matchWholeWord},set:function(n){this.m_matchWholeWord=n;r(this.context,this,"MatchWholeWord",n)},enumerable:!0,configurable:!0}),Object.defineProperty(f.prototype,"matchWildcards",{get:function(){return t("matchWildcards",this.m_matchWildcards,"SearchOptions",this._isNull),this.m_matchWildcards},set:function(n){this.m_matchWildcards=n;r(this.context,this,"MatchWildcards",n)},enumerable:!0,configurable:!0}),f.prototype._handleResult=function(n){if(u.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t.IgnorePunct)||(this.m_ignorePunct=t.IgnorePunct);i(t.IgnoreSpace)||(this.m_ignoreSpace=t.IgnoreSpace);i(t.MatchCase)||(this.m_matchCase=t.MatchCase);i(t.MatchPrefix)||(this.m_matchPrefix=t.MatchPrefix);i(t.MatchSoundsLike)||(this.m_matchSoundsLike=t.MatchSoundsLike);i(t.MatchSuffix)||(this.m_matchSuffix=t.MatchSuffix);i(t.MatchWholeWord)||(this.m_matchWholeWord=t.MatchWholeWord);i(t.MatchWildcards)||(this.m_matchWildcards=t.MatchWildcards)}},f.prototype.load=function(n){return s(this,n),this},f.newObject=function(t){return new n.SearchOptions(t,et(t,"Microsoft.WordServices.SearchOptions",!1))},f}(OfficeExtension.ClientObject);n.SearchOptions=it;rt=function(r){function u(){r.apply(this,arguments)}return __extends(u,r),Object.defineProperty(u.prototype,"items",{get:function(){return t("items",this.m__items,"SearchResultCollection",this._isNull),this.m__items},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"SearchResultCollection",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),u.prototype._GetItem=function(t){return new n.Range(this.context,a(this.context,this,[t]))},u.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},u.prototype._handleResult=function(t){var u,e,f,s;if((r.prototype._handleResult.call(this,t),!o(t))&&(u=t,h(this,u),i(u._ReferenceId)||(this.m__ReferenceId=u._ReferenceId),!o(u[OfficeExtension.Constants.items])))for(this.m__items=[],e=u[OfficeExtension.Constants.items],f=0;f<e.length;f++)s=new n.Range(this.context,v(!0,this.context,this,e[f],f)),s._handleResult(e[f]),this.m__items.push(s)},u.prototype.load=function(n){return s(this,n),this},u.prototype._initReferenceId=function(n){this.m__ReferenceId=n},u}(OfficeExtension.ClientObject);n.SearchResultCollection=rt;ut=function(r){function c(){r.apply(this,arguments)}return __extends(c,r),Object.defineProperty(c.prototype,"body",{get:function(){return this.m_body||(this.m_body=new n.Body(this.context,e(this.context,this,"Body",!1,!1))),this.m_body},enumerable:!0,configurable:!0}),Object.defineProperty(c.prototype,"_Id",{get:function(){return t("_Id",this.m__Id,"Section",this._isNull),this.m__Id},enumerable:!0,configurable:!0}),Object.defineProperty(c.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"Section",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),c.prototype.getFooter=function(t){return new n.Body(this.context,u(this.context,this,"GetFooter",1,[t],!1,!0))},c.prototype.getHeader=function(t){return new n.Body(this.context,u(this.context,this,"GetHeader",1,[t],!1,!0))},c.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},c.prototype._handleResult=function(n){if(r.prototype._handleResult.call(this,n),!o(n)){var t=n;h(this,t);i(t._Id)||(this.m__Id=t._Id);i(t._ReferenceId)||(this.m__ReferenceId=t._ReferenceId);l(this,t,["body","Body"])}},c.prototype.load=function(n){return s(this,n),this},c.prototype._initReferenceId=function(n){this.m__ReferenceId=n},c}(OfficeExtension.ClientObject);n.Section=ut;ft=function(r){function u(){r.apply(this,arguments)}return __extends(u,r),Object.defineProperty(u.prototype,"items",{get:function(){return t("items",this.m__items,"SectionCollection",this._isNull),this.m__items},enumerable:!0,configurable:!0}),Object.defineProperty(u.prototype,"_ReferenceId",{get:function(){return t("_ReferenceId",this.m__ReferenceId,"SectionCollection",this._isNull),this.m__ReferenceId},enumerable:!0,configurable:!0}),u.prototype._GetItem=function(t){return new n.Section(this.context,a(this.context,this,[t]))},u.prototype._KeepReference=function(){f(this.context,this,"_KeepReference",1,[])},u.prototype._handleResult=function(t){var u,e,f,s;if((r.prototype._handleResult.call(this,t),!o(t))&&(u=t,h(this,u),i(u._ReferenceId)||(this.m__ReferenceId=u._ReferenceId),!o(u[OfficeExtension.Constants.items])))for(this.m__items=[],e=u[OfficeExtension.Constants.items],f=0;f<e.length;f++)s=new n.Section(this.context,v(!0,this.context,this,e[f],f)),s._handleResult(e[f]),this.m__items.push(s)},u.prototype.load=function(n){return s(this,n),this},u.prototype._initReferenceId=function(n){this.m__ReferenceId=n},u}(OfficeExtension.ClientObject);n.SectionCollection=ft,function(n){n.richText="RichText"}(st=n.ContentControlType||(n.ContentControlType={})),function(n){n.boundingBox="BoundingBox";n.tags="Tags";n.hidden="Hidden"}(ht=n.ContentControlAppearance||(n.ContentControlAppearance={})),function(n){n.none="None";n.single="Single";n.word="Word";n.double="Double";n.dotted="Dotted";n.hidden="Hidden";n.thick="Thick";n.dashLine="DashLine";n.dotLine="DotLine";n.dotDashLine="DotDashLine";n.twoDotDashLine="TwoDotDashLine";n.wave="Wave"}(ct=n.UnderlineType||(n.UnderlineType={})),function(n){n.page="Page";n.column="Column";n.next="Next";n.sectionContinuous="SectionContinuous";n.sectionEven="SectionEven";n.sectionOdd="SectionOdd";n.line="Line";n.lineClearLeft="LineClearLeft";n.lineClearRight="LineClearRight";n.textWrapping="TextWrapping"}(lt=n.BreakType||(n.BreakType={})),function(n){n.before="Before";n.after="After";n.start="Start";n.end="End";n.replace="Replace"}(at=n.InsertLocation||(n.InsertLocation={})),function(n){n.unknown="Unknown";n.left="Left";n.centered="Centered";n.right="Right";n.justified="Justified"}(vt=n.Alignment||(n.Alignment={})),function(n){n.primary="Primary";n.firstPage="FirstPage";n.evenPages="EvenPages"}(yt=n.HeaderFooterType||(n.HeaderFooterType={})),function(n){n.select="Select";n.start="Start";n.end="End"}(pt=n.SelectionMode||(n.SelectionMode={})),function(n){n.accessDenied="AccessDenied";n.generalException="GeneralException";n.invalidArgument="InvalidArgument";n.itemNotFound="ItemNotFound";n.notImplemented="NotImplemented"}(wt=n.ErrorCodes||(n.ErrorCodes={}))}(Word||(Word={})),function(n){function i(t){return OfficeExtension.ClientRequestContext._run(function(){return new n.RequestContext},t)}var t=function(t){function i(i){t.call(this,i);this.m_document=new n.Document(this,OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));this._rootObject=this.m_document}return __extends(i,t),Object.defineProperty(i.prototype,"document",{get:function(){return this.m_document},enumerable:!0,configurable:!0}),i}(OfficeExtension.ClientRequestContext);n.RequestContext=t;n.run=i}(Word||(Word={}))
var OfficeExtension;
(function (OfficeExtension) {
	var Action=(function () {
		function Action(actionInfo, isWriteOperation) {
			this.m_actionInfo=actionInfo;
			this.m_isWriteOperation=isWriteOperation;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	})();
	OfficeExtension.Action=Action;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ActionFactory=(function () {
		function ActionFactory() {
		}
		ActionFactory.createSetPropertyAction=function (context, parent, propertyName, value) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4 ,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var ret=new OfficeExtension.Action(actionInfo, true);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			return ret;
		};
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3 ,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var isWriteOperation=operationType !=1 ;
			var ret=new OfficeExtension.Action(actionInfo, isWriteOperation);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			return ret;
		};
		ActionFactory.createQueryAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2 ,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			return ret;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			OfficeExtension.Utility.validateObjectPath(obj);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1 ,
				Name: "",
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(obj._objectPath);
			context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
			return ret;
		};
		ActionFactory.createTraceAction=function (context, message) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5 ,
				Name: "Trace",
				ObjectPathId: 0
			};
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addTrace(actionInfo.Id, message);
			return ret;
		};
		return ActionFactory;
	})();
	OfficeExtension.ActionFactory=ActionFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject=(function () {
		function ClientObject(context, objectPath) {
			OfficeExtension.Utility.checkArgumentNull(context, "context");
			this.m_context=context;
			this.m_objectPath=objectPath;
			if (this.m_objectPath) {
				if (!context._processingResult) {
					OfficeExtension.ActionFactory.createInstantiateAction(context, this);
					if ((context._autoCleanup) && (this._KeepReference)) {
						context.trackedObjects._autoAdd(this);
					}
				}
			}
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			set: function (value) {
				this.m_objectPath=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
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
				this.m_isNull=value;
				if (value && this.m_objectPath) {
					this.m_objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._handleResult=function (value) {
			this.m_isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			if (this.m_isNull && this.m_objectPath) {
				this.m_objectPath._updateAsNullObject();
			}
		};
		return ClientObject;
	})();
	OfficeExtension.ClientObject=ClientObject;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequest=(function () {
		function ClientRequest(context) {
			this.m_context=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_flags=0 ;
			this.m_traceInfos={};
		}
		Object.defineProperty(ClientRequest.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "traceInfos", {
			get: function () {
				return this.m_traceInfos;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype.addAction=function (action) {
			if (action.isWriteOperation) {
				this.m_flags=this.m_flags | 1 ;
			}
			this.m_actions.push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype.addReferencedObjectPath=function (objectPath) {
			if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath));
			}
			while (objectPath) {
				if (objectPath.isWriteOperation) {
					this.m_flags=this.m_flags | 1 ;
				}
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3 ) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequest.prototype.buildRequestMessageBody=function () {
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			for (var index=0; index < this.m_actions.length; index++) {
				actions.push(this.m_actions[index].actionInfo);
			}
			var ret={
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return ret;
		};
		ClientRequest.prototype.processResponse=function (msg) {
			if (msg && msg.Results) {
				for (var i=0; i < msg.Results.length; i++) {
					var actionResult=msg.Results[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequest.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		return ClientRequest;
	})();
	OfficeExtension.ClientRequest=ClientRequest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _requestExecutorFactory=function () { return new OfficeExtension.OfficeJsRequestExecutor(); };
	function _setRequestExecutorFactory(reqExecFactory) {
		_requestExecutorFactory=reqExecFactory;
	}
	OfficeExtension._setRequestExecutorFactory=_setRequestExecutorFactory;
	var ClientRequestContext=(function () {
		function ClientRequestContext(url) {
			this.m_nextId=0;
			this.m_url=url;
			if (OfficeExtension.Utility.isNullOrEmptyString(this.m_url)) {
				this.m_url=OfficeExtension.Constants.localDocument;
			}
			this._processingResult=false;
			this._customData=OfficeExtension.Constants.iterativeExecutor;
			this._requestExecutor=_requestExecutorFactory();
			this.sync=this.sync.bind(this);
		}
		Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
			get: function () {
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new OfficeExtension.TrackedObjects(this);
				}
				return this.m_trackedObjects;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestContext.prototype.load=function (clientObj, option) {
			OfficeExtension.Utility.validateContext(this, clientObj);
			var queryOption={};
			if (typeof (option)=="string") {
				var select=option;
				queryOption.Select=this.parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof (option)=="object") {
				var loadOption=option;
				if (typeof (loadOption.select)=="string") {
					queryOption.Select=this.parseSelectExpand(loadOption.select);
				}
				else if (Array.isArray(loadOption.select)) {
					queryOption.Select=loadOption.select;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.select");
				}
				if (typeof (loadOption.expand)=="string") {
					queryOption.Expand=this.parseSelectExpand(loadOption.expand);
				}
				else if (Array.isArray(loadOption.expand)) {
					queryOption.Expand=loadOption.expand;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.expand");
				}
				if (typeof (loadOption.top)=="number") {
					queryOption.Top=loadOption.top;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.top");
				}
				if (typeof (loadOption.skip)=="number") {
					queryOption.Skip=loadOption.skip;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.skip");
				}
			}
			else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
				OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option");
			}
			var action=OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			OfficeExtension.ActionFactory.createTraceAction(this, message);
		};
		ClientRequestContext.prototype.parseSelectExpand=function (select) {
			var args=[];
			if (!OfficeExtension.Utility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(",");
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=propertyName.trim();
					args.push(propertyName);
				}
			}
			return args;
		};
		ClientRequestContext.prototype.syncPrivate=function (doneCallback, failCallback) {
			var req=this._pendingRequest;
			if (!req.hasActions) {
				doneCallback();
				return;
			}
			this.m_pendingRequest=null;
			var msgBody=req.buildRequestMessageBody();
			var requestFlags=req.flags;
			var requestExecutor=this._requestExecutor;
			if (!requestExecutor) {
				requestExecutor=new OfficeExtension.OfficeJsRequestExecutor();
			}
			var requestExecutorRequestMessage={
				Url: this.m_url,
				Headers: null,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var thisObj=this;
			requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage, function (response) {
				var error;
				var traceMessages=new Array();
				if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
					error=new OfficeExtension._Internal.RuntimeError(response.ErrorCode, response.ErrorMessage, traceMessages, {});
				}
				else if (response.Body && response.Body.Error) {
					error=new OfficeExtension._Internal.RuntimeError(response.Body.Error.Code, response.Body.Error.Message, traceMessages, {
						errorLocation: response.Body.Error.Location
					});
				}
				if (response.Body && response.Body.TraceIds) {
					var traceMessageMap=req.traceInfos;
					for (var i=0; i < response.Body.TraceIds.length; i++) {
						var traceId=response.Body.TraceIds[i];
						var message=traceMessageMap[traceId];
						traceMessages.push(message);
					}
				}
				if (error) {
					failCallback(error);
					return;
				}
				else {
					thisObj._processingResult=true;
					try {
						req.processResponse(response.Body);
					}
					finally {
						thisObj._processingResult=false;
					}
					doneCallback();
					return;
				}
			});
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			var _this=this;
			return new OfficeExtension['Promise'](function (resolve, reject) {
				_this.syncPrivate(function () {
					resolve(passThroughValue);
				}, function (error) {
					reject(error);
				});
			});
		};
		ClientRequestContext._run=function (ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			var starterPromise=new OfficeExtension['Promise'](function (resolve, reject) {
				resolve();
			});
			var ctx;
			var succeeded=false;
			var resultOrError;
			return starterPromise.then(function () {
				ctx=ctxInitializer();
				ctx._autoCleanup=true;
				var batchResult=batch(ctx);
				if (OfficeExtension.Utility.isNullOrUndefined(batchResult) || (typeof batchResult.then !=='function')) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
				}
				return batchResult;
			}).then(function (batchResult) {
				return ctx.sync(batchResult);
			}).then(function (result) {
				succeeded=true;
				resultOrError=result;
			}).catch(function (error) {
				resultOrError=error;
			}).then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				attemptCleanup();
				function attemptCleanup() {
					cleanupCounter++;
					for (var key in itemsToRemove) {
						ctx.trackedObjects.remove(itemsToRemove[key]);
					}
					ctx.sync().then(function () {
						if (onCleanupSuccess) {
							onCleanupSuccess(cleanupCounter);
						}
					}).catch(function () {
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
			}).then(function () {
				if (succeeded) {
					return resultOrError;
				}
				else {
					throw resultOrError;
				}
			});
		};
		ClientRequestContext.prototype._nextId=function () {
			return++this.m_nextId;
		};
		return ClientRequestContext;
	})();
	OfficeExtension.ClientRequestContext=ClientRequestContext;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	(function (ClientRequestFlags) {
		ClientRequestFlags[ClientRequestFlags["None"]=0]="None";
		ClientRequestFlags[ClientRequestFlags["WriteOperation"]=1]="WriteOperation";
	})(OfficeExtension.ClientRequestFlags || (OfficeExtension.ClientRequestFlags={}));
	var ClientRequestFlags=OfficeExtension.ClientRequestFlags;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult=(function () {
		function ClientResult() {
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			if (typeof (value)==="object" && value && value._IsNull) {
				return;
			}
			this.m_value=value;
		};
		return ClientResult;
	})();
	OfficeExtension.ClientResult=ClientResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Constants=(function () {
		function Constants() {
		}
		Constants.getItemAt="GetItemAt";
		Constants.id="Id";
		Constants.idPrivate="_Id";
		Constants.index="_Index";
		Constants.items="_Items";
		Constants.iterativeExecutor="IterativeExecutor";
		Constants.localDocument="http://document.localhost/";
		Constants.localDocumentApiPrefix="http://document.localhost/_api/";
		Constants.referenceId="_ReferenceId";
		Constants.sourceLibHeader="X-OfficeExtension-Source";
		return Constants;
	})();
	OfficeExtension.Constants=Constants;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EmbedRequestExecutor=(function () {
		function EmbedRequestExecutor() {
		}
		EmbedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage, callback) {
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbedRequestExecutor.SourceLibHeaderValue);
			var endpoint=OfficeExtension.Embedded && OfficeExtension.Embedded._getEndpoint();
			if (!endpoint) {
				OfficeExtension.RichApiMessageUtility.sendResponseOnError(OfficeExtension.Embedded && OfficeExtension.Embedded.EmbeddedApiStatus.InternalError, "", callback);
				return;
			}
			endpoint.invoke("executeMethod", function (status, result) {
				OfficeExtension.Utility.log("Response:");
				OfficeExtension.Utility.log(JSON.stringify(result));
				if (status==OfficeExtension.Embedded.EmbeddedApiStatus.Success) {
					OfficeExtension.RichApiMessageUtility.sendResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data), callback);
				}
				else {
					OfficeExtension.RichApiMessageUtility.sendResponseOnError(result.error.Code, result.error.Message, callback);
				}
			}, EmbedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
		};
		EmbedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbedRequestExecutor.SourceLibHeaderValue="Embedded";
		return EmbedRequestExecutor;
	})();
	OfficeExtension.EmbedRequestExecutor=EmbedRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var __extends=this.__extends || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	__.prototype=b.prototype;
	d.prototype=new __();
};
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(code, message, traceMessages, debugInfo) {
				_super.call(this, message);
				this.name="OfficeExtension.Error";
				this.code=code;
				this.message=message;
				this.traceMessages=traceMessages;
				this.debugInfo=debugInfo;
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			return RuntimeError;
		})(Error);
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	OfficeExtension.Error=OfficeExtension._Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes=(function () {
		function ErrorCodes() {
		}
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.activityLimitReached="ActivityLimitReached";
		return ErrorCodes;
	})();
	OfficeExtension.ErrorCodes=ErrorCodes;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			OfficeExtension.Utility.fixObjectPathIfNecessary(this.m_clientObject, value);
			if (value && !OfficeExtension.Utility.isNullOrUndefined(value[OfficeExtension.Constants.referenceId]) && this.m_clientObject._initReferenceId) {
				this.m_clientObject._initReferenceId(value[OfficeExtension.Constants.referenceId]);
			}
		};
		return InstantiateActionResultHandler;
	})();
	OfficeExtension.InstantiateActionResultHandler=InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	(function (RichApiRequestMessageIndex) {
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["CustomData"]=0]="CustomData";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["Method"]=1]="Method";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["PathAndQuery"]=2]="PathAndQuery";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["Headers"]=3]="Headers";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["Body"]=4]="Body";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["AppPermission"]=5]="AppPermission";
		RichApiRequestMessageIndex[RichApiRequestMessageIndex["RequestFlags"]=6]="RequestFlags";
	})(OfficeExtension.RichApiRequestMessageIndex || (OfficeExtension.RichApiRequestMessageIndex={}));
	var RichApiRequestMessageIndex=OfficeExtension.RichApiRequestMessageIndex;
	(function (RichApiResponseMessageIndex) {
		RichApiResponseMessageIndex[RichApiResponseMessageIndex["StatusCode"]=0]="StatusCode";
		RichApiResponseMessageIndex[RichApiResponseMessageIndex["Headers"]=1]="Headers";
		RichApiResponseMessageIndex[RichApiResponseMessageIndex["Body"]=2]="Body";
	})(OfficeExtension.RichApiResponseMessageIndex || (OfficeExtension.RichApiResponseMessageIndex={}));
	var RichApiResponseMessageIndex=OfficeExtension.RichApiResponseMessageIndex;
	;
	(function (ActionType) {
		ActionType[ActionType["Instantiate"]=1]="Instantiate";
		ActionType[ActionType["Query"]=2]="Query";
		ActionType[ActionType["Method"]=3]="Method";
		ActionType[ActionType["SetProperty"]=4]="SetProperty";
		ActionType[ActionType["Trace"]=5]="Trace";
	})(OfficeExtension.ActionType || (OfficeExtension.ActionType={}));
	var ActionType=OfficeExtension.ActionType;
	(function (ObjectPathType) {
		ObjectPathType[ObjectPathType["GlobalObject"]=1]="GlobalObject";
		ObjectPathType[ObjectPathType["NewObject"]=2]="NewObject";
		ObjectPathType[ObjectPathType["Method"]=3]="Method";
		ObjectPathType[ObjectPathType["Property"]=4]="Property";
		ObjectPathType[ObjectPathType["Indexer"]=5]="Indexer";
		ObjectPathType[ObjectPathType["ReferenceId"]=6]="ReferenceId";
		ObjectPathType[ObjectPathType["NullObject"]=7]="NullObject";
	})(OfficeExtension.ObjectPathType || (OfficeExtension.ObjectPathType={}));
	var ObjectPathType=OfficeExtension.ObjectPathType;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isWriteOperation=false;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			set: function (value) {
				this.m_isWriteOperation=value;
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
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_objectPathInfo.ObjectPathType=7 ;
			this.m_objectPathInfo.Name="";
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_parentObjectPath=null;
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.prototype.updateUsingObjectData=function (value) {
			var referenceId=value[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				this.m_isInvalidAfterRequest=false;
				this.m_isValid=true;
				this.m_objectPathInfo.ObjectPathType=6 ;
				this.m_objectPathInfo.Name=referenceId;
				this.m_objectPathInfo.ArgumentInfo={};
				this.m_parentObjectPath=null;
				this.m_argumentObjectPaths=null;
				return;
			}
			if (this.parentObjectPath && this.parentObjectPath.isCollection) {
				var id=value[OfficeExtension.Constants.id];
				if (OfficeExtension.Utility.isNullOrUndefined(id)) {
					id=value[OfficeExtension.Constants.idPrivate];
				}
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					this.m_isInvalidAfterRequest=false;
					this.m_isValid=true;
					this.m_objectPathInfo.ObjectPathType=5 ;
					this.m_objectPathInfo.Name="";
					this.m_objectPathInfo.ArgumentInfo={};
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					this.m_argumentObjectPaths=null;
					return;
				}
			}
		};
		return ObjectPath;
	})();
	OfficeExtension.ObjectPath=ObjectPath;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 1 , Name: "" };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 2 , Name: typeName };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4 ,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5 ,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5 ,
				Name: "",
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3 ,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.isWriteOperation=(operationType !=1 );
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=childItem[OfficeExtension.Constants.id];
			if (OfficeExtension.Utility.isNullOrUndefined(id)) {
				id=childItem[OfficeExtension.Constants.idPrivate];
			}
			if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=childItem[OfficeExtension.Constants.id];
			if (OfficeExtension.Utility.isNullOrUndefined(id)) {
				id=childItem[OfficeExtension.Constants.idPrivate];
			}
			var objectPathInfo=objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5 ,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[OfficeExtension.Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3 ,
				Name: OfficeExtension.Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		return ObjectPathFactory;
	})();
	OfficeExtension.ObjectPathFactory=ObjectPathFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor() {
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage, callback) {
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
				OfficeExtension.Utility.log("Response:");
				OfficeExtension.Utility.log(JSON.stringify(result));
				if (result.status=="succeeded") {
					OfficeExtension.RichApiMessageUtility.sendResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result), callback);
				}
				else {
					OfficeExtension.RichApiMessageUtility.sendResponseOnError(result.error.code, result.error.message, callback);
				}
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue="OfficeJs";
		return OfficeJsRequestExecutor;
	})();
	OfficeExtension.OfficeJsRequestExecutor=OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeXHRSettings=(function () {
		function OfficeXHRSettings() {
		}
		return OfficeXHRSettings;
	})();
	OfficeExtension.OfficeXHRSettings=OfficeXHRSettings;
	function resetXHRFactory(oldFactory) {
		OfficeXHR.settings.oldxhr=oldFactory;
		return officeXHRFactory;
	}
	OfficeExtension.resetXHRFactory=resetXHRFactory;
	function officeXHRFactory() {
		return new OfficeXHR;
	}
	OfficeExtension.officeXHRFactory=officeXHRFactory;
	var OfficeXHR=(function () {
		function OfficeXHR() {
		}
		OfficeXHR.prototype.open=function (method, url) {
			this.m_method=method;
			this.m_url=url;
			if (this.m_url.toLowerCase().indexOf(OfficeExtension.Constants.localDocumentApiPrefix)==0) {
				this.m_url=this.m_url.substr(OfficeExtension.Constants.localDocumentApiPrefix.length);
			}
			else {
				this.m_innerXhr=OfficeXHR.settings.oldxhr();
				var thisObj=this;
				this.m_innerXhr.onreadystatechange=function () {
					thisObj.innerXhrOnreadystatechage();
				};
				this.m_innerXhr.open(method, this.m_url);
			}
		};
		OfficeXHR.prototype.abort=function () {
			if (this.m_innerXhr) {
				this.m_innerXhr.abort();
			}
		};
		OfficeXHR.prototype.send=function (body) {
			if (this.m_innerXhr) {
				this.m_innerXhr.send(body);
			}
			else {
				var thisObj=this;
				var requestFlags=0 ;
				if (!OfficeExtension.Utility.isReadonlyRestRequest(this.m_method)) {
					requestFlags=1 ;
				}
				var execFunction=OfficeXHR.settings.executeRichApiRequestAsync;
				if (!execFunction) {
					execFunction=OSF.DDA.RichApi.executeRichApiRequestAsync;
				}
				execFunction(OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray('', requestFlags, this.m_method, this.m_url, this.m_requestHeaders, body), function (asyncResult) {
					thisObj.officeContextRequestCallback(asyncResult);
				});
			}
		};
		OfficeXHR.prototype.setRequestHeader=function (header, value) {
			if (this.m_innerXhr) {
				this.m_innerXhr.setRequestHeader(header, value);
			}
			else {
				if (!this.m_requestHeaders) {
					this.m_requestHeaders={};
				}
				this.m_requestHeaders[header]=value;
			}
		};
		OfficeXHR.prototype.getResponseHeader=function (header) {
			if (this.m_responseHeaders) {
				return this.m_responseHeaders[header.toUpperCase()];
			}
			return null;
		};
		OfficeXHR.prototype.getAllResponseHeaders=function () {
			return this.m_allResponseHeaders;
		};
		OfficeXHR.prototype.overrideMimeType=function (mimeType) {
			if (this.m_innerXhr) {
				this.m_innerXhr.overrideMimeType(mimeType);
			}
		};
		OfficeXHR.prototype.innerXhrOnreadystatechage=function () {
			this.readyState=this.m_innerXhr.readyState;
			if (this.readyState==OfficeXHR.DONE) {
				this.status=this.m_innerXhr.status;
				this.statusText=this.m_innerXhr.statusText;
				this.responseText=this.m_innerXhr.responseText;
				this.response=this.m_innerXhr.response;
				this.responseType=this.m_innerXhr.responseType;
				this.setAllResponseHeaders(this.m_innerXhr.getAllResponseHeaders());
			}
			if (this.onreadystatechange) {
				this.onreadystatechange();
			}
		};
		OfficeXHR.prototype.officeContextRequestCallback=function (result) {
			this.readyState=OfficeXHR.DONE;
			if (result.status=="succeeded") {
				this.status=OfficeExtension.RichApiMessageUtility.getResponseStatusCode(result);
				this.m_responseHeaders=OfficeExtension.RichApiMessageUtility.getResponseHeaders(result);
				console.debug("ResponseHeaders="+JSON.stringify(this.m_responseHeaders));
				this.responseText=OfficeExtension.RichApiMessageUtility.getResponseBody(result);
				console.debug("ResponseText="+this.responseText);
				this.response=this.responseText;
			}
			else {
				this.status=500;
				this.statusText="Internal Error";
			}
			if (this.onreadystatechange) {
				this.onreadystatechange();
			}
		};
		OfficeXHR.prototype.setAllResponseHeaders=function (allResponseHeaders) {
			this.m_allResponseHeaders=allResponseHeaders;
			this.m_responseHeaders={};
			if (this.m_allResponseHeaders !=null) {
				var regex=new RegExp("\r?\n");
				var entries=this.m_allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=OfficeExtension.Utility.trim(key);
							value=OfficeExtension.Utility.trim(value);
							this.m_responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
		};
		OfficeXHR.UNSENT=0;
		OfficeXHR.OPENED=1;
		OfficeXHR.DONE=4;
		OfficeXHR.settings=new OfficeXHRSettings();
		return OfficeXHR;
	})();
	OfficeExtension.OfficeXHR=OfficeXHR;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				(function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' && typeof importScripts !=='undefined' && typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=setImmediate;
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
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					function lib$es6$promise$asap$$attemptVertex() {
						try {
							var r=require;
							var vertx=r('vertx');
							lib$es6$promise$asap$$vertxNext=vertx.runOnLoop || vertx.runOnContext;
							return lib$es6$promise$asap$$useVertxTimer();
						}
						catch (e) {
							return lib$es6$promise$asap$$useSetTimeout();
						}
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else if (lib$es6$promise$asap$$browserWindow===undefined && typeof require==='function') {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$attemptVertex();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() {
					}
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
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
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
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
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
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
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
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
						if (promise===value) {
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
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
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
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
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
					OfficeExtension.Promise=lib$es6$promise$promise$$default;
				}).call(this);
			}
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	if (!OfficeExtension["Promise"]) {
		if (window.Promise) {
			OfficeExtension.Promise=window.Promise;
		}
		else {
			_Internal.PromiseImpl.Init();
		}
	}
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	(function (OperationType) {
		OperationType[OperationType["Default"]=0]="Default";
		OperationType[OperationType["Read"]=1]="Read";
	})(OfficeExtension.OperationType || (OfficeExtension.OperationType={}));
	var OperationType=OfficeExtension.OperationType;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
			}
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	})();
	OfficeExtension.TrackedObjects=TrackedObjects;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStrings=(function () {
		function ResourceStrings() {
		}
		ResourceStrings.invalidObjectPath="InvalidObjectPath";
		ResourceStrings.propertyNotLoaded="PropertyNotLoaded";
		ResourceStrings.invalidRequestContext="InvalidRequestContext";
		ResourceStrings.invalidArgument="InvalidArgument";
		ResourceStrings.runMustReturnPromise="RunMustReturnPromise";
		return ResourceStrings;
	})();
	OfficeExtension.ResourceStrings=ResourceStrings;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			OfficeExtension.Utility.log("Request:");
			OfficeExtension.Utility.log(requestMessageText);
			var headers={};
			headers[OfficeExtension.Constants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.sendResponseOnSuccess=function (responseBody, responseHeaders, callback) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			callback(response);
		};
		RichApiMessageUtility.sendResponseOnError=function (errorCode, message, callback) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=OfficeExtension.ErrorCodes.generalException;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			response.ErrorMessage=message;
			callback(response);
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId="";
			var instanceId="";
			var marketplaceType="";
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
				marketplaceType
			];
		};
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2 ];
			if (typeof (ret)==="string") {
				return ret;
			}
			var arr=ret;
			return arr.join("");
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1 ];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0 ];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		return RichApiMessageUtility;
	})();
	OfficeExtension.RichApiMessageUtility=RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Utility=(function () {
		function Utility() {
		}
		Utility.checkArgumentNull=function (value, name) {
			if (Utility.isNullOrUndefined(value)) {
				Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, name);
			}
		};
		Utility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isUndefined=function (value) {
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		Utility.trim=function (str) {
			return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
		};
		Utility.caseInsensitiveCompareString=function (str1, str2) {
			if (Utility.isNullOrUndefined(str1)) {
				return Utility.isNullOrUndefined(str2);
			}
			else {
				if (Utility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		Utility.isReadonlyRestRequest=function (method) {
			return Utility.caseInsensitiveCompareString(method, "GET");
		};
		Utility.setMethodArguments=function (context, argumentInfo, args) {
			if (Utility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
				return referencedObjectPaths;
			}
			return null;
		};
		Utility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof OfficeExtension.ClientObject) {
					var clientObject=args[i];
					Utility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value);
			}
		};
		Utility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					var pathExpression=Utility.getObjectPathExpression(objectPath);
					Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, pathExpression);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		Utility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							var pathExpression=Utility.getObjectPathExpression(objectPath);
							Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, pathExpression);
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		Utility.validateContext=function (context, obj) {
			if (obj && obj.context !==context) {
				Utility.throwError(OfficeExtension.ResourceStrings.invalidRequestContext);
			}
		};
		Utility.log=function (message) {
			if (Utility._logEnabled && window.console && window.console.log) {
				window.console.log(message);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new OfficeExtension._Internal.RuntimeError(resourceId, Utility._getResourceString(resourceId, arg), new Array(), errorLocation ? { errorLocation: errorLocation } : {});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return new OfficeExtension._Internal.RuntimeError(code, message, [], { errorLocation: location });
		};
		Utility._getResourceString=function (resourceId, arg) {
			var ret=resourceId;
			if (window.Strings && window.Strings.OfficeOM) {
				var stringName="L_"+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!Utility.isNullOrUndefined(arg)) {
				ret=ret.replace("{0}", arg);
			}
			return ret;
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				Utility.throwError(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName, (entityName ? entityName+"."+propertyName : null));
			}
		};
		Utility.getObjectPathExpression=function (objectPath) {
			var ret="";
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1 :
						ret=ret;
						break;
					case 2 :
						ret="new()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 3 :
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+"()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 4 :
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? "." : "")+ret;
						break;
					case 5 :
						ret="getItem()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 6 :
						ret="_reference()"+(ret.length > 0 ? "." : "")+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		Utility._createPromiseFromResult=function (value) {
			return new OfficeExtension['Promise'](function (resolve, reject) {
				resolve(value);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!Utility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		Utility._logEnabled=false;
		Utility.s_underscoreCharCode="_".charCodeAt(0);
		return Utility;
	})();
	OfficeExtension.Utility=Utility;
})(OfficeExtension || (OfficeExtension={}));

var __extends=this.__extends || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	__.prototype=b.prototype;
	d.prototype=new __();
};
var Word;
(function (Word) {
	function _normalizeSearchOptions(context, searchOptions) {
		if (OfficeExtension.Utility.isNullOrUndefined(searchOptions)) {
			return null;
		}
		if (typeof (searchOptions) !="object") {
			OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "searchOptions");
		}
		if (searchOptions instanceof Word.SearchOptions) {
			return searchOptions;
		}
		var newSearchOptions=Word.SearchOptions.newObject(context);
		for (var property in searchOptions) {
			if (searchOptions.hasOwnProperty(property)) {
				newSearchOptions[property]=searchOptions[property];
			}
		}
		return newSearchOptions;
	}
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var Application=(function (_super) {
		__extends(Application, _super);
		function Application() {
			_super.apply(this, arguments);
		}
		Application.prototype.createDocument=function (base64File) {
			return new Word.Document(this.context, _createMethodObjectPath(this.context, this, "CreateDocument", 1 , [base64File], false, false));
		};
		Application.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		Application.newObject=function (context) {
			var ret=new Word.Application(context, _createNewObjectObjectPath(context, "Microsoft.WordServices.Application", false));
			return ret;
		};
		return Application;
	})(OfficeExtension.ClientObject);
	Word.Application=Application;
	var Body=(function (_super) {
		__extends(Body, _super);
		function Body() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Body.prototype, "contentControls", {
			get: function () {
				if (!this.m_contentControls) {
					this.m_contentControls=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false));
				}
				return this.m_contentControls;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "inlinePictures", {
			get: function () {
				if (!this.m_inlinePictures) {
					this.m_inlinePictures=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
				}
				return this.m_inlinePictures;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "lists", {
			get: function () {
				if (!this.m_lists) {
					this.m_lists=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false));
				}
				return this.m_lists;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "paragraphs", {
			get: function () {
				if (!this.m_paragraphs) {
					this.m_paragraphs=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
				}
				return this.m_paragraphs;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentBody", {
			get: function () {
				if (!this.m_parentBody) {
					this.m_parentBody=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false));
				}
				return this.m_parentBody;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "tables", {
			get: function () {
				if (!this.m_tables) {
					this.m_tables=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				return this.m_tables;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this.m_style, "Body", this._isNull);
				return this.m_style;
			},
			set: function (value) {
				this.m_style=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this.m_text, "Body", this._isNull);
				return this.m_text;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this.m_type, "Body", this._isNull);
				return this.m_type;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Body", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Body.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		Body.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Body.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Body.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		Body.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0 , [breakType, insertLocation]);
		};
		Body.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0 , [], false, true));
		};
		Body.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0 , [base64File, insertLocation], false, true));
		};
		Body.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0 , [html, insertLocation], false, true));
		};
		Body.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 , [base64EncodedImage, insertLocation], false, true));
		};
		Body.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0 , [ooxml, insertLocation], false, true));
		};
		Body.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		Body.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0 , [rowCount, columnCount, insertLocation, values], false, true));
		};
		Body.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0 , [text, insertLocation], false, true));
		};
		Body.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		Body.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		Body.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Body.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Style"])) {
				this.m_style=obj["Style"];
			}
			if (!_isUndefined(obj["Text"])) {
				this.m_text=obj["Text"];
			}
			if (!_isUndefined(obj["Type"])) {
				this.m_type=obj["Type"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "tables", "Tables"]);
		};
		Body.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Body.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Body;
	})(OfficeExtension.ClientObject);
	Word.Body=Body;
	var ContentControl=(function (_super) {
		__extends(ContentControl, _super);
		function ContentControl() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ContentControl.prototype, "contentControls", {
			get: function () {
				if (!this.m_contentControls) {
					this.m_contentControls=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false));
				}
				return this.m_contentControls;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "inlinePictures", {
			get: function () {
				if (!this.m_inlinePictures) {
					this.m_inlinePictures=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
				}
				return this.m_inlinePictures;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "lists", {
			get: function () {
				if (!this.m_lists) {
					this.m_lists=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false));
				}
				return this.m_lists;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "paragraphs", {
			get: function () {
				if (!this.m_paragraphs) {
					this.m_paragraphs=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
				}
				return this.m_paragraphs;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTableCell", {
			get: function () {
				if (!this.m_parentTableCell) {
					this.m_parentTableCell=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false));
				}
				return this.m_parentTableCell;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "tables", {
			get: function () {
				if (!this.m_tables) {
					this.m_tables=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				return this.m_tables;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "appearance", {
			get: function () {
				_throwIfNotLoaded("appearance", this.m_appearance, "ContentControl", this._isNull);
				return this.m_appearance;
			},
			set: function (value) {
				this.m_appearance=value;
				_createSetPropertyAction(this.context, this, "Appearance", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "cannotDelete", {
			get: function () {
				_throwIfNotLoaded("cannotDelete", this.m_cannotDelete, "ContentControl", this._isNull);
				return this.m_cannotDelete;
			},
			set: function (value) {
				this.m_cannotDelete=value;
				_createSetPropertyAction(this.context, this, "CannotDelete", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "cannotEdit", {
			get: function () {
				_throwIfNotLoaded("cannotEdit", this.m_cannotEdit, "ContentControl", this._isNull);
				return this.m_cannotEdit;
			},
			set: function (value) {
				this.m_cannotEdit=value;
				_createSetPropertyAction(this.context, this, "CannotEdit", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this.m_color, "ContentControl", this._isNull);
				return this.m_color;
			},
			set: function (value) {
				this.m_color=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this.m_id, "ContentControl", this._isNull);
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "placeholderText", {
			get: function () {
				_throwIfNotLoaded("placeholderText", this.m_placeholderText, "ContentControl", this._isNull);
				return this.m_placeholderText;
			},
			set: function (value) {
				this.m_placeholderText=value;
				_createSetPropertyAction(this.context, this, "PlaceholderText", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "removeWhenEdited", {
			get: function () {
				_throwIfNotLoaded("removeWhenEdited", this.m_removeWhenEdited, "ContentControl", this._isNull);
				return this.m_removeWhenEdited;
			},
			set: function (value) {
				this.m_removeWhenEdited=value;
				_createSetPropertyAction(this.context, this, "RemoveWhenEdited", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this.m_style, "ContentControl", this._isNull);
				return this.m_style;
			},
			set: function (value) {
				this.m_style=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "subtype", {
			get: function () {
				_throwIfNotLoaded("subtype", this.m_subtype, "ContentControl", this._isNull);
				return this.m_subtype;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "tag", {
			get: function () {
				_throwIfNotLoaded("tag", this.m_tag, "ContentControl", this._isNull);
				return this.m_tag;
			},
			set: function (value) {
				this.m_tag=value;
				_createSetPropertyAction(this.context, this, "Tag", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this.m_text, "ContentControl", this._isNull);
				return this.m_text;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "title", {
			get: function () {
				_throwIfNotLoaded("title", this.m_title, "ContentControl", this._isNull);
				return this.m_title;
			},
			set: function (value) {
				this.m_title=value;
				_createSetPropertyAction(this.context, this, "Title", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this.m_type, "ContentControl", this._isNull);
				return this.m_type;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "ContentControl", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		ContentControl.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		ContentControl.prototype.delete=function (keepContent) {
			_createMethodAction(this.context, this, "Delete", 0 , [keepContent]);
		};
		ContentControl.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ContentControl.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ContentControl.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		ContentControl.prototype.getTextRanges=function (punctuationMarks, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1 , [punctuationMarks, trimSpacing], true, false));
		};
		ContentControl.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0 , [breakType, insertLocation]);
		};
		ContentControl.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0 , [base64File, insertLocation], false, true));
		};
		ContentControl.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0 , [html, insertLocation], false, true));
		};
		ContentControl.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 , [base64EncodedImage, insertLocation], false, true));
		};
		ContentControl.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0 , [ooxml, insertLocation], false, true));
		};
		ContentControl.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		ContentControl.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0 , [rowCount, columnCount, insertLocation, values], false, true));
		};
		ContentControl.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0 , [text, insertLocation], false, true));
		};
		ContentControl.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		ContentControl.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		ContentControl.prototype.split=function (delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1 , [delimiters, multiParagraphs, trimDelimiters, trimSpacing], true, false));
		};
		ContentControl.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		ContentControl.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Appearance"])) {
				this.m_appearance=obj["Appearance"];
			}
			if (!_isUndefined(obj["CannotDelete"])) {
				this.m_cannotDelete=obj["CannotDelete"];
			}
			if (!_isUndefined(obj["CannotEdit"])) {
				this.m_cannotEdit=obj["CannotEdit"];
			}
			if (!_isUndefined(obj["Color"])) {
				this.m_color=obj["Color"];
			}
			if (!_isUndefined(obj["Id"])) {
				this.m_id=obj["Id"];
			}
			if (!_isUndefined(obj["PlaceholderText"])) {
				this.m_placeholderText=obj["PlaceholderText"];
			}
			if (!_isUndefined(obj["RemoveWhenEdited"])) {
				this.m_removeWhenEdited=obj["RemoveWhenEdited"];
			}
			if (!_isUndefined(obj["Style"])) {
				this.m_style=obj["Style"];
			}
			if (!_isUndefined(obj["Subtype"])) {
				this.m_subtype=obj["Subtype"];
			}
			if (!_isUndefined(obj["Tag"])) {
				this.m_tag=obj["Tag"];
			}
			if (!_isUndefined(obj["Text"])) {
				this.m_text=obj["Text"];
			}
			if (!_isUndefined(obj["Title"])) {
				this.m_title=obj["Title"];
			}
			if (!_isUndefined(obj["Type"])) {
				this.m_type=obj["Type"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentContentControl", "ParentContentControl", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "tables", "Tables"]);
		};
		ContentControl.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ContentControl.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return ContentControl;
	})(OfficeExtension.ClientObject);
	Word.ContentControl=ContentControl;
	var ContentControlCollection=(function (_super) {
		__extends(ContentControlCollection, _super);
		function ContentControlCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ContentControlCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "ContentControlCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "ContentControlCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		ContentControlCollection.prototype.getById=function (id) {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "GetById", 1 , [id], false, false));
		};
		ContentControlCollection.prototype.getByTag=function (tag) {
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTag", 1 , [tag], true, false));
		};
		ContentControlCollection.prototype.getByTitle=function (title) {
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTitle", 1 , [title], true, false));
		};
		ContentControlCollection.prototype.getByTypes=function (types) {
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTypes", 1 , [types], true, false));
		};
		ContentControlCollection.prototype.getItem=function (index) {
			return new Word.ContentControl(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ContentControlCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		ContentControlCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.ContentControl(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ContentControlCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ContentControlCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return ContentControlCollection;
	})(OfficeExtension.ClientObject);
	Word.ContentControlCollection=ContentControlCollection;
	var Document=(function (_super) {
		__extends(Document, _super);
		function Document() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Document.prototype, "body", {
			get: function () {
				if (!this.m_body) {
					this.m_body=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false));
				}
				return this.m_body;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "contentControls", {
			get: function () {
				if (!this.m_contentControls) {
					this.m_contentControls=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false));
				}
				return this.m_contentControls;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "sections", {
			get: function () {
				if (!this.m_sections) {
					this.m_sections=new Word.SectionCollection(this.context, _createPropertyObjectPath(this.context, this, "Sections", true, false));
				}
				return this.m_sections;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "saved", {
			get: function () {
				_throwIfNotLoaded("saved", this.m_saved, "Document", this._isNull);
				return this.m_saved;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Document", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Document.prototype.getSelection=function () {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetSelection", 1 , [], false, true));
		};
		Document.prototype.open=function () {
			_createMethodAction(this.context, this, "Open", 1 , []);
		};
		Document.prototype.save=function () {
			_createMethodAction(this.context, this, "Save", 0 , []);
		};
		Document.prototype._GetObjectByReferenceId=function (referenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectByReferenceId", 1 , [referenceId]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype._GetObjectTypeNameByReferenceId=function (referenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1 , [referenceId]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Document.prototype._RemoveAllReferences=function () {
			_createMethodAction(this.context, this, "_RemoveAllReferences", 1 , []);
		};
		Document.prototype._RemoveReference=function (referenceId) {
			_createMethodAction(this.context, this, "_RemoveReference", 1 , [referenceId]);
		};
		Document.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Saved"])) {
				this.m_saved=obj["Saved"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "contentControls", "ContentControls", "sections", "Sections"]);
		};
		Document.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Document.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Document;
	})(OfficeExtension.ClientObject);
	Word.Document=Document;
	var Font=(function (_super) {
		__extends(Font, _super);
		function Font() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Font.prototype, "bold", {
			get: function () {
				_throwIfNotLoaded("bold", this.m_bold, "Font", this._isNull);
				return this.m_bold;
			},
			set: function (value) {
				this.m_bold=value;
				_createSetPropertyAction(this.context, this, "Bold", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this.m_color, "Font", this._isNull);
				return this.m_color;
			},
			set: function (value) {
				this.m_color=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "doubleStrikeThrough", {
			get: function () {
				_throwIfNotLoaded("doubleStrikeThrough", this.m_doubleStrikeThrough, "Font", this._isNull);
				return this.m_doubleStrikeThrough;
			},
			set: function (value) {
				this.m_doubleStrikeThrough=value;
				_createSetPropertyAction(this.context, this, "DoubleStrikeThrough", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "highlightColor", {
			get: function () {
				_throwIfNotLoaded("highlightColor", this.m_highlightColor, "Font", this._isNull);
				return this.m_highlightColor;
			},
			set: function (value) {
				this.m_highlightColor=value;
				_createSetPropertyAction(this.context, this, "HighlightColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "italic", {
			get: function () {
				_throwIfNotLoaded("italic", this.m_italic, "Font", this._isNull);
				return this.m_italic;
			},
			set: function (value) {
				this.m_italic=value;
				_createSetPropertyAction(this.context, this, "Italic", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this.m_name, "Font", this._isNull);
				return this.m_name;
			},
			set: function (value) {
				this.m_name=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "size", {
			get: function () {
				_throwIfNotLoaded("size", this.m_size, "Font", this._isNull);
				return this.m_size;
			},
			set: function (value) {
				this.m_size=value;
				_createSetPropertyAction(this.context, this, "Size", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "strikeThrough", {
			get: function () {
				_throwIfNotLoaded("strikeThrough", this.m_strikeThrough, "Font", this._isNull);
				return this.m_strikeThrough;
			},
			set: function (value) {
				this.m_strikeThrough=value;
				_createSetPropertyAction(this.context, this, "StrikeThrough", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "subscript", {
			get: function () {
				_throwIfNotLoaded("subscript", this.m_subscript, "Font", this._isNull);
				return this.m_subscript;
			},
			set: function (value) {
				this.m_subscript=value;
				_createSetPropertyAction(this.context, this, "Subscript", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "superscript", {
			get: function () {
				_throwIfNotLoaded("superscript", this.m_superscript, "Font", this._isNull);
				return this.m_superscript;
			},
			set: function (value) {
				this.m_superscript=value;
				_createSetPropertyAction(this.context, this, "Superscript", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "underline", {
			get: function () {
				_throwIfNotLoaded("underline", this.m_underline, "Font", this._isNull);
				return this.m_underline;
			},
			set: function (value) {
				this.m_underline=value;
				_createSetPropertyAction(this.context, this, "Underline", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Font", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Font.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Font.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Bold"])) {
				this.m_bold=obj["Bold"];
			}
			if (!_isUndefined(obj["Color"])) {
				this.m_color=obj["Color"];
			}
			if (!_isUndefined(obj["DoubleStrikeThrough"])) {
				this.m_doubleStrikeThrough=obj["DoubleStrikeThrough"];
			}
			if (!_isUndefined(obj["HighlightColor"])) {
				this.m_highlightColor=obj["HighlightColor"];
			}
			if (!_isUndefined(obj["Italic"])) {
				this.m_italic=obj["Italic"];
			}
			if (!_isUndefined(obj["Name"])) {
				this.m_name=obj["Name"];
			}
			if (!_isUndefined(obj["Size"])) {
				this.m_size=obj["Size"];
			}
			if (!_isUndefined(obj["StrikeThrough"])) {
				this.m_strikeThrough=obj["StrikeThrough"];
			}
			if (!_isUndefined(obj["Subscript"])) {
				this.m_subscript=obj["Subscript"];
			}
			if (!_isUndefined(obj["Superscript"])) {
				this.m_superscript=obj["Superscript"];
			}
			if (!_isUndefined(obj["Underline"])) {
				this.m_underline=obj["Underline"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
		};
		Font.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Font.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Font;
	})(OfficeExtension.ClientObject);
	Word.Font=Font;
	var InlinePicture=(function (_super) {
		__extends(InlinePicture, _super);
		function InlinePicture() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(InlinePicture.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.InlinePicture(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "paragraph", {
			get: function () {
				if (!this.m_paragraph) {
					this.m_paragraph=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "Paragraph", false, false));
				}
				return this.m_paragraph;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTableCell", {
			get: function () {
				if (!this.m_parentTableCell) {
					this.m_parentTableCell=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false));
				}
				return this.m_parentTableCell;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "altTextDescription", {
			get: function () {
				_throwIfNotLoaded("altTextDescription", this.m_altTextDescription, "InlinePicture", this._isNull);
				return this.m_altTextDescription;
			},
			set: function (value) {
				this.m_altTextDescription=value;
				_createSetPropertyAction(this.context, this, "AltTextDescription", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "altTextTitle", {
			get: function () {
				_throwIfNotLoaded("altTextTitle", this.m_altTextTitle, "InlinePicture", this._isNull);
				return this.m_altTextTitle;
			},
			set: function (value) {
				this.m_altTextTitle=value;
				_createSetPropertyAction(this.context, this, "AltTextTitle", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "height", {
			get: function () {
				_throwIfNotLoaded("height", this.m_height, "InlinePicture", this._isNull);
				return this.m_height;
			},
			set: function (value) {
				this.m_height=value;
				_createSetPropertyAction(this.context, this, "Height", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "hyperlink", {
			get: function () {
				_throwIfNotLoaded("hyperlink", this.m_hyperlink, "InlinePicture", this._isNull);
				return this.m_hyperlink;
			},
			set: function (value) {
				this.m_hyperlink=value;
				_createSetPropertyAction(this.context, this, "Hyperlink", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "imageFormat", {
			get: function () {
				_throwIfNotLoaded("imageFormat", this.m_imageFormat, "InlinePicture", this._isNull);
				return this.m_imageFormat;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "lockAspectRatio", {
			get: function () {
				_throwIfNotLoaded("lockAspectRatio", this.m_lockAspectRatio, "InlinePicture", this._isNull);
				return this.m_lockAspectRatio;
			},
			set: function (value) {
				this.m_lockAspectRatio=value;
				_createSetPropertyAction(this.context, this, "LockAspectRatio", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this.m_width, "InlinePicture", this._isNull);
				return this.m_width;
			},
			set: function (value) {
				this.m_width=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "InlinePicture", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "InlinePicture", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		InlinePicture.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0 , []);
		};
		InlinePicture.prototype.getBase64ImageSrc=function () {
			var action=_createMethodAction(this.context, this, "GetBase64ImageSrc", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		InlinePicture.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		InlinePicture.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0 , [breakType, insertLocation]);
		};
		InlinePicture.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0 , [], false, true));
		};
		InlinePicture.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0 , [base64File, insertLocation], false, true));
		};
		InlinePicture.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0 , [html, insertLocation], false, true));
		};
		InlinePicture.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 , [base64EncodedImage, insertLocation], false, true));
		};
		InlinePicture.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0 , [ooxml, insertLocation], false, true));
		};
		InlinePicture.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		InlinePicture.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0 , [text, insertLocation], false, true));
		};
		InlinePicture.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		InlinePicture.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		InlinePicture.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["AltTextDescription"])) {
				this.m_altTextDescription=obj["AltTextDescription"];
			}
			if (!_isUndefined(obj["AltTextTitle"])) {
				this.m_altTextTitle=obj["AltTextTitle"];
			}
			if (!_isUndefined(obj["Height"])) {
				this.m_height=obj["Height"];
			}
			if (!_isUndefined(obj["Hyperlink"])) {
				this.m_hyperlink=obj["Hyperlink"];
			}
			if (!_isUndefined(obj["ImageFormat"])) {
				this.m_imageFormat=obj["ImageFormat"];
			}
			if (!_isUndefined(obj["LockAspectRatio"])) {
				this.m_lockAspectRatio=obj["LockAspectRatio"];
			}
			if (!_isUndefined(obj["Width"])) {
				this.m_width=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["next", "Next", "paragraph", "Paragraph", "parentContentControl", "ParentContentControl", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell"]);
		};
		InlinePicture.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		InlinePicture.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return InlinePicture;
	})(OfficeExtension.ClientObject);
	Word.InlinePicture=InlinePicture;
	var InlinePictureCollection=(function (_super) {
		__extends(InlinePictureCollection, _super);
		function InlinePictureCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(InlinePictureCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.InlinePicture(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "InlinePictureCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "InlinePictureCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		InlinePictureCollection.prototype._GetItem=function (index) {
			return new Word.InlinePicture(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		InlinePictureCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		InlinePictureCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.InlinePicture(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		InlinePictureCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		InlinePictureCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return InlinePictureCollection;
	})(OfficeExtension.ClientObject);
	Word.InlinePictureCollection=InlinePictureCollection;
	var List=(function (_super) {
		__extends(List, _super);
		function List() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(List.prototype, "paragraphs", {
			get: function () {
				if (!this.m_paragraphs) {
					this.m_paragraphs=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
				}
				return this.m_paragraphs;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this.m_id, "List", this._isNull);
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "levelExistences", {
			get: function () {
				_throwIfNotLoaded("levelExistences", this.m_levelExistences, "List", this._isNull);
				return this.m_levelExistences;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "levelTypes", {
			get: function () {
				_throwIfNotLoaded("levelTypes", this.m_levelTypes, "List", this._isNull);
				return this.m_levelTypes;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "List", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		List.prototype.getLevelFont=function (level) {
			return new Word.Font(this.context, _createMethodObjectPath(this.context, this, "GetLevelFont", 1 , [level], false, false));
		};
		List.prototype.getLevelParagraphs=function (level) {
			return new Word.ParagraphCollection(this.context, _createMethodObjectPath(this.context, this, "GetLevelParagraphs", 1 , [level], true, false));
		};
		List.prototype.getLevelPicture=function (level) {
			var action=_createMethodAction(this.context, this, "GetLevelPicture", 1 , [level]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		List.prototype.getLevelString=function (level) {
			var action=_createMethodAction(this.context, this, "GetLevelString", 1 , [level]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		List.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		List.prototype.resetLevelFont=function (level, resetFontName) {
			_createMethodAction(this.context, this, "ResetLevelFont", 0 , [level, resetFontName]);
		};
		List.prototype.setLevelAlignment=function (level, alignment) {
			_createMethodAction(this.context, this, "SetLevelAlignment", 0 , [level, alignment]);
		};
		List.prototype.setLevelBullet=function (level, listBullet, charCode, fontName) {
			_createMethodAction(this.context, this, "SetLevelBullet", 0 , [level, listBullet, charCode, fontName]);
		};
		List.prototype.setLevelIndents=function (level, textIndent, bulletNumberPictureIndent) {
			_createMethodAction(this.context, this, "SetLevelIndents", 0 , [level, textIndent, bulletNumberPictureIndent]);
		};
		List.prototype.setLevelNumbering=function (level, listNumbering, formatString) {
			_createMethodAction(this.context, this, "SetLevelNumbering", 0 , [level, listNumbering, formatString]);
		};
		List.prototype.setLevelPicture=function (level, base64EncodedImage) {
			_createMethodAction(this.context, this, "SetLevelPicture", 0 , [level, base64EncodedImage]);
		};
		List.prototype.setLevelStartingNumber=function (level, startingNumber) {
			_createMethodAction(this.context, this, "SetLevelStartingNumber", 0 , [level, startingNumber]);
		};
		List.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		List.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this.m_id=obj["Id"];
			}
			if (!_isUndefined(obj["LevelExistences"])) {
				this.m_levelExistences=obj["LevelExistences"];
			}
			if (!_isUndefined(obj["LevelTypes"])) {
				this.m_levelTypes=obj["LevelTypes"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["paragraphs", "Paragraphs"]);
		};
		List.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		List.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return List;
	})(OfficeExtension.ClientObject);
	Word.List=List;
	var ListCollection=(function (_super) {
		__extends(ListCollection, _super);
		function ListCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ListCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.List(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "ListCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "ListCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		ListCollection.prototype.getById=function (id) {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "GetById", 1 , [id], false, false));
		};
		ListCollection.prototype._GetItem=function (index) {
			return new Word.List(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ListCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		ListCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.List(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ListCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ListCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return ListCollection;
	})(OfficeExtension.ClientObject);
	Word.ListCollection=ListCollection;
	var ListItem=(function (_super) {
		__extends(ListItem, _super);
		function ListItem() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ListItem.prototype, "listString", {
			get: function () {
				_throwIfNotLoaded("listString", this.m_listString, "ListItem", this._isNull);
				return this.m_listString;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "siblingIndex", {
			get: function () {
				_throwIfNotLoaded("siblingIndex", this.m_siblingIndex, "ListItem", this._isNull);
				return this.m_siblingIndex;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "ListItem", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		ListItem.prototype.getAncestor=function (parentOnly) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetAncestor", 1 , [parentOnly], false, false));
		};
		ListItem.prototype.getDescendants=function (directChildrenOnly) {
			return new Word.ParagraphCollection(this.context, _createMethodObjectPath(this.context, this, "GetDescendants", 1 , [directChildrenOnly], true, false));
		};
		ListItem.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		ListItem.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["ListString"])) {
				this.m_listString=obj["ListString"];
			}
			if (!_isUndefined(obj["SiblingIndex"])) {
				this.m_siblingIndex=obj["SiblingIndex"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
		};
		ListItem.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ListItem.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return ListItem;
	})(OfficeExtension.ClientObject);
	Word.ListItem=ListItem;
	var Paragraph=(function (_super) {
		__extends(Paragraph, _super);
		function Paragraph() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Paragraph.prototype, "contentControls", {
			get: function () {
				if (!this.m_contentControls) {
					this.m_contentControls=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false));
				}
				return this.m_contentControls;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "inlinePictures", {
			get: function () {
				if (!this.m_inlinePictures) {
					this.m_inlinePictures=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
				}
				return this.m_inlinePictures;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "list", {
			get: function () {
				if (!this.m_list) {
					this.m_list=new Word.List(this.context, _createPropertyObjectPath(this.context, this, "List", false, false));
				}
				return this.m_list;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "listItem", {
			get: function () {
				if (!this.m_listItem) {
					this.m_listItem=new Word.ListItem(this.context, _createPropertyObjectPath(this.context, this, "ListItem", false, false));
				}
				return this.m_listItem;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentBody", {
			get: function () {
				if (!this.m_parentBody) {
					this.m_parentBody=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false));
				}
				return this.m_parentBody;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTableCell", {
			get: function () {
				if (!this.m_parentTableCell) {
					this.m_parentTableCell=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false));
				}
				return this.m_parentTableCell;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "previous", {
			get: function () {
				if (!this.m_previous) {
					this.m_previous=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "Previous", false, false));
				}
				return this.m_previous;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "alignment", {
			get: function () {
				_throwIfNotLoaded("alignment", this.m_alignment, "Paragraph", this._isNull);
				return this.m_alignment;
			},
			set: function (value) {
				this.m_alignment=value;
				_createSetPropertyAction(this.context, this, "Alignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "firstLineIndent", {
			get: function () {
				_throwIfNotLoaded("firstLineIndent", this.m_firstLineIndent, "Paragraph", this._isNull);
				return this.m_firstLineIndent;
			},
			set: function (value) {
				this.m_firstLineIndent=value;
				_createSetPropertyAction(this.context, this, "FirstLineIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "leftIndent", {
			get: function () {
				_throwIfNotLoaded("leftIndent", this.m_leftIndent, "Paragraph", this._isNull);
				return this.m_leftIndent;
			},
			set: function (value) {
				this.m_leftIndent=value;
				_createSetPropertyAction(this.context, this, "LeftIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineSpacing", {
			get: function () {
				_throwIfNotLoaded("lineSpacing", this.m_lineSpacing, "Paragraph", this._isNull);
				return this.m_lineSpacing;
			},
			set: function (value) {
				this.m_lineSpacing=value;
				_createSetPropertyAction(this.context, this, "LineSpacing", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineUnitAfter", {
			get: function () {
				_throwIfNotLoaded("lineUnitAfter", this.m_lineUnitAfter, "Paragraph", this._isNull);
				return this.m_lineUnitAfter;
			},
			set: function (value) {
				this.m_lineUnitAfter=value;
				_createSetPropertyAction(this.context, this, "LineUnitAfter", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineUnitBefore", {
			get: function () {
				_throwIfNotLoaded("lineUnitBefore", this.m_lineUnitBefore, "Paragraph", this._isNull);
				return this.m_lineUnitBefore;
			},
			set: function (value) {
				this.m_lineUnitBefore=value;
				_createSetPropertyAction(this.context, this, "LineUnitBefore", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "listLevel", {
			get: function () {
				_throwIfNotLoaded("listLevel", this.m_listLevel, "Paragraph", this._isNull);
				return this.m_listLevel;
			},
			set: function (value) {
				this.m_listLevel=value;
				_createSetPropertyAction(this.context, this, "ListLevel", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "outlineLevel", {
			get: function () {
				_throwIfNotLoaded("outlineLevel", this.m_outlineLevel, "Paragraph", this._isNull);
				return this.m_outlineLevel;
			},
			set: function (value) {
				this.m_outlineLevel=value;
				_createSetPropertyAction(this.context, this, "OutlineLevel", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "rightIndent", {
			get: function () {
				_throwIfNotLoaded("rightIndent", this.m_rightIndent, "Paragraph", this._isNull);
				return this.m_rightIndent;
			},
			set: function (value) {
				this.m_rightIndent=value;
				_createSetPropertyAction(this.context, this, "RightIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "spaceAfter", {
			get: function () {
				_throwIfNotLoaded("spaceAfter", this.m_spaceAfter, "Paragraph", this._isNull);
				return this.m_spaceAfter;
			},
			set: function (value) {
				this.m_spaceAfter=value;
				_createSetPropertyAction(this.context, this, "SpaceAfter", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "spaceBefore", {
			get: function () {
				_throwIfNotLoaded("spaceBefore", this.m_spaceBefore, "Paragraph", this._isNull);
				return this.m_spaceBefore;
			},
			set: function (value) {
				this.m_spaceBefore=value;
				_createSetPropertyAction(this.context, this, "SpaceBefore", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this.m_style, "Paragraph", this._isNull);
				return this.m_style;
			},
			set: function (value) {
				this.m_style=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "tableNestingLevel", {
			get: function () {
				_throwIfNotLoaded("tableNestingLevel", this.m_tableNestingLevel, "Paragraph", this._isNull);
				return this.m_tableNestingLevel;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this.m_text, "Paragraph", this._isNull);
				return this.m_text;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "Paragraph", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Paragraph", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Paragraph.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		Paragraph.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0 , []);
		};
		Paragraph.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Paragraph.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Paragraph.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		Paragraph.prototype.getTextRanges=function (punctuationMarks, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1 , [punctuationMarks, trimSpacing], true, false));
		};
		Paragraph.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0 , [breakType, insertLocation]);
		};
		Paragraph.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0 , [], false, true));
		};
		Paragraph.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0 , [base64File, insertLocation], false, true));
		};
		Paragraph.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0 , [html, insertLocation], false, true));
		};
		Paragraph.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 , [base64EncodedImage, insertLocation], false, true));
		};
		Paragraph.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0 , [ooxml, insertLocation], false, true));
		};
		Paragraph.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		Paragraph.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0 , [rowCount, columnCount, insertLocation, values], false, true));
		};
		Paragraph.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0 , [text, insertLocation], false, true));
		};
		Paragraph.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		Paragraph.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		Paragraph.prototype.split=function (delimiters, trimDelimiters, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1 , [delimiters, trimDelimiters, trimSpacing], true, false));
		};
		Paragraph.prototype.startNewList=function () {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "StartNewList", 0 , [], false, false));
		};
		Paragraph.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Paragraph.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Alignment"])) {
				this.m_alignment=obj["Alignment"];
			}
			if (!_isUndefined(obj["FirstLineIndent"])) {
				this.m_firstLineIndent=obj["FirstLineIndent"];
			}
			if (!_isUndefined(obj["LeftIndent"])) {
				this.m_leftIndent=obj["LeftIndent"];
			}
			if (!_isUndefined(obj["LineSpacing"])) {
				this.m_lineSpacing=obj["LineSpacing"];
			}
			if (!_isUndefined(obj["LineUnitAfter"])) {
				this.m_lineUnitAfter=obj["LineUnitAfter"];
			}
			if (!_isUndefined(obj["LineUnitBefore"])) {
				this.m_lineUnitBefore=obj["LineUnitBefore"];
			}
			if (!_isUndefined(obj["ListLevel"])) {
				this.m_listLevel=obj["ListLevel"];
			}
			if (!_isUndefined(obj["OutlineLevel"])) {
				this.m_outlineLevel=obj["OutlineLevel"];
			}
			if (!_isUndefined(obj["RightIndent"])) {
				this.m_rightIndent=obj["RightIndent"];
			}
			if (!_isUndefined(obj["SpaceAfter"])) {
				this.m_spaceAfter=obj["SpaceAfter"];
			}
			if (!_isUndefined(obj["SpaceBefore"])) {
				this.m_spaceBefore=obj["SpaceBefore"];
			}
			if (!_isUndefined(obj["Style"])) {
				this.m_style=obj["Style"];
			}
			if (!_isUndefined(obj["TableNestingLevel"])) {
				this.m_tableNestingLevel=obj["TableNestingLevel"];
			}
			if (!_isUndefined(obj["Text"])) {
				this.m_text=obj["Text"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "list", "List", "listItem", "ListItem", "next", "Next", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "previous", "Previous"]);
		};
		Paragraph.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Paragraph.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Paragraph;
	})(OfficeExtension.ClientObject);
	Word.Paragraph=Paragraph;
	var ParagraphCollection=(function (_super) {
		__extends(ParagraphCollection, _super);
		function ParagraphCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ParagraphCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "last", {
			get: function () {
				if (!this.m_last) {
					this.m_last=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "Last", false, false));
				}
				return this.m_last;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "ParagraphCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "ParagraphCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		ParagraphCollection.prototype._GetItem=function (index) {
			return new Word.Paragraph(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ParagraphCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		ParagraphCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First", "last", "Last"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Paragraph(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ParagraphCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ParagraphCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return ParagraphCollection;
	})(OfficeExtension.ClientObject);
	Word.ParagraphCollection=ParagraphCollection;
	var Range=(function (_super) {
		__extends(Range, _super);
		function Range() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Range.prototype, "contentControls", {
			get: function () {
				if (!this.m_contentControls) {
					this.m_contentControls=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false));
				}
				return this.m_contentControls;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "inlinePictures", {
			get: function () {
				if (!this.m_inlinePictures) {
					this.m_inlinePictures=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
				}
				return this.m_inlinePictures;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "lists", {
			get: function () {
				if (!this.m_lists) {
					this.m_lists=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false));
				}
				return this.m_lists;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "paragraphs", {
			get: function () {
				if (!this.m_paragraphs) {
					this.m_paragraphs=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
				}
				return this.m_paragraphs;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentBody", {
			get: function () {
				if (!this.m_parentBody) {
					this.m_parentBody=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false));
				}
				return this.m_parentBody;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTableCell", {
			get: function () {
				if (!this.m_parentTableCell) {
					this.m_parentTableCell=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false));
				}
				return this.m_parentTableCell;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "tables", {
			get: function () {
				if (!this.m_tables) {
					this.m_tables=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				return this.m_tables;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "hyperlink", {
			get: function () {
				_throwIfNotLoaded("hyperlink", this.m_hyperlink, "Range", this._isNull);
				return this.m_hyperlink;
			},
			set: function (value) {
				this.m_hyperlink=value;
				_createSetPropertyAction(this.context, this, "Hyperlink", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "isEmpty", {
			get: function () {
				_throwIfNotLoaded("isEmpty", this.m_isEmpty, "Range", this._isNull);
				return this.m_isEmpty;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this.m_style, "Range", this._isNull);
				return this.m_style;
			},
			set: function (value) {
				this.m_style=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this.m_text, "Range", this._isNull);
				return this.m_text;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "Range", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Range", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Range.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		Range.prototype.compareLocationWith=function (range) {
			var action=_createMethodAction(this.context, this, "CompareLocationWith", 1 , [range]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0 , []);
		};
		Range.prototype.expandTo=function (range) {
			_createMethodAction(this.context, this, "ExpandTo", 0 , [range]);
		};
		Range.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.getHyperlinkRanges=function () {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetHyperlinkRanges", 1 , [], true, false));
		};
		Range.prototype.getNextTextRange=function (punctuationMarks, trimSpacing) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetNextTextRange", 1 , [punctuationMarks, trimSpacing], false, false));
		};
		Range.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1 , []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		Range.prototype.getTextRanges=function (punctuationMarks, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1 , [punctuationMarks, trimSpacing], true, false));
		};
		Range.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0 , [breakType, insertLocation]);
		};
		Range.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0 , [], false, true));
		};
		Range.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0 , [base64File, insertLocation], false, true));
		};
		Range.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0 , [html, insertLocation], false, true));
		};
		Range.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 , [base64EncodedImage, insertLocation], false, true));
		};
		Range.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0 , [ooxml, insertLocation], false, true));
		};
		Range.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		Range.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0 , [rowCount, columnCount, insertLocation, values], false, true));
		};
		Range.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0 , [text, insertLocation], false, true));
		};
		Range.prototype.intersectWith=function (range) {
			_createMethodAction(this.context, this, "IntersectWith", 0 , [range]);
		};
		Range.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		Range.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		Range.prototype.split=function (delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1 , [delimiters, multiParagraphs, trimDelimiters, trimSpacing], true, false));
		};
		Range.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Range.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Hyperlink"])) {
				this.m_hyperlink=obj["Hyperlink"];
			}
			if (!_isUndefined(obj["IsEmpty"])) {
				this.m_isEmpty=obj["IsEmpty"];
			}
			if (!_isUndefined(obj["Style"])) {
				this.m_style=obj["Style"];
			}
			if (!_isUndefined(obj["Text"])) {
				this.m_text=obj["Text"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "tables", "Tables"]);
		};
		Range.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Range.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Range;
	})(OfficeExtension.ClientObject);
	Word.Range=Range;
	var RangeCollection=(function (_super) {
		__extends(RangeCollection, _super);
		function RangeCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(RangeCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.Range(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "RangeCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "RangeCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		RangeCollection.prototype._GetItem=function (index) {
			return new Word.Range(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		RangeCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		RangeCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Range(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		RangeCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return RangeCollection;
	})(OfficeExtension.ClientObject);
	Word.RangeCollection=RangeCollection;
	var SearchOptions=(function (_super) {
		__extends(SearchOptions, _super);
		function SearchOptions() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(SearchOptions.prototype, "matchWildCards", {
			get: function () {
				_throwIfNotLoaded("matchWildCards", this.m_matchWildcards);
				return this.m_matchWildcards;
			},
			set: function (value) {
				this.m_matchWildcards=value;
				_createSetPropertyAction(this.context, this, "MatchWildCards", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "ignorePunct", {
			get: function () {
				_throwIfNotLoaded("ignorePunct", this.m_ignorePunct, "SearchOptions", this._isNull);
				return this.m_ignorePunct;
			},
			set: function (value) {
				this.m_ignorePunct=value;
				_createSetPropertyAction(this.context, this, "IgnorePunct", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "ignoreSpace", {
			get: function () {
				_throwIfNotLoaded("ignoreSpace", this.m_ignoreSpace, "SearchOptions", this._isNull);
				return this.m_ignoreSpace;
			},
			set: function (value) {
				this.m_ignoreSpace=value;
				_createSetPropertyAction(this.context, this, "IgnoreSpace", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchCase", {
			get: function () {
				_throwIfNotLoaded("matchCase", this.m_matchCase, "SearchOptions", this._isNull);
				return this.m_matchCase;
			},
			set: function (value) {
				this.m_matchCase=value;
				_createSetPropertyAction(this.context, this, "MatchCase", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchPrefix", {
			get: function () {
				_throwIfNotLoaded("matchPrefix", this.m_matchPrefix, "SearchOptions", this._isNull);
				return this.m_matchPrefix;
			},
			set: function (value) {
				this.m_matchPrefix=value;
				_createSetPropertyAction(this.context, this, "MatchPrefix", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchSoundsLike", {
			get: function () {
				_throwIfNotLoaded("matchSoundsLike", this.m_matchSoundsLike, "SearchOptions", this._isNull);
				return this.m_matchSoundsLike;
			},
			set: function (value) {
				this.m_matchSoundsLike=value;
				_createSetPropertyAction(this.context, this, "MatchSoundsLike", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchSuffix", {
			get: function () {
				_throwIfNotLoaded("matchSuffix", this.m_matchSuffix, "SearchOptions", this._isNull);
				return this.m_matchSuffix;
			},
			set: function (value) {
				this.m_matchSuffix=value;
				_createSetPropertyAction(this.context, this, "MatchSuffix", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchWholeWord", {
			get: function () {
				_throwIfNotLoaded("matchWholeWord", this.m_matchWholeWord, "SearchOptions", this._isNull);
				return this.m_matchWholeWord;
			},
			set: function (value) {
				this.m_matchWholeWord=value;
				_createSetPropertyAction(this.context, this, "MatchWholeWord", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchWildcards", {
			get: function () {
				_throwIfNotLoaded("matchWildcards", this.m_matchWildcards, "SearchOptions", this._isNull);
				return this.m_matchWildcards;
			},
			set: function (value) {
				this.m_matchWildcards=value;
				_createSetPropertyAction(this.context, this, "MatchWildcards", value);
			},
			enumerable: true,
			configurable: true
		});
		SearchOptions.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["IgnorePunct"])) {
				this.m_ignorePunct=obj["IgnorePunct"];
			}
			if (!_isUndefined(obj["IgnoreSpace"])) {
				this.m_ignoreSpace=obj["IgnoreSpace"];
			}
			if (!_isUndefined(obj["MatchCase"])) {
				this.m_matchCase=obj["MatchCase"];
			}
			if (!_isUndefined(obj["MatchPrefix"])) {
				this.m_matchPrefix=obj["MatchPrefix"];
			}
			if (!_isUndefined(obj["MatchSoundsLike"])) {
				this.m_matchSoundsLike=obj["MatchSoundsLike"];
			}
			if (!_isUndefined(obj["MatchSuffix"])) {
				this.m_matchSuffix=obj["MatchSuffix"];
			}
			if (!_isUndefined(obj["MatchWholeWord"])) {
				this.m_matchWholeWord=obj["MatchWholeWord"];
			}
			if (!_isUndefined(obj["MatchWildcards"])) {
				this.m_matchWildcards=obj["MatchWildcards"];
			}
		};
		SearchOptions.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		SearchOptions.newObject=function (context) {
			var ret=new Word.SearchOptions(context, _createNewObjectObjectPath(context, "Microsoft.WordServices.SearchOptions", false));
			return ret;
		};
		return SearchOptions;
	})(OfficeExtension.ClientObject);
	Word.SearchOptions=SearchOptions;
	var SearchResultCollection=(function (_super) {
		__extends(SearchResultCollection, _super);
		function SearchResultCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(SearchResultCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.Range(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchResultCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "SearchResultCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchResultCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "SearchResultCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		SearchResultCollection.prototype._GetItem=function (index) {
			return new Word.Range(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		SearchResultCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		SearchResultCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Range(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		SearchResultCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		SearchResultCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return SearchResultCollection;
	})(OfficeExtension.ClientObject);
	Word.SearchResultCollection=SearchResultCollection;
	var Section=(function (_super) {
		__extends(Section, _super);
		function Section() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Section.prototype, "body", {
			get: function () {
				if (!this.m_body) {
					this.m_body=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false));
				}
				return this.m_body;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.Section(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "Section", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Section", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Section.prototype.getFooter=function (type) {
			return new Word.Body(this.context, _createMethodObjectPath(this.context, this, "GetFooter", 1 , [type], false, true));
		};
		Section.prototype.getHeader=function (type) {
			return new Word.Body(this.context, _createMethodObjectPath(this.context, this, "GetHeader", 1 , [type], false, true));
		};
		Section.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Section.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "next", "Next"]);
		};
		Section.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Section.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Section;
	})(OfficeExtension.ClientObject);
	Word.Section=Section;
	var SectionCollection=(function (_super) {
		__extends(SectionCollection, _super);
		function SectionCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(SectionCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.Section(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "SectionCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "SectionCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		SectionCollection.prototype._GetItem=function (index) {
			return new Word.Section(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		SectionCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		SectionCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Section(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		SectionCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		SectionCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return SectionCollection;
	})(OfficeExtension.ClientObject);
	Word.SectionCollection=SectionCollection;
	var Table=(function (_super) {
		__extends(Table, _super);
		function Table() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(Table.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "paragraphAfter", {
			get: function () {
				if (!this.m_paragraphAfter) {
					this.m_paragraphAfter=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "ParagraphAfter", false, false));
				}
				return this.m_paragraphAfter;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "paragraphBefore", {
			get: function () {
				if (!this.m_paragraphBefore) {
					this.m_paragraphBefore=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "ParagraphBefore", false, false));
				}
				return this.m_paragraphBefore;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentContentControl", {
			get: function () {
				if (!this.m_parentContentControl) {
					this.m_parentContentControl=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
				}
				return this.m_parentContentControl;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTableCell", {
			get: function () {
				if (!this.m_parentTableCell) {
					this.m_parentTableCell=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false));
				}
				return this.m_parentTableCell;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rows", {
			get: function () {
				if (!this.m_rows) {
					this.m_rows=new Word.TableRowCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
				}
				return this.m_rows;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "tables", {
			get: function () {
				if (!this.m_tables) {
					this.m_tables=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				return this.m_tables;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "cellPaddingBottom", {
			get: function () {
				_throwIfNotLoaded("cellPaddingBottom", this.m_cellPaddingBottom, "Table", this._isNull);
				return this.m_cellPaddingBottom;
			},
			set: function (value) {
				this.m_cellPaddingBottom=value;
				_createSetPropertyAction(this.context, this, "CellPaddingBottom", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "cellPaddingLeft", {
			get: function () {
				_throwIfNotLoaded("cellPaddingLeft", this.m_cellPaddingLeft, "Table", this._isNull);
				return this.m_cellPaddingLeft;
			},
			set: function (value) {
				this.m_cellPaddingLeft=value;
				_createSetPropertyAction(this.context, this, "CellPaddingLeft", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "cellPaddingRight", {
			get: function () {
				_throwIfNotLoaded("cellPaddingRight", this.m_cellPaddingRight, "Table", this._isNull);
				return this.m_cellPaddingRight;
			},
			set: function (value) {
				this.m_cellPaddingRight=value;
				_createSetPropertyAction(this.context, this, "CellPaddingRight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "cellPaddingTop", {
			get: function () {
				_throwIfNotLoaded("cellPaddingTop", this.m_cellPaddingTop, "Table", this._isNull);
				return this.m_cellPaddingTop;
			},
			set: function (value) {
				this.m_cellPaddingTop=value;
				_createSetPropertyAction(this.context, this, "CellPaddingTop", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "headerRowCount", {
			get: function () {
				_throwIfNotLoaded("headerRowCount", this.m_headerRowCount, "Table", this._isNull);
				return this.m_headerRowCount;
			},
			set: function (value) {
				this.m_headerRowCount=value;
				_createSetPropertyAction(this.context, this, "HeaderRowCount", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "height", {
			get: function () {
				_throwIfNotLoaded("height", this.m_height, "Table", this._isNull);
				return this.m_height;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "isUniform", {
			get: function () {
				_throwIfNotLoaded("isUniform", this.m_isUniform, "Table", this._isNull);
				return this.m_isUniform;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "nestingLevel", {
			get: function () {
				_throwIfNotLoaded("nestingLevel", this.m_nestingLevel, "Table", this._isNull);
				return this.m_nestingLevel;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rowCount", {
			get: function () {
				_throwIfNotLoaded("rowCount", this.m_rowCount, "Table", this._isNull);
				return this.m_rowCount;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this.m_shadingColor, "Table", this._isNull);
				return this.m_shadingColor;
			},
			set: function (value) {
				this.m_shadingColor=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this.m_style, "Table", this._isNull);
				return this.m_style;
			},
			set: function (value) {
				this.m_style=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleBandedColumns", {
			get: function () {
				_throwIfNotLoaded("styleBandedColumns", this.m_styleBandedColumns, "Table", this._isNull);
				return this.m_styleBandedColumns;
			},
			set: function (value) {
				this.m_styleBandedColumns=value;
				_createSetPropertyAction(this.context, this, "StyleBandedColumns", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleBandedRows", {
			get: function () {
				_throwIfNotLoaded("styleBandedRows", this.m_styleBandedRows, "Table", this._isNull);
				return this.m_styleBandedRows;
			},
			set: function (value) {
				this.m_styleBandedRows=value;
				_createSetPropertyAction(this.context, this, "StyleBandedRows", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleFirstColumn", {
			get: function () {
				_throwIfNotLoaded("styleFirstColumn", this.m_styleFirstColumn, "Table", this._isNull);
				return this.m_styleFirstColumn;
			},
			set: function (value) {
				this.m_styleFirstColumn=value;
				_createSetPropertyAction(this.context, this, "StyleFirstColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleLastColumn", {
			get: function () {
				_throwIfNotLoaded("styleLastColumn", this.m_styleLastColumn, "Table", this._isNull);
				return this.m_styleLastColumn;
			},
			set: function (value) {
				this.m_styleLastColumn=value;
				_createSetPropertyAction(this.context, this, "StyleLastColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleTotalRow", {
			get: function () {
				_throwIfNotLoaded("styleTotalRow", this.m_styleTotalRow, "Table", this._isNull);
				return this.m_styleTotalRow;
			},
			set: function (value) {
				this.m_styleTotalRow=value;
				_createSetPropertyAction(this.context, this, "StyleTotalRow", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this.m_values, "Table", this._isNull);
				return this.m_values;
			},
			set: function (value) {
				this.m_values=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this.m_verticalAlignment, "Table", this._isNull);
				return this.m_verticalAlignment;
			},
			set: function (value) {
				this.m_verticalAlignment=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this.m_width, "Table", this._isNull);
				return this.m_width;
			},
			set: function (value) {
				this.m_width=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "Table", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "Table", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Table.prototype.addColumns=function (insertLocation, columnCount, values) {
			_createMethodAction(this.context, this, "AddColumns", 0 , [insertLocation, columnCount, values]);
		};
		Table.prototype.addRows=function (insertLocation, rowCount, values) {
			_createMethodAction(this.context, this, "AddRows", 0 , [insertLocation, rowCount, values]);
		};
		Table.prototype.autoFitContents=function () {
			_createMethodAction(this.context, this, "AutoFitContents", 0 , []);
		};
		Table.prototype.autoFitWindow=function () {
			_createMethodAction(this.context, this, "AutoFitWindow", 0 , []);
		};
		Table.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		Table.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0 , []);
		};
		Table.prototype.deleteColumns=function (columnIndex, columnCount) {
			_createMethodAction(this.context, this, "DeleteColumns", 0 , [columnIndex, columnCount]);
		};
		Table.prototype.deleteRows=function (rowIndex, rowCount) {
			_createMethodAction(this.context, this, "DeleteRows", 0 , [rowIndex, rowCount]);
		};
		Table.prototype.distributeColumns=function () {
			_createMethodAction(this.context, this, "DistributeColumns", 0 , []);
		};
		Table.prototype.distributeRows=function () {
			_createMethodAction(this.context, this, "DistributeRows", 0 , []);
		};
		Table.prototype.getBorderStyle=function (borderLocation) {
			return new Word.TableBorderStyle(this.context, _createMethodObjectPath(this.context, this, "GetBorderStyle", 0 , [borderLocation], false, false));
		};
		Table.prototype.getCell=function (rowIndex, cellIndex) {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetCell", 0 , [rowIndex, cellIndex], false, false));
		};
		Table.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1 , [rangeLocation], false, false));
		};
		Table.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0 , [], false, true));
		};
		Table.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0 , [paragraphText, insertLocation], false, true));
		};
		Table.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0 , [rowCount, columnCount, insertLocation, values], false, true));
		};
		Table.prototype.mergeCells=function (topRow, firstCell, bottomRow, lastCell) {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "MergeCells", 0 , [topRow, firstCell, bottomRow, lastCell], false, true));
		};
		Table.prototype.search=function (searchText, searchOptions) {
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		Table.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		Table.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		Table.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellPaddingBottom"])) {
				this.m_cellPaddingBottom=obj["CellPaddingBottom"];
			}
			if (!_isUndefined(obj["CellPaddingLeft"])) {
				this.m_cellPaddingLeft=obj["CellPaddingLeft"];
			}
			if (!_isUndefined(obj["CellPaddingRight"])) {
				this.m_cellPaddingRight=obj["CellPaddingRight"];
			}
			if (!_isUndefined(obj["CellPaddingTop"])) {
				this.m_cellPaddingTop=obj["CellPaddingTop"];
			}
			if (!_isUndefined(obj["HeaderRowCount"])) {
				this.m_headerRowCount=obj["HeaderRowCount"];
			}
			if (!_isUndefined(obj["Height"])) {
				this.m_height=obj["Height"];
			}
			if (!_isUndefined(obj["IsUniform"])) {
				this.m_isUniform=obj["IsUniform"];
			}
			if (!_isUndefined(obj["NestingLevel"])) {
				this.m_nestingLevel=obj["NestingLevel"];
			}
			if (!_isUndefined(obj["RowCount"])) {
				this.m_rowCount=obj["RowCount"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this.m_shadingColor=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Style"])) {
				this.m_style=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBandedColumns"])) {
				this.m_styleBandedColumns=obj["StyleBandedColumns"];
			}
			if (!_isUndefined(obj["StyleBandedRows"])) {
				this.m_styleBandedRows=obj["StyleBandedRows"];
			}
			if (!_isUndefined(obj["StyleFirstColumn"])) {
				this.m_styleFirstColumn=obj["StyleFirstColumn"];
			}
			if (!_isUndefined(obj["StyleLastColumn"])) {
				this.m_styleLastColumn=obj["StyleLastColumn"];
			}
			if (!_isUndefined(obj["StyleTotalRow"])) {
				this.m_styleTotalRow=obj["StyleTotalRow"];
			}
			if (!_isUndefined(obj["Values"])) {
				this.m_values=obj["Values"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this.m_verticalAlignment=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["Width"])) {
				this.m_width=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["font", "Font", "next", "Next", "paragraphAfter", "ParagraphAfter", "paragraphBefore", "ParagraphBefore", "parentContentControl", "ParentContentControl", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "rows", "Rows", "tables", "Tables"]);
		};
		Table.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Table.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return Table;
	})(OfficeExtension.ClientObject);
	Word.Table=Table;
	var TableCollection=(function (_super) {
		__extends(TableCollection, _super);
		function TableCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "TableCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableCollection.prototype._GetItem=function (index) {
			return new Word.Table(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Table(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableCollection;
	})(OfficeExtension.ClientObject);
	Word.TableCollection=TableCollection;
	var TableRow=(function (_super) {
		__extends(TableRow, _super);
		function TableRow() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableRow.prototype, "cells", {
			get: function () {
				if (!this.m_cells) {
					this.m_cells=new Word.TableCellCollection(this.context, _createPropertyObjectPath(this.context, this, "Cells", true, false));
				}
				return this.m_cells;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "font", {
			get: function () {
				if (!this.m_font) {
					this.m_font=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this.m_font;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.TableRow(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellCount", {
			get: function () {
				_throwIfNotLoaded("cellCount", this.m_cellCount, "TableRow", this._isNull);
				return this.m_cellCount;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellPaddingBottom", {
			get: function () {
				_throwIfNotLoaded("cellPaddingBottom", this.m_cellPaddingBottom, "TableRow", this._isNull);
				return this.m_cellPaddingBottom;
			},
			set: function (value) {
				this.m_cellPaddingBottom=value;
				_createSetPropertyAction(this.context, this, "CellPaddingBottom", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellPaddingLeft", {
			get: function () {
				_throwIfNotLoaded("cellPaddingLeft", this.m_cellPaddingLeft, "TableRow", this._isNull);
				return this.m_cellPaddingLeft;
			},
			set: function (value) {
				this.m_cellPaddingLeft=value;
				_createSetPropertyAction(this.context, this, "CellPaddingLeft", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellPaddingRight", {
			get: function () {
				_throwIfNotLoaded("cellPaddingRight", this.m_cellPaddingRight, "TableRow", this._isNull);
				return this.m_cellPaddingRight;
			},
			set: function (value) {
				this.m_cellPaddingRight=value;
				_createSetPropertyAction(this.context, this, "CellPaddingRight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellPaddingTop", {
			get: function () {
				_throwIfNotLoaded("cellPaddingTop", this.m_cellPaddingTop, "TableRow", this._isNull);
				return this.m_cellPaddingTop;
			},
			set: function (value) {
				this.m_cellPaddingTop=value;
				_createSetPropertyAction(this.context, this, "CellPaddingTop", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "isHeader", {
			get: function () {
				_throwIfNotLoaded("isHeader", this.m_isHeader, "TableRow", this._isNull);
				return this.m_isHeader;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "preferredHeight", {
			get: function () {
				_throwIfNotLoaded("preferredHeight", this.m_preferredHeight, "TableRow", this._isNull);
				return this.m_preferredHeight;
			},
			set: function (value) {
				this.m_preferredHeight=value;
				_createSetPropertyAction(this.context, this, "PreferredHeight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "rowIndex", {
			get: function () {
				_throwIfNotLoaded("rowIndex", this.m_rowIndex, "TableRow", this._isNull);
				return this.m_rowIndex;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this.m_shadingColor, "TableRow", this._isNull);
				return this.m_shadingColor;
			},
			set: function (value) {
				this.m_shadingColor=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this.m_values, "TableRow", this._isNull);
				return this.m_values;
			},
			set: function (value) {
				this.m_values=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this.m_verticalAlignment, "TableRow", this._isNull);
				return this.m_verticalAlignment;
			},
			set: function (value) {
				this.m_verticalAlignment=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "TableRow", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableRow", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableRow.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0 , []);
		};
		TableRow.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0 , []);
		};
		TableRow.prototype.getBorderStyle=function (borderLocation) {
			return new Word.TableBorderStyle(this.context, _createMethodObjectPath(this.context, this, "GetBorderStyle", 0 , [borderLocation], false, false));
		};
		TableRow.prototype.insertRows=function (insertLocation, rowCount, values) {
			_createMethodAction(this.context, this, "InsertRows", 1 , [insertLocation, rowCount, values]);
		};
		TableRow.prototype.merge=function () {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "Merge", 0 , [], false, false));
		};
		TableRow.prototype.search=function (searchText, searchOptions) {
			return new Word.SearchResultCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1 , [searchText, searchOptions], true, true));
		};
		TableRow.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1 , [selectionMode]);
		};
		TableRow.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableRow.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellCount"])) {
				this.m_cellCount=obj["CellCount"];
			}
			if (!_isUndefined(obj["CellPaddingBottom"])) {
				this.m_cellPaddingBottom=obj["CellPaddingBottom"];
			}
			if (!_isUndefined(obj["CellPaddingLeft"])) {
				this.m_cellPaddingLeft=obj["CellPaddingLeft"];
			}
			if (!_isUndefined(obj["CellPaddingRight"])) {
				this.m_cellPaddingRight=obj["CellPaddingRight"];
			}
			if (!_isUndefined(obj["CellPaddingTop"])) {
				this.m_cellPaddingTop=obj["CellPaddingTop"];
			}
			if (!_isUndefined(obj["IsHeader"])) {
				this.m_isHeader=obj["IsHeader"];
			}
			if (!_isUndefined(obj["PreferredHeight"])) {
				this.m_preferredHeight=obj["PreferredHeight"];
			}
			if (!_isUndefined(obj["RowIndex"])) {
				this.m_rowIndex=obj["RowIndex"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this.m_shadingColor=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Values"])) {
				this.m_values=obj["Values"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this.m_verticalAlignment=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["cells", "Cells", "font", "Font", "next", "Next", "parentTable", "ParentTable"]);
		};
		TableRow.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableRow.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableRow;
	})(OfficeExtension.ClientObject);
	Word.TableRow=TableRow;
	var TableRowCollection=(function (_super) {
		__extends(TableRowCollection, _super);
		function TableRowCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableRowCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.TableRow(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "TableRowCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableRowCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableRowCollection.prototype._GetItem=function (index) {
			return new Word.TableRow(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableRowCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableRowCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.TableRow(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableRowCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableRowCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableRowCollection;
	})(OfficeExtension.ClientObject);
	Word.TableRowCollection=TableRowCollection;
	var TableCell=(function (_super) {
		__extends(TableCell, _super);
		function TableCell() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableCell.prototype, "body", {
			get: function () {
				if (!this.m_body) {
					this.m_body=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false));
				}
				return this.m_body;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "next", {
			get: function () {
				if (!this.m_next) {
					this.m_next=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "Next", false, false));
				}
				return this.m_next;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "parentRow", {
			get: function () {
				if (!this.m_parentRow) {
					this.m_parentRow=new Word.TableRow(this.context, _createPropertyObjectPath(this.context, this, "ParentRow", false, false));
				}
				return this.m_parentRow;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "parentTable", {
			get: function () {
				if (!this.m_parentTable) {
					this.m_parentTable=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false));
				}
				return this.m_parentTable;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellIndex", {
			get: function () {
				_throwIfNotLoaded("cellIndex", this.m_cellIndex, "TableCell", this._isNull);
				return this.m_cellIndex;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellPaddingBottom", {
			get: function () {
				_throwIfNotLoaded("cellPaddingBottom", this.m_cellPaddingBottom, "TableCell", this._isNull);
				return this.m_cellPaddingBottom;
			},
			set: function (value) {
				this.m_cellPaddingBottom=value;
				_createSetPropertyAction(this.context, this, "CellPaddingBottom", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellPaddingLeft", {
			get: function () {
				_throwIfNotLoaded("cellPaddingLeft", this.m_cellPaddingLeft, "TableCell", this._isNull);
				return this.m_cellPaddingLeft;
			},
			set: function (value) {
				this.m_cellPaddingLeft=value;
				_createSetPropertyAction(this.context, this, "CellPaddingLeft", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellPaddingRight", {
			get: function () {
				_throwIfNotLoaded("cellPaddingRight", this.m_cellPaddingRight, "TableCell", this._isNull);
				return this.m_cellPaddingRight;
			},
			set: function (value) {
				this.m_cellPaddingRight=value;
				_createSetPropertyAction(this.context, this, "CellPaddingRight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellPaddingTop", {
			get: function () {
				_throwIfNotLoaded("cellPaddingTop", this.m_cellPaddingTop, "TableCell", this._isNull);
				return this.m_cellPaddingTop;
			},
			set: function (value) {
				this.m_cellPaddingTop=value;
				_createSetPropertyAction(this.context, this, "CellPaddingTop", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "columnWidth", {
			get: function () {
				_throwIfNotLoaded("columnWidth", this.m_columnWidth, "TableCell", this._isNull);
				return this.m_columnWidth;
			},
			set: function (value) {
				this.m_columnWidth=value;
				_createSetPropertyAction(this.context, this, "ColumnWidth", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "rowIndex", {
			get: function () {
				_throwIfNotLoaded("rowIndex", this.m_rowIndex, "TableCell", this._isNull);
				return this.m_rowIndex;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this.m_shadingColor, "TableCell", this._isNull);
				return this.m_shadingColor;
			},
			set: function (value) {
				this.m_shadingColor=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, "TableCell", this._isNull);
				return this.m_value;
			},
			set: function (value) {
				this.m_value=value;
				_createSetPropertyAction(this.context, this, "Value", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this.m_verticalAlignment, "TableCell", this._isNull);
				return this.m_verticalAlignment;
			},
			set: function (value) {
				this.m_verticalAlignment=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this.m_width, "TableCell", this._isNull);
				return this.m_width;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.m__Id, "TableCell", this._isNull);
				return this.m__Id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableCell", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableCell.prototype.deleteColumn=function () {
			_createMethodAction(this.context, this, "DeleteColumn", 0 , []);
		};
		TableCell.prototype.deleteRow=function () {
			_createMethodAction(this.context, this, "DeleteRow", 0 , []);
		};
		TableCell.prototype.getBorderStyle=function (borderLocation) {
			return new Word.TableBorderStyle(this.context, _createMethodObjectPath(this.context, this, "GetBorderStyle", 0 , [borderLocation], false, false));
		};
		TableCell.prototype.insertColumns=function (insertLocation, columnCount, values) {
			_createMethodAction(this.context, this, "InsertColumns", 0 , [insertLocation, columnCount, values]);
		};
		TableCell.prototype.insertRows=function (insertLocation, rowCount, values) {
			_createMethodAction(this.context, this, "InsertRows", 0 , [insertLocation, rowCount, values]);
		};
		TableCell.prototype.split=function (rowCount, columnCount) {
			_createMethodAction(this.context, this, "Split", 0 , [rowCount, columnCount]);
		};
		TableCell.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableCell.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellIndex"])) {
				this.m_cellIndex=obj["CellIndex"];
			}
			if (!_isUndefined(obj["CellPaddingBottom"])) {
				this.m_cellPaddingBottom=obj["CellPaddingBottom"];
			}
			if (!_isUndefined(obj["CellPaddingLeft"])) {
				this.m_cellPaddingLeft=obj["CellPaddingLeft"];
			}
			if (!_isUndefined(obj["CellPaddingRight"])) {
				this.m_cellPaddingRight=obj["CellPaddingRight"];
			}
			if (!_isUndefined(obj["CellPaddingTop"])) {
				this.m_cellPaddingTop=obj["CellPaddingTop"];
			}
			if (!_isUndefined(obj["ColumnWidth"])) {
				this.m_columnWidth=obj["ColumnWidth"];
			}
			if (!_isUndefined(obj["RowIndex"])) {
				this.m_rowIndex=obj["RowIndex"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this.m_shadingColor=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this.m_verticalAlignment=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["Width"])) {
				this.m_width=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.m__Id=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "next", "Next", "parentRow", "ParentRow", "parentTable", "ParentTable"]);
		};
		TableCell.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableCell.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableCell;
	})(OfficeExtension.ClientObject);
	Word.TableCell=TableCell;
	var TableCellCollection=(function (_super) {
		__extends(TableCellCollection, _super);
		function TableCellCollection() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableCellCollection.prototype, "first", {
			get: function () {
				if (!this.m_first) {
					this.m_first=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "First", false, false));
				}
				return this.m_first;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, "TableCellCollection", this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableCellCollection", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableCellCollection.prototype._GetItem=function (index) {
			return new Word.TableCell(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableCellCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableCellCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["first", "First"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.TableCell(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableCellCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableCellCollection.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableCellCollection;
	})(OfficeExtension.ClientObject);
	Word.TableCellCollection=TableCellCollection;
	var TableBorderStyle=(function (_super) {
		__extends(TableBorderStyle, _super);
		function TableBorderStyle() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(TableBorderStyle.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this.m_color, "TableBorderStyle", this._isNull);
				return this.m_color;
			},
			set: function (value) {
				this.m_color=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorderStyle.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this.m_type, "TableBorderStyle", this._isNull);
				return this.m_type;
			},
			set: function (value) {
				this.m_type=value;
				_createSetPropertyAction(this.context, this, "Type", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorderStyle.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this.m_width, "TableBorderStyle", this._isNull);
				return this.m_width;
			},
			set: function (value) {
				this.m_width=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorderStyle.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, "TableBorderStyle", this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		TableBorderStyle.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1 , []);
		};
		TableBorderStyle.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this.m_color=obj["Color"];
			}
			if (!_isUndefined(obj["Type"])) {
				this.m_type=obj["Type"];
			}
			if (!_isUndefined(obj["Width"])) {
				this.m_width=obj["Width"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
		};
		TableBorderStyle.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableBorderStyle.prototype._initReferenceId=function (value) {
			this.m__ReferenceId=value;
		};
		return TableBorderStyle;
	})(OfficeExtension.ClientObject);
	Word.TableBorderStyle=TableBorderStyle;
	var ContentControlType;
	(function (ContentControlType) {
		ContentControlType.unknown="Unknown";
		ContentControlType.richTextInline="RichTextInline";
		ContentControlType.richTextParagraphs="RichTextParagraphs";
		ContentControlType.richTextTableCell="RichTextTableCell";
		ContentControlType.richTextTableRow="RichTextTableRow";
		ContentControlType.richTextTable="RichTextTable";
		ContentControlType.plainTextInline="PlainTextInline";
		ContentControlType.plainTextParagraph="PlainTextParagraph";
		ContentControlType.picture="Picture";
		ContentControlType.buildingBlockGallery="BuildingBlockGallery";
		ContentControlType.checkBox="CheckBox";
		ContentControlType.comboBox="ComboBox";
		ContentControlType.dropDownList="DropDownList";
		ContentControlType.datePicker="DatePicker";
		ContentControlType.repeatingSection="RepeatingSection";
		ContentControlType.richText="RichText";
		ContentControlType.plainText="PlainText";
	})(ContentControlType=Word.ContentControlType || (Word.ContentControlType={}));
	var ContentControlAppearance;
	(function (ContentControlAppearance) {
		ContentControlAppearance.boundingBox="BoundingBox";
		ContentControlAppearance.tags="Tags";
		ContentControlAppearance.hidden="Hidden";
	})(ContentControlAppearance=Word.ContentControlAppearance || (Word.ContentControlAppearance={}));
	var UnderlineType;
	(function (UnderlineType) {
		UnderlineType.none="None";
		UnderlineType.single="Single";
		UnderlineType.word="Word";
		UnderlineType.double="Double";
		UnderlineType.dotted="Dotted";
		UnderlineType.hidden="Hidden";
		UnderlineType.thick="Thick";
		UnderlineType.dashLine="DashLine";
		UnderlineType.dotLine="DotLine";
		UnderlineType.dotDashLine="DotDashLine";
		UnderlineType.twoDotDashLine="TwoDotDashLine";
		UnderlineType.wave="Wave";
	})(UnderlineType=Word.UnderlineType || (Word.UnderlineType={}));
	var BreakType;
	(function (BreakType) {
		BreakType.page="Page";
		BreakType.column="Column";
		BreakType.next="Next";
		BreakType.sectionContinuous="SectionContinuous";
		BreakType.sectionEven="SectionEven";
		BreakType.sectionOdd="SectionOdd";
		BreakType.line="Line";
		BreakType.lineClearLeft="LineClearLeft";
		BreakType.lineClearRight="LineClearRight";
		BreakType.textWrapping="TextWrapping";
	})(BreakType=Word.BreakType || (Word.BreakType={}));
	var InsertLocation;
	(function (InsertLocation) {
		InsertLocation.before="Before";
		InsertLocation.after="After";
		InsertLocation.start="Start";
		InsertLocation.end="End";
		InsertLocation.replace="Replace";
	})(InsertLocation=Word.InsertLocation || (Word.InsertLocation={}));
	var Alignment;
	(function (Alignment) {
		Alignment.unknown="Unknown";
		Alignment.left="Left";
		Alignment.centered="Centered";
		Alignment.right="Right";
		Alignment.justified="Justified";
	})(Alignment=Word.Alignment || (Word.Alignment={}));
	var HeaderFooterType;
	(function (HeaderFooterType) {
		HeaderFooterType.primary="Primary";
		HeaderFooterType.firstPage="FirstPage";
		HeaderFooterType.evenPages="EvenPages";
	})(HeaderFooterType=Word.HeaderFooterType || (Word.HeaderFooterType={}));
	var BodyType;
	(function (BodyType) {
		BodyType.unknown="Unknown";
		BodyType.mainDoc="MainDoc";
		BodyType.section="Section";
		BodyType.header="Header";
		BodyType.footer="Footer";
		BodyType.tableCell="TableCell";
	})(BodyType=Word.BodyType || (Word.BodyType={}));
	var SelectionMode;
	(function (SelectionMode) {
		SelectionMode.select="Select";
		SelectionMode.start="Start";
		SelectionMode.end="End";
	})(SelectionMode=Word.SelectionMode || (Word.SelectionMode={}));
	var ImageFormat;
	(function (ImageFormat) {
		ImageFormat.unsupported="Unsupported";
		ImageFormat.undefined="Undefined";
		ImageFormat.bmp="Bmp";
		ImageFormat.jpeg="Jpeg";
		ImageFormat.gif="Gif";
		ImageFormat.tiff="Tiff";
		ImageFormat.png="Png";
		ImageFormat.icon="Icon";
		ImageFormat.exif="Exif";
		ImageFormat.wmf="Wmf";
		ImageFormat.emf="Emf";
		ImageFormat.pict="Pict";
		ImageFormat.pdf="Pdf";
	})(ImageFormat=Word.ImageFormat || (Word.ImageFormat={}));
	var RangeLocation;
	(function (RangeLocation) {
		RangeLocation.whole="Whole";
		RangeLocation.start="Start";
		RangeLocation.end="End";
	})(RangeLocation=Word.RangeLocation || (Word.RangeLocation={}));
	var LocationRelation;
	(function (LocationRelation) {
		LocationRelation.unrelated="Unrelated";
		LocationRelation.equal="Equal";
		LocationRelation.containsStart="ContainsStart";
		LocationRelation.containsEnd="ContainsEnd";
		LocationRelation.contains="Contains";
		LocationRelation.insideStart="InsideStart";
		LocationRelation.insideEnd="InsideEnd";
		LocationRelation.inside="Inside";
		LocationRelation.adjacentBefore="AdjacentBefore";
		LocationRelation.overlapsBefore="OverlapsBefore";
		LocationRelation.before="Before";
		LocationRelation.adjacentAfter="AdjacentAfter";
		LocationRelation.overlapsAfter="OverlapsAfter";
		LocationRelation.after="After";
	})(LocationRelation=Word.LocationRelation || (Word.LocationRelation={}));
	var BorderLocation;
	(function (BorderLocation) {
		BorderLocation.top="Top";
		BorderLocation.left="Left";
		BorderLocation.bottom="Bottom";
		BorderLocation.right="Right";
		BorderLocation.insideHorizontal="InsideHorizontal";
		BorderLocation.insideVertical="InsideVertical";
		BorderLocation.inside="Inside";
		BorderLocation.outside="Outside";
		BorderLocation.all="All";
	})(BorderLocation=Word.BorderLocation || (Word.BorderLocation={}));
	var BorderType;
	(function (BorderType) {
		BorderType.mixed="Mixed";
		BorderType.none="None";
		BorderType.single="Single";
		BorderType.thick="Thick";
		BorderType.double="Double";
		BorderType.hairline="Hairline";
		BorderType.dotted="Dotted";
		BorderType.dashed="Dashed";
		BorderType.dotDashed="DotDashed";
		BorderType.dot2Dashed="Dot2Dashed";
		BorderType.triple="Triple";
		BorderType.thinThickSmall="ThinThickSmall";
		BorderType.thickThinSmall="ThickThinSmall";
		BorderType.thinThickThinSmall="ThinThickThinSmall";
		BorderType.thinThickMed="ThinThickMed";
		BorderType.thickThinMed="ThickThinMed";
		BorderType.thinThickThinMed="ThinThickThinMed";
		BorderType.thinThickLarge="ThinThickLarge";
		BorderType.thickThinLarge="ThickThinLarge";
		BorderType.thinThickThinLarge="ThinThickThinLarge";
		BorderType.wave="Wave";
		BorderType.doubleWave="DoubleWave";
		BorderType.dashedSmall="DashedSmall";
		BorderType.dashDotStroked="DashDotStroked";
		BorderType.threeDEmboss="ThreeDEmboss";
		BorderType.threeDEngrave="ThreeDEngrave";
	})(BorderType=Word.BorderType || (Word.BorderType={}));
	var VerticalAlignment;
	(function (VerticalAlignment) {
		VerticalAlignment.mixed="Mixed";
		VerticalAlignment.top="Top";
		VerticalAlignment.center="Center";
		VerticalAlignment.bottom="Bottom";
	})(VerticalAlignment=Word.VerticalAlignment || (Word.VerticalAlignment={}));
	var ListLevelType;
	(function (ListLevelType) {
		ListLevelType.bullet="Bullet";
		ListLevelType.number="Number";
		ListLevelType.picture="Picture";
	})(ListLevelType=Word.ListLevelType || (Word.ListLevelType={}));
	var ListBullet;
	(function (ListBullet) {
		ListBullet.custom="Custom";
		ListBullet.solid="Solid";
		ListBullet.hollow="Hollow";
		ListBullet.square="Square";
		ListBullet.diamonds="Diamonds";
		ListBullet.arrow="Arrow";
		ListBullet.checkmark="Checkmark";
	})(ListBullet=Word.ListBullet || (Word.ListBullet={}));
	var ListNumbering;
	(function (ListNumbering) {
		ListNumbering.none="None";
		ListNumbering.arabic="Arabic";
		ListNumbering.upperRoman="UpperRoman";
		ListNumbering.lowerRoman="LowerRoman";
		ListNumbering.upperLetter="UpperLetter";
		ListNumbering.lowerLetter="LowerLetter";
		ListNumbering.ordinal="Ordinal";
		ListNumbering.cardinalText="CardinalText";
		ListNumbering.ordinalText="OrdinalText";
	})(ListNumbering=Word.ListNumbering || (Word.ListNumbering={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.itemNotFound="ItemNotFound";
		ErrorCodes.notImplemented="NotImplemented";
	})(ErrorCodes=Word.ErrorCodes || (Word.ErrorCodes={}));
})(Word || (Word={}));
var Word;
(function (Word) {
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			_super.call(this, url);
			this.m_document=new Word.Document(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
			this._rootObject=this.m_document;
		}
		Object.defineProperty(RequestContext.prototype, "document", {
			get: function () {
				return this.m_document;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "application", {
			get: function () {
				if (this.m_application==null) {
					this.m_application=new Word.Application(this, OfficeExtension.ObjectPathFactory.createNewObjectObjectPath(this, "Microsoft.WordServices.Application", false));
				}
				return this.m_application;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	})(OfficeExtension.ClientRequestContext);
	Word.RequestContext=RequestContext;
	function run(batch) {
		return OfficeExtension.ClientRequestContext._run(function () { return new Word.RequestContext(); }, batch);
	}
	Word.run=run;
})(Word || (Word={}));


