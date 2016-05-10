/*!---------------------------------------------------------
* Copyright (C) Microsoft Corporation. All rights reserved.
*--------------------------------------------------------*/
"use strict";var _amdLoaderGlobal=this,AMDLoader;!function(e){var t=_amdLoaderGlobal,n="function"==typeof t.importScripts;if("function"!=typeof t.define||!t.define.amd){!function(){t.console||(t.console={}),t.console.log||(t.console.log=function(){}),t.console.warn||(t.console.warn=t.console.log),t.console.error||(t.console.error=t.console.log)}();var r=function(){function e(){}return e.startsWith=function(e,t){return e.length>=t.length&&e.substr(0,t.length)===t},e.endsWith=function(e,t){return e.length>=t.length&&e.substr(e.length-t.length)===t},e.containsQueryString=function(e){return e.indexOf("?")>=0},e.isAbsolutePath=function(t){return e.startsWith(t,"http://")||e.startsWith(t,"https://")||e.startsWith(t,"/")},e.forEachProperty=function(e,t){if(e){var n;for(n in e)e.hasOwnProperty(n)&&t(n,e[n])}},e.isEmpty=function(t){var n=!0;return e.forEachProperty(t,function(){n=!1}),n},e.isArray=function(e){return Array.isArray?Array.isArray(e):"[object Array]"===Object.prototype.toString.call(e)},e.recursiveClone=function(t){if(!t||"object"!=typeof t)return t;var n=e.isArray(t)?[]:{};return e.forEachProperty(t,function(t,r){n[t]=r&&"object"==typeof r?e.recursiveClone(r):r}),n},e.generateAnonymousModule=function(){return"===anonymous"+e.NEXT_ANONYMOUS_ID++ +"==="},e.isAnonymousModule=function(e){return 0===e.indexOf("===anonymous")},e.NEXT_ANONYMOUS_ID=1,e}(),i=function(){function e(){}return e.validateConfigurationOptions=function(e){function t(e){return"load"===e.errorCode?(console.error('Loading "'+e.moduleId+'" failed'),console.error("Detail: ",e.detail),console.error("Here are the modules that depend on it:"),console.error(e.neededBy),void 0):"factory"===e.errorCode?(console.error('The factory method of "'+e.moduleId+'" has thrown an exception'),console.error(e.detail),void 0):void 0}return e=e||{},"string"!=typeof e.baseUrl&&(e.baseUrl=""),"object"!=typeof e.paths&&(e.paths={}),"object"!=typeof e.bundles&&(e.bundles=[]),"object"!=typeof e.shim&&(e.shim={}),"object"!=typeof e.config&&(e.config={}),"undefined"==typeof e.catchError&&(e.catchError=n),"string"!=typeof e.urlArgs&&(e.urlArgs=""),"function"!=typeof e.onError&&(e.onError=t),"object"==typeof e.ignoreDuplicateModules&&r.isArray(e.ignoreDuplicateModules)||(e.ignoreDuplicateModules=[]),e.baseUrl.length>0&&(r.endsWith(e.baseUrl,"/")||(e.baseUrl+="/")),e},e.mergeConfigurationOptions=function(t,n){"undefined"==typeof t&&(t=null),"undefined"==typeof n&&(n=null);var i=r.recursiveClone(n||{});return r.forEachProperty(t,function(e,t){"bundles"===e&&"undefined"!=typeof i.bundles?r.isArray(t)?i.bundles=i.bundles.concat(t):r.forEachProperty(t,function(e,t){var n={location:e,modules:t};i.bundles.push(n)}):"ignoreDuplicateModules"===e&&"undefined"!=typeof i.ignoreDuplicateModules?i.ignoreDuplicateModules=i.ignoreDuplicateModules.concat(t):"paths"===e&&"undefined"!=typeof i.paths?r.forEachProperty(t,function(e,t){return i.paths[e]=t}):"shim"===e&&"undefined"!=typeof i.shim?r.forEachProperty(t,function(e,t){return i.shim[e]=t}):"config"===e&&"undefined"!=typeof i.config?r.forEachProperty(t,function(e,t){return i.config[e]=t}):i[e]=r.recursiveClone(t)}),e.validateConfigurationOptions(i)},e}();e.ConfigurationOptionsUtil=i;var o=function(){function e(e){this.options=i.mergeConfigurationOptions(e),this._createIgnoreDuplicateModulesMap(),this._createSortedPathsRules(),this._createShimModules(),this._createOverwriteModuleIdToPath()}return e.prototype._createOverwriteModuleIdToPath=function(){this.overwriteModuleIdToPath={};for(var e=0;e<this.options.bundles.length;e++){var t=this.options.bundles[e],n=t.location;if(t.modules)for(var r=0;r<t.modules.length;r++)this.overwriteModuleIdToPath[t.modules[r]]=n}},e.prototype._createIgnoreDuplicateModulesMap=function(){this.ignoreDuplicateModulesMap={};for(var e=0;e<this.options.ignoreDuplicateModules.length;e++)this.ignoreDuplicateModulesMap[this.options.ignoreDuplicateModules[e]]=!0},e.prototype._createSortedPathsRules=function(){var e=this;this.sortedPathsRules=[],r.forEachProperty(this.options.paths,function(t,n){r.isArray(n)?e.sortedPathsRules.push({from:t,to:n}):e.sortedPathsRules.push({from:t,to:[n]})}),this.sortedPathsRules.sort(function(e,t){return t.from.length-e.from.length})},e.prototype._ensureShimModule1=function(e,t){for(var n=0;n<t.length;n++){var r=t[n];this.shimModules.hasOwnProperty(r)||this._ensureShimModule1(r,[])}this.shimModules[e]={dependencies:t,callback:null}},e.prototype._ensureShimModule2=function(e,n){this.shimModules[e]={dependencies:n.deps||[],callback:function(){for(var e=[],r=0;r<arguments.length-0;r++)e[r]=arguments[r+0];if("function"==typeof n.init){var i=n.init.apply(t,e);if("undefined"!=typeof i)return i}if("function"==typeof n.exports)return n.exports.apply(t,e);if("string"==typeof n.exports){for(var o=n.exports.split("."),s=t,a=0;a<o.length;a++)s&&(s=s[o[a]]);return s}return n.exports||{}}}},e.prototype._createShimModules=function(){var e=this;this.shimModules={},r.forEachProperty(this.options.shim,function(t,n){return n?r.isArray(n)?(e._ensureShimModule1(t,n),void 0):(e._ensureShimModule2(t,n),void 0):void 0})},e.prototype.cloneAndMerge=function(t){return new e(i.mergeConfigurationOptions(t,this.options))},e.prototype.getOptionsLiteral=function(){return this.options},e.prototype._applyPaths=function(e){for(var t,n=0,i=this.sortedPathsRules.length;i>n;n++)if(t=this.sortedPathsRules[n],r.startsWith(e,t.from)){for(var o=[],s=0,a=t.to.length;a>s;s++)o.push(t.to[s]+e.substr(t.from.length));return o}return[e]},e.prototype._addUrlArgsToUrl=function(e){var t=e.indexOf("?")>=0;return t?e+"&"+this.options.urlArgs:e+"?"+this.options.urlArgs},e.prototype._addUrlArgsIfNecessaryToUrl=function(e){return this.options.urlArgs?this._addUrlArgsToUrl(e):e},e.prototype._addUrlArgsIfNecessaryToUrls=function(e){if(this.options.urlArgs)for(var t=0,n=e.length;n>t;t++)e[t]=this._addUrlArgsToUrl(e[t]);return e},e.prototype.moduleIdToPaths=function(e){var t=e;this.overwriteModuleIdToPath.hasOwnProperty(t)&&(t=this.overwriteModuleIdToPath[t]);var n;if(r.endsWith(t,".js")||r.isAbsolutePath(t))n=[t];else{n=this._applyPaths(t);for(var i=0,o=n.length;o>i;i++)r.isAbsolutePath(n[i])||(n[i]=this.options.baseUrl+n[i]),r.endsWith(n[i],".js")||r.containsQueryString(n[i])||(n[i]=n[i]+".js")}return this._addUrlArgsIfNecessaryToUrls(n)},e.prototype.requireToUrl=function(e){var t=e;return r.isAbsolutePath(t)||(t=this._applyPaths(t)[0],r.isAbsolutePath(t)||(t=this.options.baseUrl+t)),this._addUrlArgsIfNecessaryToUrl(t)},e.prototype.isShimmed=function(e){return this.shimModules.hasOwnProperty(e)},e.prototype.getShimmedModuleDefine=function(e){return this.shimModules[e]},e.prototype.isDuplicateMessageIgnoredFor=function(e){return this.ignoreDuplicateModulesMap.hasOwnProperty(e)},e.prototype.getConfigForModule=function(e){return this.options.config?this.options.config[e]:void 0},e.prototype.shouldCatchError=function(){return this.options.catchError},e.prototype.onError=function(e){this.options.onError(e)},e}();e.Configuration=o;var s=function(){function e(e,t){this._config=e;var n=t.lastIndexOf("/");this.fromModulePath=-1!==n?t.substr(0,n+1):""}return e.prototype._normalizeModuleId=function(e){for(var t=e.split("/"),n=0;n<t.length;n++)"."===t[n]||""===t[n]?(t.splice(n,1),n--):n>0&&".."===t[n]&&".."!==t[n-1]&&(t.splice(n-1,2),n-=2);return t.join("/")},e.prototype.resolveModule=function(e){var t=e;return r.isAbsolutePath(t)||(r.startsWith(t,"./")||r.startsWith(t,"../"))&&(t=this._normalizeModuleId(this.fromModulePath+t)),t},e.prototype.moduleIdToPaths=function(e){return this._config.moduleIdToPaths(e)},e.prototype.requireToUrl=function(e){return this._config.requireToUrl(e)},e.prototype.shouldCatchError=function(){return this._config.shouldCatchError()},e.prototype.onError=function(e){this._config.onError(e)},e}();e.ModuleIdResolver=s;var a=function(){function e(e,t,n,r,i,o){this._id=e,this._dependencies=t,this._dependenciesValues=[],this._callback=n,this._errorback=r,this._moduleIdResolver=i,this._exports={},this._exportsPassedIn=!1,this._config=o,this._digestDependencies(),0===this._unresolvedDependenciesCount&&this._complete()}return e.prototype._digestDependencies=function(){var e=this;this._unresolvedDependenciesCount=this._dependencies.length,this._managerDependencies=[],this._managerDependenciesMap={};var t,n,r;for(t=0,n=this._dependencies.length;n>t;t++)if(r=this._dependencies[t])if("exports"===r)this._exportsPassedIn=!0,this._dependenciesValues[t]=this._exports,this._unresolvedDependenciesCount--;else if("module"===r)this._dependenciesValues[t]={id:this._id,config:function(){return e._config}},this._unresolvedDependenciesCount--;else if("require"===r)this.addManagerDependency(r,t);else{var i=r.indexOf("!");if(i>=0){var o=r.substring(0,i),s=r.substring(i+1,r.length);r=this._moduleIdResolver.resolveModule(o)+"!"+s}else r=this._moduleIdResolver.resolveModule(r);this.addManagerDependency(r,t)}else console.warn("Please check module "+this._id+", the dependency list looks broken"),this._dependenciesValues[t]=null,this._unresolvedDependenciesCount--},e.prototype.addManagerDependency=function(e,t){if(this._managerDependenciesMap.hasOwnProperty(e))throw new Error("Module "+this._id+" contains multiple times a dependency to "+e);this._managerDependencies.push(e),this._managerDependenciesMap[e]=t},e.prototype.renameDependency=function(e,t){if(!this._managerDependenciesMap.hasOwnProperty(e))throw new Error("Loader: Cannot rename an unknown dependency!");var n=this._managerDependenciesMap[e];delete this._managerDependenciesMap[e],this._managerDependenciesMap[t]=n},e.prototype.getId=function(){return this._id},e.prototype.getModuleIdResolver=function(){return this._moduleIdResolver},e.prototype.isExportsPassedIn=function(){return this._exportsPassedIn},e.prototype.getExports=function(){return this._exports},e.prototype.getDependencies=function(){return this._managerDependencies},e.prototype._invokeFactory=function(){var e=null,n=null;if(this._moduleIdResolver.shouldCatchError())try{n=this._callback.apply(t,this._dependenciesValues)}catch(r){e=r}else n=this._callback.apply(t,this._dependenciesValues);return{returnedValue:n,producedError:e}},e.prototype._complete=function(){var e=null;if(this._callback)if("function"==typeof this._callback){var t=this._invokeFactory();e=t.producedError,e||"undefined"==typeof t.returnedValue||this._exportsPassedIn&&!r.isEmpty(this._exports)||(this._exports=t.returnedValue)}else this._exports=this._callback;e&&this.getModuleIdResolver().onError({errorCode:"factory",moduleId:this._id,detail:e})},e.prototype.cleanUp=function(){this._dependencies=null,this._dependenciesValues=null,this._callback=null,this._moduleIdResolver=null,this._managerDependencies=null,this._managerDependenciesMap=null},e.prototype.onDependencyError=function(e){return this._errorback?(this._errorback(e),!0):!1},e.prototype.resolveDependency=function(e,t){if(!this._managerDependenciesMap.hasOwnProperty(e))throw new Error("Cannot resolve a dependency I do not have!");this._dependenciesValues[this._managerDependenciesMap[e]]=t,delete this._managerDependenciesMap[e],this._unresolvedDependenciesCount--,0===this._unresolvedDependenciesCount&&this._complete()},e.prototype.isComplete=function(){return 0===this._unresolvedDependenciesCount},e}();e.Module=a;var l=function(){function e(e){this._config=new o,this._scriptLoader=e,this._modules={},this._knownModules={},this._inverseDependencies={},this._dependencies={},this._inversePluginDependencies={},this._queuedDefineCalls=[],this._loadingScriptsCount=0}return e.prototype.enqueueDefineModule=function(e,t,n){0===this._loadingScriptsCount?this.defineModule(e,t,n,null):this._queuedDefineCalls.push({id:e,dependencies:t,callback:n})},e.prototype.enqueueDefineAnonymousModule=function(e,t){this._queuedDefineCalls.push({id:null,dependencies:e,callback:t})},e.prototype.defineModule=function(e,t,n,r,i){if("undefined"==typeof i&&(i=new s(this._config,e)),this._modules.hasOwnProperty(e))return this._config.isDuplicateMessageIgnoredFor(e)||console.warn("Duplicate definition of module '"+e+"'"),void 0;var o=this._config.getConfigForModule(e),l=new a(e,t,n,r,i,o);this._modules[e]=l,this._resolve(l)},e.prototype._relativeRequire=function(e,t,n,i){return"string"==typeof t?this.synchronousRequire(t,e):(this.defineModule(r.generateAnonymousModule(),t,n,i,e),void 0)},e.prototype.synchronousRequire=function(e,t){"undefined"==typeof t&&(t=new s(this._config,e));var n=t.resolveModule(e),r=n.indexOf("!");if(r>=0){var i=n.substring(0,r),o=n.substring(r+1,n.length),a={};this._modules.hasOwnProperty(i)&&(a=this._modules[i]);var l=function(e){return t.resolveModule(e)};o="function"==typeof a.normalize?a.normalize(o,l):l(o),n=i+"!"+o}if(!this._modules.hasOwnProperty(n))throw new Error("Check dependency list! Synchronous require cannot resolve module '"+n+"'. This is the first mention of this module!");var c=this._modules[n];if(!c.isComplete())throw new Error("Check dependency list! Synchronous require cannot resolve module '"+n+"'. This module has not been resolved completely yet.");return c.getExports()},e.prototype.configure=function(e,t){this._config=t?new o(e):this._config.cloneAndMerge(e)},e.prototype.getConfigurationOptions=function(){return this._config.getOptionsLiteral()},e.prototype._onLoad=function(e){var t;if(this._loadingScriptsCount--,this._config.isShimmed(e))t=this._config.getShimmedModuleDefine(e),this.defineModule(e,t.dependencies,t.callback,null);else if(0===this._queuedDefineCalls.length)console.warn("No define call received from module "+e+". This might be a problem.");else for(;this._queuedDefineCalls.length>0;){if(t=this._queuedDefineCalls.shift(),t.id===e||null===t.id){t.id=e,this.defineModule(t.id,t.dependencies,t.callback,null);break}this.defineModule(t.id,t.dependencies,t.callback,null)}if(0===this._loadingScriptsCount)for(;this._queuedDefineCalls.length>0;)t=this._queuedDefineCalls.shift(),null===t.id?console.warn("Found an unmatched anonymous define call in the define queue. Ignoring it!"):this.defineModule(t.id,t.dependencies,t.callback,null)},e.prototype._onLoadError=function(e,t){this._loadingScriptsCount--;var n,r={errorCode:"load",moduleId:e,neededBy:this._inverseDependencies[e]?this._inverseDependencies[e].slice(0):[],detail:t},i={},o=!1,s=[];for(s.push(e),i[e]=!0;s.length>0;)if(n=s.shift(),this._modules[n]&&(o=this._modules[n].onDependencyError(r)||o),this._inverseDependencies[n])for(var a=0,l=this._inverseDependencies[n].length;l>a;a++)i.hasOwnProperty(this._inverseDependencies[n][a])||(s.push(this._inverseDependencies[n][a]),i[this._inverseDependencies[n][a]]=!0);o||this._config.onError(r)},e.prototype._onModuleComplete=function(e,t){var n,i,o,s;if(this._inverseDependencies.hasOwnProperty(e)){var a=this._inverseDependencies[e];for(delete this._inverseDependencies[e],n=0,i=a.length;i>n;n++)o=a[n],s=this._modules[o],s.resolveDependency(e,t),s.isComplete()&&this._onModuleComplete(o,s.getExports())}if(this._inversePluginDependencies.hasOwnProperty(e)){var l=this._inversePluginDependencies[e];for(delete this._inversePluginDependencies[e],n=0,i=l.length;i>n;n++){var c=l[n].moduleId,u=this._modules[c];this._resolvePluginDependencySync(c,l[n].dependencyId,t),u.isComplete()&&this._onModuleComplete(c,u.getExports())}}r.isAnonymousModule(e)?(delete this._modules[e],delete this._dependencies[e]):this._modules[e].cleanUp()},e.prototype._hasDependencyPath=function(e,t){var n,r,i,o,s,a={},l=[];for(l.push(e),a[e]=!0;l.length>0;)if(i=l.shift(),this._dependencies.hasOwnProperty(i))for(o=this._dependencies[i],n=0,r=o.length;r>n;n++){if(s=o[n],s===t)return!0;a.hasOwnProperty(s)||(a[s]=!0,l.push(s))}return!1},e.prototype._findCyclePath=function(e,t,n){if(e===t||50===n)return[e];if(!this._dependencies.hasOwnProperty(e))return null;for(var r,i=this._dependencies[e],o=0,s=i.length;s>o;o++)if(r=this._findCyclePath(i[o],t,n+1),null!==r)return r.push(e),r;return null},e.prototype._createRequire=function(e){var t=this,n=function(n,r,i){return t._relativeRequire(e,n,r,i)};return n.toUrl=function(t){return e.requireToUrl(e.resolveModule(t))},n},e.prototype._resolvePluginDependencySync=function(e,t,n){var r=this,i=this._modules[e],o=i.getModuleIdResolver(),s=t.indexOf("!"),a=t.substring(0,s),l=t.substring(s+1,t.length),c=function(e){return o.resolveModule(e)};if(l="function"==typeof n.normalize?n.normalize(l,c):c(l),n.dynamic){var u=function(n){i.resolveDependency(t,n),i.isComplete()&&r._onModuleComplete(e,i.getExports())};u.error=function(n){r._config.onError({errorCode:"load",moduleId:t,neededBy:[e],detail:n})},n.load(l,this._createRequire(o),u,this._config.getOptionsLiteral())}else{var p=t;t=a+"!"+l,i.renameDependency(p,t),this._resolveDependency(e,t,function(){var e=function(e){r.defineModule(t,[],e,null)};e.error=function(e){r._config.onError({errorCode:"load",moduleId:t,neededBy:r._inverseDependencies[t]?r._inverseDependencies[t].slice(0):[],detail:e})},n.load(l,r._createRequire(o),e,r._config.getOptionsLiteral())})}},e.prototype._resolvePluginDependencyAsync=function(e,t){var n=this._modules[e],r=t.indexOf("!"),i=t.substring(0,r);this._inversePluginDependencies[i]=this._inversePluginDependencies[i]||[],this._inversePluginDependencies[i].push({moduleId:e,dependencyId:t}),this._modules.hasOwnProperty(i)||this._knownModules.hasOwnProperty(i)||(this._knownModules[i]=!0,this._loadModule(n.getModuleIdResolver(),i))},e.prototype._resolvePluginDependency=function(e,t){var n=t.indexOf("!"),r=t.substring(0,n);this._modules.hasOwnProperty(r)&&this._modules[r].isComplete()?this._resolvePluginDependencySync(e,t,this._modules[r].getExports()):this._resolvePluginDependencyAsync(e,t)},e.prototype._resolveShimmedDependency=function(e,t,n){var i=this._config.getShimmedModuleDefine(t);i.dependencies.length>0?this.defineModule(r.generateAnonymousModule(),i.dependencies,function(){return n(t)},null,new s(this._config,t)):n(t)},e.prototype._resolveDependency=function(e,t,n){var r=this._modules[e];if(this._modules.hasOwnProperty(t)&&this._modules[t].isComplete())r.resolveDependency(t,this._modules[t].getExports());else if(this._dependencies[e].push(t),this._hasDependencyPath(t,e)){console.warn("There is a dependency cycle between '"+t+"' and '"+e+"'. The cyclic path follows:");var i=this._findCyclePath(t,e,0);i.reverse(),i.push(t),console.warn(i.join(" => \n"));var o,s=this._modules.hasOwnProperty(t)?this._modules[t]:null;s&&s.isExportsPassedIn()&&(o=s.getExports()),r.resolveDependency(t,o)}else this._inverseDependencies[t]=this._inverseDependencies[t]||[],this._inverseDependencies[t].push(e),this._modules.hasOwnProperty(t)||this._knownModules.hasOwnProperty(t)||(this._knownModules[t]=!0,this._config.isShimmed(t)?this._resolveShimmedDependency(e,t,n):n(t))},e.prototype._loadModule=function(e,t){var n=this;this._loadingScriptsCount++;var r=e.moduleIdToPaths(t),i=-1,o=function(e){i++,i>=r.length?n._onLoadError(t,e):n._scriptLoader.load(r[i],function(){return n._onLoad(t)},o)};o(null)},e.prototype._resolve=function(e){var t,n,r,i,o,s,a=this;r=e.getId(),i=e.getDependencies(),s=e.getModuleIdResolver(),this._dependencies[r]=[];var l=function(e){return a._loadModule(s,e)};for(t=0,n=i.length;n>t;t++)o=i[t],"require"!==o?o.indexOf("!")>=0?this._resolvePluginDependency(r,o):this._resolveDependency(r,o,l):e.resolveDependency(o,this._createRequire(s));e.isComplete()&&this._onModuleComplete(r,e.getExports())},e}();e.ModuleManager=l;var c=function(){function e(e){this.actualScriptLoader=e,this.callbackMap={}}return e.prototype.load=function(e,t,n){var r=this,i={callback:t,errorback:n};return this.callbackMap.hasOwnProperty(e)?(this.callbackMap[e].push(i),void 0):(this.callbackMap[e]=[i],this.actualScriptLoader.load(e,function(){return r.triggerCallback(e)},function(t){return r.triggerErrorback(e,t)}),void 0)},e.prototype.triggerCallback=function(e){var t=this.callbackMap[e];delete this.callbackMap[e];for(var n=0;n<t.length;n++)t[n].callback()},e.prototype.triggerErrorback=function(e,t){var n=this.callbackMap[e];delete this.callbackMap[e];for(var r=0;r<n.length;r++)n[r].errorback(t)},e}(),u=function(){function e(){}return e.prototype.attachListenersV1=function(e,t,n){var r=function(){e.detachEvent("onreadystatechange",i),e.addEventListener&&e.removeEventListener("error",o)},i=function(){("loaded"===e.readyState||"complete"===e.readyState)&&(r(),t())},o=function(e){r(),n(e)};e.attachEvent("onreadystatechange",i),e.addEventListener&&e.addEventListener("error",o)},e.prototype.attachListenersV2=function(e,t,n){var r=function(){e.removeEventListener("load",i),e.removeEventListener("error",o)},i=function(){r(),t()},o=function(e){r(),n(e)};e.addEventListener("load",i),e.addEventListener("error",o)},e.prototype.load=function(e,n,r){var i=document.createElement("script");i.setAttribute("async","async"),i.setAttribute("type","text/javascript"),t.attachEvent?this.attachListenersV1(i,n,r):this.attachListenersV2(i,n,r),i.setAttribute("src",e),document.getElementsByTagName("head")[0].appendChild(i)},e}(),p=function(){function e(){this.loadCalls=[],this.loadTimeout=-1}return e.prototype.load=function(e,t,n){var r=this;this.loadCalls.push({scriptSrc:e,callback:t,errorback:n}),navigator.userAgent.indexOf("Firefox")>=0?this._load():-1===this.loadTimeout&&(this.loadTimeout=setTimeout(function(){r.loadTimeout=-1,r._load()},0))},e.prototype._load=function(){var e=this.loadCalls;this.loadCalls=[];var t,n=e.length,r=[];for(t=0;n>t;t++)r.push(e[t].scriptSrc);var i=!1;try{importScripts.apply(null,r)}catch(o){for(i=!0,t=0;n>t;t++)e[t].errorback(o)}if(!i)for(t=0;n>t;t++)e[t].callback()},e}(),h=null;h=n?new c(new p):new c(new u);var d=new l(h),m=function(){function e(e,t,n){"string"!=typeof e&&(n=t,t=e,e=null),"object"==typeof t&&r.isArray(t)||(n=t,t=null),t||(t=["require","exports","module"]),e?d.enqueueDefineModule(e,t,n):d.enqueueDefineAnonymousModule(t,n)}return e.amd={jQuery:!0},e}(),f=function(){function e(){if(1===arguments.length){if(arguments[0]instanceof Object&&!r.isArray(arguments[0]))return e.config(arguments[0]),void 0;if("string"==typeof arguments[0])return d.synchronousRequire(arguments[0])}if((2===arguments.length||3===arguments.length)&&r.isArray(arguments[0]))return d.defineModule(r.generateAnonymousModule(),arguments[0],arguments[1],arguments[2]),void 0;throw new Error("Unrecognized require call")}return e.config=function(e,t){"undefined"==typeof t&&(t=!1),d.configure(e,t)},e.getConfig=function(){return d.getConfigurationOptions()},e.reset=function(){d=new l(h)},e}();n||(window.onload=function(){var e,t,n,i=document.getElementsByTagName("script");for(e=0,t=i.length;t>e&&!(n=i[e].getAttribute("data-main"));e++);n&&d.defineModule(r.generateAnonymousModule(),[n],null,null,new s(new o,""))}),"undefined"!=typeof t.require&&"function"!=typeof t.require&&f.config(t.require),t.define=m,t.require=f}}(AMDLoader||(AMDLoader={})),/*---------------------------------------------------------
 * Copyright (C) Microsoft Corporation. All rights reserved.
 *--------------------------------------------------------*/
function(e){function t(e,t){var n=e.replace(/\{(\d+)\}/g,function(e,n){var r=n[0];return"undefined"!=typeof t[r]?t[r]:e});return self&&self.document&&self.document.URL.match(/[^\?]*\?[^\#]*pseudo=true/)&&(n="［"+n.replace(/[aouei]/g,"$&$&")+"］"),n}function n(e,t){var n=e[t];return n?n:(n=e["*"],n?n:null)}function r(e,n){for(var r=[],i=0;i<arguments.length-2;i++)r[i]=arguments[i+2];return t(n,r)}var i=e.Plugin&&e.Plugin.Resources?e.Plugin.Resources:void 0,o="i-default";define("vs/nls",{load:function(e,s,a,l){if(l=l||{},!e||0===e.length||l.isBuild)a({localize:r});else{var c;if(i)c=".nls.keys",s([e+c],function(e){a({localize:function(t,n){if(!e[t])return"NLS error: unkown key "+t;var r=e[t].keys;if(n>=r.length)return"NLS error unknow index "+n;var o=r[n],s=[];s[0]=t+"_"+o;for(var a=0;a<arguments.length-2;a++)s[a+1]=arguments[a+2];return i.getString.apply(i,s)}})});else{var u=l["vs/nls"]||{},p=u.availableLanguages?n(u.availableLanguages,e):null;c=".nls",null!==p&&p!==o&&(c=c+"."+p),s([e+c],function(e){a({localize:function(n,r){for(var i=[],o=0;o<arguments.length-2;o++)i[o]=arguments[o+2];if(!e[n])return"NLS error: unkown key "+n;var s=e[n];return r>=s.length?"NLS error unknow index "+r:t(s[r],i)}})})}}},localize:r})}(this);var __extends=this.__extends||function(e,t){function n(){this.constructor=e}for(var r in t)t.hasOwnProperty(r)&&(e[r]=t[r]);n.prototype=t.prototype,e.prototype=new n},_cssPluginGlobal=this,CSSLoaderPlugin;!function(e){var t=_cssPluginGlobal,n=function(){function e(){this._pendingLoads=0}return e.prototype.attachListeners=function(e,t,n,r){var i=function(){t.removeEventListener("load",o),t.removeEventListener("error",s)},o=function(){i(),n()},s=function(e){i(),r(e)};t.addEventListener("load",o),t.addEventListener("error",s)},e.prototype._onLoad=function(e,t){this._pendingLoads--,t()},e.prototype._onLoadError=function(e,t,n){this._pendingLoads--,t(n)},e.prototype._insertLinkNode=function(e){this._pendingLoads++;var t=document.head||document.getElementsByTagName("head")[0],n=t.getElementsByTagName("link")||document.head.getElementsByTagName("script");n.length>0?t.insertBefore(e,n[n.length-1]):t.appendChild(e)},e.prototype.createLinkTag=function(e,t,n,r){var i=this,o=document.createElement("link");o.setAttribute("rel","stylesheet"),o.setAttribute("type","text/css"),o.setAttribute("data-name",e);var s=function(){return i._onLoad(e,n)},a=function(t){return i._onLoadError(e,r,t)};return this.attachListeners(e,o,s,a),o.setAttribute("href",t),o},e.prototype._linkTagExists=function(e,t){var n,r,i,o,s=document.getElementsByTagName("link");for(n=0,r=s.length;r>n;n++)if(i=s[n].getAttribute("data-name"),o=s[n].getAttribute("href"),i===e||o===t)return!0;return!1},e.prototype.load=function(e,t,n,r){if(this._linkTagExists(e,t))return n(),void 0;var i=this.createLinkTag(e,t,n,r);this._insertLinkNode(i)},e}(),r=function(e){function t(){e.call(this),this._blockedLoads=[],this._mergeStyleSheetsTimeout=-1}return __extends(t,e),t.prototype.load=function(e,t,n,r){if(this._linkTagExists(e,t))return n(),void 0;var i=this.createLinkTag(e,t,n,r);this._styleSheetCount()<31?this._insertLinkNode(i):(this._blockedLoads.push(i),this._handleBlocked())},t.prototype._styleSheetCount=function(){var e=document.getElementsByTagName("link").length,t=document.getElementsByTagName("style").length;return e+t},t.prototype._onLoad=function(t,n){e.prototype._onLoad.call(this,t,n),this._handleBlocked()},t.prototype._onLoadError=function(t,n,r){e.prototype._onLoadError.call(this,t,n,r),this._handleBlocked()},t.prototype._handleBlocked=function(){var e=this,t=this._blockedLoads.length;t>0&&-1===this._mergeStyleSheetsTimeout&&(this._mergeStyleSheetsTimeout=window.setTimeout(function(){return e._mergeStyleSheets()},0))},t.prototype._mergeStyleSheet=function(e,t,n,r){for(var i=r.rules.length-1;i>=0;i--)t.insertRule(a.rewriteUrls(n,e,r.rules[i].cssText),0)},t.prototype._mergeStyleSheets=function(){this._mergeStyleSheetsTimeout=-1;var e,t=this._blockedLoads.length,n=document.getElementsByTagName("link"),r=n.length,i=[];for(e=0;r>e;e++)("loaded"===n[e].readyState||"complete"===n[e].readyState)&&i.push({linkNode:n[e],rulesLength:n[e].styleSheet.rules.length});var o=i.length,s=Math.min(Math.floor(o/2),t);i.sort(function(e,t){return t.rulesLength-e.rulesLength});var a,l;for(e=0;s>e;e++)a=i.length-1-e,l=e%(i.length-s),this._mergeStyleSheet(i[l].linkNode.href,i[l].linkNode.styleSheet,i[a].linkNode.href,i[a].linkNode.styleSheet),i[a].linkNode.parentNode.removeChild(i[a].linkNode),r--;for(var c=this._styleSheetCount();31>c&&this._blockedLoads.length>0;)this._insertLinkNode(this._blockedLoads.shift()),c++},t}(n),i=function(e){function t(){e.call(this)}return __extends(t,e),t.prototype.attachListeners=function(e,t,n){t.onload=function(){t.onload=null,n()}},t}(r),o=function(){function e(){this.fs=require.nodeRequire("fs")}return e.prototype.load=function(t,n,r){var i=this.fs.readFileSync(n,"utf8");i.charCodeAt(0)===e.BOM_CHAR_CODE&&(i=i.substring(1)),r(i)},e.BOM_CHAR_CODE=65279,e}(),s=function(){function e(e){this.cssLoader=e}return e.prototype.load=function(t,n,r,i){i=i||{};var o=n.toUrl(t+".css");this.cssLoader.load(t,o,function(n){i.isBuild&&(e.BUILD_MAP[t]=n),r({})},function(){"function"==typeof r.error&&r.error("Could not find "+o+" or it was empty")})},e.prototype.write=function(n,r,i){var o=i.getEntryPoint();t.entryPoints=t.entryPoints||{},t.entryPoints[o]=t.entryPoints[o]||[],t.entryPoints[o].push({moduleName:r,contents:e.BUILD_MAP[r]}),i.asModule(n+"!"+r,"define(['vs/css!"+o+"'], {});")},e.prototype.writeFile=function(e,n,r,i){if(t.entryPoints&&t.entryPoints.hasOwnProperty(n)){for(var o=r.toUrl(n+".css"),s=["/*---------------------------------------------------------"," * Copyright (C) Microsoft Corporation. All rights reserved."," *--------------------------------------------------------*/"],l=t.entryPoints[n],c=0;c<l.length;c++)s.push(a.rewriteUrls(l[c].moduleName,n,l[c].contents));i(o,s.join("\r\n"))}},e.BUILD_MAP={},e}();e.CSSPlugin=s;var a=function(){function e(){}return e.startsWith=function(e,t){return e.length>=t.length&&e.substr(0,t.length)===t},e.pathOf=function(e){var t=e.lastIndexOf("/");return-1!==t?e.substr(0,t+1):""},e.joinPaths=function(t,n){function r(t,n){return e.startsWith(t,n)?Math.max(n.length,t.indexOf("/",n.length)):0}function i(e,t){if("./"!==t){if("../"===t){var n=e.length>0?e[e.length-1]:null;if(n&&"/"===n)return;if(n&&"../"!==n)return e.pop(),void 0}e.push(t)}}function o(e,t){for(;t.length>0;){var n=t.indexOf("/"),r=n>=0?t.substring(0,n+1):t;t=n>=0?t.substring(n+1):"",i(e,r)}}var s=0;s=s||r(t,"//"),s=s||r(t,"http://"),s=s||r(t,"https://");var a=[];return o(a,t.substr(s)),n.length>0&&"/"===n.charAt(0)&&(a=[]),o(a,n),t.substring(0,s)+a.join("")},e.commonPrefix=function(e,t){for(var n=Math.min(e.length,t.length),r=0;n>r&&e.charCodeAt(r)===t.charCodeAt(r);r++);return e.substring(0,r)},e.commonFolderPrefix=function(t,n){var r=e.commonPrefix(t,n),i=r.lastIndexOf("/");return-1===i?"":r.substring(0,i+1)},e.relativePath=function(t,n){if(e.startsWith(n,"/")||e.startsWith(n,"http://")||e.startsWith(n,"https://"))return n;var r=e.commonFolderPrefix(t,n);t=t.substr(r.length),n=n.substr(r.length);for(var i=t.split("/").length,o="",s=1;i>s;s++)o+="../";return o+n},e.rewriteUrls=function(t,n,r){return r.replace(/url\(\s*([^\)]+)\s*\)?/g,function(){for(var r=[],i=0;i<arguments.length-1;i++)r[i]=arguments[i+1];var o=r[0];for(('"'===o.charAt(0)||"'"===o.charAt(0))&&(o=o.substring(1));o.length>0&&(" "===o.charAt(o.length-1)||"	"===o.charAt(o.length-1));)o=o.substring(0,o.length-1);if(('"'===o.charAt(o.length-1)||"'"===o.charAt(o.length-1))&&(o=o.substring(0,o.length-1)),!e.startsWith(o,"data:")&&!e.startsWith(o,"http://")&&!e.startsWith(o,"https://")){var s=e.joinPaths(e.pathOf(t),o);o=e.relativePath(n,s)}return"url("+o+")"})},e}();e.Utilities=a,function(){var e=null;e="undefined"!=typeof process&&process.versions&&process.versions.node&&!process.__node_webkit&&!process.__atom_shell?new o:navigator.userAgent.indexOf("MSIE 9")>=0?new r:navigator.userAgent.indexOf("MSIE 8")>=0?new i:new n,define("vs/css",new s(e))}()}(CSSLoaderPlugin||(CSSLoaderPlugin={}));var TextLoaderPlugin;!function(e){function t(e,t){var n=65279,r=e.readFileSync(t,"utf8");return r.charCodeAt(0)===n&&(r=r.substring(1)),r}var n=function(){function e(){}return e.prototype.load=function(e,t,n,r){var i=new XMLHttpRequest;i.onreadystatechange=function(){4===i.readyState&&(i.status>=200&&i.status<300||1223===i.status||0===i.status&&i.responseText&&i.responseText.length>0?n(i.responseText):r(i),i.onreadystatechange=null)},i.open("GET",t,!0),i.responseType="",i.setRequestHeader("X-Requested-With","XMLHttpRequest"),i.send(null)},e}(),r=function(){function e(){this.fs=require.nodeRequire("fs")}return e.prototype.load=function(e,n,r){r(t(this.fs,n))},e}(),i=function(){function e(e){this.textLoader=e}return e.prototype.load=function(t,n,r,i){i=i||{};var o=n.toUrl(t);this.textLoader.load(t,o,function(n){i.isBuild&&(e.BUILD_MAP[t]=n),r(n)},function(){"function"==typeof r.error&&r.error("Could not find "+o)})},e.prototype.write=function(t,n,r){if(e.BUILD_MAP.hasOwnProperty(n)){var i=o.escapeText(e.BUILD_MAP[n]);r('define("'+t+"!"+n+'", function () { return "'+i+'"; });')}},e.BUILD_MAP={},e}();e.TextPlugin=i;var o=function(){function e(){}return e.escapeText=function(e){for(var t,n="\b".charCodeAt(0),r="\f".charCodeAt(0),i="\n".charCodeAt(0),o=0,s="\r".charCodeAt(0),a="	".charCodeAt(0),l="".charCodeAt(0),c="\\".charCodeAt(0),u='"'.charCodeAt(0),p=0,h=null,d=[],m=0,f=e.length;f>m;m++){switch(t=e.charCodeAt(m)){case n:h="\\b";break;case r:h="\\f";break;case i:h="\\n";break;case o:h="\\0";break;case s:h="\\r";break;case a:h="\\t";break;case l:h="\\v";break;case c:h="\\\\";break;case u:h='\\"'}null!==h&&(d.push(e.substring(p,m)),d.push(h),p=m+1,h=null)}return d.push(e.substring(p,f)),d.join("")},e}();e.Utilities=o,function(){var e=null;e="undefined"!=typeof process&&process.versions&&process.versions.node&&!process.__node_webkit&&!process.__atom_shell?new r:new n,define("vs/text",new i(e))}()}(TextLoaderPlugin||(TextLoaderPlugin={})),define("vs/native",{load:function(e,t,n){self.MonacoEnvironment.appRoot||n.onError(new Error("missing appRoot")),self.require.nodeRequire||n.onError(new Error("missing nodeRequire"));try{var r=self.require.nodeRequire(self.MonacoEnvironment.appRoot+"/"+e);n(r)}catch(i){n.onError(i)}}});