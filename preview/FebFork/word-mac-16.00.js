/* Word Mac specific JavaScript API library */
/* Version: 16.0.6328.1000 */
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

var __extends = this.__extends || function (n, t) {
	function r() {
		this.constructor = n
	}
	for (var i in t)
		t.hasOwnProperty(i) && (n[i] = t[i]);
	r.prototype = t.prototype;
	n.prototype = new r
}, OsfMsAjaxFactory, OfficeExt, msAjaxCDNPath, OSF, OSFWebkit, OSFLog, Logger, OSFAppTelemetry, OfficeExtension, Word;
(function (n) {
	var t = function () {
		function t() {}

		var i = null,
		n = !0;
		return t.prototype.isMsAjaxLoaded = function () {
			var t = "function",
			i = "undefined";
			return typeof Sys !== i && typeof Type !== i && Sys.StringBuilder && typeof Sys.StringBuilder === t && Type.registerNamespace && typeof Type.registerNamespace === t && Type.registerClass && typeof Type.registerClass === t && typeof Function._validateParams === t ? n : !1
		},
		t.prototype.loadMsAjaxFull = function (n) {
			var t = (window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";
			OSF.OUtil.loadScript(t, n)
		},
		Object.defineProperty(t.prototype, "msAjaxError", {
			get : function () {
				var n = this;
				return n._msAjaxError == i && n.isMsAjaxLoaded() && (n._msAjaxError = Error),
				n._msAjaxError
			},
			set : function (n) {
				this._msAjaxError = n
			},
			enumerable : n,
			configurable : n
		}),
		Object.defineProperty(t.prototype, "msAjaxSerializer", {
			get : function () {
				var n = this;
				return n._msAjaxSerializer == i && n.isMsAjaxLoaded() && (n._msAjaxSerializer = Sys.Serialization.JavaScriptSerializer),
				n._msAjaxSerializer
			},
			set : function (n) {
				this._msAjaxSerializer = n
			},
			enumerable : n,
			configurable : n
		}),
		Object.defineProperty(t.prototype, "msAjaxString", {
			get : function () {
				var n = this;
				return n._msAjaxString == i && n.isMsAjaxLoaded() && (n._msAjaxSerializer = String),
				n._msAjaxString
			},
			set : function (n) {
				this._msAjaxString = n
			},
			enumerable : n,
			configurable : n
		}),
		Object.defineProperty(t.prototype, "msAjaxDebug", {
			get : function () {
				var n = this;
				return n._msAjaxDebug == i && n.isMsAjaxLoaded() && (n._msAjaxDebug = Sys.Debug),
				n._msAjaxDebug
			},
			set : function (n) {
				this._msAjaxDebug = n
			},
			enumerable : n,
			configurable : n
		}),
		t
	}
	();
	n.MicrosoftAjaxFactory = t
})(OfficeExt || (OfficeExt = {}));
OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory;
OSF = OSF || {}, function (n) {
	var t = function () {
		function n(n) {
			this._internalStorage = n
		}
		return n.prototype.getItem = function (n) {
			try {
				return this._internalStorage && this._internalStorage.getItem(n)
			} catch (t) {
				return null
			}
		},
		n.prototype.setItem = function (n, t) {
			try {
				this._internalStorage && this._internalStorage.setItem(n, t)
			} catch (i) {}

		},
		n.prototype.clear = function () {
			try {
				this._internalStorage && this._internalStorage.clear()
			} catch (n) {}

		},
		n.prototype.removeItem = function (n) {
			try {
				this._internalStorage && this._internalStorage.removeItem(n)
			} catch (t) {}

		},
		n.prototype.getKeysWithPrefix = function (n) {
			var r = [],
			u,
			t,
			i;
			try {
				for (u = this._internalStorage && this._internalStorage.length || 0, t = 0; t < u; t++)
					i = this._internalStorage.key(t), i.indexOf(n) === 0 && r.push(i)
			} catch (f) {}

			return r
		},
		n
	}
	();
	n.SafeStorage = t
}
(OfficeExt || (OfficeExt = {}));
OSF.OUtil = function () {
	function k() {
		var n = o * Math.random();
		return n ^= r^(new Date).getMilliseconds() << Math.floor(Math.random() * 21),
		n.toString(16)
	}
	function d() {
		if (!l) {
			try {
				var n = window.sessionStorage
			} catch (i) {
				n = t
			}
			l = new OfficeExt.SafeStorage(n)
		}
		return l
	}
	var u = "on",
	v = "configurable",
	y = "writable",
	f = "enumerable",
	e = "undefined",
	i = !0,
	n = !1,
	o = 2147483647,
	t = null,
	s = -1,
	p = "&_xdm_Info=",
	w = "&_serializer_version=",
	b = "_xdm_",
	g = "_serializer_version=",
	h = "#",
	c = {},
	nt = 3e4,
	l = t,
	a = t,
	r = (new Date).getTime();
	return {
		set_entropy : function (n) {
			var t,
			u,
			i;
			if (typeof n == "string")
				for (t = 0; t < n.length; t += 4) {
					for (u = 0, i = 0; i < 4 && t + i < n.length; i++)
						u = (u << 8) + n.charCodeAt(t + i);
					r ^= u
				}
			else
				r ^= typeof n == "number" ? n : o * Math.random();
			r &= o
		},
		extend : function (n, t) {
			var i = function () {};
			i.prototype = t.prototype;
			n.prototype = new i;
			n.prototype.constructor = n;
			n.uber = t.prototype;
			t.prototype.constructor === Object.prototype.constructor && (t.prototype.constructor = t)
		},
		setNamespace : function (n, t) {
			t && n && !t[n] && (t[n] = {})
		},
		unsetNamespace : function (n, t) {
			t && n && t[n] && delete t[n]
		},
		loadScript : function (r, u, f) {
			var s,
			e,
			o,
			h,
			l;
			r && u && (s = window.document, e = c[r], e ? e.loaded ? u() : e.pendingCallbacks.push(u) : (o = s.createElement("script"), o.type = "text/javascript", e = {
						loaded : n,
						pendingCallbacks : [u],
						timer : t
					}, c[r] = e, h = function () {
					var r,
					n,
					u;
					for (e.timer != t && (clearTimeout(e.timer), delete e.timer), e.loaded = i, r = e.pendingCallbacks.length, n = 0; n < r; n++)
						u = e.pendingCallbacks.shift(), u()
				}, l = function () {
					var i,
					n,
					u;
					for (delete c[r], e.timer != t && (clearTimeout(e.timer), delete e.timer), i = e.pendingCallbacks.length, n = 0; n < i; n++)
						u = e.pendingCallbacks.shift(), u()
				}, o.readyState ? o.onreadystatechange = function () {
					(o.readyState == "loaded" || o.readyState == "complete") && (o.onreadystatechange = t, h())
				}
					 : o.onload = h, o.onerror = l, f = f || nt, e.timer = setTimeout(l, f), o.src = r, s.getElementsByTagName("head")[0].appendChild(o)))
		},
		loadCSS : function (n) {
			if (n) {
				var i = window.document,
				t = i.createElement("link");
				t.type = "text/css";
				t.rel = "stylesheet";
				t.href = n;
				i.getElementsByTagName("head")[0].appendChild(t)
			}
		},
		parseEnum : function (n, t) {
			var i = t[n.trim()];
			if (typeof i == e) {
				OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + n);
				throw OsfMsAjaxFactory.msAjaxError.argument("str");
			}
			return i
		},
		delayExecutionAndCache : function () {
			var n = {
				calc : arguments[0]
			};
			return function () {
				return n.calc && (n.val = n.calc.apply(this, arguments), delete n.calc),
				n.val
			}
		},
		getUniqueId : function () {
			return s = s + 1,
			s.toString()
		},
		formatString : function () {
			var n = arguments,
			t = n[0];
			return t.replace(/{(\d+)}/gm, function (t, i) {
				var r = parseInt(i, 10) + 1;
				return n[r] === undefined ? "{" + i + "}" : n[r]
			})
		},
		generateConversationId : function () {
			return [k(), k(), (new Date).getTime().toString()].join("_")
		},
		getFrameNameAndConversationId : function (n, t) {
			var i = b + n + this.generateConversationId();
			return t.setAttribute("name", i),
			this.generateConversationId()
		},
		addXdmInfoAsHash : function (n, t) {
			return OSF.OUtil.addInfoAsHash(n, p, t)
		},
		addSerializerVersionAsHash : function (n, t) {
			return OSF.OUtil.addInfoAsHash(n, w, t)
		},
		addInfoAsHash : function (n, t, i) {
			n = n.trim() || "";
			var r = n.split(h),
			u = r.shift(),
			f = r.join(h);
			return [u, h, f, t, i].join("")
		},
		parseXdmInfo : function (n) {
			return OSF.OUtil.parseXdmInfoWithGivenFragment(n, window.location.hash)
		},
		parseXdmInfoWithGivenFragment : function (n, t) {
			return OSF.OUtil.parseInfoWithGivenFragment(p, b, n, t)
		},
		parseSerializerVersion : function (n) {
			return OSF.OUtil.parseSerializerVersionWithGivenFragment(n, window.location.hash)
		},
		parseSerializerVersionWithGivenFragment : function (n, t) {
			return parseInt(OSF.OUtil.parseInfoWithGivenFragment(w, g, n, t))
		},
		parseInfoWithGivenFragment : function (n, i, r, u) {
			var s = u.split(n),
			f = s.length > 1 ? s[s.length - 1] : t,
			h = d(),
			e,
			o,
			c;
			return !r && h && (e = window.name.indexOf(i), e > -1 && (o = window.name.indexOf(";", e), o == -1 && (o = window.name.length), c = window.name.substring(e, o), f ? h.setItem(c, f) : f = h.getItem(c))),
			f
		},
		getConversationId : function () {
			var i = window.location.search,
			n = t,
			r;
			return i && (r = i.indexOf("&"), n = r > 0 ? i.substring(1, r) : i.substr(1), n && n.charAt(n.length - 1) === "=" && (n = n.substring(0, n.length - 1), n && (n = decodeURIComponent(n)))),
			n
		},
		getInfoItems : function (n) {
			var t = n.split("$");
			return typeof t[1] == e && (t = n.split("|")),
			t
		},
		getConversationUrl : function () {
			var t = "",
			r = OSF.OUtil.parseXdmInfo(i),
			n;
			return r && (n = OSF.OUtil.getInfoItems(r), n != undefined && n.length >= 3 && (t = n[2])),
			t
		},
		validateParamObject : function (t, r) {
			var u = Function._validateParams(arguments, [{
							name : "params",
							type : Object,
							mayBeNull : n
						}, {
							name : "expectedProperties",
							type : Object,
							mayBeNull : n
						}, {
							name : "callback",
							type : Function,
							mayBeNull : i
						}
					]),
			f;
			if (u)
				throw u;
			for (f in r)
				if (u = Function._validateParameter(t[f], r[f], f), u)
					throw u;
		},
		writeProfilerMark : function (n) {
			window.msWriteProfilerMark && (window.msWriteProfilerMark(n), OsfMsAjaxFactory.msAjaxDebug.trace(n))
		},
		outputDebug : function (n) {
			typeof Sys !== e && Sys && Sys.Debug && OsfMsAjaxFactory.msAjaxDebug.trace(n)
		},
		defineNondefaultProperty : function (n, t, r, u) {
			var e,
			f;
			r = r || {};
			for (e in u)
				f = u[e], r[f] == undefined && (r[f] = i);
			return Object.defineProperty(n, t, r),
			n
		},
		defineNondefaultProperties : function (n, t, i) {
			t = t || {};
			for (var r in t)
				OSF.OUtil.defineNondefaultProperty(n, r, t[r], i);
			return n
		},
		defineEnumerableProperty : function (n, t, i) {
			return OSF.OUtil.defineNondefaultProperty(n, t, i, [f])
		},
		defineEnumerableProperties : function (n, t) {
			return OSF.OUtil.defineNondefaultProperties(n, t, [f])
		},
		defineMutableProperty : function (n, t, i) {
			return OSF.OUtil.defineNondefaultProperty(n, t, i, [y, f, v])
		},
		defineMutableProperties : function (n, t) {
			return OSF.OUtil.defineNondefaultProperties(n, t, [y, f, v])
		},
		finalizeProperties : function (t, r) {
			var e,
			u;
			r = r || {};
			for (var o = Object.getOwnPropertyNames(t), s = o.length, f = 0; f < s; f++)
				e = o[f], u = Object.getOwnPropertyDescriptor(t, e), u.get || u.set || (u.writable = r.writable || n), u.configurable = r.configurable || n, u.enumerable = r.enumerable || i, Object.defineProperty(t, e, u);
			return t
		},
		mapList : function (n, t) {
			var i = [],
			r;
			if (n)
				for (r in n)
					i.push(t(n[r]));
			return i
		},
		listContainsKey : function (t, r) {
			for (var u in t)
				if (r == u)
					return i;
			return n
		},
		listContainsValue : function (t, r) {
			for (var u in t)
				if (r == t[u])
					return i;
			return n
		},
		augmentList : function (n, t) {
			var r = n.push ? function (t, i) {
				n.push(i)
			}
			 : function (t, i) {
				n[t] = i
			};
			for (var i in t)
				r(i, t[i])
		},
		redefineList : function (n, t) {
			var r,
			i;
			for (r in n)
				delete n[r];
			for (i in t)
				n[i] = t[i]
		},
		isArray : function (n) {
			return Object.prototype.toString.apply(n) === "[object Array]"
		},
		isFunction : function (n) {
			return Object.prototype.toString.apply(n) === "[object Function]"
		},
		isDate : function (n) {
			return Object.prototype.toString.apply(n) === "[object Date]"
		},
		addEventListener : function (t, i, r) {
			t.addEventListener ? t.addEventListener(i, r, n) : Sys.Browser.agent === Sys.Browser.InternetExplorer && t.attachEvent ? t.attachEvent(u + i, r) : t[u + i] = r
		},
		removeEventListener : function (i, r, f) {
			i.removeEventListener ? i.removeEventListener(r, f, n) : Sys.Browser.agent === Sys.Browser.InternetExplorer && i.detachEvent ? i.detachEvent(u + r, f) : i[u + r] = t
		},
		getCookieValue : function (n) {
			var t = RegExp(n + "[^;]+").exec(document.cookie);
			return t.toString().replace(/^[^=]+./, "")
		},
		xhrGet : function (n, t, r) {
			var u;
			try {
				u = new XMLHttpRequest;
				u.onreadystatechange = function () {
					u.readyState == 4 && (u.status == 200 ? t(u.responseText) : r(u.status))
				};
				u.open("GET", n, i);
				u.send()
			} catch (f) {
				r(f)
			}
		},
		xhrGetFull : function (n, t, r, u) {
			var f,
			e = t;
			try {
				f = new XMLHttpRequest;
				f.onreadystatechange = function () {
					f.readyState == 4 && (f.status == 200 ? r(f, e) : u(f.status))
				};
				f.open("GET", n, i);
				f.send()
			} catch (o) {
				u(o)
			}
		},
		encodeBase64 : function (n) {
			var h;
			if (!n)
				return n;
			var a = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",
			l = [],
			i = [],
			o = 0,
			c,
			e,
			s,
			r,
			u,
			f,
			t,
			v = n.length;
			do
				for (c = n.charCodeAt(o++), e = n.charCodeAt(o++), s = n.charCodeAt(o++), t = 0, r = c & 255, u = c >> 8, f = e & 255, i[t++] = r >> 2, i[t++] = (r & 3) << 4 | u >> 4, i[t++] = (u & 15) << 2 | f >> 6, i[t++] = f & 63, isNaN(e) || (r = e >> 8, u = s & 255, f = s >> 8, i[t++] = r >> 2, i[t++] = (r & 3) << 4 | u >> 4, i[t++] = (u & 15) << 2 | f >> 6, i[t++] = f & 63), isNaN(e) ? i[t - 1] = 64 : isNaN(s) && (i[t - 2] = 64, i[t - 1] = 64), h = 0; h < t; h++)
					l.push(a.charAt(i[h]));
			while (o < v);
			return l.join("")
		},
		getSessionStorage : function () {
			return d()
		},
		getLocalStorage : function () {
			if (!a) {
				try {
					var n = window.localStorage
				} catch (i) {
					n = t
				}
				a = new OfficeExt.SafeStorage(n)
			}
			return a
		},
		convertIntToCssHexColor : function (n) {
			return "#" + (Number(n) + 16777216).toString(16).slice(-6)
		},
		attachClickHandler : function (n, t) {
			n.onclick = function () {
				t()
			};
			n.ontouchend = function (n) {
				t();
				n.preventDefault()
			}
		},
		getQueryStringParamValue : function (t, i) {
			var u = Function._validateParams(arguments, [{
							name : "queryString",
							type : String,
							mayBeNull : n
						}, {
							name : "paramName",
							type : String,
							mayBeNull : n
						}
					]),
			r;
			return u ? (OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null."), "") : (r = new RegExp("[\\?&]" + i + "=([^&#]*)", "i"), !r.test(t)) ? (OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found."), "") : r.exec(t)[1]
		},
		isiOS : function () {
			return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? i : n
		},
		shallowCopy : function (n) {
			var i = n.constructor();
			for (var t in n)
				n.hasOwnProperty(t) && (i[t] = n[t]);
			return i
		}
	}
}
();
OSF.OUtil.Guid = function () {
	var n = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
	return {
		generateNewGuid : function () {
			for (var i = "", r = (new Date).getTime(), t = 0; t < 32 && r > 0; t++)
				(t == 8 || t == 12 || t == 16 || t == 20) && (i += "-"), i += n[r % 16], r = Math.floor(r / 16);
			for (; t < 32; t++)
				(t == 8 || t == 12 || t == 16 || t == 20) && (i += "-"), i += n[Math.floor(Math.random() * 16)];
			return i
		}
	}
}
();
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.AppName = {
	Unsupported : 0,
	Excel : 1,
	Word : 2,
	PowerPoint : 4,
	Outlook : 8,
	ExcelWebApp : 16,
	WordWebApp : 32,
	OutlookWebApp : 64,
	Project : 128,
	AccessWebApp : 256,
	PowerpointWebApp : 512,
	ExcelIOS : 1024,
	Sway : 2048,
	WordIOS : 4096,
	PowerPointIOS : 8192,
	Access : 16384,
	Lync : 32768,
	OutlookIOS : 65536,
	OneNoteWebApp : 131072
};
OSF.InternalPerfMarker = {
	DataCoercionBegin : "Agave.HostCall.CoerceDataStart",
	DataCoercionEnd : "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
	IssueCall : "Agave.HostCall.IssueCall",
	ReceiveResponse : "Agave.HostCall.ReceiveResponse",
	RuntimeExceptionRaised : "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
	Select : 0,
	UnSelect : 1,
	CancelDialog : 2,
	InsertAgave : 3,
	CtrlF6In : 4,
	CtrlF6Exit : 5,
	CtrlF6ExitShift : 6,
	SelectWithError : 7,
	NotifyHostError : 8
};
OSF.SharedConstants = {
	NotificationConversationIdSuffix : "_ntf"
};
OSF.OfficeAppContext = function (n, t, i, r, u, f, e, o, s, h, c, l, a, v, y, p, w) {
	var b = this;
	b._id = n;
	b._appName = t;
	b._appVersion = i;
	b._appUILocale = r;
	b._dataLocale = u;
	b._docUrl = f;
	b._clientMode = e;
	b._settings = o;
	b._reason = s;
	b._osfControlType = h;
	b._eToken = c;
	b._correlationId = l;
	b._appInstanceId = a;
	b._touchEnabled = v;
	b._commerceAllowed = y;
	b._appMinorVersion = p;
	b._requirementMatrix = w;
	b.get_id = function () {
		return this._id
	};
	b.get_appName = function () {
		return this._appName
	};
	b.get_appVersion = function () {
		return this._appVersion
	};
	b.get_appUILocale = function () {
		return this._appUILocale
	};
	b.get_dataLocale = function () {
		return this._dataLocale
	};
	b.get_docUrl = function () {
		return this._docUrl
	};
	b.get_clientMode = function () {
		return this._clientMode
	};
	b.get_bindings = function () {
		return this._bindings
	};
	b.get_settings = function () {
		return this._settings
	};
	b.get_reason = function () {
		return this._reason
	};
	b.get_osfControlType = function () {
		return this._osfControlType
	};
	b.get_eToken = function () {
		return this._eToken
	};
	b.get_correlationId = function () {
		return this._correlationId
	};
	b.get_appInstanceId = function () {
		return this._appInstanceId
	};
	b.get_touchEnabled = function () {
		return this._touchEnabled
	};
	b.get_commerceAllowed = function () {
		return this._commerceAllowed
	};
	b.get_appMinorVersion = function () {
		return this._appMinorVersion
	};
	b.get_requirementMatrix = function () {
		return this._requirementMatrix
	}
};
OSF.OsfControlType = {
	DocumentLevel : 0,
	ContainerLevel : 1
};
OSF.ClientMode = {
	ReadOnly : 0,
	ReadWrite : 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
	Inserted : "inserted",
	DocumentOpened : "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat = {
	Unformatted : "unformatted",
	Formatted : "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
	All : "all"
};
Microsoft.Office.WebExtension.Parameters = {
	BindingType : "bindingType",
	CoercionType : "coercionType",
	ValueFormat : "valueFormat",
	FilterType : "filterType",
	Columns : "columns",
	SampleData : "sampleData",
	GoToType : "goToType",
	SelectionMode : "selectionMode",
	Id : "id",
	PromptText : "promptText",
	ItemName : "itemName",
	FailOnCollision : "failOnCollision",
	StartRow : "startRow",
	StartColumn : "startColumn",
	RowCount : "rowCount",
	ColumnCount : "columnCount",
	Callback : "callback",
	AsyncContext : "asyncContext",
	Data : "data",
	Rows : "rows",
	OverwriteIfStale : "overwriteIfStale",
	FileType : "fileType",
	EventType : "eventType",
	Handler : "handler",
	SliceSize : "sliceSize",
	SliceIndex : "sliceIndex",
	ActiveView : "activeView",
	Status : "status",
	Xml : "xml",
	Namespace : "namespace",
	Prefix : "prefix",
	XPath : "xPath",
	Text : "text",
	ImageLeft : "imageLeft",
	ImageTop : "imageTop",
	ImageWidth : "imageWidth",
	ImageHeight : "imageHeight",
	TaskId : "taskId",
	FieldId : "fieldId",
	FieldValue : "fieldValue",
	ServerUrl : "serverUrl",
	ListName : "listName",
	ResourceId : "resourceId",
	ViewType : "viewType",
	ViewName : "viewName",
	GetRawValue : "getRawValue",
	CellFormat : "cellFormat",
	TableOptions : "tableOptions",
	TaskIndex : "taskIndex",
	ResourceIndex : "resourceIndex"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
	ReadOnly : 1,
	ReadWrite : 0
};
OSF.DDA.PropertyDescriptors = {
	AsyncResultStatus : "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.getXdmEventName = function (n, t) {
	return t == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || t == Microsoft.Office.WebExtension.EventType.BindingDataChanged ? n + "_" + t : t
};
OSF.DDA.MethodDispId = {
	dispidMethodMin : 64,
	dispidGetSelectedDataMethod : 64,
	dispidSetSelectedDataMethod : 65,
	dispidAddBindingFromSelectionMethod : 66,
	dispidAddBindingFromPromptMethod : 67,
	dispidGetBindingMethod : 68,
	dispidReleaseBindingMethod : 69,
	dispidGetBindingDataMethod : 70,
	dispidSetBindingDataMethod : 71,
	dispidAddRowsMethod : 72,
	dispidClearAllRowsMethod : 73,
	dispidGetAllBindingsMethod : 74,
	dispidLoadSettingsMethod : 75,
	dispidSaveSettingsMethod : 76,
	dispidGetDocumentCopyMethod : 77,
	dispidAddBindingFromNamedItemMethod : 78,
	dispidAddColumnsMethod : 79,
	dispidGetDocumentCopyChunkMethod : 80,
	dispidReleaseDocumentCopyMethod : 81,
	dispidNavigateToMethod : 82,
	dispidGetActiveViewMethod : 83,
	dispidGetDocumentThemeMethod : 84,
	dispidGetOfficeThemeMethod : 85,
	dispidGetFilePropertiesMethod : 86,
	dispidClearFormatsMethod : 87,
	dispidSetTableOptionsMethod : 88,
	dispidSetFormatsMethod : 89,
	dispidExecuteRichApiRequestMethod : 93,
	dispidAppCommandInvocationCompletedMethod : 94,
	dispidAddDataPartMethod : 128,
	dispidGetDataPartByIdMethod : 129,
	dispidGetDataPartsByNamespaceMethod : 130,
	dispidGetDataPartXmlMethod : 131,
	dispidGetDataPartNodesMethod : 132,
	dispidDeleteDataPartMethod : 133,
	dispidGetDataNodeValueMethod : 134,
	dispidGetDataNodeXmlMethod : 135,
	dispidGetDataNodesMethod : 136,
	dispidSetDataNodeValueMethod : 137,
	dispidSetDataNodeXmlMethod : 138,
	dispidAddDataNamespaceMethod : 139,
	dispidGetDataUriByPrefixMethod : 140,
	dispidGetDataPrefixByUriMethod : 141,
	dispidGetDataNodeTextMethod : 142,
	dispidSetDataNodeTextMethod : 143,
	dispidMethodMax : 143,
	dispidGetSelectedTaskMethod : 110,
	dispidGetSelectedResourceMethod : 111,
	dispidGetTaskMethod : 112,
	dispidGetResourceFieldMethod : 113,
	dispidGetWSSUrlMethod : 114,
	dispidGetTaskFieldMethod : 115,
	dispidGetProjectFieldMethod : 116,
	dispidGetSelectedViewMethod : 117,
	dispidGetTaskByIndexMethod : 118,
	dispidGetResourceByIndexMethod : 119,
	dispidSetTaskFieldMethod : 120,
	dispidSetResourceFieldMethod : 121,
	dispidGetMaxTaskIndexMethod : 122,
	dispidGetMaxResourceIndexMethod : 123
};
OSF.DDA.EventDispId = {
	dispidEventMin : 0,
	dispidInitializeEvent : 0,
	dispidSettingsChangedEvent : 1,
	dispidDocumentSelectionChangedEvent : 2,
	dispidBindingSelectionChangedEvent : 3,
	dispidBindingDataChangedEvent : 4,
	dispidDocumentOpenEvent : 5,
	dispidDocumentCloseEvent : 6,
	dispidActiveViewChangedEvent : 7,
	dispidDocumentThemeChangedEvent : 8,
	dispidOfficeThemeChangedEvent : 9,
	dispidActivationStatusChangedEvent : 32,
	dispidAppCommandInvokedEvent : 39,
	dispidTaskSelectionChangedEvent : 56,
	dispidResourceSelectionChangedEvent : 57,
	dispidViewSelectionChangedEvent : 58,
	dispidDataNodeAddedEvent : 60,
	dispidDataNodeReplacedEvent : 61,
	dispidDataNodeDeletedEvent : 62,
	dispidEventMax : 63
};
OSF.DDA.ErrorCodeManager = function () {
	var n = {};
	return {
		getErrorArgs : function (t) {
			return n[t] || n[this.errorCodes.ooeInternalError]
		},
		addErrorMessage : function (t, i) {
			n[t] = i
		},
		errorCodes : {
			ooeSuccess : 0,
			ooeChunkResult : 1,
			ooeCoercionTypeNotSupported : 1e3,
			ooeGetSelectionNotMatchDataType : 1001,
			ooeCoercionTypeNotMatchBinding : 1002,
			ooeInvalidGetRowColumnCounts : 1003,
			ooeSelectionNotSupportCoercionType : 1004,
			ooeInvalidGetStartRowColumn : 1005,
			ooeNonUniformPartialGetNotSupported : 1006,
			ooeGetDataIsTooLarge : 1008,
			ooeFileTypeNotSupported : 1009,
			ooeGetDataParametersConflict : 1010,
			ooeInvalidGetColumns : 1011,
			ooeInvalidGetRows : 1012,
			ooeInvalidReadForBlankRow : 1013,
			ooeUnsupportedDataObject : 2e3,
			ooeCannotWriteToSelection : 2001,
			ooeDataNotMatchSelection : 2002,
			ooeOverwriteWorksheetData : 2003,
			ooeDataNotMatchBindingSize : 2004,
			ooeInvalidSetStartRowColumn : 2005,
			ooeInvalidDataFormat : 2006,
			ooeDataNotMatchCoercionType : 2007,
			ooeDataNotMatchBindingType : 2008,
			ooeSetDataIsTooLarge : 2009,
			ooeNonUniformPartialSetNotSupported : 2010,
			ooeInvalidSetColumns : 2011,
			ooeInvalidSetRows : 2012,
			ooeSetDataParametersConflict : 2013,
			ooeCellDataAmountBeyondLimits : 2014,
			ooeSelectionCannotBound : 3e3,
			ooeBindingNotExist : 3002,
			ooeBindingToMultipleSelection : 3003,
			ooeInvalidSelectionForBindingType : 3004,
			ooeOperationNotSupportedOnThisBindingType : 3005,
			ooeNamedItemNotFound : 3006,
			ooeMultipleNamedItemFound : 3007,
			ooeInvalidNamedItemForBindingType : 3008,
			ooeUnknownBindingType : 3009,
			ooeOperationNotSupportedOnMatrixData : 3010,
			ooeInvalidColumnsForBinding : 3011,
			ooeSettingNameNotExist : 4e3,
			ooeSettingsCannotSave : 4001,
			ooeSettingsAreStale : 4002,
			ooeOperationNotSupported : 5e3,
			ooeInternalError : 5001,
			ooeDocumentReadOnly : 5002,
			ooeEventHandlerNotExist : 5003,
			ooeInvalidApiCallInContext : 5004,
			ooeShuttingDown : 5005,
			ooeUnsupportedEnumeration : 5007,
			ooeIndexOutOfRange : 5008,
			ooeBrowserAPINotSupported : 5009,
			ooeInvalidParam : 5010,
			ooeRequestTimeout : 5011,
			ooeTooManyIncompleteRequests : 5100,
			ooeRequestTokenUnavailable : 5101,
			ooeActivityLimitReached : 5102,
			ooeCustomXmlNodeNotFound : 6e3,
			ooeCustomXmlError : 6100,
			ooeCustomXmlExceedQuota : 6101,
			ooeCustomXmlOutOfDate : 6102,
			ooeNoCapability : 7e3,
			ooeCannotNavTo : 7001,
			ooeSpecifiedIdNotExist : 7002,
			ooeNavOutOfBound : 7004,
			ooeElementMissing : 8e3,
			ooeProtectedError : 8001,
			ooeInvalidCellsValue : 8010,
			ooeInvalidTableOptionValue : 8011,
			ooeInvalidFormatValue : 8012,
			ooeRowIndexOutOfRange : 8020,
			ooeColIndexOutOfRange : 8021,
			ooeFormatValueOutOfRange : 8022,
			ooeCellFormatAmountBeyondLimits : 8023,
			ooeMemoryFileLimit : 11e3,
			ooeNetworkProblemRetrieveFile : 11001,
			ooeInvalidSliceSize : 11002,
			ooeInvalidCallback : 11101
		},
		initializeErrorMessages : function (t) {
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = {
				name : t.L_InvalidCoercion,
				message : t.L_CoercionTypeNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = {
				name : t.L_DataReadError,
				message : t.L_GetSelectionNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = {
				name : t.L_InvalidCoercion,
				message : t.L_CoercionTypeNotMatchBinding
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = {
				name : t.L_DataReadError,
				message : t.L_InvalidGetRowColumnCounts
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = {
				name : t.L_DataReadError,
				message : t.L_SelectionNotSupportCoercionType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = {
				name : t.L_DataReadError,
				message : t.L_InvalidGetStartRowColumn
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = {
				name : t.L_DataReadError,
				message : t.L_NonUniformPartialGetNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = {
				name : t.L_DataReadError,
				message : t.L_GetDataIsTooLarge
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = {
				name : t.L_DataReadError,
				message : t.L_FileTypeNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = {
				name : t.L_DataReadError,
				message : t.L_GetDataParametersConflict
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = {
				name : t.L_DataReadError,
				message : t.L_InvalidGetColumns
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = {
				name : t.L_DataReadError,
				message : t.L_InvalidGetRows
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = {
				name : t.L_DataReadError,
				message : t.L_InvalidReadForBlankRow
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = {
				name : t.L_DataWriteError,
				message : t.L_UnsupportedDataObject
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = {
				name : t.L_DataWriteError,
				message : t.L_CannotWriteToSelection
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = {
				name : t.L_DataWriteError,
				message : t.L_DataNotMatchSelection
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = {
				name : t.L_DataWriteError,
				message : t.L_OverwriteWorksheetData
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = {
				name : t.L_DataWriteError,
				message : t.L_DataNotMatchBindingSize
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = {
				name : t.L_DataWriteError,
				message : t.L_InvalidSetStartRowColumn
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = {
				name : t.L_InvalidFormat,
				message : t.L_InvalidDataFormat
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = {
				name : t.L_InvalidDataObject,
				message : t.L_DataNotMatchCoercionType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = {
				name : t.L_InvalidDataObject,
				message : t.L_DataNotMatchBindingType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = {
				name : t.L_DataWriteError,
				message : t.L_SetDataIsTooLarge
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = {
				name : t.L_DataWriteError,
				message : t.L_NonUniformPartialSetNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = {
				name : t.L_DataWriteError,
				message : t.L_InvalidSetColumns
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = {
				name : t.L_DataWriteError,
				message : t.L_InvalidSetRows
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = {
				name : t.L_DataWriteError,
				message : t.L_SetDataParametersConflict
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = {
				name : t.L_BindingCreationError,
				message : t.L_SelectionCannotBound
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = {
				name : t.L_InvalidBindingError,
				message : t.L_BindingNotExist
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = {
				name : t.L_BindingCreationError,
				message : t.L_BindingToMultipleSelection
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = {
				name : t.L_BindingCreationError,
				message : t.L_InvalidSelectionForBindingType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = {
				name : t.L_InvalidBindingOperation,
				message : t.L_OperationNotSupportedOnThisBindingType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = {
				name : t.L_BindingCreationError,
				message : t.L_NamedItemNotFound
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = {
				name : t.L_BindingCreationError,
				message : t.L_MultipleNamedItemFound
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = {
				name : t.L_BindingCreationError,
				message : t.L_InvalidNamedItemForBindingType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = {
				name : t.L_InvalidBinding,
				message : t.L_UnknownBindingType
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = {
				name : t.L_InvalidBindingOperation,
				message : t.L_OperationNotSupportedOnMatrixData
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = {
				name : t.L_InvalidBinding,
				message : t.L_InvalidColumnsForBinding
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = {
				name : t.L_ReadSettingsError,
				message : t.L_SettingNameNotExist
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = {
				name : t.L_SaveSettingsError,
				message : t.L_SettingsCannotSave
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = {
				name : t.L_SettingsStaleError,
				message : t.L_SettingsAreStale
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = {
				name : t.L_HostError,
				message : t.L_OperationNotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = {
				name : t.L_InternalError,
				message : t.L_InternalErrorDescription
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = {
				name : t.L_PermissionDenied,
				message : t.L_DocumentReadOnly
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = {
				name : t.L_EventRegistrationError,
				message : t.L_EventHandlerNotExist
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = {
				name : t.L_InvalidAPICall,
				message : t.L_InvalidApiCallInContext
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = {
				name : t.L_ShuttingDown,
				message : t.L_ShuttingDown
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = {
				name : t.L_UnsupportedEnumeration,
				message : t.L_UnsupportedEnumerationMessage
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = {
				name : t.L_IndexOutOfRange,
				message : t.L_IndexOutOfRange
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = {
				name : t.L_APINotSupported,
				message : t.L_BrowserAPINotSupported
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = {
				name : t.L_APICallFailed,
				message : t.L_RequestTimeout
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = {
				name : t.L_APICallFailed,
				message : t.L_TooManyIncompleteRequests
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = {
				name : t.L_APICallFailed,
				message : t.L_RequestTokenUnavailable
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = {
				name : t.L_APICallFailed,
				message : t.L_ActivityLimitReached
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = {
				name : t.L_InvalidNode,
				message : t.L_CustomXmlNodeNotFound
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = {
				name : t.L_CustomXmlError,
				message : t.L_CustomXmlError
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = {
				name : t.L_CustomXmlError,
				message : t.L_CustomXmlError
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = {
				name : t.L_CustomXmlError,
				message : t.L_CustomXmlError
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = {
				name : t.L_PermissionDenied,
				message : t.L_NoCapability
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = {
				name : t.L_CannotNavigateTo,
				message : t.L_CannotNavigateTo
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = {
				name : t.L_SpecifiedIdNotExist,
				message : t.L_SpecifiedIdNotExist
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = {
				name : t.L_NavOutOfBound,
				message : t.L_NavOutOfBound
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = {
				name : t.L_DataWriteReminder,
				message : t.L_CellDataAmountBeyondLimits
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = {
				name : t.L_MissingParameter,
				message : t.L_ElementMissing
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = {
				name : t.L_PermissionDenied,
				message : t.L_NoCapability
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = {
				name : t.L_InvalidValue,
				message : t.L_InvalidCellsValue
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = {
				name : t.L_InvalidValue,
				message : t.L_InvalidTableOptionValue
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = {
				name : t.L_InvalidValue,
				message : t.L_InvalidFormatValue
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = {
				name : t.L_OutOfRange,
				message : t.L_RowIndexOutOfRange
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = {
				name : t.L_OutOfRange,
				message : t.L_ColIndexOutOfRange
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = {
				name : t.L_OutOfRange,
				message : t.L_FormatValueOutOfRange
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = {
				name : t.L_FormattingReminder,
				message : t.L_CellFormatAmountBeyondLimits
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = {
				name : t.L_MemoryLimit,
				message : t.L_CloseFileBeforeRetrieve
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = {
				name : t.L_NetworkProblem,
				message : t.L_NetworkProblemRetrieveFile
			};
			n[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = {
				name : t.L_InvalidValue,
				message : t.L_SliceSizeNotSupported
			}
		}
	}
}
(), function (n) {
	(function (n) {
		var t = 1.1,
		r = function () {
			function n(n) {
				this.isSetSupported = function (n, t) {
					var u,
					i,
					r;
					return n == undefined ? !1 : (t == undefined && (t = 0), u = this._setMap, i = u._sets, i.hasOwnProperty(n.toLowerCase()) ? (r = i[n.toLowerCase()], r > 0 && r >= t) : !1)
				};
				this._setMap = n
			}
			return n
		}
		(),
		i,
		f,
		e,
		o,
		s,
		h,
		c,
		l,
		a,
		u,
		v,
		y,
		p,
		w,
		b,
		k,
		d;
		n.RequirementMatrix = r;
		i = function () {
			function n(n) {
				this._sets = n
			}
			return n
		}
		();
		n.DefaultSetRequirement = i;
		f = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					documentevents : t,
					excelapi : t,
					matrixbindings : t,
					matrixcoercion : t,
					selection : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t,
					textbindings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.ExcelClientDefaultSetRequirement = f;
		e = function (n) {
			function t() {
				n.call(this, {
					mailbox : 1.3
				})
			}
			return __extends(t, n),
			t
		}
		(i);
		n.OutlookClientDefaultSetRequirement = e;
		o = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					compressedfile : t,
					customxmlparts : t,
					documentevents : t,
					file : t,
					htmlcoercion : t,
					matrixbindings : t,
					matrixcoercion : t,
					ooxmlcoercion : t,
					pdffile : t,
					selection : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t,
					textbindings : t,
					textcoercion : t,
					textfile : t,
					wordapi : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.WordClientDefaultSetRequirement = o;
		s = function (n) {
			function i() {
				n.call(this, {
					activeview : t,
					compressedfile : t,
					documentevents : t,
					file : t,
					pdffile : t,
					selection : t,
					settings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.PowerpointClientDefaultSetRequirement = s;
		h = function (n) {
			function i() {
				n.call(this, {
					selection : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.ProjectClientDefaultSetRequirement = h;
		c = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					documentevents : t,
					matrixbindings : t,
					matrixcoercion : t,
					selection : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t,
					textbindings : t,
					textcoercion : t,
					file : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.ExcelWebDefaultSetRequirement = c;
		l = function (n) {
			function i() {
				n.call(this, {
					customxmlparts : t,
					documentevents : t,
					file : t,
					ooxmlcoercion : t,
					selection : t,
					settings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.WordWebDefaultSetRequirement = l;
		a = function (n) {
			function i() {
				n.call(this, {
					activeview : t,
					settings : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.PowerpointWebDefaultSetRequirement = a;
		u = function (n) {
			function t() {
				n.call(this, {
					mailbox : 1.3
				})
			}
			return __extends(t, n),
			t
		}
		(i);
		n.OutlookWebDefaultSetRequirement = u;
		v = function (n) {
			function i() {
				n.call(this, {
					activeview : t,
					documentevents : t,
					selection : t,
					settings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.SwayWebDefaultSetRequirement = v;
		y = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					partialtablebindings : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.AccessWebDefaultSetRequirement = y;
		p = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					documentevents : t,
					matrixbindings : t,
					matrixcoercion : t,
					selection : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t,
					textbindings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.ExcelIOSDefaultSetRequirement = p;
		w = function (n) {
			function i() {
				n.call(this, {
					bindingevents : t,
					compressedfile : t,
					customxmlparts : t,
					documentevents : t,
					file : t,
					htmlcoercion : t,
					matrixbindings : t,
					matrixcoercion : t,
					ooxmlcoercion : t,
					pdffile : t,
					selection : t,
					settings : t,
					tablebindings : t,
					tablecoercion : t,
					textbindings : t,
					textcoercion : t,
					textfile : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.WordIOSDefaultSetRequirement = w;
		b = function (n) {
			function i() {
				n.call(this, {
					activeview : t,
					compressedfile : t,
					documentevents : t,
					file : t,
					pdffile : t,
					selection : t,
					settings : t,
					textcoercion : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.PowerpointIOSDefaultSetRequirement = b;
		k = function (n) {
			function i() {
				n.call(this, {
					mailbox : t
				})
			}
			return __extends(i, n),
			i
		}
		(i);
		n.OutlookIOSDefaultSetRequirement = k;
		d = function () {
			function n() {}

			return n.initializeOsfDda = function () {
				OSF.OUtil.setNamespace("Requirement", OSF.DDA)
			},
			n.getDefaultRequirementMatrix = function (t) {
				var u,
				s,
				f,
				e,
				h,
				o;
				return this.initializeDefaultSetMatrix(),
				u = undefined,
				t.get_requirementMatrix() != undefined && typeof JSON != "undefined" ? (s = JSON.parse(t.get_requirementMatrix()), u = new r(new i(s))) : (f = t.get_appMinorVersion(), e = "", e = f < 10 ? "0" + f : "" + f, h = t.get_appVersion() + "." + e, o = t.get_appName() + "-" + h, u = n.DefaultSetArrayMatrix != undefined && n.DefaultSetArrayMatrix[o] != undefined ? new r(n.DefaultSetArrayMatrix[o]) : new r(new i({}))),
				u
			},
			n.initializeDefaultSetMatrix = function () {
				n.DefaultSetArrayMatrix[n.Excel_RCLIENT_1600] = new f;
				n.DefaultSetArrayMatrix[n.Word_RCLIENT_1600] = new o;
				n.DefaultSetArrayMatrix[n.PowerPoint_RCLIENT_1600] = new s;
				n.DefaultSetArrayMatrix[n.Outlook_RCLIENT_1600] = new e;
				n.DefaultSetArrayMatrix[n.Excel_WAC_1600] = new c;
				n.DefaultSetArrayMatrix[n.Word_WAC_1600] = new l;
				n.DefaultSetArrayMatrix[n.Outlook_WAC_1600] = new u;
				n.DefaultSetArrayMatrix[n.Outlook_WAC_1601] = new u;
				n.DefaultSetArrayMatrix[n.Project_RCLIENT_1600] = new h;
				n.DefaultSetArrayMatrix[n.Access_WAC_1600] = new y;
				n.DefaultSetArrayMatrix[n.PowerPoint_WAC_1600] = new a;
				n.DefaultSetArrayMatrix[n.Excel_IOS_1600] = new p;
				n.DefaultSetArrayMatrix[n.SWAY_WAC_1600] = new v;
				n.DefaultSetArrayMatrix[n.Word_IOS_1600] = new w;
				n.DefaultSetArrayMatrix[n.PowerPoint_IOS_1600] = new b;
				n.DefaultSetArrayMatrix[n.Outlook_IOS_1600] = new k
			},
			n.Excel_RCLIENT_1600 = "1-16.00",
			n.Word_RCLIENT_1600 = "2-16.00",
			n.PowerPoint_RCLIENT_1600 = "4-16.00",
			n.Outlook_RCLIENT_1600 = "8-16.00",
			n.Excel_WAC_1600 = "16-16.00",
			n.Word_WAC_1600 = "32-16.00",
			n.Outlook_WAC_1600 = "64-16.00",
			n.Outlook_WAC_1601 = "64-16.01",
			n.Project_RCLIENT_1600 = "128-16.00",
			n.Access_WAC_1600 = "256-16.00",
			n.PowerPoint_WAC_1600 = "512-16.00",
			n.Excel_IOS_1600 = "1024-16.01",
			n.SWAY_WAC_1600 = "2048-16.00",
			n.Word_IOS_1600 = "4096-16.00",
			n.PowerPoint_IOS_1600 = "8192-16.00",
			n.Outlook_IOS_1600 = "65536-16.00",
			n.DefaultSetArrayMatrix = {},
			n
		}
		();
		n.RequirementsMatrixFactory = d
	})(n.Requirement || (n.Requirement = {}));
	var t = n.Requirement
}
(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
	WebEditor : "webEditor",
	WebViewer : "webViewer",
	Client : "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
	ReadOnly : "readOnly",
	ReadWrite : "readWrite"
};
OSF.NamespaceManager = function () {
	var t,
	n = !1;
	return {
		enableShortcut : function () {
			n || (window.Office ? t = window.Office : OSF.OUtil.setNamespace("Office", window), window.Office = Microsoft.Office.WebExtension, n = !0)
		},
		disableShortcut : function () {
			n && (t ? window.Office = t : OSF.OUtil.unsetNamespace("Office", window), n = !1)
		}
	}
}
();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function (n) {
	n ? OSF.NamespaceManager.enableShortcut() : OSF.NamespaceManager.disableShortcut()
};
Microsoft.Office.WebExtension.select = function (n, t) {
	var i,
	r,
	o,
	u,
	f,
	e;
	if (n && typeof n == "string" && (r = n.indexOf("#"), r != -1)) {
		o = n.substring(0, r);
		u = n.substring(r + 1);
		switch (o) {
		case "binding":
		case "bindings":
			u && (i = new OSF.DDA.BindingPromise(u))
		}
	}
	if (i)
		return i.onFail = t, i;
	else if (t)
		if (f = typeof t, f == "function")
			e = {},
	e[Microsoft.Office.WebExtension.Parameters.Callback] = t,
	OSF.DDA.issueAsyncResult(e, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
	else
		throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, f);
};
OSF.DDA.Context = function (n, t, i, r, u) {
	var f = this,
	e,
	o;
	OSF.OUtil.defineEnumerableProperties(f, {
		contentLanguage : {
			value : n.get_dataLocale()
		},
		displayLanguage : {
			value : n.get_appUILocale()
		},
		touchEnabled : {
			value : n.get_touchEnabled()
		},
		commerceAllowed : {
			value : n.get_commerceAllowed()
		}
	});
	t && OSF.OUtil.defineEnumerableProperty(f, "document", {
		value : t
	});
	i && OSF.OUtil.defineEnumerableProperty(f, "license", {
		value : i
	});
	r && (e = r.displayName || "appOM", delete r.displayName, OSF.OUtil.defineEnumerableProperty(f, e, {
			value : r
		}));
	u && OSF.OUtil.defineEnumerableProperty(f, "officeTheme", {
		get : function () {
			return u()
		}
	});
	o = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(n);
	OSF.OUtil.defineEnumerableProperty(f, "requirements", {
		value : o
	})
};
OSF.DDA.OutlookContext = function (n, t, i, r, u) {
	OSF.DDA.OutlookContext.uber.constructor.call(this, n, null, i, r, u);
	t && OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
		value : t
	})
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function () {};
OSF.DDA.Document = function (n, t) {
	var i;
	switch (n.get_clientMode()) {
	case OSF.ClientMode.ReadOnly:
		i = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
		break;
	case OSF.ClientMode.ReadWrite:
		i = Microsoft.Office.WebExtension.DocumentMode.ReadWrite
	}
	t && OSF.OUtil.defineEnumerableProperty(this, "settings", {
		value : t
	});
	OSF.OUtil.defineMutableProperties(this, {
		mode : {
			value : i
		},
		url : {
			value : n.get_docUrl()
		}
	})
};
OSF.DDA.JsomDocument = function (n, t, i) {
	var r = this,
	u;
	OSF.DDA.JsomDocument.uber.constructor.call(r, n, i);
	t && OSF.OUtil.defineEnumerableProperty(r, "bindings", {
		get : function () {
			return t
		}
	});
	u = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(r, [u.GetSelectedDataAsync, u.SetSelectedDataAsync]);
	OSF.DDA.DispIdHost.addEventSupport(r, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
	get : function () {
		var n;
		return OSF && OSF._OfficeAppFactory && (n = OSF._OfficeAppFactory.getContext()),
		n
	}
});
OSF.DDA.License = function (n) {
	OSF.OUtil.defineEnumerableProperty(this, "value", {
		value : n
	})
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
	Context : "Context",
	Value : "Value",
	Status : "Status",
	Error : "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
	Succeeded : "succeeded",
	Failed : "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
	Success : 0,
	Failed : 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
	Name : "Name",
	Message : "Message",
	Code : "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function (n) {
	var t,
	i;
	for (t in n)
		i = {},
	OSF.OUtil.defineEnumerableProperties(i, {
		id : {
			value : t
		},
		displayName : {
			value : n[t]
		}
	}),
	OSF.DDA.AsyncMethodNames[t] = i
};
OSF.DDA.AsyncMethodCall = function (n, t, i, r, u, f, e) {
	function c(n, t) {
		var u,
		i,
		r;
		for (u in n) {
			if (i = n[u], r = t[u], i["enum"])
				switch (typeof r) {
				case "string":
					if (OSF.OUtil.listContainsValue(i["enum"], r))
						break;
				case "undefined":
					throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
					break;
				default:
					throw s();
				}
			if (i.types && !OSF.OUtil.listContainsValue(i.types, typeof r))
				throw s();
		}
	}
	function l(t, i, r) {
		var f,
		u,
		h,
		e,
		l,
		a;
		if (t.length < o)
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
		for (f = [], u = 0; u < o; u++)
			f.push(t[u]);
		for (c(n, f), h = {}, u = 0; u < o; u++) {
			if (e = n[u], l = f[u], e.verify && (a = e.verify(l, i, r), !a))
				throw s();
			h[e.name] = l
		}
		return h
	}
	function a(n, i, r, u) {
		var f,
		l,
		a,
		v,
		y,
		e,
		s;
		if (n.length > o + 2)
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		for (a = n.length - 1; a >= o; a--) {
			v = n[a];
			switch (typeof v) {
			case "object":
				if (f)
					throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
				else
					f = v;
				break;
			case h:
				if (l)
					throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
				else
					l = v;
				break;
			default:
				throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
			}
		}
		f = f || {};
		for (y in t)
			OSF.OUtil.listContainsKey(f, y) || (e = undefined, s = t[y], s.calculate && i && (e = s.calculate(i, r, u)), e || s.defaultValue === undefined || (e = s.defaultValue), f[y] = e);
		if (l)
			if (f[Microsoft.Office.WebExtension.Parameters.Callback])
				throw Strings.OfficeOM.L_RedundantCallbackSpecification;
			else
				f[Microsoft.Office.WebExtension.Parameters.Callback] = l;
		return c(t, f),
		f
	}
	var h = "function",
	o = n.length,
	s = OSF.OUtil.delayExecutionAndCache(function () {
			return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, e)
		});
	this.verifyAndExtractCall = function (n, t, r) {
		var e = l(n, t, r),
		c = a(n, e, t, r),
		u = {},
		o,
		s,
		h;
		for (o in e)
			u[o] = e[o];
		for (s in c)
			u[s] = c[s];
		for (h in i)
			u[h] = i[h](t, r);
		return f && (u = f(u, t, r)),
		u
	};
	this.processResponse = function (n, t, i, f) {
		return n == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess ? r ? r(t, i, f) : t : u ? u(n, t) : OSF.DDA.ErrorCodeManager.getErrorArgs(n)
	};
	this.getCallArgs = function (n) {
		for (var t, u, r, i = n.length - 1; i >= o; i--) {
			r = n[i];
			switch (typeof r) {
			case "object":
				t = r;
				break;
			case h:
				u = r
			}
		}
		return t = t || {},
		u && (t[Microsoft.Office.WebExtension.Parameters.Callback] = u),
		t
	}
};
OSF.DDA.AsyncMethodCallFactory = function () {
	function n(n) {
		var i = null,
		r,
		t;
		if (n)
			for (i = {}, r = n.length, t = 0; t < r; t++)
				i[n[t].name] = n[t].value;
		return i
	}
	return {
		manufacture : function (t) {
			var i = t.supportedOptions ? n(t.supportedOptions) : [],
			r = t.privateStateCallbacks ? n(t.privateStateCallbacks) : [];
			return new OSF.DDA.AsyncMethodCall(t.requiredArguments || [], i, r, t.onSucceeded, t.onFailed, t.checkCallArgs, t.method.displayName)
		}
	}
}
();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function (n) {
	OSF.DDA.AsyncMethodCalls[n.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(n)
};
OSF.DDA.Error = function (n, t, i) {
	OSF.OUtil.defineEnumerableProperties(this, {
		name : {
			value : n
		},
		message : {
			value : t
		},
		code : {
			value : i
		}
	})
};
OSF.DDA.AsyncResult = function (n, t) {
	OSF.OUtil.defineEnumerableProperties(this, {
		value : {
			value : n[OSF.DDA.AsyncResultEnum.Properties.Value]
		},
		status : {
			value : t ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
		}
	});
	n[OSF.DDA.AsyncResultEnum.Properties.Context] && OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
		value : n[OSF.DDA.AsyncResultEnum.Properties.Context]
	});
	t && OSF.OUtil.defineEnumerableProperty(this, "error", {
		value : new OSF.DDA.Error(t[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], t[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], t[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
	})
};
OSF.DDA.issueAsyncResult = function (n, t, i) {
	var f = n[Microsoft.Office.WebExtension.Parameters.Callback],
	u,
	r;
	f && (u = {}, u[OSF.DDA.AsyncResultEnum.Properties.Context] = n[Microsoft.Office.WebExtension.Parameters.AsyncContext], t == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess ? u[OSF.DDA.AsyncResultEnum.Properties.Value] = i : (r = {}, i = i || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError), r[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = t || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, r[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = i.name || i, r[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = i.message || i), f(new OSF.DDA.AsyncResult(u, r)))
};
OSF.DDA.ListType = function () {
	var n = {};
	return {
		setListType : function (t, i) {
			n[t] = i
		},
		isListType : function (t) {
			return OSF.OUtil.listContainsKey(n, t)
		},
		getDescriptor : function (t) {
			return n[t]
		}
	}
}
();
OSF.DDA.HostParameterMap = function (n, t) {
	function e(i, f) {
		var a = i ? {}

		 : undefined,
		s,
		h,
		o,
		v,
		c,
		l;
		for (s in i) {
			if (h = i[s], OSF.DDA.ListType.isListType(s)) {
				o = [];
				for (v in h)
					o.push(e(h[v], f))
			} else
				OSF.OUtil.listContainsKey(r, s) ? o = r[s][f](h) : f == u && n.preserveNesting(s) ? o = e(h, f) : (c = t[s], c ? (l = c[f], l && (o = l[h], o === undefined && (o = h))) : o = h);
			a[s] = o
		}
		return a
	}
	function c(i, r) {
		var e,
		u,
		h,
		s;
		for (u in r)
			h = n.isComplexType(u) ? c(i, t[u][o]) : i[u], h != undefined && (e || (e = {}), s = r[u], s == f && (s = u), e[s] = n.pack(u, h));
		return e
	}
	function s(i, r, e) {
		var o,
		l,
		h,
		c,
		v,
		a;
		e || (e = {});
		for (o in r)
			if (l = r[o], h = l == f ? i : i[l], h === null || h === undefined)
				e[o] = undefined;
			else if (h = n.unpack(o, h), n.isComplexType(o))
				c = t[o][u], n.preserveNesting(o) ? e[o] = s(h, c) : s(h, c, e);
			else {
				if (OSF.DDA.ListType.isListType(o)) {
					c = {};
					v = OSF.DDA.ListType.getDescriptor(o);
					c[v] = f;
					for (a in h)
						h[a] = s(h[a], c)
				}
				e[o] = h
			}
		return e
	}
	function l(n, i, r) {
		var f = t[n][r],
		u,
		o,
		l;
		return r == "toHost" ? (o = e(i, r), u = c(o, f)) : r == h && (l = s(i, f), u = e(l, r)),
		u
	}
	var h = "fromHost",
	i = this,
	o = "toHost",
	u = h,
	f = "self",
	r = {};
	r[Microsoft.Office.WebExtension.Parameters.Data] = {
		toHost : function (n) {
			if (n != null && n.rows !== undefined) {
				var t = {};
				t[OSF.DDA.TableDataProperties.TableRows] = n.rows;
				t[OSF.DDA.TableDataProperties.TableHeaders] = n.headers;
				n = t
			}
			return n
		},
		fromHost : function (n) {
			return n
		}
	};
	r[Microsoft.Office.WebExtension.Parameters.SampleData] = r[Microsoft.Office.WebExtension.Parameters.Data];
	t || (t = {});
	i.addMapping = function (n, i) {
		var e,
		h,
		c,
		l,
		r,
		s,
		a,
		v;
		if (i.map) {
			e = i.map;
			h = {};
			for (c in e)
				l = e[c], l == f && (l = c), h[l] = c
		} else
			e = i.toHost, h = i.fromHost;
		if (r = t[n], r) {
			s = r[o];
			for (a in s)
				e[a] = s[a];
			s = r[u];
			for (v in s)
				h[v] = s[v]
		} else
			r = t[n] = {};
		r[o] = e;
		r[u] = h
	};
	i.toHost = function (n, t) {
		return l(n, t, o)
	};
	i.fromHost = function (n, t) {
		return l(n, t, u)
	};
	i.self = f;
	i.addComplexType = function (t) {
		n.addComplexType(t)
	};
	i.getDynamicType = function (t) {
		return n.getDynamicType(t)
	};
	i.setDynamicType = function (t, i) {
		n.setDynamicType(t, i)
	};
	i.dynamicTypes = r;
	i.doMapValues = function (n, t) {
		return e(n, t)
	}
};
OSF.DDA.SpecialProcessor = function (n, t) {
	var i = this;
	i.addComplexType = function (t) {
		n.push(t)
	};
	i.getDynamicType = function (n) {
		return t[n]
	};
	i.setDynamicType = function (n, i) {
		t[n] = i
	};
	i.isComplexType = function (t) {
		return OSF.OUtil.listContainsValue(n, t)
	};
	i.isDynamicType = function (n) {
		return OSF.OUtil.listContainsKey(t, n)
	};
	i.preserveNesting = function (n) {
		var t = [];
		return OSF.DDA.PropertyDescriptors && t.push(OSF.DDA.PropertyDescriptors.Subset),
		OSF.DDA.DataNodeEventProperties && (t = t.concat([OSF.DDA.DataNodeEventProperties.OldNode, OSF.DDA.DataNodeEventProperties.NewNode, OSF.DDA.DataNodeEventProperties.NextSiblingNode])),
		OSF.OUtil.listContainsValue(t, n)
	};
	i.pack = function (n, i) {
		return this.isDynamicType(n) ? t[n].toHost(i) : i
	};
	i.unpack = function (n, i) {
		return this.isDynamicType(n) ? t[n].fromHost(i) : i
	}
};
OSF.DDA.getDecoratedParameterMap = function (n, t) {
	function r(n) {
		var i = null,
		r,
		t;
		if (n)
			for (i = {}, r = n.length, t = 0; t < r; t++)
				i[n[t].name] = n[t].value;
		return i
	}
	var i = new OSF.DDA.HostParameterMap(n),
	f = i.self,
	u;
	i.define = function (n) {
		var t = {},
		u = r(n.toHost);
		n.invertible ? t.map = u : n.canonical ? t.toHost = t.fromHost = u : (t.toHost = u, t.fromHost = r(n.fromHost));
		i.addMapping(n.type, t);
		n.isComplexType && i.addComplexType(n.type)
	};
	for (u in t)
		i.define(t[u]);
	return i
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
	InvokeMethod : "invokeMethod",
	AddEventHandler : "addEventHandler",
	RemoveEventHandler : "removeEventHandler"
};
OSF.DDA.DispIdHost.Delegates = {
	ExecuteAsync : "executeAsync",
	RegisterEventAsync : "registerEventAsync",
	UnregisterEventAsync : "unregisterEventAsync",
	ParameterMap : "parameterMap"
};
OSF.DDA.DispIdHost.Facade = function (n, t) {
	function s(n, t, i, r) {
		if (typeof n == "number")
			r || (r = t.getCallArgs(i)), OSF.DDA.issueAsyncResult(r, n, OSF.DDA.ErrorCodeManager.getErrorArgs(n));
		else
			throw n;
	}
	var r = {},
	u = OSF.DDA.AsyncMethodNames,
	i = OSF.DDA.MethodDispId,
	h = {
		GoToByIdAsync : i.dispidNavigateToMethod,
		GetSelectedDataAsync : i.dispidGetSelectedDataMethod,
		SetSelectedDataAsync : i.dispidSetSelectedDataMethod,
		GetDocumentCopyChunkAsync : i.dispidGetDocumentCopyChunkMethod,
		ReleaseDocumentCopyAsync : i.dispidReleaseDocumentCopyMethod,
		GetDocumentCopyAsync : i.dispidGetDocumentCopyMethod,
		AddFromSelectionAsync : i.dispidAddBindingFromSelectionMethod,
		AddFromPromptAsync : i.dispidAddBindingFromPromptMethod,
		AddFromNamedItemAsync : i.dispidAddBindingFromNamedItemMethod,
		GetAllAsync : i.dispidGetAllBindingsMethod,
		GetByIdAsync : i.dispidGetBindingMethod,
		ReleaseByIdAsync : i.dispidReleaseBindingMethod,
		GetDataAsync : i.dispidGetBindingDataMethod,
		SetDataAsync : i.dispidSetBindingDataMethod,
		AddRowsAsync : i.dispidAddRowsMethod,
		AddColumnsAsync : i.dispidAddColumnsMethod,
		DeleteAllDataValuesAsync : i.dispidClearAllRowsMethod,
		RefreshAsync : i.dispidLoadSettingsMethod,
		SaveAsync : i.dispidSaveSettingsMethod,
		GetActiveViewAsync : i.dispidGetActiveViewMethod,
		GetFilePropertiesAsync : i.dispidGetFilePropertiesMethod,
		GetOfficeThemeAsync : i.dispidGetOfficeThemeMethod,
		GetDocumentThemeAsync : i.dispidGetDocumentThemeMethod,
		ClearFormatsAsync : i.dispidClearFormatsMethod,
		SetTableOptionsAsync : i.dispidSetTableOptionsMethod,
		SetFormatsAsync : i.dispidSetFormatsMethod,
		ExecuteRichApiRequestAsync : i.dispidExecuteRichApiRequestMethod,
		AppCommandInvocationCompletedAsync : i.dispidAppCommandInvocationCompletedMethod,
		AddDataPartAsync : i.dispidAddDataPartMethod,
		GetDataPartByIdAsync : i.dispidGetDataPartByIdMethod,
		GetDataPartsByNameSpaceAsync : i.dispidGetDataPartsByNamespaceMethod,
		GetPartXmlAsync : i.dispidGetDataPartXmlMethod,
		GetPartNodesAsync : i.dispidGetDataPartNodesMethod,
		DeleteDataPartAsync : i.dispidDeleteDataPartMethod,
		GetNodeValueAsync : i.dispidGetDataNodeValueMethod,
		GetNodeXmlAsync : i.dispidGetDataNodeXmlMethod,
		GetRelativeNodesAsync : i.dispidGetDataNodesMethod,
		SetNodeValueAsync : i.dispidSetDataNodeValueMethod,
		SetNodeXmlAsync : i.dispidSetDataNodeXmlMethod,
		AddDataPartNamespaceAsync : i.dispidAddDataNamespaceMethod,
		GetDataPartNamespaceAsync : i.dispidGetDataUriByPrefixMethod,
		GetDataPartPrefixAsync : i.dispidGetDataPrefixByUriMethod,
		GetNodeTextAsync : i.dispidGetDataNodeTextMethod,
		SetNodeTextAsync : i.dispidSetDataNodeTextMethod,
		GetSelectedTask : i.dispidGetSelectedTaskMethod,
		GetTask : i.dispidGetTaskMethod,
		GetWSSUrl : i.dispidGetWSSUrlMethod,
		GetTaskField : i.dispidGetTaskFieldMethod,
		GetSelectedResource : i.dispidGetSelectedResourceMethod,
		GetResourceField : i.dispidGetResourceFieldMethod,
		GetProjectField : i.dispidGetProjectFieldMethod,
		GetSelectedView : i.dispidGetSelectedViewMethod,
		GetTaskByIndex : i.dispidGetTaskByIndexMethod,
		GetResourceByIndex : i.dispidGetResourceByIndexMethod,
		SetTaskField : i.dispidSetTaskFieldMethod,
		SetResourceField : i.dispidSetResourceFieldMethod,
		GetMaxTaskIndex : i.dispidGetMaxTaskIndexMethod,
		GetMaxResourceIndex : i.dispidGetMaxResourceIndexMethod
	},
	f,
	o,
	e;
	for (f in h)
		u[f] && (r[u[f].id] = h[f]);
	u = Microsoft.Office.WebExtension.EventType;
	i = OSF.DDA.EventDispId;
	o = {
		SettingsChanged : i.dispidSettingsChangedEvent,
		DocumentSelectionChanged : i.dispidDocumentSelectionChangedEvent,
		BindingSelectionChanged : i.dispidBindingSelectionChangedEvent,
		BindingDataChanged : i.dispidBindingDataChangedEvent,
		ActiveViewChanged : i.dispidActiveViewChangedEvent,
		OfficeThemeChanged : i.dispidOfficeThemeChangedEvent,
		DocumentThemeChanged : i.dispidDocumentThemeChangedEvent,
		AppCommandInvoked : i.dispidAppCommandInvokedEvent,
		TaskSelectionChanged : i.dispidTaskSelectionChangedEvent,
		ResourceSelectionChanged : i.dispidResourceSelectionChangedEvent,
		ViewSelectionChanged : i.dispidViewSelectionChangedEvent,
		DataNodeInserted : i.dispidDataNodeAddedEvent,
		DataNodeReplaced : i.dispidDataNodeReplacedEvent,
		DataNodeDeleted : i.dispidDataNodeDeletedEvent
	};
	for (e in o)
		u[e] && (r[u[e]] = o[e]);
	this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function (i, u, f, e) {
		var o,
		h,
		c,
		l,
		a,
		v;
		try {
			h = i.id;
			c = OSF.DDA.AsyncMethodCalls[h];
			o = c.verifyAndExtractCall(u, f, e);
			l = r[h];
			a = n(h);
			v = t.toHost ? t.toHost(l, o) : o;
			a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
				dispId : l,
				hostCallArgs : v,
				onCalling : function () {
					OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
				},
				onReceiving : function () {
					OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
				},
				onComplete : function (n, i) {
					var r,
					u;
					r = n == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess ? t.fromHost ? t.fromHost(l, i) : i : i;
					u = c.processResponse(n, r, f, o);
					OSF.DDA.issueAsyncResult(o, n, u)
				}
			})
		} catch (y) {
			s(y, c, u, o)
		}
	};
	this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function (i, u, f) {
		function a(n) {
			var t,
			i;
			n == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess && (t = u.addEventHandler(e, l), t || (n = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed));
			n != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess && (i = OSF.DDA.ErrorCodeManager.getErrorArgs(n));
			OSF.DDA.issueAsyncResult(o, n, i)
		}
		var o,
		e,
		l,
		h,
		c,
		v;
		try {
			h = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
			o = h.verifyAndExtractCall(i, f, u);
			e = o[Microsoft.Office.WebExtension.Parameters.EventType];
			l = o[Microsoft.Office.WebExtension.Parameters.Handler];
			u.getEventHandlerCount(e) == 0 ? (c = r[e], v = n(e)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync], v({
					eventType : e,
					dispId : c,
					targetId : f.id || "",
					onCalling : function () {
						OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
					},
					onReceiving : function () {
						OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
					},
					onComplete : a,
					onEvent : function (n) {
						var i = t.fromHost(c, n);
						u.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(e, f, i))
					}
				})) : a(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
		} catch (y) {
			s(y, h, i, o)
		}
	};
	this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function (t, i, u) {
		function a(n) {
			var t;
			n != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess && (t = OSF.DDA.ErrorCodeManager.getErrorArgs(n));
			OSF.DDA.issueAsyncResult(e, n, t)
		}
		var e,
		f,
		h,
		c,
		l,
		o,
		v,
		y;
		try {
			c = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
			e = c.verifyAndExtractCall(t, u, i);
			f = e[Microsoft.Office.WebExtension.Parameters.EventType];
			h = e[Microsoft.Office.WebExtension.Parameters.Handler];
			h === null ? (o = i.clearEventHandlers(f), l = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) : (o = i.removeEventHandler(f, h), l = o ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist);
			o && i.getEventHandlerCount(f) == 0 ? (v = r[f], y = n(f)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync], y({
					eventType : f,
					dispId : v,
					targetId : u.id || "",
					onCalling : function () {
						OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
					},
					onReceiving : function () {
						OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
					},
					onComplete : a
				})) : a(l)
		} catch (p) {
			s(p, c, t, e)
		}
	}
};
OSF.DDA.DispIdHost.addAsyncMethods = function (n, t, i) {
	var f,
	r,
	u;
	for (f in t)
		r = t[f], u = r.displayName, n[u] || OSF.OUtil.defineEnumerableProperty(n, u, {
			value : function (t) {
				return function () {
					var r = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
					r(t, arguments, n, i)
				}
			}
			(r)
		})
};
OSF.DDA.DispIdHost.addEventSupport = function (n, t) {
	var i = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName,
	r = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
	n[i] || OSF.OUtil.defineEnumerableProperty(n, i, {
		value : function () {
			var i = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
			i(arguments, t, n)
		}
	});
	n[r] || OSF.OUtil.defineEnumerableProperty(n, r, {
		value : function () {
			var i = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
			i(arguments, t, n)
		}
	})
};
OsfMsAjaxFactory.isMsAjaxLoaded() || OSF._OfficeAppFactory && OSF._OfficeAppFactory && OSF._OfficeAppFactory.getLoadScriptHelper && OSF._OfficeAppFactory.getLoadScriptHelper().isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId) || (msAjaxCDNPath = (window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js", OsfMsAjaxFactory.loadMsAjaxFull(function () {
		if (!OsfMsAjaxFactory.isMsAjaxLoaded())
			throw "Not able to load MicrosoftAjax.js.";
	}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response = {
	Status : 0,
	Payload : 1
};
OSF.DDA.SafeArray.UniqueArguments = {
	Offset : "offset",
	Run : "run",
	BindingSpecificData : "bindingSpecificData",
	MergedCellGuid : "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException = function (n, t) {
	var i,
	r = n.number;
	if (r)
		switch (r) {
		case -2146828218:
			i = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
			break;
		case -2146827850:
		default:
			i = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError
		}
	t.onComplete && t.onComplete(i || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor = function () {
	function n(n) {
		var t,
		e,
		i,
		r,
		u,
		f,
		o;
		try {
			if (e = n.ubound(1), i = n.ubound(2), n = n.toArray(), e == 1 && i == 1)
				t = [n];
			else
				for (t = [], r = 0; r < e; r++) {
					for (u = [], f = 0; f < i; f++)
						o = n[r * i + f], o != OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid && u.push(o);
					u.length > 0 && t.push(u)
				}
		} catch (s) {}

		return t
	}
	var i = [],
	t = {};
	t[Microsoft.Office.WebExtension.Parameters.Data] = function () {
		var t = 0,
		i = 1;
		return {
			toHost : function (n) {
				if (typeof n != "string" && n[OSF.DDA.TableDataProperties.TableRows] !== undefined) {
					var r = [];
					r[t] = n[OSF.DDA.TableDataProperties.TableRows];
					r[i] = n[OSF.DDA.TableDataProperties.TableHeaders];
					n = r
				}
				return n
			},
			fromHost : function (r) {
				var f,
				e,
				u;
				return r.toArray ? (e = r.dimensions(), e === 2 ? f = n(r) : (u = r.toArray(), u.length === 2 && (u[0] != null && u[0].toArray || u[1] != null && u[1].toArray) ? (f = {}, f[OSF.DDA.TableDataProperties.TableRows] = n(u[t]), f[OSF.DDA.TableDataProperties.TableHeaders] = n(u[i])) : f = u)) : f = r,
				f
			}
		}
	}
	();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, i, t);
	this.unpack = function (n, i) {
		var r,
		u;
		return this.isComplexType(n) || OSF.DDA.ListType.isListType(n) ? (u = (i || typeof i == "unknown") && i.toArray, r = u ? i.toArray() : i || {}) : r = this.isDynamicType(n) ? t[n].fromHost(i) : i,
		r
	}
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor, [{
				type : Microsoft.Office.WebExtension.Parameters.ValueFormat,
				toHost : [{
						name : Microsoft.Office.WebExtension.ValueFormat.Unformatted,
						value : 0
					}, {
						name : Microsoft.Office.WebExtension.ValueFormat.Formatted,
						value : 1
					}
				]
			}, {
				type : Microsoft.Office.WebExtension.Parameters.FilterType,
				toHost : [{
						name : Microsoft.Office.WebExtension.FilterType.All,
						value : 0
					}
				]
			}
		]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.AsyncResultStatus,
	fromHost : [{
			name : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.AsyncResultStatus.Failed,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.executeAsync = function (n) {
	function t(n) {
		var i = n,
		f,
		r,
		u;
		if (OSF.OUtil.isArray(n))
			for (f = i.length, r = 0; r < f; r++)
				i[r] = t(i[r]);
		else if (OSF.OUtil.isDate(n))
			i = n.getVarDate();
		else if (typeof n == "object" && !OSF.OUtil.isArray(n)) {
			i = [];
			for (u in n)
				OSF.OUtil.isFunction(n[u]) || (i[u] = t(n[u]))
		}
		return i
	}
	function i(n) {
		var u = n,
		t,
		r;
		if (n != null && n.toArray) {
			for (t = n.toArray(), r = 0; r < t.length; r++)
				t[r] = i(t[r]);
			u = t
		}
		return u
	}
	try {
		n.onCalling && n.onCalling();
		var r = (new Date).getTime();
		OSF.ClientHostController.execute(n.dispId, t(n.hostCallArgs), function (t) {
			var f = t.toArray(),
			e = f[OSF.DDA.SafeArray.Response.Status],
			u,
			o,
			s;
			if (e == OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult)
				return u = f[OSF.DDA.SafeArray.Response.Payload], u = i(u), u != null && (n._chunkResultData || (n._chunkResultData = []), n._chunkResultData[u[0]] = u[1]), !1;
			if (n.onReceiving && n.onReceiving(), n.onComplete) {
				if (e == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
					if (f.length > 2)
						for (u = [], o = 1; o < f.length; o++)
							u[o - 1] = f[o];
					else
						u = f[OSF.DDA.SafeArray.Response.Payload];
					n._chunkResultData && (u = i(u), u != null && (s = u[u.length - 1], n._chunkResultData.length == s ? u[u.length - 1] = n._chunkResultData : e = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError))
				} else
					u = f[OSF.DDA.SafeArray.Response.Payload];
				n.onComplete(e, u)
			}
			return OSF.AppTelemetry && OSF.AppTelemetry.onMethodDone(n.dispId, n.hostCallArgs, Math.abs((new Date).getTime() - r), e),
			!0
		})
	} catch (u) {
		OSF.DDA.SafeArray.Delegate._onException(u, n)
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent = function (n, t) {
	var i = (new Date).getTime();
	return function (r) {
		t.onReceiving && t.onReceiving();
		var u = r.toArray ? r.toArray()[OSF.DDA.SafeArray.Response.Status] : r;
		t.onComplete && t.onComplete(u);
		OSF.AppTelemetry && OSF.AppTelemetry.onRegisterDone(n, t.dispId, Math.abs((new Date).getTime() - i), u)
	}
};
OSF.DDA.SafeArray.Delegate.registerEventAsync = function (n) {
	n.onCalling && n.onCalling();
	var t = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(!0, n);
	try {
		OSF.ClientHostController.registerEvent(n.dispId, n.targetId, function (t, i) {
			n.onEvent && n.onEvent(i);
			OSF.AppTelemetry && OSF.AppTelemetry.onEventDone(n.dispId)
		}, t)
	} catch (i) {
		OSF.DDA.SafeArray.Delegate._onException(i, n)
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync = function (n) {
	n.onCalling && n.onCalling();
	var t = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(!1, n);
	try {
		OSF.ClientHostController.unregisterEvent(n.dispId, n.targetId, t)
	} catch (i) {
		OSF.DDA.SafeArray.Delegate._onException(i, n)
	}
};
OSF.ClientMode = {
	ReadWrite : 0,
	ReadOnly : 1
};
OSF.DDA.RichInitializationReason = {
	1 : Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2 : Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper = function (n, t, i, r, u) {
	var f = this;
	f._hostInfo = n;
	f._webAppState = t;
	f._context = i;
	f._settings = r;
	f._hostFacade = u;
	f._initializeSettings = f.initializeSettings
};
OSF.InitializationHelper.prototype.deserializeSettings = function (n, t) {
	var f = "undefined",
	r = OSF.OUtil.getSessionStorage(),
	i,
	u;
	return r && (i = r.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey()), i ? n = typeof JSON !== f ? JSON.parse(i) : OsfMsAjaxFactory.msAjaxSerializer.deserialize(i, !0) : (i = typeof JSON !== f ? JSON.stringify(n) : OsfMsAjaxFactory.msAjaxSerializer.serialize(n), r.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), i))),
	u = OSF.DDA.SettingsManager.deserializeSettings(n),
	t ? new OSF.DDA.RefreshableSettings(u) : new OSF.DDA.Settings(u)
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function () {};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function (n) {
	this.prepareApiSurface(n);
	Microsoft.Office.WebExtension.initialize(this.getInitializationReason(n))
};
OSF.InitializationHelper.prototype.prepareApiSurface = function (n) {
	var r = new OSF.DDA.License(n.get_eToken()),
	u = OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme ? OSF.DDA.OfficeTheme.getOfficeTheme : null,
	t,
	i;
	OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(n, n.doc, r, null, u));
	t = OSF.DDA.DispIdHost.getClientDelegateMethods;
	i = OSF.DDA.SafeArray.Delegate.ParameterMap;
	OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(t, i))
};
OSF.InitializationHelper.prototype.getInitializationReason = function (n) {
	return OSF.DDA.RichInitializationReason[n.get_reason()]
};
OSF.DDA.DispIdHost.getClientDelegateMethods = function (n) {
	var t = {},
	i,
	r;
	return t[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.SafeArray.Delegate.executeAsync,
	t[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.SafeArray.Delegate.registerEventAsync,
	t[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.SafeArray.Delegate.unregisterEventAsync,
	OSF.DDA.AsyncMethodNames.RefreshAsync && n == OSF.DDA.AsyncMethodNames.RefreshAsync.id && (i = function (n, t, i) {
		return OSF.DDA.ClientSettingsManager.read(t, i)
	}, t[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(i)),
	OSF.DDA.AsyncMethodNames.SaveAsync && n == OSF.DDA.AsyncMethodNames.SaveAsync.id && (r = function (n, t, i) {
		return OSF.DDA.ClientSettingsManager.write(n[OSF.DDA.SettingsManager.SerializedSettings], n[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], t, i)
	}, t[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(r)),
	t
};
OSF = OSF || {}, function (n) {
	var t = function () {
		function n(n) {
			this.data = n;
			this.safeArrayFlag = this.isSafeArray(n)
		}
		var t = !1;
		return n.prototype.dimensions = function () {
			var n = 0;
			return this.safeArrayFlag ? n = this.data[0][0] : this.isArray() && (n = 2),
			n
		},
		n.prototype.getItem = function () {
			var i = [],
			n = null,
			t;
			for (i = this.safeArrayFlag ? this.toArray() : this.data, n = i, t = 0; t < arguments.length; t++)
				n = n[arguments[t]];
			return n
		},
		n.prototype.lbound = function () {
			return 0
		},
		n.prototype.ubound = function (n) {
			var t = this,
			i = 0;
			if (t.safeArrayFlag)
				i = t.data[0][n];
			else if (t.isArray())
				if (n == 1)
					return t.data.length;
				else if (n == 2)
					if (OSF.OUtil.isArray(t.data[0]))
						return t.data[0].length;
					else if (t.data[0] != null)
						return 1;
			return i
		},
		n.prototype.toArray = function () {
			var i = this,
			r;
			if (i.isArray() == t)
				return i.data;
			for (var u = [], e = i.safeArrayFlag ? 1 : 0, f = e; f < i.data.length; f++)
				r = i.data[f], i.isSafeArray(r) ? u.push(new n(r)) : u.push(r);
			return u
		},
		n.prototype.isArray = function () {
			return OSF.OUtil.isArray(this.data)
		},
		n.prototype.isSafeArray = function (n) {
			var e = t,
			i,
			o,
			r,
			u,
			f;
			if (OSF.OUtil.isArray(n) && OSF.OUtil.isArray(n[0])) {
				if (i = n[0], o = i[0], i.length != o + 1)
					return t;
				for (r = 1, u = 1; u < i.length; u++) {
					if (f = i[u], isFinite(f) == t)
						return t;
					r = r * f
				}
				r++;
				e = r == n.length
			}
			return e
		},
		n
	}
	();
	n.WebkitSafeArray = t
}
(OSFWebkit || (OSFWebkit = {})), function (n) {
	(function (n) {
		function u(n, i) {
			t.agaveHostCallback(n, i)
		}
		function f(n, i) {
			t.agaveHostEventCallback(n, i)
		}
		function e() {
			return t == null && (t = new i("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback")),
			t
		}
		var t = null,
		r,
		i;
		n.agaveHostCallback = u;
		n.agaveHostEventCallback = f;
		n.GetScriptMessenger = e;
		r = function () {
			function n(n, t, i) {
				this.id = n;
				this.targetId = t;
				this.handler = i
			}
			return n
		}
		();
		i = function () {
			function n(t, i) {
				var r = this;
				r.callingIndex = 0;
				r.callbackList = {};
				r.eventHandlerList = {};
				r.asyncMethodCallbackFunctionName = t;
				r.eventCallbackFunctionName = i;
				r.conversationId = n.getCurrentTimeMS().toString()
			}
			return n.prototype.invokeMethod = function (n, t, i, r) {
				var u = {};
				this.postWebkitMessage(u, n, t, i, r)
			},
			n.prototype.registerEvent = function (n, t, i, u, f, e) {
				var o = {
					eventCallbackFunction : this.eventCallbackFunctionName
				},
				s = {
					id : i,
					targetId : u
				},
				h = this.postWebkitMessage(o, n, t, s, e);
				this.eventHandlerList[h] = new r(i, u, f)
			},
			n.prototype.unregisterEvent = function (n, t, i, r, u) {
				var f = this,
				s = {
					id : i,
					targetId : r
				},
				e,
				o;
				for (e in f.eventHandlerList)
					f.eventHandlerList.hasOwnProperty(e) && (o = f.eventHandlerList[e], o.id == i && o.targetId == r && delete f.eventHandlerList[e]);
				f.invokeMethod(n, t, s, u)
			},
			n.prototype.agaveHostCallback = function (n, t) {
				var r = this.callbackList[n],
				i;
				r && (i = r(t), (i === undefined || i === !0) && delete this.callbackList[n])
			},
			n.prototype.agaveHostEventCallback = function (n, t) {
				var i = this.eventHandlerList[n];
				i && i.handler(t)
			},
			n.prototype.postWebkitMessage = function (t, i, r, u, f) {
				var e = this,
				s = e.generateCorrelationId(),
				h,
				o;
				return e.callbackList[s] = f,
				t.methodId = r,
				t.params = u,
				t.callbackId = s,
				t.callbackFunction = e.asyncMethodCallbackFunctionName,
				h = function () {
					window.webkit.messageHandlers[i].postMessage(JSON.stringify(t))
				},
				o = n.getCurrentTimeMS(),
				e.lastMessageTimestamp == null || o - e.lastMessageTimestamp >= n.MESSAGE_TIME_DELTA ? (h(), e.lastMessageTimestamp = o) : (e.lastMessageTimestamp += n.MESSAGE_TIME_DELTA, setTimeout(function () {
						h()
					}, e.lastMessageTimestamp - o)),
				s
			},
			n.prototype.generateCorrelationId = function () {
				return ++this.callingIndex,
				this.conversationId + this.callingIndex
			},
			n.getCurrentTimeMS = function () {
				return (new Date).getTime()
			},
			n.MESSAGE_TIME_DELTA = 10,
			n
		}
		();
		n.WebkitScriptMessaging = i
	})(n.ScriptMessaging || (n.ScriptMessaging = {}));
	var t = n.ScriptMessaging
}
(OSFWebkit || (OSFWebkit = {}));
OSF.ScriptMessaging = OSFWebkit.ScriptMessaging, function (n) {
	var i,
	r,
	t;
	n.MessageHandlerName = "Agave",
	function (n) {
		var t = "RequirementMatrix",
		i = "CommerceAllowed",
		r = "TouchEnabled",
		u = "InstanceId",
		f = "CorrelationId",
		e = "APISetVersion",
		o = "SolutionToken",
		s = "ControlIntegrationLevel",
		h = "ActivationMode",
		c = "DocumentUrl",
		l = "BindingCount",
		a = "AppDataLocale",
		v = "APPUILocale",
		y = "AppCapabilities",
		p = "APIVersionSequence",
		w = "RevisionVersion",
		b = "MinorVersion",
		k = "MajorVersion",
		d = "SolutionReferenceId",
		g = "Settings";
		n[n[g] = 0] = g;
		n[n[d] = 1] = d;
		n[n.AppType = 2] = "AppType";
		n[n[k] = 3] = k;
		n[n[b] = 4] = b;
		n[n[w] = 5] = w;
		n[n[p] = 6] = p;
		n[n[y] = 7] = y;
		n[n[v] = 8] = v;
		n[n[a] = 9] = a;
		n[n[l] = 10] = l;
		n[n[c] = 11] = c;
		n[n[h] = 12] = h;
		n[n[s] = 13] = s;
		n[n[o] = 14] = o;
		n[n[e] = 15] = e;
		n[n[f] = 16] = f;
		n[n[u] = 17] = u;
		n[n[r] = 18] = r;
		n[n[i] = 19] = i;
		n[n[t] = 20] = t
	}
	(n.AppContextProperties || (n.AppContextProperties = {}));
	i = n.AppContextProperties,
	function (n) {
		var t = "GetContext",
		i = "WriteSettings",
		r = "UnregisterEvent",
		u = "RegisterEvent";
		n[n.Execute = 1] = "Execute";
		n[n[u] = 2] = u;
		n[n[r] = 3] = r;
		n[n[i] = 4] = i;
		n[n[t] = 5] = t
	}
	(n.MethodId || (n.MethodId = {}));
	r = n.MethodId;
	t = function () {
		function t(n) {
			this.hostScriptProxy = n
		}
		return t.prototype.execute = function (t, i, r) {
			var u = i,
			f,
			e;
			u == null && (u = []);
			f = {
				id : t,
				apiArgs : u
			};
			e = function (t) {
				var i = t,
				u;
				return OSF.OUtil.isArray(t) && t.length >= 2 && (u = t[0], i = t[1]),
				r ? r(new n.WebkitSafeArray(i)) : void 0
			};
			this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.Execute, f, e)
		},
		t.prototype.registerEvent = function (t, i, r, u) {
			var f = function (t) {
				var i = t,
				u = 0;
				OSF.OUtil.isArray(t) && t.length >= 2 && (i = t[0], u = t[1]);
				r && r(u, new n.WebkitSafeArray(i))
			},
			e = function (t) {
				if (u)
					return u(new n.WebkitSafeArray(t))
			};
			this.hostScriptProxy.registerEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.RegisterEvent, t, i, f, e)
		},
		t.prototype.unregisterEvent = function (t, i, r) {
			var u = function (t) {
				return r(new n.WebkitSafeArray(t))
			};
			this.hostScriptProxy.unregisterEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.UnregisterEvent, t, i, u)
		},
		t
	}
	();
	n.WebkitHostController = t
}
(OSFWebkit || (OSFWebkit = {}));
OSF.Webkit = OSFWebkit;
OSF.ClientHostController = new OSFWebkit.WebkitHostController(OSF.ScriptMessaging.GetScriptMessenger());
OSF.DDA.ClientSettingsManager = {
	getSettingsExecuteMethod : function (n) {
		return function (t) {
			var i,
			r,
			u = function (n, i) {
				t.onReceiving && t.onReceiving();
				t.onComplete && t.onComplete(n, i)
			};
			try {
				n(t.hostCallArgs, t.onCalling, u)
			} catch (f) {
				i = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				r = {
					name : Strings.OfficeOM.L_InternalError,
					message : f
				};
				u(i, r)
			}
		}
	},
	read : function (n, t) {
		var u = [],
		f = [],
		i,
		r;
		n && n();
		i = OSF._OfficeAppFactory.getInitializationHelper();
		r = function (n) {
			t && t(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, n.get_settings())
		};
		i.getAppContext(null, r)
	},
	write : function (n, t, i, r) {
		var u = {},
		e = [],
		o = [],
		f,
		s;
		for (f in n)
			e.push(f), o.push(n[f]);
		u.keys = e;
		u.values = o;
		i && i();
		s = function (n) {
			r && r(n[0], null)
		};
		OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.WriteSettings, u, s)
	}
};
OSF.InitializationHelper.prototype.initializeSettings = function (n, t) {
	var i = n.get_settings();
	return this.deserializeSettings(i, t)
};
OSF.InitializationHelper.prototype.getAppContext = function (n, t) {
	var i = function (n) {
		var f,
		i = OSF.Webkit.AppContextProperties,
		e = n[i.AppType],
		o = !1;
		for (var l in OSF.AppName)
			if (OSF.AppName[l] == e) {
				o = !0;
				break
			}
		if (!o)
			throw "Unsupported client type " + e;
		for (var s = n[i.Settings], h = {}, c = s[0], a = s[1], r = 0; r < c.length; r++)
			h[c[r]] = a[r];
		var v = n[i.SolutionReferenceId],
		y = n[i.MajorVersion],
		p = n[i.AppCapabilities],
		w = n[i.APPUILocale],
		b = n[i.AppDataLocale],
		k = n[i.DocumentUrl],
		d = n[i.ActivationMode],
		g = n[i.ControlIntegrationLevel],
		u = n[i.SolutionToken];
		u = u ? u.toString() : "";
		var nt = n[i.CorrelationId],
		tt = n[i.InstanceId],
		it = n[i.TouchEnabled],
		rt = n[i.CommerceAllowed],
		ut = n[i.MinorVersion],
		ft = n[i.RequirementMatrix];
		f = new OSF.OfficeAppContext(v, e, y, w, b, k, p, h, d, g, u, nt, tt, it, rt, ut, ft);
		OSF.AppTelemetry && OSF.AppTelemetry.initialize(f);
		t(f)
	};
	OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.GetContext, [], i)
}, function (n) {
	var u = "ResponseTime",
	i = "SessionId",
	r = "CorrelationId",
	t = !0,
	f = function () {
		function n(n) {
			this._table = n;
			this._fields = {}

		}
		return Object.defineProperty(n.prototype, "Fields", {
			get : function () {
				return this._fields
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(n.prototype, "Table", {
			get : function () {
				return this._table
			},
			enumerable : t,
			configurable : t
		}),
		n.prototype.SerializeFields = function () {},
		n.prototype.SetSerializedField = function (n, t) {
			typeof t != "undefined" && t !== null && (this._serializedFields[n] = t.toString())
		},
		n.prototype.SerializeRow = function () {
			var n = this;
			return n._serializedFields = {},
			n.SetSerializedField("Table", n._table),
			n.SerializeFields(),
			JSON.stringify(n._serializedFields)
		},
		n
	}
	(),
	e,
	o,
	s,
	h,
	c;
	n.BaseUsageData = f;
	e = function (n) {
		function u() {
			n.call(this, "AppActivated")
		}
		var f = "AppSizeHeight",
		e = "AppSizeWidth",
		o = "ClientId",
		s = "HostVersion",
		h = "Host",
		c = "UserId",
		l = "Browser",
		a = "AssetId",
		v = "AppURL",
		y = "AppInstanceId",
		p = "AppId";
		return __extends(u, n),
		Object.defineProperty(u.prototype, r, {
			get : function () {
				return this.Fields[r]
			},
			set : function (n) {
				this.Fields[r] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, i, {
			get : function () {
				return this.Fields[i]
			},
			set : function (n) {
				this.Fields[i] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, p, {
			get : function () {
				return this.Fields[p]
			},
			set : function (n) {
				this.Fields[p] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, y, {
			get : function () {
				return this.Fields[y]
			},
			set : function (n) {
				this.Fields[y] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, v, {
			get : function () {
				return this.Fields[v]
			},
			set : function (n) {
				this.Fields[v] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, a, {
			get : function () {
				return this.Fields[a]
			},
			set : function (n) {
				this.Fields[a] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, l, {
			get : function () {
				return this.Fields[l]
			},
			set : function (n) {
				this.Fields[l] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, c, {
			get : function () {
				return this.Fields[c]
			},
			set : function (n) {
				this.Fields[c] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, h, {
			get : function () {
				return this.Fields[h]
			},
			set : function (n) {
				this.Fields[h] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, s, {
			get : function () {
				return this.Fields[s]
			},
			set : function (n) {
				this.Fields[s] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, o, {
			get : function () {
				return this.Fields[o]
			},
			set : function (n) {
				this.Fields[o] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, e, {
			get : function () {
				return this.Fields[e]
			},
			set : function (n) {
				this.Fields[e] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, f, {
			get : function () {
				return this.Fields[f]
			},
			set : function (n) {
				this.Fields[f] = n
			},
			enumerable : t,
			configurable : t
		}),
		u.prototype.SerializeFields = function () {
			var n = this;
			n.SetSerializedField(r, n.CorrelationId);
			n.SetSerializedField(i, n.SessionId);
			n.SetSerializedField(p, n.AppId);
			n.SetSerializedField(y, n.AppInstanceId);
			n.SetSerializedField(v, n.AppURL);
			n.SetSerializedField(a, n.AssetId);
			n.SetSerializedField(l, n.Browser);
			n.SetSerializedField(c, n.UserId);
			n.SetSerializedField(h, n.Host);
			n.SetSerializedField(s, n.HostVersion);
			n.SetSerializedField(o, n.ClientId);
			n.SetSerializedField(e, n.AppSizeWidth);
			n.SetSerializedField(f, n.AppSizeHeight)
		},
		u
	}
	(f);
	n.AppActivatedUsageData = e;
	o = function (n) {
		function f() {
			n.call(this, "ScriptLoad")
		}
		var e = "StartTime",
		o = "ScriptId";
		return __extends(f, n),
		Object.defineProperty(f.prototype, r, {
			get : function () {
				return this.Fields[r]
			},
			set : function (n) {
				this.Fields[r] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, i, {
			get : function () {
				return this.Fields[i]
			},
			set : function (n) {
				this.Fields[i] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, o, {
			get : function () {
				return this.Fields[o]
			},
			set : function (n) {
				this.Fields[o] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, e, {
			get : function () {
				return this.Fields[e]
			},
			set : function (n) {
				this.Fields[e] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, u, {
			get : function () {
				return this.Fields[u]
			},
			set : function (n) {
				this.Fields[u] = n
			},
			enumerable : t,
			configurable : t
		}),
		f.prototype.SerializeFields = function () {
			var n = this;
			n.SetSerializedField(r, n.CorrelationId);
			n.SetSerializedField(i, n.SessionId);
			n.SetSerializedField(o, n.ScriptId);
			n.SetSerializedField(e, n.StartTime);
			n.SetSerializedField(u, n.ResponseTime)
		},
		f
	}
	(f);
	n.ScriptLoadUsageData = o;
	s = function (n) {
		function u() {
			n.call(this, "AppClosed")
		}
		var f = "CloseMethod",
		e = "OpenTime",
		o = "AppSizeFinalHeight",
		s = "AppSizeFinalWidth",
		h = "FocusTime";
		return __extends(u, n),
		Object.defineProperty(u.prototype, r, {
			get : function () {
				return this.Fields[r]
			},
			set : function (n) {
				this.Fields[r] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, i, {
			get : function () {
				return this.Fields[i]
			},
			set : function (n) {
				this.Fields[i] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, h, {
			get : function () {
				return this.Fields[h]
			},
			set : function (n) {
				this.Fields[h] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, s, {
			get : function () {
				return this.Fields[s]
			},
			set : function (n) {
				this.Fields[s] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, o, {
			get : function () {
				return this.Fields[o]
			},
			set : function (n) {
				this.Fields[o] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, e, {
			get : function () {
				return this.Fields[e]
			},
			set : function (n) {
				this.Fields[e] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, f, {
			get : function () {
				return this.Fields[f]
			},
			set : function (n) {
				this.Fields[f] = n
			},
			enumerable : t,
			configurable : t
		}),
		u.prototype.SerializeFields = function () {
			var n = this;
			n.SetSerializedField(r, n.CorrelationId);
			n.SetSerializedField(i, n.SessionId);
			n.SetSerializedField(h, n.FocusTime);
			n.SetSerializedField(s, n.AppSizeFinalWidth);
			n.SetSerializedField(o, n.AppSizeFinalHeight);
			n.SetSerializedField(e, n.OpenTime);
			n.SetSerializedField(f, n.CloseMethod)
		},
		u
	}
	(f);
	n.AppClosedUsageData = s;
	h = function (n) {
		function f() {
			n.call(this, "APIUsage")
		}
		var e = "ErrorType",
		o = "Parameters",
		s = "APIID",
		h = "APIType";
		return __extends(f, n),
		Object.defineProperty(f.prototype, r, {
			get : function () {
				return this.Fields[r]
			},
			set : function (n) {
				this.Fields[r] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, i, {
			get : function () {
				return this.Fields[i]
			},
			set : function (n) {
				this.Fields[i] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, h, {
			get : function () {
				return this.Fields[h]
			},
			set : function (n) {
				this.Fields[h] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, s, {
			get : function () {
				return this.Fields[s]
			},
			set : function (n) {
				this.Fields[s] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, o, {
			get : function () {
				return this.Fields[o]
			},
			set : function (n) {
				this.Fields[o] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, u, {
			get : function () {
				return this.Fields[u]
			},
			set : function (n) {
				this.Fields[u] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(f.prototype, e, {
			get : function () {
				return this.Fields[e]
			},
			set : function (n) {
				this.Fields[e] = n
			},
			enumerable : t,
			configurable : t
		}),
		f.prototype.SerializeFields = function () {
			var n = this;
			n.SetSerializedField(r, n.CorrelationId);
			n.SetSerializedField(i, n.SessionId);
			n.SetSerializedField(h, n.APIType);
			n.SetSerializedField(s, n.APIID);
			n.SetSerializedField(o, n.Parameters);
			n.SetSerializedField(u, n.ResponseTime);
			n.SetSerializedField(e, n.ErrorType)
		},
		f
	}
	(f);
	n.APIUsageUsageData = h;
	c = function (n) {
		function u() {
			n.call(this, "AppInitialization")
		}
		var f = "Message",
		e = "SuccessCode";
		return __extends(u, n),
		Object.defineProperty(u.prototype, r, {
			get : function () {
				return this.Fields[r]
			},
			set : function (n) {
				this.Fields[r] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, i, {
			get : function () {
				return this.Fields[i]
			},
			set : function (n) {
				this.Fields[i] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, e, {
			get : function () {
				return this.Fields[e]
			},
			set : function (n) {
				this.Fields[e] = n
			},
			enumerable : t,
			configurable : t
		}),
		Object.defineProperty(u.prototype, f, {
			get : function () {
				return this.Fields[f]
			},
			set : function (n) {
				this.Fields[f] = n
			},
			enumerable : t,
			configurable : t
		}),
		u.prototype.SerializeFields = function () {
			var n = this;
			n.SetSerializedField(r, n.CorrelationId);
			n.SetSerializedField(i, n.SessionId);
			n.SetSerializedField(e, n.SuccessCode);
			n.SetSerializedField(f, n.Message)
		},
		u
	}
	(f);
	n.AppInitializationUsageData = c
}
(OSFLog || (OSFLog = {})), function (n) {
	"use strict";
	function u() {
		OSF.Logger && OSF.Logger.ulsEndpoint && OSF.Logger.ulsEndpoint.loadProxyFrame()
	}
	function f(n, t, i) {
		if (OSF.Logger && OSF.Logger.ulsEndpoint) {
			var r = {
				traceLevel : n,
				message : t,
				flag : i,
				internalLog : !0
			},
			u = JSON.stringify(r);
			OSF.Logger.ulsEndpoint.writeLog(u)
		}
	}
	function e() {
		try {
			return new t
		} catch (n) {
			return null
		}
	}
	var i,
	r,
	t;
	(function (n) {
		n[n.info = 0] = "info";
		n[n.warning = 1] = "warning";
		n[n.error = 2] = "error"
	})(n.TraceLevel || (n.TraceLevel = {}));
	i = n.TraceLevel,
	function (n) {
		n[n.none = 0] = "none";
		n[n.flush = 1] = "flush"
	}
	(n.SendFlag || (n.SendFlag = {}));
	r = n.SendFlag;
	n.allowUploadingData = u;
	n.sendLog = f;
	t = function () {
		function n() {
			var n = this,
			t = n;
			n.proxyFrame = null;
			n.telemetryEndPoint = "https://telemetryservice.firstpartyapps.oaspapps.com/telemetryservice/telemetryproxy.html";
			n.buffer = [];
			n.proxyFrameReady = !1;
			OSF.OUtil.addEventListener(window, "message", function (n) {
				return t.tellProxyFrameReady(n)
			});
			setTimeout(function () {
				t.loadProxyFrame()
			}, 3e3)
		}
		return n.prototype.writeLog = function (n) {
			var t = this;
			t.proxyFrameReady === !0 ? t.proxyFrame.contentWindow.postMessage(n, "*") : t.buffer.length < 128 && t.buffer.push(n)
		},
		n.prototype.loadProxyFrame = function () {
			var n = this;
			n.proxyFrame == null && (n.proxyFrame = document.createElement("iframe"), n.proxyFrame.setAttribute("style", "display:none"), n.proxyFrame.setAttribute("src", n.telemetryEndPoint), document.head.appendChild(n.proxyFrame))
		},
		n.prototype.tellProxyFrameReady = function (n) {
			var t = this,
			f = t,
			i,
			r,
			u;
			if (n.data === "ProxyFrameReadyToLog") {
				for (t.proxyFrameReady = !0, i = 0; i < t.buffer.length; i++)
					t.writeLog(t.buffer[i]);
				t.buffer.length = 0;
				OSF.OUtil.removeEventListener(window, "message", function (n) {
					return f.tellProxyFrameReady(n)
				})
			} else
				n.data === "ProxyFrameReadyToInit" && (r = {
						appName : "Office APPs",
						sessionId : OSF.OUtil.Guid.generateNewGuid()
					}, u = JSON.stringify(r), t.proxyFrame.contentWindow.postMessage(u, "*"))
		},
		n
	}
	();
	OSF.Logger || (OSF.Logger = n);
	n.ulsEndpoint = e()
}
(Logger || (Logger = {})), function (n) {
	function c(r) {
		if (OSF.Logger && !t) {
			t = new h;
			t.hostVersion = r.get_appVersion();
			t.appId = r.get_id();
			t.host = r.get_appName();
			t.browser = window.navigator.userAgent;
			t.correlationId = r.get_correlationId();
			t.clientId = (new o).getClientId();
			t.appInstanceId = r.get_appInstanceId();
			t.appInstanceId && (t.appInstanceId = t.appInstanceId.replace(/[{}]/g, "").toLowerCase());
			var f = location.href.indexOf("?");
			t.appURL = f == -1 ? location.href : location.href.substring(0, f),
			function (n, t) {
				var u,
				f,
				r;
				t.assetId = "";
				t.userId = "";
				try {
					u = decodeURIComponent(n);
					f = new DOMParser;
					r = f.parseFromString(u, "text/xml");
					t.userId = r.getElementsByTagName("t")[0].attributes.getNamedItem("cid").nodeValue;
					t.assetId = r.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue
				} catch (e) {}

				finally {
					u = i;
					r = i;
					f = i
				}
			}
			(r.get_eToken(), t),
			function () {
				var c = new Date,
				r = i,
				o = 0,
				h = !1,
				f = function () {
					document.hasFocus() ? r == i && (r = new Date) : r && (o += Math.abs((new Date).getTime() - r.getTime()), r = i)
				},
				t = [],
				s,
				e;
				for (t.push(new u("focus", f)), t.push(new u("blur", f)), t.push(new u("focusout", f)), t.push(new u("focusin", f)), s = function () {
					for (var u = 0; u < t.length; u++)
						OSF.OUtil.removeEventListener(window, t[u].name, t[u].handler);
						if (t.length = 0, !h) {
							document.hasFocus() && r && (o += Math.abs((new Date).getTime() - r.getTime()), r = i);
							n.onAppClosed(Math.abs((new Date).getTime() - c.getTime()), o);
							h = !0
						}
					}, t.push(new u("beforeunload", s)), t.push(new u("unload", s)), e = 0; e < t.length; e++)OSF.OUtil.addEventListener(window, t[e].name, t[e].handler);
				f()
			}
			();
			n.onAppActivated()
		}
	}
	function l() {
		if (t) {
			(new o).enumerateLog(function (n, t) {
				return (new f).LogRawData(t)
			}, !0);
			var n = new OSFLog.AppActivatedUsageData;
			n.SessionId = r;
			n.AppId = t.appId;
			n.AssetId = t.assetId;
			n.AppURL = t.appURL;
			n.UserId = t.userId;
			n.ClientId = t.clientId;
			n.Browser = t.browser;
			n.Host = t.host;
			n.HostVersion = t.hostVersion;
			n.CorrelationId = t.correlationId;
			n.AppSizeWidth = window.innerWidth;
			n.AppSizeHeight = window.innerHeight;
			n.AppInstanceId = t.appInstanceId;
			(new f).LogData(n);
			setTimeout(function () {
				OSF.Logger && OSF.Logger.allowUploadingData()
			}, 100)
		}
	}
	function a(n, t, i, u) {
		var e = new OSFLog.ScriptLoadUsageData;
		e.CorrelationId = u;
		e.SessionId = r;
		e.ScriptId = n;
		e.StartTime = t;
		e.ResponseTime = i;
		(new f).LogData(e)
	}
	function v(n, i, u, o, s) {
		if (t) {
			var h = new OSFLog.APIUsageUsageData;
			h.CorrelationId = e;
			h.SessionId = r;
			h.APIType = n;
			h.APIID = i;
			h.Parameters = u;
			h.ResponseTime = o;
			h.ErrorType = s;
			(new f).LogData(h)
		}
	}
	function y(n, t, r, u) {
		var f = i,
		e;
		if (t)
			if (typeof t == "number")
				f = String(t);
			else if (typeof t == "object")
				for (e in t)
					f !== i ? f += "," : f = "", typeof t[e] == "number" && (f += String(t[e]));
			else
				f = "";
		OSF.AppTelemetry.onCallDone("method", n, f, r, u)
	}
	function p(n, t) {
		OSF.AppTelemetry.onCallDone("property", -1, n, t)
	}
	function w(n, t) {
		OSF.AppTelemetry.onCallDone("event", n, i, 0, t)
	}
	function b(n, t, r, u) {
		OSF.AppTelemetry.onCallDone(n ? "registerevent" : "unregisterevent", t, i, r, u)
	}
	function k(n, i) {
		if (t) {
			var u = new OSFLog.AppClosedUsageData;
			u.CorrelationId = e;
			u.SessionId = r;
			u.FocusTime = i;
			u.OpenTime = n;
			u.AppSizeFinalWidth = window.innerWidth;
			u.AppSizeFinalHeight = window.innerHeight;
			(new o).saveLog(r, u.SerializeRow())
		}
	}
	function d(n) {
		e = n
	}
	function s(n, t) {
		var i = new OSFLog.AppInitializationUsageData;
		i.CorrelationId = e;
		i.SessionId = r;
		i.SuccessCode = n ? 1 : 0;
		i.Message = t;
		(new f).LogData(i)
	}
	function g(n) {
		s(!1, n)
	}
	function nt(n) {
		s(!0, n)
	}
	var i = null;
	"use strict";
	var t,
	r = OSF.OUtil.Guid.generateNewGuid(),
	e = "",
	h = function () {
		function n() {}

		return n
	}
	(),
	u = function () {
		function n(n, t) {
			this.name = n;
			this.handler = t
		}
		return n
	}
	(),
	o = function () {
		function n() {
			this.clientIDKey = "Office API client";
			this.logIdSetKey = "Office App Log Id Set"
		}
		return n.prototype.getClientId = function () {
			var t = this,
			n = t.getValue(t.clientIDKey);
			return (!n || n.length <= 0 || n.length > 40) && (n = OSF.OUtil.Guid.generateNewGuid(), t.setValue(t.clientIDKey, n)),
			n
		},
		n.prototype.saveLog = function (n, t) {
			var i = this,
			r = i.getValue(i.logIdSetKey);
			r = (r && r.length > 0 ? r + ";" : "") + n;
			i.setValue(i.logIdSetKey, r);
			i.setValue(n, t)
		},
		n.prototype.enumerateLog = function (n, t) {
			var i = this,
			e = i.getValue(i.logIdSetKey),
			u,
			o,
			r,
			f;
			if (e) {
				u = e.split(";");
				for (o in u)
					r = u[o], f = i.getValue(r), f && (n && n(r, f), t && i.remove(r));
				t && i.remove(i.logIdSetKey)
			}
		},
		n.prototype.getValue = function (n) {
			var t = OSF.OUtil.getLocalStorage(),
			i = "";
			return t && (i = t.getItem(n)),
			i
		},
		n.prototype.setValue = function (n, t) {
			var i = OSF.OUtil.getLocalStorage();
			i && i.setItem(n, t)
		},
		n.prototype.remove = function (n) {
			var t = OSF.OUtil.getLocalStorage();
			if (t)
				try {
					t.removeItem(n)
				} catch (i) {}

		},
		n
	}
	(),
	f = function () {
		function n() {}

		return n.prototype.LogData = function (n) {
			OSF.Logger && OSF.Logger.sendLog(OSF.Logger.TraceLevel.info, n.SerializeRow(), OSF.Logger.SendFlag.none)
		},
		n.prototype.LogRawData = function (n) {
			OSF.Logger && OSF.Logger.sendLog(OSF.Logger.TraceLevel.info, n, OSF.Logger.SendFlag.none)
		},
		n
	}
	();
	n.initialize = c;
	n.onAppActivated = l;
	n.onScriptDone = a;
	n.onCallDone = v;
	n.onMethodDone = y;
	n.onPropertyDone = p;
	n.onEventDone = w;
	n.onRegisterDone = b;
	n.onAppClosed = k;
	n.setOsfControlAppCorrelationId = d;
	n.doAppInitializationLogging = s;
	n.logAppCommonMessage = g;
	n.logAppException = nt;
	OSF.AppTelemetry = n
}
(OSFAppTelemetry || (OSFAppTelemetry = {}));
Microsoft.Office.WebExtension.FileType = {
	Text : "text",
	Compressed : "compressed",
	Pdf : "pdf"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FileProperties : "FileProperties",
	FileSliceProperties : "FileSliceProperties"
});
OSF.DDA.FileProperties = {
	Handle : "FileHandle",
	FileSize : "FileSize",
	SliceSize : Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.File = function (n, t, i) {
	var r,
	u;
	OSF.OUtil.defineEnumerableProperties(this, {
		size : {
			value : t
		},
		sliceCount : {
			value : Math.ceil(t / i)
		}
	});
	r = {};
	r[OSF.DDA.FileProperties.Handle] = n;
	r[OSF.DDA.FileProperties.SliceSize] = i;
	u = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [u.GetDocumentCopyChunkAsync, u.ReleaseDocumentCopyAsync], r)
};
OSF.DDA.FileSliceOffset = "fileSliceoffset";
OSF.DDA.AsyncMethodNames.addNames({
	GetDocumentCopyAsync : "getFileAsync",
	GetDocumentCopyChunkAsync : "getSliceAsync",
	ReleaseDocumentCopyAsync : "closeAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.FileType,
			"enum" : Microsoft.Office.WebExtension.FileType
		}
	],
	supportedOptions : [{
			name : Microsoft.Office.WebExtension.Parameters.SliceSize,
			value : {
				types : ["number"],
				defaultValue : 4194304
			}
		}
	],
	checkCallArgs : function (n) {
		var t = n[Microsoft.Office.WebExtension.Parameters.SliceSize];
		if (t <= 0 || t > 4194304)
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;
		return n
	},
	onSucceeded : function (n, t, i) {
		return new OSF.DDA.File(n[OSF.DDA.FileProperties.Handle], n[OSF.DDA.FileProperties.FileSize], i[Microsoft.Office.WebExtension.Parameters.SliceSize])
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.SliceIndex,
			types : ["number"]
		}
	],
	privateStateCallbacks : [{
			name : OSF.DDA.FileProperties.Handle,
			value : function (n, t) {
				return t[OSF.DDA.FileProperties.Handle]
			}
		}, {
			name : OSF.DDA.FileProperties.SliceSize,
			value : function (n, t) {
				return t[OSF.DDA.FileProperties.SliceSize]
			}
		}
	],
	checkCallArgs : function (n, t, i) {
		var r = n[Microsoft.Office.WebExtension.Parameters.SliceIndex];
		if (r < 0 || r >= t.sliceCount)
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
		return n[OSF.DDA.FileSliceOffset] = parseInt((r * i[OSF.DDA.FileProperties.SliceSize]).toString()),
		n
	},
	onSucceeded : function (n, t, i) {
		var r = {};
		return OSF.OUtil.defineEnumerableProperties(r, {
			data : {
				value : n[Microsoft.Office.WebExtension.Parameters.Data]
			},
			index : {
				value : i[Microsoft.Office.WebExtension.Parameters.SliceIndex]
			},
			size : {
				value : n[OSF.DDA.FileProperties.SliceSize]
			}
		}),
		r
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
	privateStateCallbacks : [{
			name : OSF.DDA.FileProperties.Handle,
			value : function (n, t) {
				return t[OSF.DDA.FileProperties.Handle]
			}
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FileType, {
	Text : "text",
	Pdf : "pdf"
});
OSF.DDA.FilePropertiesDescriptor = {
	Url : "Url"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FilePropertiesDescriptor : "FilePropertiesDescriptor"
});
Microsoft.Office.WebExtension.FileProperties = function (n) {
	OSF.OUtil.defineEnumerableProperties(this, {
		url : {
			value : n[OSF.DDA.FilePropertiesDescriptor.Url]
		}
	})
};
OSF.DDA.AsyncMethodNames.addNames({
	GetFilePropertiesAsync : "getFilePropertiesAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
			value : 0
		}
	],
	requiredArguments : [],
	supportedOptions : [],
	onSucceeded : function (n) {
		return new Microsoft.Office.WebExtension.FileProperties(n)
	}
});
Microsoft.Office.WebExtension.GoToType = {
	Binding : "binding",
	NamedItem : "namedItem",
	Slide : "slide",
	Index : "index"
};
Microsoft.Office.WebExtension.SelectionMode = {
	Default : "default",
	Selected : "selected",
	None : "none"
};
Microsoft.Office.WebExtension.Index = {
	First : "first",
	Last : "last",
	Next : "next",
	Previous : "previous"
};
OSF.DDA.AsyncMethodNames.addNames({
	GoToByIdAsync : "goToByIdAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.GoToByIdAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			types : ["string", "number"]
		}, {
			name : Microsoft.Office.WebExtension.Parameters.GoToType,
			"enum" : Microsoft.Office.WebExtension.GoToType
		}
	],
	supportedOptions : [{
			name : Microsoft.Office.WebExtension.Parameters.SelectionMode,
			value : {
				"enum" : Microsoft.Office.WebExtension.SelectionMode,
				defaultValue : Microsoft.Office.WebExtension.SelectionMode.Default
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.GoToType,
	toHost : [{
			name : Microsoft.Office.WebExtension.GoToType.Binding,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.GoToType.NamedItem,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.GoToType.Slide,
			value : 2
		}, {
			name : Microsoft.Office.WebExtension.GoToType.Index,
			value : 3
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.SelectionMode,
	toHost : [{
			name : Microsoft.Office.WebExtension.SelectionMode.Default,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.SelectionMode.Selected,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.SelectionMode.None,
			value : 2
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidNavigateToMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.GoToType,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.SelectionMode,
			value : 2
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.FileProperties,
	fromHost : [{
			name : OSF.DDA.FileProperties.Handle,
			value : 0
		}, {
			name : OSF.DDA.FileProperties.FileSize,
			value : 1
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.FileSliceProperties,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 0
		}, {
			name : OSF.DDA.FileProperties.SliceSize,
			value : 1
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.FileType,
	toHost : [{
			name : Microsoft.Office.WebExtension.FileType.Text,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.FileType.Compressed,
			value : 5
		}, {
			name : Microsoft.Office.WebExtension.FileType.Pdf,
			value : 6
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.FileType,
			value : 0
		}
	],
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.FileProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,
	toHost : [{
			name : OSF.DDA.FileProperties.Handle,
			value : 0
		}, {
			name : OSF.DDA.FileSliceOffset,
			value : 1
		}, {
			name : OSF.DDA.FileProperties.SliceSize,
			value : 2
		}
	],
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.FileSliceProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,
	toHost : [{
			name : OSF.DDA.FileProperties.Handle,
			value : 0
		}
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	GetSelectedDataAsync : "getSelectedDataAsync",
	SetSelectedDataAsync : "setSelectedDataAsync"
}), function () {
	function r(n, t, i) {
		var r = n[Microsoft.Office.WebExtension.Parameters.Data];
		return OSF.DDA.TableDataProperties && r && (r[OSF.DDA.TableDataProperties.TableRows] != undefined || r[OSF.DDA.TableDataProperties.TableHeaders] != undefined) && (r = OSF.DDA.OMFactory.manufactureTableData(r)),
		r = OSF.DDA.DataCoercion.coerceData(r, i[Microsoft.Office.WebExtension.Parameters.CoercionType]),
		r == undefined ? null : r
	}
	var i = !1,
	n = "boolean",
	t = "number";
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.CoercionType,
				"enum" : Microsoft.Office.WebExtension.CoercionType
			}
		],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value : {
					"enum" : Microsoft.Office.WebExtension.ValueFormat,
					defaultValue : Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.FilterType,
				value : {
					"enum" : Microsoft.Office.WebExtension.FilterType,
					defaultValue : Microsoft.Office.WebExtension.FilterType.All
				}
			}
		],
		privateStateCallbacks : [],
		onSucceeded : r
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Data,
				types : ["string", "object", t, n]
			}
		],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.CoercionType,
				value : {
					"enum" : Microsoft.Office.WebExtension.CoercionType,
					calculate : function (n) {
						return OSF.DDA.DataCoercion.determineCoercionType(n[Microsoft.Office.WebExtension.Parameters.Data])
					}
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ImageLeft,
				value : {
					types : [t, n],
					defaultValue : i
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ImageTop,
				value : {
					types : [t, n],
					defaultValue : i
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value : {
					types : [t, n],
					defaultValue : i
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value : {
					types : [t, n],
					defaultValue : i
				}
			}
		],
		privateStateCallbacks : []
	})
}
();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.CoercionType,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ValueFormat,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.FilterType,
			value : 2
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.CoercionType,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ImageLeft,
			value : 2
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ImageTop,
			value : 3
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ImageWidth,
			value : 4
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ImageHeight,
			value : 5
		}
	]
});
OSF.DDA.SettingsManager = {
	SerializedSettings : "serializedSettings",
	DateJSONPrefix : "Date(",
	DataJSONSuffix : ")",
	serializeSettings : function (n) {
		var r = {},
		i,
		t;
		for (i in n) {
			t = n[i];
			try {
				t = JSON ? JSON.stringify(t, function (n, t) {
						return OSF.OUtil.isDate(this[n]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[n].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : t
					}) : Sys.Serialization.JavaScriptSerializer.serialize(t);
				r[i] = t
			} catch (u) {}

		}
		return r
	},
	deserializeSettings : function (n) {
		var r = {},
		i,
		t;
		n = n || {};
		for (i in n) {
			t = n[i];
			try {
				t = JSON ? JSON.parse(t, function (n, t) {
						var i;
						return typeof t == "string" && t && t.length > 6 && t.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && t.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix && (i = new Date(parseInt(t.slice(5, -1))), i) ? i : t
					}) : Sys.Serialization.JavaScriptSerializer.deserialize(t, !0);
				r[i] = t
			} catch (u) {}

		}
		return r
	}
};
OSF.DDA.Settings = function (n) {
	var t = "name",
	i;
	n = n || {};
	i = function (n) {
		var i = OSF.OUtil.getSessionStorage(),
		t,
		r;
		i && (t = OSF.DDA.SettingsManager.serializeSettings(n), r = JSON ? JSON.stringify(t) : Sys.Serialization.JavaScriptSerializer.serialize(t), i.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), r))
	};
	OSF.OUtil.defineEnumerableProperties(this, {
		get : {
			value : function (i) {
				var u = Function._validateParams(arguments, [{
								name : t,
								type : String,
								mayBeNull : !1
							}
						]),
				r;
				if (u)
					throw u;
				return r = n[i],
				typeof r == "undefined" ? null : r
			}
		},
		set : {
			value : function (r, u) {
				var f = Function._validateParams(arguments, [{
								name : t,
								type : String,
								mayBeNull : !1
							}, {
								name : "value",
								mayBeNull : !0
							}
						]);
				if (f)
					throw f;
				n[r] = u;
				i(n)
			}
		},
		remove : {
			value : function (r) {
				var u = Function._validateParams(arguments, [{
								name : t,
								type : String,
								mayBeNull : !1
							}
						]);
				if (u)
					throw u;
				delete n[r];
				i(n)
			}
		}
	});
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], n)
};
OSF.DDA.RefreshableSettings = function (n) {
	OSF.DDA.RefreshableSettings.uber.constructor.call(this, n);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], n);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]))
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function (n) {
	var t,
	i;
	this._eventHandlers = {};
	for (t in n)
		i = n[t], this._eventHandlers[i] = []
};
OSF.EventDispatch.prototype = {
	getSupportedEvents : function () {
		var n = [];
		for (var t in this._eventHandlers)
			n.push(t);
		return n
	},
	supportsEvent : function (n) {
		var t = !1;
		for (var i in this._eventHandlers)
			if (n == i) {
				t = !0;
				break
			}
		return t
	},
	hasEventHandler : function (n, t) {
		var i = this._eventHandlers[n],
		r;
		if (i && i.length > 0)
			for (r in i)
				if (i[r] === t)
					return !0;
		return !1
	},
	addEventHandler : function (n, t) {
		if (typeof t != "function")
			return !1;
		var i = this._eventHandlers[n];
		return i && !this.hasEventHandler(n, t) ? (i.push(t), !0) : !1
	},
	removeEventHandler : function (n, t) {
		var i = this._eventHandlers[n],
		r;
		if (i && i.length > 0)
			for (r = 0; r < i.length; r++)
				if (i[r] === t)
					return i.splice(r, 1), !0;
		return !1
	},
	clearEventHandlers : function (n) {
		return typeof this._eventHandlers[n] != "undefined" && this._eventHandlers[n].length > 0 ? (this._eventHandlers[n] = [], !0) : !1
	},
	getEventHandlerCount : function (n) {
		return this._eventHandlers[n] != undefined ? this._eventHandlers[n].length : -1
	},
	fireEvent : function (n) {
		var t,
		i,
		r;
		if (n.type == undefined)
			return !1;
		if (t = n.type, t && this._eventHandlers[t]) {
			i = this._eventHandlers[t];
			for (r in i)
				i[r](n);
			return !0
		} else
			return !1
	}
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function (n, t, i) {
	var u = this,
	r;
	switch (n) {
	case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
		r = new OSF.DDA.DocumentSelectionChangedEventArgs(t);
		break;
	case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
		r = new OSF.DDA.BindingSelectionChangedEventArgs(u.manufactureBinding(i, t.document), i[OSF.DDA.PropertyDescriptors.Subset]);
		break;
	case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
		r = new OSF.DDA.BindingDataChangedEventArgs(u.manufactureBinding(i, t.document));
		break;
	case Microsoft.Office.WebExtension.EventType.SettingsChanged:
		r = new OSF.DDA.SettingsChangedEventArgs(t);
		break;
	case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
		r = new OSF.DDA.ActiveViewChangedEventArgs(i);
		break;
	case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
		r = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(i);
		break;
	case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
		r = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(i);
		break;
	case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
		r = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(i);
		break;
	case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
		r = new OSF.DDA.NodeInsertedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NewNode]), i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
		break;
	case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
		r = new OSF.DDA.NodeReplacedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.OldNode]), u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NewNode]), i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
		break;
	case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
		r = new OSF.DDA.NodeDeletedEventArgs(u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.OldNode]), u.manufactureDataNode(i[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), i[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
		break;
	case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
		r = new OSF.DDA.TaskSelectionChangedEventArgs(t);
		break;
	case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
		r = new OSF.DDA.ResourceSelectionChangedEventArgs(t);
		break;
	case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
		r = new OSF.DDA.ViewSelectionChangedEventArgs(t);
		break;
	default:
		throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, n));
	}
	return r
};
OSF.DDA.AsyncMethodNames.addNames({
	AddHandlerAsync : "addHandlerAsync",
	RemoveHandlerAsync : "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.AddHandlerAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.EventType,
			"enum" : Microsoft.Office.WebExtension.EventType,
			verify : function (n, t, i) {
				return i.supportsEvent(n)
			}
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Handler,
			types : ["function"]
		}
	],
	supportedOptions : [],
	privateStateCallbacks : []
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.EventType,
			"enum" : Microsoft.Office.WebExtension.EventType,
			verify : function (n, t, i) {
				return i.supportsEvent(n)
			}
		}
	],
	supportedOptions : [{
			name : Microsoft.Office.WebExtension.Parameters.Handler,
			value : {
				types : ["function", "object"],
				defaultValue : null
			}
		}
	],
	privateStateCallbacks : []
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	SettingsChanged : "settingsChanged"
});
OSF.DDA.SettingsChangedEventArgs = function (n) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.SettingsChanged
		},
		settings : {
			value : n
		}
	})
};
OSF.DDA.AsyncMethodNames.addNames({
	RefreshAsync : "refreshAsync",
	SaveAsync : "saveAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.RefreshAsync,
	requiredArguments : [],
	supportedOptions : [],
	privateStateCallbacks : [],
	onSucceeded : function (n) {
		var t = n[OSF.DDA.SettingsManager.SerializedSettings];
		return OSF.DDA.SettingsManager.deserializeSettings(t)
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.SaveAsync,
	requiredArguments : [],
	supportedOptions : [{
			name : Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
			value : {
				types : ["boolean"],
				defaultValue : !0
			}
		}
	],
	privateStateCallbacks : [{
			name : OSF.DDA.SettingsManager.SerializedSettings,
			value : function (n, t) {
				return OSF.DDA.SettingsManager.serializeSettings(t)
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidLoadSettingsMethod,
	fromHost : [{
			name : OSF.DDA.SettingsManager.SerializedSettings,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSaveSettingsMethod,
	toHost : [{
			name : OSF.DDA.SettingsManager.SerializedSettings,
			value : OSF.DDA.SettingsManager.SerializedSettings
		}, {
			name : Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
			value : Microsoft.Office.WebExtension.Parameters.OverwriteIfStale
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidSettingsChangedEvent
});
Microsoft.Office.WebExtension.BindingType = {
	Table : "table",
	Text : "text",
	Matrix : "matrix"
};
OSF.DDA.BindingProperties = {
	Id : "BindingId",
	Type : Microsoft.Office.WebExtension.Parameters.BindingType
};
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, {
	BindingList : "BindingList"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	Subset : "subset",
	BindingProperties : "BindingProperties"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList, OSF.DDA.PropertyDescriptors.BindingProperties);
OSF.DDA.BindingPromise = function (n, t) {
	this._id = n;
	OSF.OUtil.defineEnumerableProperty(this, "onFail", {
		get : function () {
			return t
		},
		set : function (n) {
			var i = typeof n;
			if (i != "undefined" && i != "function")
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, i);
			t = n
		}
	})
};
OSF.DDA.BindingPromise.prototype = {
	_fetch : function (n) {
		var t = this,
		i;
		return t.binding ? n && n(t.binding) : t._binding || (i = t, Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(t._id, function (t) {
				t.status == Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded ? (OSF.OUtil.defineEnumerableProperty(i, "binding", {
						value : t.value
					}), n && n(i.binding)) : i.onFail && i.onFail(t)
			})),
		t
	},
	getDataAsync : function () {
		var n = arguments;
		return this._fetch(function (t) {
			t.getDataAsync.apply(t, n)
		}),
		this
	},
	setDataAsync : function () {
		var n = arguments;
		return this._fetch(function (t) {
			t.setDataAsync.apply(t, n)
		}),
		this
	},
	addHandlerAsync : function () {
		var n = arguments;
		return this._fetch(function (t) {
			t.addHandlerAsync.apply(t, n)
		}),
		this
	},
	removeHandlerAsync : function () {
		var n = arguments;
		return this._fetch(function (t) {
			t.removeHandlerAsync.apply(t, n)
		}),
		this
	}
};
OSF.DDA.BindingFacade = function (n) {
	this._eventDispatches = [];
	OSF.OUtil.defineEnumerableProperty(this, "document", {
		value : n
	});
	var t = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [t.AddFromSelectionAsync, t.AddFromNamedItemAsync, t.GetAllAsync, t.GetByIdAsync, t.ReleaseByIdAsync])
};
OSF.DDA.UnknownBinding = function (n, t) {
	OSF.OUtil.defineEnumerableProperties(this, {
		document : {
			value : t
		},
		id : {
			value : n
		}
	})
};
OSF.DDA.Binding = function (n, t) {
	var r,
	u,
	i,
	f;
	OSF.OUtil.defineEnumerableProperties(this, {
		document : {
			value : t
		},
		id : {
			value : n
		}
	});
	r = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [r.GetDataAsync, r.SetDataAsync]);
	u = Microsoft.Office.WebExtension.EventType;
	i = t.bindings._eventDispatches;
	i[n] || (i[n] = new OSF.EventDispatch([u.BindingSelectionChanged, u.BindingDataChanged]));
	f = i[n];
	OSF.DDA.DispIdHost.addEventSupport(this, f)
};
OSF.DDA.generateBindingId = function () {
	return "UnnamedBinding_" + OSF.OUtil.getUniqueId() + "_" + (new Date).getTime()
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureBinding = function (n, t) {
	var r = n[OSF.DDA.BindingProperties.Id],
	u = n[OSF.DDA.BindingProperties.RowCount],
	f = n[OSF.DDA.BindingProperties.ColumnCount],
	s = n[OSF.DDA.BindingProperties.HasHeaders],
	i,
	e,
	o;
	switch (n[OSF.DDA.BindingProperties.Type]) {
	case Microsoft.Office.WebExtension.BindingType.Text:
		i = new OSF.DDA.TextBinding(r, t);
		break;
	case Microsoft.Office.WebExtension.BindingType.Matrix:
		i = new OSF.DDA.MatrixBinding(r, t, u, f);
		break;
	case Microsoft.Office.WebExtension.BindingType.Table:
		e = function () {
			return OSF.DDA.ExcelDocument && Microsoft.Office.WebExtension.context.document && Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument
		};
		o = e() && OSF.DDA.ExcelTableBinding ? OSF.DDA.ExcelTableBinding : OSF.DDA.TableBinding;
		i = new o(r, t, u, f, s);
		break;
	default:
		i = new OSF.DDA.UnknownBinding(r, t)
	}
	return i
};
OSF.DDA.AsyncMethodNames.addNames({
	AddFromSelectionAsync : "addFromSelectionAsync",
	AddFromNamedItemAsync : "addFromNamedItemAsync",
	GetAllAsync : "getAllAsync",
	GetByIdAsync : "getByIdAsync",
	ReleaseByIdAsync : "releaseByIdAsync",
	GetDataAsync : "getDataAsync",
	SetDataAsync : "setDataAsync"
}), function () {
	function u(n) {
		return OSF.DDA.OMFactory.manufactureBinding(n, Microsoft.Office.WebExtension.context.document)
	}
	function f(n) {
		return n.id
	}
	function e(n, t, i) {
		var u = n[Microsoft.Office.WebExtension.Parameters.Data];
		return OSF.DDA.TableDataProperties && u && (u[OSF.DDA.TableDataProperties.TableRows] != undefined || u[OSF.DDA.TableDataProperties.TableHeaders] != undefined) && (u = OSF.DDA.OMFactory.manufactureTableData(u)),
		u = OSF.DDA.DataCoercion.coerceData(u, i[Microsoft.Office.WebExtension.Parameters.CoercionType]),
		u == undefined ? r : u
	}
	var t = "number",
	i = "object",
	n = "string",
	r = null;
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum" : Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : {
					types : [n],
					calculate : OSF.DDA.generateBindingId
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Columns,
				value : {
					types : [i],
					defaultValue : r
				}
			}
		],
		privateStateCallbacks : [],
		onSucceeded : u
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.ItemName,
				types : [n]
			}, {
				name : Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum" : Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : {
					types : [n],
					calculate : OSF.DDA.generateBindingId
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Columns,
				value : {
					types : [i],
					defaultValue : r
				}
			}
		],
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.FailOnCollision,
				value : function () {
					return !0
				}
			}
		],
		onSucceeded : u
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetAllAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : function (n) {
			return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.BindingList], u)
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetByIdAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : u
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : function (n, t, i) {
			var r = i[Microsoft.Office.WebExtension.Parameters.Id];
			delete t._eventDispatches[r]
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetDataAsync,
		requiredArguments : [],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.CoercionType,
				value : {
					"enum" : Microsoft.Office.WebExtension.CoercionType,
					calculate : function (n, t) {
						return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type)
					}
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value : {
					"enum" : Microsoft.Office.WebExtension.ValueFormat,
					defaultValue : Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.FilterType,
				value : {
					"enum" : Microsoft.Office.WebExtension.FilterType,
					defaultValue : Microsoft.Office.WebExtension.FilterType.All
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Rows,
				value : {
					types : [i, n],
					defaultValue : r
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Columns,
				value : {
					types : [i],
					defaultValue : r
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.StartRow,
				value : {
					types : [t],
					defaultValue : 0
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.StartColumn,
				value : {
					types : [t],
					defaultValue : 0
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.RowCount,
				value : {
					types : [t],
					defaultValue : 0
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.ColumnCount,
				value : {
					types : [t],
					defaultValue : 0
				}
			}
		],
		checkCallArgs : function (n, t) {
			if (n[Microsoft.Office.WebExtension.Parameters.StartRow] == 0 && n[Microsoft.Office.WebExtension.Parameters.StartColumn] == 0 && n[Microsoft.Office.WebExtension.Parameters.RowCount] == 0 && n[Microsoft.Office.WebExtension.Parameters.ColumnCount] == 0 && (delete n[Microsoft.Office.WebExtension.Parameters.StartRow], delete n[Microsoft.Office.WebExtension.Parameters.StartColumn], delete n[Microsoft.Office.WebExtension.Parameters.RowCount], delete n[Microsoft.Office.WebExtension.Parameters.ColumnCount]), n[Microsoft.Office.WebExtension.Parameters.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type) && (n[Microsoft.Office.WebExtension.Parameters.StartRow] || n[Microsoft.Office.WebExtension.Parameters.StartColumn] || n[Microsoft.Office.WebExtension.Parameters.RowCount] || n[Microsoft.Office.WebExtension.Parameters.ColumnCount]))
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			return n
		},
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : f
			}
		],
		onSucceeded : e
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Data,
				types : [n, i, t, "boolean"]
			}
		],
		supportedOptions : [{
				name : Microsoft.Office.WebExtension.Parameters.CoercionType,
				value : {
					"enum" : Microsoft.Office.WebExtension.CoercionType,
					calculate : function (n) {
						return OSF.DDA.DataCoercion.determineCoercionType(n[Microsoft.Office.WebExtension.Parameters.Data])
					}
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Rows,
				value : {
					types : [i, n],
					defaultValue : r
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Columns,
				value : {
					types : [i],
					defaultValue : r
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.StartRow,
				value : {
					types : [t],
					defaultValue : 0
				}
			}, {
				name : Microsoft.Office.WebExtension.Parameters.StartColumn,
				value : {
					types : [t],
					defaultValue : 0
				}
			}
		],
		checkCallArgs : function (n, t) {
			if (n[Microsoft.Office.WebExtension.Parameters.StartRow] == 0 && n[Microsoft.Office.WebExtension.Parameters.StartColumn] == 0 && (delete n[Microsoft.Office.WebExtension.Parameters.StartRow], delete n[Microsoft.Office.WebExtension.Parameters.StartColumn]), n[Microsoft.Office.WebExtension.Parameters.CoercionType] != OSF.DDA.DataCoercion.getCoercionDefaultForBinding(t.type) && (n[Microsoft.Office.WebExtension.Parameters.StartRow] || n[Microsoft.Office.WebExtension.Parameters.StartColumn]))
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			return n
		},
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : f
			}
		]
	})
}
();
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount : "BindingRowCount",
	ColumnCount : "BindingColumnCount",
	HasHeaders : "HasHeaders"
});
OSF.DDA.MatrixBinding = function (n, t, i, r) {
	OSF.DDA.MatrixBinding.uber.constructor.call(this, n, t);
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.BindingType.Matrix
		},
		rowCount : {
			value : i ? i : 0
		},
		columnCount : {
			value : r ? r : 0
		}
	})
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.BindingProperties,
	fromHost : [{
			name : OSF.DDA.BindingProperties.Id,
			value : 0
		}, {
			name : OSF.DDA.BindingProperties.Type,
			value : 1
		}, {
			name : OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
			value : 2
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.BindingType,
	toHost : [{
			name : Microsoft.Office.WebExtension.BindingType.Text,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.BindingType.Matrix,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.BindingType.Table,
			value : 2
		}
	],
	invertible : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.BindingProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.BindingType,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.BindingProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.ItemName,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.BindingType,
			value : 2
		}, {
			name : Microsoft.Office.WebExtension.Parameters.FailOnCollision,
			value : 3
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidReleaseBindingMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetBindingMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.BindingProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,
	fromHost : [{
			name : OSF.DDA.ListDescriptors.BindingList,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetBindingDataMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.CoercionType,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ValueFormat,
			value : 2
		}, {
			name : Microsoft.Office.WebExtension.Parameters.FilterType,
			value : 3
		}, {
			name : OSF.DDA.PropertyDescriptors.Subset,
			value : 4
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.CoercionType,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 2
		}, {
			name : OSF.DDA.SafeArray.UniqueArguments.Offset,
			value : 3
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
	fromHost : [{
			name : OSF.DDA.BindingProperties.RowCount,
			value : 0
		}, {
			name : OSF.DDA.BindingProperties.ColumnCount,
			value : 1
		}, {
			name : OSF.DDA.BindingProperties.HasHeaders,
			value : 2
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.Subset,
	toHost : [{
			name : OSF.DDA.SafeArray.UniqueArguments.Offset,
			value : 0
		}, {
			name : OSF.DDA.SafeArray.UniqueArguments.Run,
			value : 1
		}
	],
	canonical : !0,
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.SafeArray.UniqueArguments.Offset,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.StartRow,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.StartColumn,
			value : 1
		}
	],
	canonical : !0,
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.SafeArray.UniqueArguments.Run,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.RowCount,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.ColumnCount,
			value : 1
		}
	],
	canonical : !0,
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddRowsMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddColumnsMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidClearAllRowsMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
Microsoft.Office.WebExtension.TableData = function (n, t) {
	function i(n) {
		if (n == null || n == undefined)
			return null;
		try {
			for (var t = OSF.DDA.DataCoercion.findArrayDimensionality(n, 2); t < 2; t++)
				n = [n];
			return n
		} catch (i) {}

	}
	OSF.OUtil.defineEnumerableProperties(this, {
		headers : {
			get : function () {
				return t
			},
			set : function (n) {
				t = i(n)
			}
		},
		rows : {
			get : function () {
				return n
			},
			set : function (t) {
				n = t == null || OSF.OUtil.isArray(t) && t.length == 0 ? [] : i(t)
			}
		}
	});
	this.headers = t;
	this.rows = n
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureTableData = function (n) {
	return new Microsoft.Office.WebExtension.TableData(n[OSF.DDA.TableDataProperties.TableRows], n[OSF.DDA.TableDataProperties.TableHeaders])
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	TableDataProperties : "TableDataProperties"
});
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount : "BindingRowCount",
	ColumnCount : "BindingColumnCount",
	HasHeaders : "HasHeaders"
});
OSF.DDA.TableDataProperties = {
	TableRows : "TableRows",
	TableHeaders : "TableHeaders"
};
OSF.DDA.TableBinding = function (n, t, i, r, u) {
	OSF.DDA.TableBinding.uber.constructor.call(this, n, t);
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.BindingType.Table
		},
		rowCount : {
			value : i ? i : 0
		},
		columnCount : {
			value : r ? r : 0
		},
		hasHeaders : {
			value : u ? u : !1
		}
	});
	var f = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [f.AddRowsAsync, f.AddColumnsAsync, f.DeleteAllDataValuesAsync])
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({
	AddRowsAsync : "addRowsAsync",
	AddColumnsAsync : "addColumnsAsync",
	DeleteAllDataValuesAsync : "deleteAllDataValuesAsync"
}), function () {
	function n(n) {
		return n.id
	}
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddRowsAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Data,
				types : ["object"]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : n
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddColumnsAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Data,
				types : ["object"]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : n
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				value : n
			}
		]
	})
}
();
OSF.DDA.TextBinding = function (n, t) {
	OSF.DDA.TextBinding.uber.constructor.call(this, n, t);
	OSF.OUtil.defineEnumerableProperty(this, "type", {
		value : Microsoft.Office.WebExtension.BindingType.Text
	})
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DocumentSelectionChanged : "documentSelectionChanged"
});
OSF.DDA.DocumentSelectionChangedEventArgs = function (n) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged
		},
		document : {
			value : n
		}
	})
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	BindingSelectionChanged : "bindingSelectionChanged",
	BindingDataChanged : "bindingDataChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	BindingSelectionChangedEvent : "BindingSelectionChangedEvent"
});
OSF.DDA.BindingSelectionChangedEventArgs = function (n, t) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.BindingSelectionChanged
		},
		binding : {
			value : n
		}
	});
	for (var i in t)
		OSF.OUtil.defineEnumerableProperty(this, i, {
			value : t[i]
		})
};
OSF.DDA.BindingDataChangedEventArgs = function (n) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.BindingDataChanged
		},
		binding : {
			value : n
		}
	})
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.BindingProperties,
			value : 0
		}, {
			name : OSF.DDA.PropertyDescriptors.Subset,
			value : 1
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent,
	fromHost : [{
			name : OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidBindingDataChangedEvent,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.BindingProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
Microsoft.Office.WebExtension.CoercionType = {
	Text : "text",
	Matrix : "matrix",
	Table : "table"
};
OSF.DDA.DataCoercion = function () {
	var n = null;
	return {
		findArrayDimensionality : function (n) {
			if (OSF.OUtil.isArray(n)) {
				for (var t = 0, i = 0; i < n.length; i++)
					t = Math.max(t, OSF.DDA.DataCoercion.findArrayDimensionality(n[i]));
				return t + 1
			} else
				return 0
		},
		getCoercionDefaultForBinding : function (n) {
			switch (n) {
			case Microsoft.Office.WebExtension.BindingType.Matrix:
				return Microsoft.Office.WebExtension.CoercionType.Matrix;
			case Microsoft.Office.WebExtension.BindingType.Table:
				return Microsoft.Office.WebExtension.CoercionType.Table;
			case Microsoft.Office.WebExtension.BindingType.Text:
			default:
				return Microsoft.Office.WebExtension.CoercionType.Text
			}
		},
		getBindingDefaultForCoercion : function (n) {
			switch (n) {
			case Microsoft.Office.WebExtension.CoercionType.Matrix:
				return Microsoft.Office.WebExtension.BindingType.Matrix;
			case Microsoft.Office.WebExtension.CoercionType.Table:
				return Microsoft.Office.WebExtension.BindingType.Table;
			case Microsoft.Office.WebExtension.CoercionType.Text:
			case Microsoft.Office.WebExtension.CoercionType.Html:
			case Microsoft.Office.WebExtension.CoercionType.Ooxml:
			default:
				return Microsoft.Office.WebExtension.BindingType.Text
			}
		},
		determineCoercionType : function (t) {
			if (t == n || t == undefined)
				return n;
			var i = n,
			r = typeof t;
			if (t.rows !== undefined)
				i = Microsoft.Office.WebExtension.CoercionType.Table;
			else if (OSF.OUtil.isArray(t))
				i = Microsoft.Office.WebExtension.CoercionType.Matrix;
			else if (r == "string" || r == "number" || r == "boolean" || OSF.OUtil.isDate(t))
				i = Microsoft.Office.WebExtension.CoercionType.Text;
			else
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;
			return i
		},
		coerceData : function (n, t, i) {
			return i = i || OSF.DDA.DataCoercion.determineCoercionType(n),
			i && i != t && (OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin), n = OSF.DDA.DataCoercion._coerceDataFromTable(t, OSF.DDA.DataCoercion._coerceDataToTable(n, i)), OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd)),
			n
		},
		_matrixToText : function (n) {
			if (n.length == 1 && n[0].length == 1)
				return "" + n[0][0];
			for (var t = "", i = 0; i < n.length; i++)
				t += n[i].join("\t") + "\n";
			return t.substring(0, t.length - 1)
		},
		_textToMatrix : function (n) {
			for (var t = n.split("\n"), i = 0; i < t.length; i++)
				t[i] = t[i].split("\t");
			return t
		},
		_tableToText : function (t) {
			var i = "",
			r;
			return t.headers != n && (i = OSF.DDA.DataCoercion._matrixToText([t.headers]) + "\n"),
			r = OSF.DDA.DataCoercion._matrixToText(t.rows),
			r == "" && (i = i.substring(0, i.length - 1)),
			i + r
		},
		_tableToMatrix : function (t) {
			var i = t.rows;
			return t.headers != n && i.unshift(t.headers),
			i
		},
		_coerceDataFromTable : function (t, i) {
			var r;
			switch (t) {
			case Microsoft.Office.WebExtension.CoercionType.Table:
				r = i;
				break;
			case Microsoft.Office.WebExtension.CoercionType.Matrix:
				r = OSF.DDA.DataCoercion._tableToMatrix(i);
				break;
			case Microsoft.Office.WebExtension.CoercionType.SlideRange:
				r = n;
				OSF.DDA.OMFactory.manufactureSlideRange && (r = OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(i)));
				r == n && (r = OSF.DDA.DataCoercion._tableToText(i));
				break;
			case Microsoft.Office.WebExtension.CoercionType.Text:
			case Microsoft.Office.WebExtension.CoercionType.Html:
			case Microsoft.Office.WebExtension.CoercionType.Ooxml:
			default:
				r = OSF.DDA.DataCoercion._tableToText(i)
			}
			return r
		},
		_coerceDataToTable : function (n, t) {
			t == undefined && (t = OSF.DDA.DataCoercion.determineCoercionType(n));
			var i;
			switch (t) {
			case Microsoft.Office.WebExtension.CoercionType.Table:
				i = n;
				break;
			case Microsoft.Office.WebExtension.CoercionType.Matrix:
				i = new Microsoft.Office.WebExtension.TableData(n);
				break;
			case Microsoft.Office.WebExtension.CoercionType.Text:
			case Microsoft.Office.WebExtension.CoercionType.Html:
			case Microsoft.Office.WebExtension.CoercionType.Ooxml:
			default:
				i = new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(n))
			}
			return i
		}
	}
}
();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.Text,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.CoercionType.Matrix,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.CoercionType.Table,
			value : 2
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, {
	Html : "html"
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.Html,
			value : 3
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, {
	Ooxml : "ooxml"
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.Ooxml,
			value : 4
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, {
	OoxmlPackage : "ooxmlPackage"
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.OoxmlPackage,
			value : 5
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, {
	PdfFile : "pdf"
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.PdfFile,
			value : 6
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, {
	SlideRange : "slideRange"
});
OSF.DDA.SlideProperties = {
	Id : 0,
	Title : 1,
	Index : 2
};
OSF.DDA.Slide = function (n, t, i) {
	OSF.OUtil.defineEnumerableProperties(this, {
		id : {
			value : n
		},
		title : {
			value : t
		},
		index : {
			value : i
		}
	})
};
OSF.DDA.SlideRange = function (n) {
	OSF.OUtil.defineEnumerableProperties(this, {
		slides : {
			value : n
		}
	})
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureSlideRange = function (n) {
	var r = null,
	t = r,
	f,
	e;
	if (t = JSON ? JSON.parse(n) : Sys.Serialization.JavaScriptSerializer.deserialize(n), t == r)
		return r;
	f = 0;
	for (e in OSF.DDA.SlideProperties)
		OSF.DDA.SlideProperties.hasOwnProperty(e) && f++;
	for (var o = [], u = !0, i = 0; i < t.length && u; i++)
		if (u = !1, t[i].length == f) {
			var s = parseInt(t[i][OSF.DDA.SlideProperties.Id]),
			c = t[i][OSF.DDA.SlideProperties.Title],
			h = parseInt(t[i][OSF.DDA.SlideProperties.Index]);
			isNaN(s) || isNaN(h) || (u = !0, o.push(new OSF.DDA.Slide(s, c, h)))
		}
	return u ? new OSF.DDA.SlideRange(o) : r
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost : [{
			name : Microsoft.Office.WebExtension.CoercionType.SlideRange,
			value : 7
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FilterType, {
	OnlyVisible : "onlyVisible"
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : Microsoft.Office.WebExtension.Parameters.FilterType,
	toHost : [{
			name : Microsoft.Office.WebExtension.FilterType.OnlyVisible,
			value : 1
		}
	]
});
OSF.DDA.DataPartProperties = {
	Id : Microsoft.Office.WebExtension.Parameters.Id,
	BuiltIn : "DataPartBuiltIn"
};
OSF.DDA.DataNodeProperties = {
	Handle : "DataNodeHandle",
	BaseName : "DataNodeBaseName",
	NamespaceUri : "DataNodeNamespaceUri",
	NodeType : "DataNodeType"
};
OSF.DDA.DataNodeEventProperties = {
	OldNode : "OldNode",
	NewNode : "NewNode",
	NextSiblingNode : "NextSiblingNode",
	InUndoRedo : "InUndoRedo"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	DataPartProperties : "DataPartProperties",
	DataNodeProperties : "DataNodeProperties"
});
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, {
	DataPartList : "DataPartList",
	DataNodeList : "DataNodeList"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataPartList, OSF.DDA.PropertyDescriptors.DataPartProperties);
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataNodeList, OSF.DDA.PropertyDescriptors.DataNodeProperties);
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	DataNodeInsertedEvent : "DataNodeInsertedEvent",
	DataNodeReplacedEvent : "DataNodeReplacedEvent",
	DataNodeDeletedEvent : "DataNodeDeletedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DataNodeDeleted : "nodeDeleted",
	DataNodeInserted : "nodeInserted",
	DataNodeReplaced : "nodeReplaced"
});
OSF.DDA.CustomXmlParts = function () {
	this._eventDispatches = [];
	var n = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [n.AddDataPartAsync, n.GetDataPartByIdAsync, n.GetDataPartsByNameSpaceAsync])
};
OSF.DDA.CustomXmlPart = function (n, t, i) {
	var u,
	e,
	r,
	f;
	OSF.OUtil.defineEnumerableProperties(this, {
		builtIn : {
			value : i
		},
		id : {
			value : t
		},
		namespaceManager : {
			value : new OSF.DDA.CustomXmlPrefixMappings(t)
		}
	});
	u = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [u.DeleteDataPartAsync, u.GetPartNodesAsync, u.GetPartXmlAsync]);
	e = n._eventDispatches;
	r = e[t];
	r || (f = Microsoft.Office.WebExtension.EventType, r = new OSF.EventDispatch([f.DataNodeDeleted, f.DataNodeInserted, f.DataNodeReplaced]), e[t] = r);
	OSF.DDA.DispIdHost.addEventSupport(this, r)
};
OSF.DDA.CustomXmlPrefixMappings = function (n) {
	var t = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [t.AddDataPartNamespaceAsync, t.GetDataPartNamespaceAsync, t.GetDataPartPrefixAsync], n)
};
OSF.DDA.CustomXmlNode = function (n, t, i, r) {
	OSF.OUtil.defineEnumerableProperties(this, {
		baseName : {
			value : r
		},
		namespaceUri : {
			value : i
		},
		nodeType : {
			value : t
		}
	});
	var u = OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [u.GetRelativeNodesAsync, u.GetNodeValueAsync, u.GetNodeXmlAsync, u.SetNodeValueAsync, u.SetNodeXmlAsync, u.GetNodeTextAsync, u.SetNodeTextAsync], n)
};
OSF.DDA.NodeInsertedEventArgs = function (n, t) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.DataNodeInserted
		},
		newNode : {
			value : n
		},
		inUndoRedo : {
			value : t
		}
	})
};
OSF.DDA.NodeReplacedEventArgs = function (n, t, i) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.DataNodeReplaced
		},
		oldNode : {
			value : n
		},
		newNode : {
			value : t
		},
		inUndoRedo : {
			value : i
		}
	})
};
OSF.DDA.NodeDeletedEventArgs = function (n, t, i) {
	OSF.OUtil.defineEnumerableProperties(this, {
		type : {
			value : Microsoft.Office.WebExtension.EventType.DataNodeDeleted
		},
		oldNode : {
			value : n
		},
		oldNextSibling : {
			value : t
		},
		inUndoRedo : {
			value : i
		}
	})
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureDataNode = function (n) {
	if (n)
		return new OSF.DDA.CustomXmlNode(n[OSF.DDA.DataNodeProperties.Handle], n[OSF.DDA.DataNodeProperties.NodeType], n[OSF.DDA.DataNodeProperties.NamespaceUri], n[OSF.DDA.DataNodeProperties.BaseName])
};
OSF.DDA.OMFactory.manufactureDataPart = function (n, t) {
	return new OSF.DDA.CustomXmlPart(t, n[OSF.DDA.DataPartProperties.Id], n[OSF.DDA.DataPartProperties.BuiltIn])
};
OSF.DDA.AsyncMethodNames.addNames({
	AddDataPartAsync : "addAsync",
	GetDataPartByIdAsync : "getByIdAsync",
	GetDataPartsByNameSpaceAsync : "getByNamespaceAsync",
	DeleteDataPartAsync : "deleteAsync",
	GetPartNodesAsync : "getNodesAsync",
	GetPartXmlAsync : "getXmlAsync",
	AddDataPartNamespaceAsync : "addNamespaceAsync",
	GetDataPartNamespaceAsync : "getNamespaceAsync",
	GetDataPartPrefixAsync : "getPrefixAsync",
	GetRelativeNodesAsync : "getNodesAsync",
	GetNodeValueAsync : "getNodeValueAsync",
	GetNodeXmlAsync : "getXmlAsync",
	SetNodeValueAsync : "setNodeValueAsync",
	SetNodeXmlAsync : "setXmlAsync",
	GetNodeTextAsync : "getTextAsync",
	SetNodeTextAsync : "setTextAsync"
}), function () {
	function r(n) {
		return OSF.DDA.OMFactory.manufactureDataPart(n, Microsoft.Office.WebExtension.context.document.customXmlParts)
	}
	function e(n) {
		return OSF.DDA.OMFactory.manufactureDataNode(n)
	}
	function i(n) {
		var t = n[Microsoft.Office.WebExtension.Parameters.Data];
		return t == undefined ? null : t
	}
	function u(n) {
		return n.id
	}
	function f(n, t) {
		return t
	}
	function t(n, t) {
		return t
	}
	var n = "string";
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddDataPartAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Xml,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : r
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetDataPartByIdAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Id,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : r
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetDataPartsByNameSpaceAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Namespace,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [],
		onSucceeded : function (n) {
			return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataPartList], r)
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.DeleteDataPartAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : u
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetPartNodesAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.XPath,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : u
			}
		],
		onSucceeded : function (n) {
			return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataNodeList], e)
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetPartXmlAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : u
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.AddDataPartNamespaceAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Prefix,
				types : [n]
			}, {
				name : Microsoft.Office.WebExtension.Parameters.Namespace,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : f
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetDataPartNamespaceAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Prefix,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : f
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetDataPartPrefixAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Namespace,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataPartProperties.Id,
				value : f
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetRelativeNodesAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.XPath,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		],
		onSucceeded : function (n) {
			return OSF.OUtil.mapList(n[OSF.DDA.ListDescriptors.DataNodeList], e)
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetNodeValueAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetNodeXmlAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.SetNodeValueAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Data,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.SetNodeXmlAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Xml,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.GetNodeTextAsync,
		requiredArguments : [],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		],
		onSucceeded : i
	});
	OSF.DDA.AsyncMethodCalls.define({
		method : OSF.DDA.AsyncMethodNames.SetNodeTextAsync,
		requiredArguments : [{
				name : Microsoft.Office.WebExtension.Parameters.Text,
				types : [n]
			}
		],
		supportedOptions : [],
		privateStateCallbacks : [{
				name : OSF.DDA.DataNodeProperties.Handle,
				value : t
			}
		]
	})
}
();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.DataPartProperties,
	fromHost : [{
			name : OSF.DDA.DataPartProperties.Id,
			value : 0
		}, {
			name : OSF.DDA.DataPartProperties.BuiltIn,
			value : 1
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.DataNodeProperties,
	fromHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}, {
			name : OSF.DDA.DataNodeProperties.BaseName,
			value : 1
		}, {
			name : OSF.DDA.DataNodeProperties.NamespaceUri,
			value : 2
		}, {
			name : OSF.DDA.DataNodeProperties.NodeType,
			value : 3
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
	fromHost : [{
			name : OSF.DDA.DataNodeEventProperties.InUndoRedo,
			value : 0
		}, {
			name : OSF.DDA.DataNodeEventProperties.NewNode,
			value : 1
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
	fromHost : [{
			name : OSF.DDA.DataNodeEventProperties.InUndoRedo,
			value : 0
		}, {
			name : OSF.DDA.DataNodeEventProperties.OldNode,
			value : 1
		}, {
			name : OSF.DDA.DataNodeEventProperties.NewNode,
			value : 2
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
	fromHost : [{
			name : OSF.DDA.DataNodeEventProperties.InUndoRedo,
			value : 0
		}, {
			name : OSF.DDA.DataNodeEventProperties.OldNode,
			value : 1
		}, {
			name : OSF.DDA.DataNodeEventProperties.NextSiblingNode,
			value : 2
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.DataNodeEventProperties.OldNode,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.DataNodeProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.DataNodeEventProperties.NewNode,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.DataNodeProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.DataNodeEventProperties.NextSiblingNode,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.DataNodeProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddDataPartMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.DataPartProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Xml,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataPartByIdMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.DataPartProperties,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataPartsByNamespaceMethod,
	fromHost : [{
			name : OSF.DDA.ListDescriptors.DataPartList,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Namespace,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataPartXmlMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataPartNodesMethod,
	fromHost : [{
			name : OSF.DDA.ListDescriptors.DataNodeList,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.XPath,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidDeleteDataPartMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Id,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataNodeValueMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataNodeXmlMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataNodesMethod,
	fromHost : [{
			name : OSF.DDA.ListDescriptors.DataNodeList,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.XPath,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSetDataNodeValueMethod,
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSetDataNodeXmlMethod,
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Xml,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidAddDataNamespaceMethod,
	toHost : [{
			name : OSF.DDA.DataPartProperties.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Prefix,
			value : 1
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Namespace,
			value : 2
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataUriByPrefixMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataPartProperties.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Prefix,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataPrefixByUriMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataPartProperties.Id,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Namespace,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetDataNodeTextMethod,
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	],
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidSetDataNodeTextMethod,
	toHost : [{
			name : OSF.DDA.DataNodeProperties.Handle,
			value : 0
		}, {
			name : Microsoft.Office.WebExtension.Parameters.Text,
			value : 1
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidDataNodeAddedEvent,
	fromHost : [{
			name : OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidDataNodeReplacedEvent,
	fromHost : [{
			name : OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.EventDispId.dispidDataNodeDeletedEvent,
	fromHost : [{
			name : OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
	fromHost : [{
			name : OSF.DDA.FilePropertiesDescriptor.Url,
			value : 0
		}
	],
	isComplexType : !0
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,
	fromHost : [{
			name : OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	ExecuteRichApiRequestAsync : "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method : OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
	requiredArguments : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			types : ["object"]
		}
	],
	supportedOptions : []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type : OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
	toHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : 0
		}
	],
	fromHost : [{
			name : Microsoft.Office.WebExtension.Parameters.Data,
			value : OSF.DDA.SafeArray.Delegate.ParameterMap.self
		}
	]
});
OSF.DDA.WordDocument = function (n, t) {
	var i = this;
	OSF.DDA.WordDocument.uber.constructor.call(i, n, new OSF.DDA.BindingFacade(i), t);
	OSF.DDA.DispIdHost.addAsyncMethods(i, [OSF.DDA.AsyncMethodNames.GoToByIdAsync, OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync, OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync]);
	OSF.OUtil.defineEnumerableProperty(i, "customXmlParts", {
		value : new OSF.DDA.CustomXmlParts
	});
	OSF.OUtil.finalizeProperties(i)
};
OSF.OUtil.extend(OSF.DDA.WordDocument, OSF.DDA.JsomDocument);
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function (n, t) {
	OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
	n.doc = new OSF.DDA.WordDocument(n, this._initializeSettings(n, !1));
	OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
	t()
}, function (n) {
	var t = function () {
		function n(n, t) {
			this.m_actionInfo = n;
			this.m_isWriteOperation = t
		}
		return Object.defineProperty(n.prototype, "actionInfo", {
			get : function () {
				return this.m_actionInfo
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(n.prototype, "isWriteOperation", {
			get : function () {
				return this.m_isWriteOperation
			},
			enumerable : !0,
			configurable : !0
		}),
		n
	}
	();
	n.Action = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t() {}

		return t.createSetPropertyAction = function (t, i, r, u) {
			var f;
			n.Utility.validateObjectPath(i);
			var e = {
				Id : t._nextId(),
				ActionType : 4,
				Name : r,
				ObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			},
			s = [u],
			o = n.Utility.setMethodArguments(t, e.ArgumentInfo, s);
			return n.Utility.validateReferencedObjectPaths(o),
			f = new n.Action(e, !0),
			t._pendingRequest.addAction(f),
			t._pendingRequest.addReferencedObjectPath(i._objectPath),
			t._pendingRequest.addReferencedObjectPaths(o),
			f
		},
		t.createMethodAction = function (t, i, r, u, f) {
			var e,
			o,
			h,
			s;
			return n.Utility.validateObjectPath(i),
			e = {
				Id : t._nextId(),
				ActionType : 3,
				Name : r,
				ObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			},
			o = n.Utility.setMethodArguments(t, e.ArgumentInfo, f),
			n.Utility.validateReferencedObjectPaths(o),
			h = u != 1,
			s = new n.Action(e, h),
			t._pendingRequest.addAction(s),
			t._pendingRequest.addReferencedObjectPath(i._objectPath),
			t._pendingRequest.addReferencedObjectPaths(o),
			s
		},
		t.createQueryAction = function (t, i, r) {
			var u,
			f;
			return n.Utility.validateObjectPath(i),
			u = {
				Id : t._nextId(),
				ActionType : 2,
				Name : "",
				ObjectPathId : i._objectPath.objectPathInfo.Id
			},
			u.QueryInfo = r,
			f = new n.Action(u, !1),
			t._pendingRequest.addAction(f),
			t._pendingRequest.addReferencedObjectPath(i._objectPath),
			f
		},
		t.createInstantiateAction = function (t, i) {
			n.Utility.validateObjectPath(i);
			var u = {
				Id : t._nextId(),
				ActionType : 1,
				Name : "",
				ObjectPathId : i._objectPath.objectPathInfo.Id
			},
			r = new n.Action(u, !1);
			return t._pendingRequest.addAction(r),
			t._pendingRequest.addReferencedObjectPath(i._objectPath),
			t._pendingRequest.addActionResultHandler(r, new n.InstantiateActionResultHandler(i)),
			r
		},
		t.createTraceAction = function (t, i) {
			var r = {
				Id : t._nextId(),
				ActionType : 5,
				Name : "Trace",
				ObjectPathId : 0
			},
			u = new n.Action(r, !1);
			return t._pendingRequest.addAction(u),
			t._pendingRequest.addTrace(r.Id, i),
			u
		},
		t
	}
	();
	n.ActionFactory = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(t, i) {
			n.Utility.checkArgumentNull(t, "context");
			this.m_context = t;
			this.m_objectPath = i;
			this.m_objectPath && (t._processingResult || (n.ActionFactory.createInstantiateAction(t, this), t._autoCleanup && this._KeepReference && t.trackedObjects._autoAdd(this)))
		}
		return Object.defineProperty(t.prototype, "context", {
			get : function () {
				return this.m_context
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "_objectPath", {
			get : function () {
				return this.m_objectPath
			},
			set : function (n) {
				this.m_objectPath = n
			},
			enumerable : !0,
			configurable : !0
		}),
		t.prototype._handleResult = function (n) {},
		t
	}
	();
	n.ClientObject = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(n) {
			this.m_context = n;
			this.m_actions = [];
			this.m_actionResultHandler = {};
			this.m_referencedObjectPaths = {};
			this.m_flags = 0;
			this.m_traceInfos = {}

		}
		return Object.defineProperty(t.prototype, "flags", {
			get : function () {
				return this.m_flags
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "traceInfos", {
			get : function () {
				return this.m_traceInfos
			},
			enumerable : !0,
			configurable : !0
		}),
		t.prototype.addAction = function (n) {
			n.isWriteOperation && (this.m_flags = this.m_flags | 1);
			this.m_actions.push(n)
		},
		Object.defineProperty(t.prototype, "hasActions", {
			get : function () {
				return this.m_actions.length > 0
			},
			enumerable : !0,
			configurable : !0
		}),
		t.prototype.addTrace = function (n, t) {
			this.m_traceInfos[n] = t
		},
		t.prototype.addReferencedObjectPath = function (t) {
			if (!this.m_referencedObjectPaths[t.objectPathInfo.Id])
				for (t.isValid || n.Utility.throwError(n.ResourceStrings.invalidObjectPath, n.Utility.getObjectPathExpression(t)); t; )
					t.isWriteOperation && (this.m_flags = this.m_flags | 1), this.m_referencedObjectPaths[t.objectPathInfo.Id] = t, t.objectPathInfo.ObjectPathType == 3 && this.addReferencedObjectPaths(t.argumentObjectPaths), t = t.parentObjectPath
		},
		t.prototype.addReferencedObjectPaths = function (n) {
			if (n)
				for (var t = 0; t < n.length; t++)
					this.addReferencedObjectPath(n[t])
		},
		t.prototype.addActionResultHandler = function (n, t) {
			this.m_actionResultHandler[n.actionInfo.Id] = t
		},
		t.prototype.buildRequestMessageBody = function () {
			var r = {},
			t,
			i,
			n;
			for (t in this.m_referencedObjectPaths)
				r[t] = this.m_referencedObjectPaths[t].objectPathInfo;
			for (i = [], n = 0; n < this.m_actions.length; n++)
				i.push(this.m_actions[n].actionInfo);
			return {
				Actions : i,
				ObjectPaths : r
			}
		},
		t.prototype.processResponse = function (n) {
			var t,
			i,
			r;
			if (n && n.Results)
				for (t = 0; t < n.Results.length; t++)
					i = n.Results[t], r = this.m_actionResultHandler[i.ActionId], r && r._handleResult(i.Value)
		},
		t.prototype.invalidatePendingInvalidObjectPaths = function () {
			for (var n in this.m_referencedObjectPaths)
				this.m_referencedObjectPaths[n].isInvalidAfterRequest && (this.m_referencedObjectPaths[n].isValid = !1)
		},
		t
	}
	();
	n.ClientRequest = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(t) {
			this.m_nextId = 0;
			this.m_url = t;
			n.Utility.isNullOrEmptyString(this.m_url) && (this.m_url = n.Constants.localDocument);
			this._processingResult = !1;
			this._customData = n.Constants.iterativeExecutor;
			this._requestExecutor = new n.OfficeJsRequestExecutor;
			this.sync = this.sync.bind(this)
		}
		return Object.defineProperty(t.prototype, "_pendingRequest", {
			get : function () {
				return this.m_pendingRequest == null && (this.m_pendingRequest = new n.ClientRequest(this)),
				this.m_pendingRequest
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "trackedObjects", {
			get : function () {
				return this.m_trackedObjects || (this.m_trackedObjects = new n.TrackedObjects(this)),
				this.m_trackedObjects
			},
			enumerable : !0,
			configurable : !0
		}),
		t.prototype.load = function (t, i) {
			var u,
			f,
			r,
			e;
			n.Utility.validateContext(this, t);
			u = {};
			typeof i == "string" ? (f = i, u.Select = this.parseSelectExpand(f)) : Array.isArray(i) ? u.Select = i : typeof i == "object" ? (r = i, typeof r.select == "string" ? u.Select = this.parseSelectExpand(r.select) : Array.isArray(r.select) ? u.Select = r.select : n.Utility.isNullOrUndefined(r.select) || n.Utility.throwError(n.ResourceStrings.invalidArgument, "option.select"), typeof r.expand == "string" ? u.Expand = this.parseSelectExpand(r.expand) : Array.isArray(r.expand) ? u.Expand = r.expand : n.Utility.isNullOrUndefined(r.expand) || n.Utility.throwError(n.ResourceStrings.invalidArgument, "option.expand"), typeof r.top == "number" ? u.Top = r.top : n.Utility.isNullOrUndefined(r.top) || n.Utility.throwError(n.ResourceStrings.invalidArgument, "option.top"), typeof r.skip == "number" ? u.Skip = r.skip : n.Utility.isNullOrUndefined(r.skip) || n.Utility.throwError(n.ResourceStrings.invalidArgument, "option.skip")) : n.Utility.isNullOrUndefined(i) || n.Utility.throwError(n.ResourceStrings.invalidArgument, "option");
			e = n.ActionFactory.createQueryAction(this, t, u);
			this._pendingRequest.addActionResultHandler(e, t)
		},
		t.prototype.trace = function (t) {
			n.ActionFactory.createTraceAction(this, t)
		},
		t.prototype.parseSelectExpand = function (t) {
			var f = [],
			u,
			i,
			r;
			if (!n.Utility.isNullOrEmptyString(t))
				for (u = t.split(","), i = 0; i < u.length; i++)
					r = u[i], r = r.trim(), f.push(r);
			return f
		},
		t.prototype.syncPrivate = function (t, i) {
			var r = this._pendingRequest,
			e,
			f;
			if (!r.hasActions) {
				t();
				return
			}
			this.m_pendingRequest = null;
			var o = r.buildRequestMessageBody(),
			s = r.flags,
			u = this._requestExecutor;
			u || (u = new n.OfficeJsRequestExecutor);
			e = {
				Url : this.m_url,
				Headers : null,
				Body : o
			};
			r.invalidatePendingInvalidObjectPaths();
			f = this;
			u.executeAsync(this._customData, s, e, function (u) {
				var e,
				s = [],
				h,
				o,
				c,
				l;
				if (n.Utility.isNullOrEmptyString(u.ErrorCode) ? u.Body && u.Body.Error && (e = new n._Internal.RuntimeError(u.Body.Error.Code, u.Body.Error.Message, s, {
								errorLocation : u.Body.Error.Location
							})) : e = new n._Internal.RuntimeError(u.ErrorCode, u.ErrorMessage, s, {}), u.Body && u.Body.TraceIds)
					for (h = r.traceInfos, o = 0; o < u.Body.TraceIds.length; o++)
						c = u.Body.TraceIds[o], l = h[c], s.push(l);
				if (e) {
					i(e);
					return
				} else {
					f._processingResult = !0;
					try {
						r.processResponse(u.Body)
					}
					finally {
						f._processingResult = !1
					}
					t();
					return
				}
			})
		},
		t.prototype.sync = function (t) {
			var i = this;
			return n._EnsurePromise(),
			new n.Promise(function (n, r) {
				i.syncPrivate(function () {
					n(t)
				}, function (n) {
					r(n)
				})
			})
		},
		t._run = function (t, i, r, u, f, e) {
			r === void 0 && (r = 3);
			u === void 0 && (u = 5e3);
			n._EnsurePromise();
			var c = new n.Promise(function (n, t) {
					n()
				}),
			o,
			h = !1,
			s;
			return c.then(function () {
				o = t();
				o._autoCleanup = !0;
				var r = i(o);
				return (n.Utility.isNullOrUndefined(r) || typeof r.then != "function") && n.Utility.throwError(n.ResourceStrings.runMustReturnPromise),
				r
			}).then(function (n) {
				return o.sync(n)
			}).then(function (n) {
				h = !0;
				s = n
			}).catch (function (n) {
				s = n
			})
				.then(function () {
					function s() {
						n++;
						for (var i in t)
							o.trackedObjects.remove(t[i]);
						o.sync().then(function () {
							f && f(n)
						}).catch (function () {
							e && e(n);
							n < r && setTimeout(function () {
								s()
							}, u)
						})
					}
					var t = o.trackedObjects._retrieveAndClearAutoCleanupList(),
					i,
					n;
					o._autoCleanup = !1;
					for (i in t)
						t[i]._objectPath.isValid = !1;
					n = 0;
					s()
				}).then(function () {
					if (h)
						return s;
					else
						throw s;
				})
		},
		t.prototype._nextId = function () {
			return ++this.m_nextId
		},
		t
	}
	();
	n.ClientRequestContext = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	(function (n) {
		n[n.None = 0] = "None";
		n[n.WriteOperation = 1] = "WriteOperation"
	})(n.ClientRequestFlags || (n.ClientRequestFlags = {}));
	var t = n.ClientRequestFlags
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function n() {}

		return Object.defineProperty(n.prototype, "value", {
			get : function () {
				return this.m_value
			},
			enumerable : !0,
			configurable : !0
		}),
		n.prototype._handleResult = function (n) {
			this.m_value = n
		},
		n
	}
	();
	n.ClientResult = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function n() {}

		return n.getItemAt = "GetItemAt",
		n.id = "Id",
		n.idPrivate = "_Id",
		n.index = "_Index",
		n.items = "_Items",
		n.iterativeExecutor = "IterativeExecutor",
		n.localDocument = "http://document.localhost/",
		n.localDocumentApiPrefix = "http://document.localhost/_api/",
		n.referenceId = "_ReferenceId",
		n
	}
	();
	n.Constants = t
}
(OfficeExtension || (OfficeExtension = {}));
__extends = this.__extends || function (n, t) {
	function r() {
		this.constructor = n
	}
	for (var i in t)
		t.hasOwnProperty(i) && (n[i] = t[i]);
	r.prototype = t.prototype;
	n.prototype = new r
}, function (n) {
	var t;
	(function (n) {
		var t = function (n) {
			function t(t, i, r, u) {
				n.call(this, i);
				this.name = "OfficeExtension.Error";
				this.code = t;
				this.message = i;
				this.traceMessages = r;
				this.debugInfo = u
			}
			return __extends(t, n),
			t.prototype.toString = function () {
				return this.code + ": " + this.message
			},
			t
		}
		(Error);
		n.RuntimeError = t
	})(t = n._Internal || (n._Internal = {}));
	n.Error = n._Internal.RuntimeError
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function n() {}

		return n.accessDenied = "AccessDenied",
		n.generalException = "GeneralException",
		n.activityLimitReached = "ActivityLimitReached",
		n
	}
	();
	n.ErrorCodes = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(n) {
			this.m_clientObject = n
		}
		return t.prototype._handleResult = function (t) {
			n.Utility.fixObjectPathIfNecessary(this.m_clientObject, t);
			t && !n.Utility.isNullOrUndefined(t[n.Constants.referenceId]) && this.m_clientObject._initReferenceId && this.m_clientObject._initReferenceId(t[n.Constants.referenceId])
		},
		t
	}
	();
	n.InstantiateActionResultHandler = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t,
	i,
	r,
	u;
	(function (n) {
		n[n.CustomData = 0] = "CustomData";
		n[n.Method = 1] = "Method";
		n[n.PathAndQuery = 2] = "PathAndQuery";
		n[n.Headers = 3] = "Headers";
		n[n.Body = 4] = "Body";
		n[n.AppPermission = 5] = "AppPermission";
		n[n.RequestFlags = 6] = "RequestFlags"
	})(n.RichApiRequestMessageIndex || (n.RichApiRequestMessageIndex = {}));
	t = n.RichApiRequestMessageIndex,
	function (n) {
		n[n.StatusCode = 0] = "StatusCode";
		n[n.Headers = 1] = "Headers";
		n[n.Body = 2] = "Body"
	}
	(n.RichApiResponseMessageIndex || (n.RichApiResponseMessageIndex = {}));
	i = n.RichApiResponseMessageIndex,
	function (n) {
		n[n.Instantiate = 1] = "Instantiate";
		n[n.Query = 2] = "Query";
		n[n.Method = 3] = "Method";
		n[n.SetProperty = 4] = "SetProperty";
		n[n.Trace = 5] = "Trace"
	}
	(n.ActionType || (n.ActionType = {}));
	r = n.ActionType,
	function (n) {
		n[n.GlobalObject = 1] = "GlobalObject";
		n[n.NewObject = 2] = "NewObject";
		n[n.Method = 3] = "Method";
		n[n.Property = 4] = "Property";
		n[n.Indexer = 5] = "Indexer";
		n[n.ReferenceId = 6] = "ReferenceId"
	}
	(n.ObjectPathType || (n.ObjectPathType = {}));
	u = n.ObjectPathType
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(n, t, i, r) {
			this.m_objectPathInfo = n;
			this.m_parentObjectPath = t;
			this.m_isWriteOperation = !1;
			this.m_isCollection = i;
			this.m_isInvalidAfterRequest = r;
			this.m_isValid = !0
		}
		return Object.defineProperty(t.prototype, "objectPathInfo", {
			get : function () {
				return this.m_objectPathInfo
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "isWriteOperation", {
			get : function () {
				return this.m_isWriteOperation
			},
			set : function (n) {
				this.m_isWriteOperation = n
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "isCollection", {
			get : function () {
				return this.m_isCollection
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "isInvalidAfterRequest", {
			get : function () {
				return this.m_isInvalidAfterRequest
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "parentObjectPath", {
			get : function () {
				return this.m_parentObjectPath
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "argumentObjectPaths", {
			get : function () {
				return this.m_argumentObjectPaths
			},
			set : function (n) {
				this.m_argumentObjectPaths = n
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(t.prototype, "isValid", {
			get : function () {
				return this.m_isValid
			},
			set : function (n) {
				this.m_isValid = n
			},
			enumerable : !0,
			configurable : !0
		}),
		t.prototype.updateUsingObjectData = function (t) {
			var r = t[n.Constants.referenceId],
			i;
			if (!n.Utility.isNullOrEmptyString(r)) {
				this.m_isInvalidAfterRequest = !1;
				this.m_isValid = !0;
				this.m_objectPathInfo.ObjectPathType = 6;
				this.m_objectPathInfo.Name = r;
				this.m_objectPathInfo.ArgumentInfo = {};
				this.m_parentObjectPath = null;
				this.m_argumentObjectPaths = null;
				return
			}
			if (this.parentObjectPath && this.parentObjectPath.isCollection && (i = t[n.Constants.id], n.Utility.isNullOrUndefined(i) && (i = t[n.Constants.idPrivate]), !n.Utility.isNullOrUndefined(i))) {
				this.m_isInvalidAfterRequest = !1;
				this.m_isValid = !0;
				this.m_objectPathInfo.ObjectPathType = 5;
				this.m_objectPathInfo.Name = "";
				this.m_objectPathInfo.ArgumentInfo = {};
				this.m_objectPathInfo.ArgumentInfo.Arguments = [i];
				this.m_argumentObjectPaths = null;
				return
			}
		},
		t
	}
	();
	n.ObjectPath = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t() {}

		return t.createGlobalObjectObjectPath = function (t) {
			var i = {
				Id : t._nextId(),
				ObjectPathType : 1,
				Name : ""
			};
			return new n.ObjectPath(i, null, !1, !1)
		},
		t.createNewObjectObjectPath = function (t, i, r) {
			var u = {
				Id : t._nextId(),
				ObjectPathType : 2,
				Name : i
			};
			return new n.ObjectPath(u, null, r, !1)
		},
		t.createPropertyObjectPath = function (t, i, r, u, f) {
			var e = {
				Id : t._nextId(),
				ObjectPathType : 4,
				Name : r,
				ParentObjectPathId : i._objectPath.objectPathInfo.Id
			};
			return new n.ObjectPath(e, i._objectPath, u, f)
		},
		t.createIndexerObjectPath = function (t, i, r) {
			var u = {
				Id : t._nextId(),
				ObjectPathType : 5,
				Name : "",
				ParentObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			};
			return u.ArgumentInfo.Arguments = r,
			new n.ObjectPath(u, i._objectPath, !1, !1)
		},
		t.createIndexerObjectPathUsingParentPath = function (t, i, r) {
			var u = {
				Id : t._nextId(),
				ObjectPathType : 5,
				Name : "",
				ParentObjectPathId : i.objectPathInfo.Id,
				ArgumentInfo : {}

			};
			return u.ArgumentInfo.Arguments = r,
			new n.ObjectPath(u, i, !1, !1)
		},
		t.createMethodObjectPath = function (t, i, r, u, f, e, o) {
			var h = {
				Id : t._nextId(),
				ObjectPathType : 3,
				Name : r,
				ParentObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			},
			c = n.Utility.setMethodArguments(t, h.ArgumentInfo, f),
			s = new n.ObjectPath(h, i._objectPath, e, o);
			return s.argumentObjectPaths = c,
			s.isWriteOperation = u != 1,
			s
		},
		t.createChildItemObjectPathUsingIndexer = function (t, i, r) {
			var f = r[n.Constants.id],
			u;
			return n.Utility.isNullOrUndefined(f) && (f = r[n.Constants.idPrivate]),
			u = u = {
				Id : t._nextId(),
				ObjectPathType : 5,
				Name : "",
				ParentObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			},
			u.ArgumentInfo.Arguments = [f],
			new n.ObjectPath(u, i._objectPath, !1, !1)
		},
		t.createChildItemObjectPathUsingGetItemAt = function (t, i, r, u) {
			var e = r[n.Constants.index],
			f;
			return e && (u = e),
			f = {
				Id : t._nextId(),
				ObjectPathType : 3,
				Name : n.Constants.getItemAt,
				ParentObjectPathId : i._objectPath.objectPathInfo.Id,
				ArgumentInfo : {}

			},
			f.ArgumentInfo.Arguments = [u],
			new n.ObjectPath(f, i._objectPath, !1, !1)
		},
		t
	}
	();
	n.ObjectPathFactory = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t() {}

		return t.prototype.executeAsync = function (i, r, u, f) {
			var e = JSON.stringify(u.Body),
			o;
			n.Utility.log("Request:");
			n.Utility.log(e);
			o = n.RichApiMessageUtility.buildRequestMessageSafeArray(i, r, "POST", "ProcessQuery", null, e);
			OSF.DDA.RichApi.executeRichApiRequestAsync(o, function (i) {
				var r,
				u;
				n.Utility.log("Response:");
				n.Utility.log(JSON.stringify(i));
				r = {
					ErrorCode : "",
					ErrorMessage : "",
					Headers : null,
					Body : null
				};
				i.status == "succeeded" ? (u = n.RichApiMessageUtility.getResponseBody(i), r.Body = JSON.parse(u), r.Headers = n.RichApiMessageUtility.getResponseHeaders(i), f(r)) : (r.ErrorCode = n.ErrorCodes.generalException, i.error.code == t.OfficeJsErrorCode_ooeNoCapability ? r.ErrorCode = n.ErrorCodes.accessDenied : i.error.code == t.OfficeJsErrorCode_ooeActivityLimitReached && (r.ErrorCode = n.ErrorCodes.activityLimitReached), r.ErrorMessage = i.error.message, f(r))
			})
		},
		t.OfficeJsErrorCode_ooeNoCapability = 7e3,
		t.OfficeJsErrorCode_ooeActivityLimitReached = 5102,
		t
	}
	();
	n.OfficeJsRequestExecutor = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	function u(n) {
		return t.settings.oldxhr = n,
		r
	}
	function r() {
		return new t
	}
	var i = function () {
		function n() {}

		return n
	}
	(),
	t;
	n.OfficeXHRSettings = i;
	n.resetXHRFactory = u;
	n.officeXHRFactory = r;
	t = function () {
		function t() {}

		return t.prototype.open = function (i, r) {
			if (this.m_method = i, this.m_url = r, this.m_url.toLowerCase().indexOf(n.Constants.localDocumentApiPrefix) == 0)
				this.m_url = this.m_url.substr(n.Constants.localDocumentApiPrefix.length);
			else {
				this.m_innerXhr = t.settings.oldxhr();
				var u = this;
				this.m_innerXhr.onreadystatechange = function () {
					u.innerXhrOnreadystatechage()
				};
				this.m_innerXhr.open(i, this.m_url)
			}
		},
		t.prototype.abort = function () {
			this.m_innerXhr && this.m_innerXhr.abort()
		},
		t.prototype.send = function (i) {
			var f,
			u,
			r;
			this.m_innerXhr ? this.m_innerXhr.send(i) : (f = this, u = 0, n.Utility.isReadonlyRestRequest(this.m_method) || (u = 1), r = t.settings.executeRichApiRequestAsync, r || (r = OSF.DDA.RichApi.executeRichApiRequestAsync), r(n.RichApiMessageUtility.buildRequestMessageSafeArray("", u, this.m_method, this.m_url, this.m_requestHeaders, i), function (n) {
					f.officeContextRequestCallback(n)
				}))
		},
		t.prototype.setRequestHeader = function (n, t) {
			this.m_innerXhr ? this.m_innerXhr.setRequestHeader(n, t) : (this.m_requestHeaders || (this.m_requestHeaders = {}), this.m_requestHeaders[n] = t)
		},
		t.prototype.getResponseHeader = function (n) {
			return this.m_responseHeaders ? this.m_responseHeaders[n.toUpperCase()] : null
		},
		t.prototype.getAllResponseHeaders = function () {
			return this.m_allResponseHeaders
		},
		t.prototype.overrideMimeType = function (n) {
			this.m_innerXhr && this.m_innerXhr.overrideMimeType(n)
		},
		t.prototype.innerXhrOnreadystatechage = function () {
			this.readyState = this.m_innerXhr.readyState;
			this.readyState == t.DONE && (this.status = this.m_innerXhr.status, this.statusText = this.m_innerXhr.statusText, this.responseText = this.m_innerXhr.responseText, this.response = this.m_innerXhr.response, this.responseType = this.m_innerXhr.responseType, this.setAllResponseHeaders(this.m_innerXhr.getAllResponseHeaders()));
			this.onreadystatechange && this.onreadystatechange()
		},
		t.prototype.officeContextRequestCallback = function (i) {
			this.readyState = t.DONE;
			i.status == "succeeded" ? (this.status = n.RichApiMessageUtility.getResponseStatusCode(i), this.m_responseHeaders = n.RichApiMessageUtility.getResponseHeaders(i), console.debug("ResponseHeaders=" + JSON.stringify(this.m_responseHeaders)), this.responseText = n.RichApiMessageUtility.getResponseBody(i), console.debug("ResponseText=" + this.responseText), this.response = this.responseText) : (this.status = 500, this.statusText = "Internal Error");
			this.onreadystatechange && this.onreadystatechange()
		},
		t.prototype.setAllResponseHeaders = function (t) {
			var s,
			o,
			r,
			i,
			u,
			f,
			e;
			if (this.m_allResponseHeaders = t, this.m_responseHeaders = {}, this.m_allResponseHeaders != null)
				for (s = new RegExp("\r?\n"), o = this.m_allResponseHeaders.split(s), r = 0; r < o.length; r++)
					i = o[r], i != null && (u = i.indexOf(":"), u > 0 && (f = i.substr(0, u), e = i.substr(u + 1), f = n.Utility.trim(f), e = n.Utility.trim(e), this.m_responseHeaders[f.toUpperCase()] = e))
		},
		t.UNSENT = 0,
		t.OPENED = 1,
		t.DONE = 4,
		t.settings = new i,
		t
	}
	();
	n.OfficeXHR = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	function i() {
		n.Promise || t.Init()
	}
	n._EnsurePromise = i;
	var t;
	(function (t) {
		function i() {
			(function () {
				"use strict";
				function wt(n) {
					return typeof n == "function" || typeof n == "object" && n !== null
				}
				function b(n) {
					return typeof n == "function"
				}
				function bt(n) {
					return typeof n == "object" && n !== null
				}
				function kt(n) {
					k = n
				}
				function dt(n) {
					f = n
				}
				function ti() {
					var t = process.nextTick,
					n = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
					return Array.isArray(n) && n[1] === "0" && n[2] === "10" && (t = setImmediate),
					function () {
						t(c)
					}
				}
				function ii() {
					return function () {
						tt(c)
					}
				}
				function ri() {
					var n = 0,
					i = new ut(c),
					t = document.createTextNode("");
					return i.observe(t, {
						characterData : !0
					}),
					function () {
						t.data = n = ++n % 2
					}
				}
				function ui() {
					var n = new MessageChannel;
					return n.port1.onmessage = c,
					function () {
						n.port2.postMessage(0)
					}
				}
				function ft() {
					return function () {
						setTimeout(c, 1)
					}
				}
				function c() {
					for (var t, i, n = 0; n < h; n += 2)
						t = e[n], i = e[n + 1], t(i), e[n] = undefined, e[n + 1] = undefined;
					h = 0
				}
				function fi() {
					try {
						var t = require,
						n = t("vertx");
						return tt = n.runOnLoop || n.runOnContext,
						ii()
					} catch (i) {
						return ft()
					}
				}
				function l() {}

				function ei() {
					return new TypeError("You cannot resolve a promise with itself")
				}
				function oi() {
					return new TypeError("A promises callback cannot return that same promise.")
				}
				function si(n) {
					try {
						return n.then
					} catch (t) {
						return y.error = t,
						y
					}
				}
				function hi(n, t, i, r) {
					try {
						n.call(t, i, r)
					} catch (u) {
						return u
					}
				}
				function ci(n, i, u) {
					f(function (n) {
						var f = !1,
						e = hi(u, i, function (t) {
								f || (f = !0, i !== t ? v(n, t) : r(n, t))
							}, function (i) {
								f || (f = !0, t(n, i))
							}, "Settle: " + (n._label || " unknown promise"));
						!f && e && (f = !0, t(n, e))
					}, n)
				}
				function li(n, i) {
					i._state === a ? r(n, i._result) : i._state === s ? t(n, i._result) : p(i, undefined, function (t) {
						v(n, t)
					}, function (i) {
						t(n, i)
					})
				}
				function ai(n, i) {
					if (i.constructor === n.constructor)
						li(n, i);
					else {
						var u = si(i);
						u === y ? t(n, y.error) : u === undefined ? r(n, i) : b(u) ? ci(n, i, u) : r(n, i)
					}
				}
				function v(n, i) {
					n === i ? t(n, ei()) : wt(i) ? ai(n, i) : r(n, i)
				}
				function vi(n) {
					n._onerror && n._onerror(n._result);
					d(n)
				}
				function r(n, t) {
					n._state === o && (n._result = t, n._state = a, n._subscribers.length !== 0 && f(d, n))
				}
				function t(n, t) {
					n._state === o && (n._state = s, n._result = t, f(vi, n))
				}
				function p(n, t, i, r) {
					var u = n._subscribers,
					e = u.length;
					n._onerror = null;
					u[e] = t;
					u[e + a] = i;
					u[e + s] = r;
					e === 0 && n._state && f(d, n)
				}
				function d(n) {
					var i = n._subscribers,
					e = n._state,
					r,
					u,
					f,
					t;
					if (i.length !== 0) {
						for (f = n._result, t = 0; t < i.length; t += 3)
							r = i[t], u = i[t + e], r ? st(e, r, u, f) : u(f);
						n._subscribers.length = 0
					}
				}
				function ot() {
					this.error = null
				}
				function yi(n, t) {
					try {
						return n(t)
					} catch (i) {
						return w.error = i,
						w
					}
				}
				function st(n, i, u, f) {
					var c = b(u),
					e,
					l,
					h,
					y;
					if (c) {
						if (e = yi(u, f), e === w ? (y = !0, l = e.error, e = null) : h = !0, i === e) {
							t(i, oi());
							return
						}
					} else
						e = f, h = !0;
					i._state !== o || (c && h ? v(i, e) : y ? t(i, l) : n === a ? r(i, e) : n === s && t(i, e))
				}
				function pi(n, i) {
					try {
						i(function (t) {
							v(n, t)
						}, function (i) {
							t(n, i)
						})
					} catch (r) {
						t(n, r)
					}
				}
				function u(n, i) {
					var u = this;
					u._instanceConstructor = n;
					u.promise = new n(l);
					u._validateInput(i) ? (u._input = i, u.length = i.length, u._remaining = i.length, u._init(), u.length === 0 ? r(u.promise, u._result) : (u.length = u.length || 0, u._enumerate(), u._remaining === 0 && r(u.promise, u._result))) : t(u.promise, u._validationError())
				}
				function wi(n) {
					return new ht(this, n).promise
				}
				function bi(n) {
					function e(n) {
						v(i, n)
					}
					function s(n) {
						t(i, n)
					}
					var u = this,
					i = new u(l),
					f,
					r;
					if (!nt(n))
						return t(i, new TypeError("You must pass an array to race.")), i;
					for (f = n.length, r = 0; i._state === o && r < f; r++)
						p(u.resolve(n[r]), undefined, e, s);
					return i
				}
				function ki(n) {
					var i = this,
					t;
					return n && typeof n == "object" && n.constructor === i ? n : (t = new i(l), v(t, n), t)
				}
				function di(n) {
					var r = this,
					i = new r(l);
					return t(i, n),
					i
				}
				function gi() {
					throw new TypeError("You must pass a resolver function as the first argument to the promise constructor");
				}
				function nr() {
					throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
				}
				function i(n) {
					this._id = yt++;
					this._state = undefined;
					this._result = undefined;
					this._subscribers = [];
					l !== n && (b(n) || gi(), this instanceof i || nr(), pi(this, n))
				}
				var g,
				e,
				et,
				w,
				ht,
				ct,
				lt,
				at,
				vt,
				yt,
				pt;
				g = Array.isArray ? Array.isArray : function (n) {
					return Object.prototype.toString.call(n) === "[object Array]"
				};
				var nt = g,
				h = 0,
				tr = {}

				.toString,
				tt,
				k,
				f = function (n, t) {
					e[h] = n;
					e[h + 1] = t;
					h += 2;
					h === 2 && (k ? k(c) : et())
				};
				var it = typeof window != "undefined" ? window : undefined,
				rt = it || {},
				ut = rt.MutationObserver || rt.WebKitMutationObserver,
				gt = typeof process != "undefined" && {}

				.toString.call(process) === "[object process]",
				ni = typeof Uint8ClampedArray != "undefined" && typeof importScripts != "undefined" && typeof MessageChannel != "undefined";
				e = new Array(1e3);
				et = gt ? ti() : ut ? ri() : ni ? ui() : it === undefined && typeof require == "function" ? fi() : ft();
				var o = void 0,
				a = 1,
				s = 2,
				y = new ot;
				w = new ot;
				u.prototype._validateInput = function (n) {
					return nt(n)
				};
				u.prototype._validationError = function () {
					return new n.Error("Array Methods must be provided an Array")
				};
				u.prototype._init = function () {
					this._result = new Array(this.length)
				};
				ht = u;
				u.prototype._enumerate = function () {
					for (var n = this, i = n.length, r = n.promise, u = n._input, t = 0; r._state === o && t < i; t++)
						n._eachEntry(u[t], t)
				};
				u.prototype._eachEntry = function (n, t) {
					var i = this,
					r = i._instanceConstructor;
					bt(n) ? n.constructor === r && n._state !== o ? (n._onerror = null, i._settledAt(n._state, t, n._result)) : i._willSettleAt(r.resolve(n), t) : (i._remaining--, i._result[t] = n)
				};
				u.prototype._settledAt = function (n, i, u) {
					var f = this,
					e = f.promise;
					e._state === o && (f._remaining--, n === s ? t(e, u) : f._result[i] = u);
					f._remaining === 0 && r(e, f._result)
				};
				u.prototype._willSettleAt = function (n, t) {
					var i = this;
					p(n, undefined, function (n) {
						i._settledAt(a, t, n)
					}, function (n) {
						i._settledAt(s, t, n)
					})
				};
				ct = wi;
				lt = bi;
				at = ki;
				vt = di;
				yt = 0;
				pt = i;
				i.all = ct;
				i.race = lt;
				i.resolve = at;
				i.reject = vt;
				i._setScheduler = kt;
				i._setAsap = dt;
				i._asap = f;
				i.prototype = {
					constructor : i,
					then : function (n, t) {
						var u = this,
						i = u._state,
						r,
						e,
						o;
						return i === a && !n || i === s && !t ? this : (r = new this.constructor(l), e = u._result, i ? (o = arguments[i - 1], f(function () {
									st(i, r, o, e)
								})) : p(u, r, n, t), r)
					},
					"catch" : function (n) {
						return this.then(null, n)
					}
				};
				n.Promise = pt
			}).call(this)
		}
		t.Init = i
	})(t || (t = {}))
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	(function (n) {
		n[n.Default = 0] = "Default";
		n[n.Read = 1] = "Read"
	})(n.OperationType || (n.OperationType = {}));
	var t = n.OperationType
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t(n) {
			this._autoCleanupList = {};
			this.m_context = n
		}
		return t.prototype.add = function (n) {
			var t = this;
			Array.isArray(n) ? n.forEach(function (n) {
				return t._addCommon(n, !0)
			}) : this._addCommon(n, !0)
		},
		t.prototype._autoAdd = function (n) {
			this._addCommon(n, !1);
			this._autoCleanupList[n._objectPath.objectPathInfo.Id] = n
		},
		t.prototype._addCommon = function (t, i) {
			var r = t[n.Constants.referenceId];
			n.Utility.isNullOrEmptyString(r) && t._KeepReference && (t._KeepReference(), n.ActionFactory.createInstantiateAction(this.m_context, t), i && this.m_context._autoCleanup && delete this._autoCleanupList[t._objectPath.objectPathInfo.Id])
		},
		t.prototype.remove = function (n) {
			var t = this;
			Array.isArray(n) ? n.forEach(function (n) {
				return t._removeCommon(n)
			}) : this._removeCommon(n)
		},
		t.prototype._removeCommon = function (t) {
			var r = t[n.Constants.referenceId],
			i;
			n.Utility.isNullOrEmptyString(r) || (i = this.m_context._rootObject, i._RemoveReference && i._RemoveReference(r))
		},
		t.prototype._retrieveAndClearAutoCleanupList = function () {
			var n = this._autoCleanupList;
			return this._autoCleanupList = {},
			n
		},
		t
	}
	();
	n.TrackedObjects = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function n() {}

		return n.invalidObjectPath = "InvalidObjectPath",
		n.propertyNotLoaded = "PropertyNotLoaded",
		n.invalidRequestContext = "InvalidRequestContext",
		n.invalidArgument = "InvalidArgument",
		n.runMustReturnPromise = "RunMustReturnPromise",
		n
	}
	();
	n.ResourceStrings = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function n() {}

		return n.buildRequestMessageSafeArray = function (n, t, i, r, u, f) {
			var e = [],
			o;
			if (u)
				for (o in u)
					e.push(o), e.push(u[o]);
			var s = 0,
			h = "",
			c = "",
			l = "";
			return [n, i, r, e, f, s, t, h, c, l]
		},
		n.getResponseBody = function (t) {
			return n.getResponseBodyFromSafeArray(t.value.data)
		},
		n.getResponseHeaders = function (t) {
			return n.getResponseHeadersFromSafeArray(t.value.data)
		},
		n.getResponseBodyFromSafeArray = function (n) {
			var t = n[2],
			i;
			return typeof t == "string" ? t : (i = t, i.join(""))
		},
		n.getResponseHeadersFromSafeArray = function (n) {
			var i = n[1],
			r,
			t;
			if (!i)
				return null;
			for (r = {}, t = 0; t < i.length - 1; t += 2)
				r[i[t]] = i[t + 1];
			return r
		},
		n.getResponseStatusCode = function (t) {
			return n.getResponseStatusCodeFromSafeArray(t.value.data)
		},
		n.getResponseStatusCodeFromSafeArray = function (n) {
			return n[0]
		},
		n
	}
	();
	n.RichApiMessageUtility = t
}
(OfficeExtension || (OfficeExtension = {})), function (n) {
	var t = function () {
		function t() {}

		return t.checkArgumentNull = function (i, r) {
			t.isNullOrUndefined(i) && t.throwError(n.ResourceStrings.invalidArgument, r)
		},
		t.isNullOrUndefined = function (n) {
			return n === null ? !0 : typeof n == "undefined" ? !0 : !1
		},
		t.isUndefined = function (n) {
			return typeof n == "undefined" ? !0 : !1
		},
		t.isNullOrEmptyString = function (n) {
			return n === null ? !0 : typeof n == "undefined" ? !0 : n.length == 0 ? !0 : !1
		},
		t.trim = function (n) {
			return n.replace(new RegExp("^\\s+|\\s+$", "g"), "")
		},
		t.caseInsensitiveCompareString = function (n, i) {
			return t.isNullOrUndefined(n) ? t.isNullOrUndefined(i) : t.isNullOrUndefined(i) ? !1 : n.toUpperCase() == i.toUpperCase()
		},
		t.isReadonlyRestRequest = function (n) {
			return t.caseInsensitiveCompareString(n, "GET")
		},
		t.setMethodArguments = function (n, i, r) {
			if (t.isNullOrUndefined(r))
				return null;
			var u = [],
			f = [],
			e = t.collectObjectPathInfos(n, r, u, f);
			return (i.Arguments = r, e) ? (i.ReferencedObjectPathIds = f, u) : null
		},
		t.collectObjectPathInfos = function (i, r, u, f) {
			for (var o, h, c, s = !1, e = 0; e < r.length; e++)
				r[e]instanceof n.ClientObject ? (o = r[e], t.validateContext(i, o), r[e] = o._objectPath.objectPathInfo.Id, f.push(o._objectPath.objectPathInfo.Id), u.push(o._objectPath), s = !0) : Array.isArray(r[e]) ? (h = [], c = t.collectObjectPathInfos(i, r[e], u, h), c ? (f.push(h), s = !0) : f.push(0)) : f.push(0);
			return s
		},
		t.fixObjectPathIfNecessary = function (n, t) {
			n && n._objectPath && t && n._objectPath.updateUsingObjectData(t)
		},
		t.validateObjectPath = function (i) {
			for (var r = i._objectPath, u; r; )
				r.isValid || (u = t.getObjectPathExpression(r), t.throwError(n.ResourceStrings.invalidObjectPath, u)), r = r.parentObjectPath
		},
		t.validateReferencedObjectPaths = function (i) {
			var u,
			r,
			f;
			if (i)
				for (u = 0; u < i.length; u++)
					for (r = i[u]; r; )
						r.isValid || (f = t.getObjectPathExpression(r), t.throwError(n.ResourceStrings.invalidObjectPath, f)), r = r.parentObjectPath
		},
		t.validateContext = function (i, r) {
			r && r.context !== i && t.throwError(n.ResourceStrings.invalidRequestContext)
		},
		t.log = function (n) {
			t._logEnabled && window.console && window.console.log && window.console.log(n)
		},
		t.load = function (n, t) {
			n.context.load(n, t)
		},
		t.throwError = function (i, r) {
			throw new n._Internal.RuntimeError(i, t._getResourceString(i, r), [], {});
		},
		t.createRuntimeError = function (t, i, r) {
			return new n._Internal.RuntimeError(t, i, [], {
				errorLocation : r
			})
		},
		t._getResourceString = function (n, i) {
			var r = n,
			f,
			u;
			return window.Strings && window.Strings.OfficeOM && (f = "L_" + n, u = window.Strings.OfficeOM[f], u && (r = u)),
			t.isNullOrUndefined(i) || (r = r.replace("{0}", i)),
			r
		},
		t.throwIfNotLoaded = function (i, r) {
			t.isUndefined(r) && i.charCodeAt(0) != t.s_underscoreCharCode && t.throwError(n.ResourceStrings.propertyNotLoaded, i)
		},
		t.getObjectPathExpression = function (n) {
			for (var i = ""; n; ) {
				switch (n.objectPathInfo.ObjectPathType) {
				case 1:
					i = i;
					break;
				case 2:
					i = "new()" + (i.length > 0 ? "." : "") + i;
					break;
				case 3:
					i = t.normalizeName(n.objectPathInfo.Name) + "()" + (i.length > 0 ? "." : "") + i;
					break;
				case 4:
					i = t.normalizeName(n.objectPathInfo.Name) + (i.length > 0 ? "." : "") + i;
					break;
				case 5:
					i = "getItem()" + (i.length > 0 ? "." : "") + i;
					break;
				case 6:
					i = "_reference()" + (i.length > 0 ? "." : "") + i;
					break
				}
				n = n.parentObjectPath
			}
			return i
		},
		t._createPromiseFromResult = function (t) {
			return n._EnsurePromise(),
			new n.Promise(function (n, i) {
				n(t)
			})
		},
		t._addActionResultHandler = function (n, t, i) {
			n.context._pendingRequest.addActionResultHandler(t, i)
		},
		t._handleNavigationPropertyResults = function (n, i, r) {
			for (var u = 0; u < r.length - 1; u += 2)
				t.isUndefined(i[r[u + 1]]) || n[r[u]]._handleResult(i[r[u + 1]])
		},
		t.normalizeName = function (n) {
			return n.substr(0, 1).toLowerCase() + n.substr(1)
		},
		t._logEnabled = !1,
		t.s_underscoreCharCode = "_".charCodeAt(0),
		t
	}
	();
	n.Utility = t
}
(OfficeExtension || (OfficeExtension = {}));
__extends = this.__extends || function (n, t) {
	function r() {
		this.constructor = n
	}
	for (var i in t)
		t.hasOwnProperty(i) && (n[i] = t[i]);
	r.prototype = t.prototype;
	n.prototype = new r
}, function (n) {
	function y(t, i) {
		var u,
		r;
		if (OfficeExtension.Utility.isNullOrUndefined(i))
			return null;
		if (typeof i != "object" && OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "searchOptions"), i instanceof n.SearchOptions)
			return i;
		u = n.SearchOptions.newObject(t);
		for (r in i)
			i.hasOwnProperty(r) && (u[r] = i[r]);
		return u
	}
	var e = OfficeExtension.ObjectPathFactory.createPropertyObjectPath,
	u = OfficeExtension.ObjectPathFactory.createMethodObjectPath,
	a = OfficeExtension.ObjectPathFactory.createIndexerObjectPath,
	ot = OfficeExtension.ObjectPathFactory.createNewObjectObjectPath,
	v = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer,
	kt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt,
	f = OfficeExtension.ActionFactory.createMethodAction,
	r = OfficeExtension.ActionFactory.createSetPropertyAction,
	o = OfficeExtension.Utility.isNullOrUndefined,
	i = OfficeExtension.Utility.isUndefined,
	t = OfficeExtension.Utility.throwIfNotLoaded,
	s = OfficeExtension.Utility.load,
	h = OfficeExtension.Utility.fixObjectPathIfNecessary,
	c = OfficeExtension.Utility._addActionResultHandler,
	l = OfficeExtension.Utility._handleNavigationPropertyResults,
	st = function (a) {
		function v() {
			a.apply(this, arguments)
		}
		return __extends(v, a),
		Object.defineProperty(v.prototype, "contentControls", {
			get : function () {
				return this.m_contentControls || (this.m_contentControls = new n.ContentControlCollection(this.context, e(this.context, this, "ContentControls", !0, !1))),
				this.m_contentControls
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "font", {
			get : function () {
				return this.m_font || (this.m_font = new n.Font(this.context, e(this.context, this, "Font", !1, !1))),
				this.m_font
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "inlinePictures", {
			get : function () {
				return this.m_inlinePictures || (this.m_inlinePictures = new n.InlinePictureCollection(this.context, e(this.context, this, "InlinePictures", !0, !1))),
				this.m_inlinePictures
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "paragraphs", {
			get : function () {
				return this.m_paragraphs || (this.m_paragraphs = new n.ParagraphCollection(this.context, e(this.context, this, "Paragraphs", !0, !1))),
				this.m_paragraphs
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "parentContentControl", {
			get : function () {
				return this.m_parentContentControl || (this.m_parentContentControl = new n.ContentControl(this.context, e(this.context, this, "ParentContentControl", !1, !1))),
				this.m_parentContentControl
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "style", {
			get : function () {
				return t("style", this.m_style),
				this.m_style
			},
			set : function (n) {
				this.m_style = n;
				r(this.context, this, "Style", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "text", {
			get : function () {
				return t("text", this.m_text),
				this.m_text
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		v.prototype.clear = function () {
			f(this.context, this, "Clear", 0, [])
		},
		v.prototype.getHtml = function () {
			var t = f(this.context, this, "GetHtml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.getOoxml = function () {
			var t = f(this.context, this, "GetOoxml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.insertBreak = function (n, t) {
			f(this.context, this, "InsertBreak", 0, [n, t])
		},
		v.prototype.insertContentControl = function () {
			return new n.ContentControl(this.context, u(this.context, this, "InsertContentControl", 0, [], !1, !0))
		},
		v.prototype.insertFileFromBase64 = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertFileFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertHtml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertHtml", 0, [t, i], !1, !0))
		},
		v.prototype.insertInlinePictureFromBase64 = function (t, i) {
			return new n.InlinePicture(this.context, u(this.context, this, "InsertInlinePictureFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertOoxml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertOoxml", 0, [t, i], !1, !0))
		},
		v.prototype.insertParagraph = function (t, i) {
			return new n.Paragraph(this.context, u(this.context, this, "InsertParagraph", 0, [t, i], !1, !0))
		},
		v.prototype.insertText = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertText", 0, [t, i], !1, !0))
		},
		v.prototype.search = function (t, i) {
			return i = y(this.context, i),
			new n.SearchResultCollection(this.context, u(this.context, this, "Search", 1, [t, i], !0, !0))
		},
		v.prototype.select = function (n) {
			f(this.context, this, "Select", 1, [n])
		},
		v.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		v.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Style) || (this.m_style = t.Style);
				i(t.Text) || (this.m_text = t.Text);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "paragraphs", "Paragraphs", "parentContentControl", "ParentContentControl"])
			}
		},
		v.prototype.load = function (n) {
			return s(this, n),
			this
		},
		v.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		v
	}
	(OfficeExtension.ClientObject),
	p,
	w,
	b,
	k,
	d,
	g,
	nt,
	tt,
	it,
	rt,
	ut,
	ft,
	et,
	ht,
	ct,
	lt,
	at,
	vt,
	yt,
	pt,
	wt,
	bt;
	n.Body = st;
	p = function (a) {
		function v() {
			a.apply(this, arguments)
		}
		return __extends(v, a),
		Object.defineProperty(v.prototype, "contentControls", {
			get : function () {
				return this.m_contentControls || (this.m_contentControls = new n.ContentControlCollection(this.context, e(this.context, this, "ContentControls", !0, !1))),
				this.m_contentControls
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "font", {
			get : function () {
				return this.m_font || (this.m_font = new n.Font(this.context, e(this.context, this, "Font", !1, !1))),
				this.m_font
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "inlinePictures", {
			get : function () {
				return this.m_inlinePictures || (this.m_inlinePictures = new n.InlinePictureCollection(this.context, e(this.context, this, "InlinePictures", !0, !1))),
				this.m_inlinePictures
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "paragraphs", {
			get : function () {
				return this.m_paragraphs || (this.m_paragraphs = new n.ParagraphCollection(this.context, e(this.context, this, "Paragraphs", !0, !1))),
				this.m_paragraphs
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "parentContentControl", {
			get : function () {
				return this.m_parentContentControl || (this.m_parentContentControl = new n.ContentControl(this.context, e(this.context, this, "ParentContentControl", !1, !1))),
				this.m_parentContentControl
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "appearance", {
			get : function () {
				return t("appearance", this.m_appearance),
				this.m_appearance
			},
			set : function (n) {
				this.m_appearance = n;
				r(this.context, this, "Appearance", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "cannotDelete", {
			get : function () {
				return t("cannotDelete", this.m_cannotDelete),
				this.m_cannotDelete
			},
			set : function (n) {
				this.m_cannotDelete = n;
				r(this.context, this, "CannotDelete", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "cannotEdit", {
			get : function () {
				return t("cannotEdit", this.m_cannotEdit),
				this.m_cannotEdit
			},
			set : function (n) {
				this.m_cannotEdit = n;
				r(this.context, this, "CannotEdit", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "color", {
			get : function () {
				return t("color", this.m_color),
				this.m_color
			},
			set : function (n) {
				this.m_color = n;
				r(this.context, this, "Color", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "id", {
			get : function () {
				return t("id", this.m_id),
				this.m_id
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "placeholderText", {
			get : function () {
				return t("placeholderText", this.m_placeholderText),
				this.m_placeholderText
			},
			set : function (n) {
				this.m_placeholderText = n;
				r(this.context, this, "PlaceholderText", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "removeWhenEdited", {
			get : function () {
				return t("removeWhenEdited", this.m_removeWhenEdited),
				this.m_removeWhenEdited
			},
			set : function (n) {
				this.m_removeWhenEdited = n;
				r(this.context, this, "RemoveWhenEdited", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "style", {
			get : function () {
				return t("style", this.m_style),
				this.m_style
			},
			set : function (n) {
				this.m_style = n;
				r(this.context, this, "Style", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "tag", {
			get : function () {
				return t("tag", this.m_tag),
				this.m_tag
			},
			set : function (n) {
				this.m_tag = n;
				r(this.context, this, "Tag", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "text", {
			get : function () {
				return t("text", this.m_text),
				this.m_text
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "title", {
			get : function () {
				return t("title", this.m_title),
				this.m_title
			},
			set : function (n) {
				this.m_title = n;
				r(this.context, this, "Title", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "type", {
			get : function () {
				return t("type", this.m_type),
				this.m_type
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		v.prototype.clear = function () {
			f(this.context, this, "Clear", 0, [])
		},
		v.prototype.delete  = function (n) {
			f(this.context, this, "Delete", 0, [n])
		},
		v.prototype.getHtml = function () {
			var t = f(this.context, this, "GetHtml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.getOoxml = function () {
			var t = f(this.context, this, "GetOoxml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.insertBreak = function (n, t) {
			f(this.context, this, "InsertBreak", 0, [n, t])
		},
		v.prototype.insertFileFromBase64 = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertFileFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertHtml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertHtml", 0, [t, i], !1, !0))
		},
		v.prototype.insertInlinePictureFromBase64 = function (t, i) {
			return new n.InlinePicture(this.context, u(this.context, this, "InsertInlinePictureFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertOoxml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertOoxml", 0, [t, i], !1, !0))
		},
		v.prototype.insertParagraph = function (t, i) {
			return new n.Paragraph(this.context, u(this.context, this, "InsertParagraph", 0, [t, i], !1, !0))
		},
		v.prototype.insertText = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertText", 0, [t, i], !1, !0))
		},
		v.prototype.search = function (t, i) {
			return i = y(this.context, i),
			new n.SearchResultCollection(this.context, u(this.context, this, "Search", 1, [t, i], !0, !0))
		},
		v.prototype.select = function (n) {
			f(this.context, this, "Select", 1, [n])
		},
		v.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		v.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Appearance) || (this.m_appearance = t.Appearance);
				i(t.CannotDelete) || (this.m_cannotDelete = t.CannotDelete);
				i(t.CannotEdit) || (this.m_cannotEdit = t.CannotEdit);
				i(t.Color) || (this.m_color = t.Color);
				i(t.Id) || (this.m_id = t.Id);
				i(t.PlaceholderText) || (this.m_placeholderText = t.PlaceholderText);
				i(t.RemoveWhenEdited) || (this.m_removeWhenEdited = t.RemoveWhenEdited);
				i(t.Style) || (this.m_style = t.Style);
				i(t.Tag) || (this.m_tag = t.Tag);
				i(t.Text) || (this.m_text = t.Text);
				i(t.Title) || (this.m_title = t.Title);
				i(t.Type) || (this.m_type = t.Type);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "paragraphs", "Paragraphs", "parentContentControl", "ParentContentControl"])
			}
		},
		v.prototype.load = function (n) {
			return s(this, n),
			this
		},
		v.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		v
	}
	(OfficeExtension.ClientObject);
	n.ContentControl = p;
	w = function (r) {
		function e() {
			r.apply(this, arguments)
		}
		return __extends(e, r),
		Object.defineProperty(e.prototype, "items", {
			get : function () {
				return t("items", this.m__items),
				this.m__items
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(e.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		e.prototype.getById = function (t) {
			return new n.ContentControl(this.context, u(this.context, this, "GetById", 1, [t], !1, !1))
		},
		e.prototype.getByTag = function (t) {
			return new n.ContentControlCollection(this.context, u(this.context, this, "GetByTag", 1, [t], !0, !1))
		},
		e.prototype.getByTitle = function (t) {
			return new n.ContentControlCollection(this.context, u(this.context, this, "GetByTitle", 1, [t], !0, !1))
		},
		e.prototype.getItem = function (t) {
			return new n.ContentControl(this.context, a(this.context, this, [t]))
		},
		e.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		e.prototype._handleResult = function (t) {
			var r,
			f,
			u,
			e;
			if (!o(t) && (r = t, h(this, r), i(r._ReferenceId) || (this.m__ReferenceId = r._ReferenceId), !o(r[OfficeExtension.Constants.items])))
				for (this.m__items = [], f = r[OfficeExtension.Constants.items], u = 0; u < f.length; u++)
					e = new n.ContentControl(this.context, v(this.context, this, f[u])), e._handleResult(f[u]), this.m__items.push(e)
		},
		e.prototype.load = function (n) {
			return s(this, n),
			this
		},
		e.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		e
	}
	(OfficeExtension.ClientObject);
	n.ContentControlCollection = w;
	b = function (r) {
		function a() {
			r.apply(this, arguments)
		}
		return __extends(a, r),
		Object.defineProperty(a.prototype, "body", {
			get : function () {
				return this.m_body || (this.m_body = new n.Body(this.context, e(this.context, this, "Body", !1, !1))),
				this.m_body
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(a.prototype, "contentControls", {
			get : function () {
				return this.m_contentControls || (this.m_contentControls = new n.ContentControlCollection(this.context, e(this.context, this, "ContentControls", !0, !1))),
				this.m_contentControls
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(a.prototype, "sections", {
			get : function () {
				return this.m_sections || (this.m_sections = new n.SectionCollection(this.context, e(this.context, this, "Sections", !0, !1))),
				this.m_sections
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(a.prototype, "saved", {
			get : function () {
				return t("saved", this.m_saved),
				this.m_saved
			},
			enumerable : !0,
			configurable : !0
		}),
		a.prototype.getSelection = function () {
			return new n.Range(this.context, u(this.context, this, "GetSelection", 1, [], !1, !0))
		},
		a.prototype.save = function () {
			f(this.context, this, "Save", 0, [])
		},
		a.prototype._GetObjectByReferenceId = function (n) {
			var i = f(this.context, this, "_GetObjectByReferenceId", 1, [n]),
			t = new OfficeExtension.ClientResult;
			return c(this, i, t),
			t
		},
		a.prototype._GetObjectTypeNameByReferenceId = function (n) {
			var i = f(this.context, this, "_GetObjectTypeNameByReferenceId", 1, [n]),
			t = new OfficeExtension.ClientResult;
			return c(this, i, t),
			t
		},
		a.prototype._RemoveAllReferences = function () {
			f(this.context, this, "_RemoveAllReferences", 1, [])
		},
		a.prototype._RemoveReference = function (n) {
			f(this.context, this, "_RemoveReference", 1, [n])
		},
		a.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Saved) || (this.m_saved = t.Saved);
				l(this, t, ["body", "Body", "contentControls", "ContentControls", "sections", "Sections"])
			}
		},
		a.prototype.load = function (n) {
			return s(this, n),
			this
		},
		a
	}
	(OfficeExtension.ClientObject);
	n.Document = b;
	k = function (n) {
		function u() {
			n.apply(this, arguments)
		}
		return __extends(u, n),
		Object.defineProperty(u.prototype, "bold", {
			get : function () {
				return t("bold", this.m_bold),
				this.m_bold
			},
			set : function (n) {
				this.m_bold = n;
				r(this.context, this, "Bold", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "color", {
			get : function () {
				return t("color", this.m_color),
				this.m_color
			},
			set : function (n) {
				this.m_color = n;
				r(this.context, this, "Color", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "doubleStrikeThrough", {
			get : function () {
				return t("doubleStrikeThrough", this.m_doubleStrikeThrough),
				this.m_doubleStrikeThrough
			},
			set : function (n) {
				this.m_doubleStrikeThrough = n;
				r(this.context, this, "DoubleStrikeThrough", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "highlightColor", {
			get : function () {
				return t("highlightColor", this.m_highlightColor),
				this.m_highlightColor
			},
			set : function (n) {
				this.m_highlightColor = n;
				r(this.context, this, "HighlightColor", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "italic", {
			get : function () {
				return t("italic", this.m_italic),
				this.m_italic
			},
			set : function (n) {
				this.m_italic = n;
				r(this.context, this, "Italic", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "name", {
			get : function () {
				return t("name", this.m_name),
				this.m_name
			},
			set : function (n) {
				this.m_name = n;
				r(this.context, this, "Name", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "size", {
			get : function () {
				return t("size", this.m_size),
				this.m_size
			},
			set : function (n) {
				this.m_size = n;
				r(this.context, this, "Size", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "strikeThrough", {
			get : function () {
				return t("strikeThrough", this.m_strikeThrough),
				this.m_strikeThrough
			},
			set : function (n) {
				this.m_strikeThrough = n;
				r(this.context, this, "StrikeThrough", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "subscript", {
			get : function () {
				return t("subscript", this.m_subscript),
				this.m_subscript
			},
			set : function (n) {
				this.m_subscript = n;
				r(this.context, this, "Subscript", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "superscript", {
			get : function () {
				return t("superscript", this.m_superscript),
				this.m_superscript
			},
			set : function (n) {
				this.m_superscript = n;
				r(this.context, this, "Superscript", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "underline", {
			get : function () {
				return t("underline", this.m_underline),
				this.m_underline
			},
			set : function (n) {
				this.m_underline = n;
				r(this.context, this, "Underline", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		u.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		u.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Bold) || (this.m_bold = t.Bold);
				i(t.Color) || (this.m_color = t.Color);
				i(t.DoubleStrikeThrough) || (this.m_doubleStrikeThrough = t.DoubleStrikeThrough);
				i(t.HighlightColor) || (this.m_highlightColor = t.HighlightColor);
				i(t.Italic) || (this.m_italic = t.Italic);
				i(t.Name) || (this.m_name = t.Name);
				i(t.Size) || (this.m_size = t.Size);
				i(t.StrikeThrough) || (this.m_strikeThrough = t.StrikeThrough);
				i(t.Subscript) || (this.m_subscript = t.Subscript);
				i(t.Superscript) || (this.m_superscript = t.Superscript);
				i(t.Underline) || (this.m_underline = t.Underline);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId)
			}
		},
		u.prototype.load = function (n) {
			return s(this, n),
			this
		},
		u.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		u
	}
	(OfficeExtension.ClientObject);
	n.Font = k;
	d = function (a) {
		function v() {
			a.apply(this, arguments)
		}
		return __extends(v, a),
		Object.defineProperty(v.prototype, "paragraph", {
			get : function () {
				return this.m_paragraph || (this.m_paragraph = new n.Paragraph(this.context, e(this.context, this, "Paragraph", !1, !1))),
				this.m_paragraph
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "parentContentControl", {
			get : function () {
				return this.m_parentContentControl || (this.m_parentContentControl = new n.ContentControl(this.context, e(this.context, this, "ParentContentControl", !1, !1))),
				this.m_parentContentControl
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "altTextDescription", {
			get : function () {
				return t("altTextDescription", this.m_altTextDescription),
				this.m_altTextDescription
			},
			set : function (n) {
				this.m_altTextDescription = n;
				r(this.context, this, "AltTextDescription", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "altTextTitle", {
			get : function () {
				return t("altTextTitle", this.m_altTextTitle),
				this.m_altTextTitle
			},
			set : function (n) {
				this.m_altTextTitle = n;
				r(this.context, this, "AltTextTitle", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "height", {
			get : function () {
				return t("height", this.m_height),
				this.m_height
			},
			set : function (n) {
				this.m_height = n;
				r(this.context, this, "Height", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "hyperlink", {
			get : function () {
				return t("hyperlink", this.m_hyperlink),
				this.m_hyperlink
			},
			set : function (n) {
				this.m_hyperlink = n;
				r(this.context, this, "Hyperlink", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "lockAspectRatio", {
			get : function () {
				return t("lockAspectRatio", this.m_lockAspectRatio),
				this.m_lockAspectRatio
			},
			set : function (n) {
				this.m_lockAspectRatio = n;
				r(this.context, this, "LockAspectRatio", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "width", {
			get : function () {
				return t("width", this.m_width),
				this.m_width
			},
			set : function (n) {
				this.m_width = n;
				r(this.context, this, "Width", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_Id", {
			get : function () {
				return t("_Id", this.m__Id),
				this.m__Id
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		v.prototype.delete  = function () {
			f(this.context, this, "Delete", 0, [])
		},
		v.prototype.getBase64ImageSrc = function () {
			var t = f(this.context, this, "GetBase64ImageSrc", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.insertBreak = function (n, t) {
			f(this.context, this, "InsertBreak", 0, [n, t])
		},
		v.prototype.insertContentControl = function () {
			return new n.ContentControl(this.context, u(this.context, this, "InsertContentControl", 0, [], !1, !0))
		},
		v.prototype.insertFileFromBase64 = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertFileFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertHtml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertHtml", 0, [t, i], !1, !0))
		},
		v.prototype.insertInlinePictureFromBase64 = function (t, i) {
			return new n.InlinePicture(this.context, u(this.context, this, "InsertInlinePictureFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertOoxml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertOoxml", 0, [t, i], !1, !0))
		},
		v.prototype.insertParagraph = function (t, i) {
			return new n.Paragraph(this.context, u(this.context, this, "InsertParagraph", 0, [t, i], !1, !0))
		},
		v.prototype.insertText = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertText", 0, [t, i], !1, !0))
		},
		v.prototype.select = function (n) {
			f(this.context, this, "Select", 1, [n])
		},
		v.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		v.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.AltTextDescription) || (this.m_altTextDescription = t.AltTextDescription);
				i(t.AltTextTitle) || (this.m_altTextTitle = t.AltTextTitle);
				i(t.Height) || (this.m_height = t.Height);
				i(t.Hyperlink) || (this.m_hyperlink = t.Hyperlink);
				i(t.LockAspectRatio) || (this.m_lockAspectRatio = t.LockAspectRatio);
				i(t.Width) || (this.m_width = t.Width);
				i(t._Id) || (this.m__Id = t._Id);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["paragraph", "Paragraph", "parentContentControl", "ParentContentControl"])
			}
		},
		v.prototype.load = function (n) {
			return s(this, n),
			this
		},
		v.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		v
	}
	(OfficeExtension.ClientObject);
	n.InlinePicture = d;
	g = function (r) {
		function u() {
			r.apply(this, arguments)
		}
		return __extends(u, r),
		Object.defineProperty(u.prototype, "items", {
			get : function () {
				return t("items", this.m__items),
				this.m__items
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		u.prototype._GetItem = function (t) {
			return new n.InlinePicture(this.context, a(this.context, this, [t]))
		},
		u.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		u.prototype._handleResult = function (t) {
			var r,
			f,
			u,
			e;
			if (!o(t) && (r = t, h(this, r), i(r._ReferenceId) || (this.m__ReferenceId = r._ReferenceId), !o(r[OfficeExtension.Constants.items])))
				for (this.m__items = [], f = r[OfficeExtension.Constants.items], u = 0; u < f.length; u++)
					e = new n.InlinePicture(this.context, v(this.context, this, f[u])), e._handleResult(f[u]), this.m__items.push(e)
		},
		u.prototype.load = function (n) {
			return s(this, n),
			this
		},
		u.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		u
	}
	(OfficeExtension.ClientObject);
	n.InlinePictureCollection = g;
	nt = function (a) {
		function v() {
			a.apply(this, arguments)
		}
		return __extends(v, a),
		Object.defineProperty(v.prototype, "contentControls", {
			get : function () {
				return this.m_contentControls || (this.m_contentControls = new n.ContentControlCollection(this.context, e(this.context, this, "ContentControls", !0, !1))),
				this.m_contentControls
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "font", {
			get : function () {
				return this.m_font || (this.m_font = new n.Font(this.context, e(this.context, this, "Font", !1, !1))),
				this.m_font
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "inlinePictures", {
			get : function () {
				return this.m_inlinePictures || (this.m_inlinePictures = new n.InlinePictureCollection(this.context, e(this.context, this, "InlinePictures", !0, !1))),
				this.m_inlinePictures
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "parentContentControl", {
			get : function () {
				return this.m_parentContentControl || (this.m_parentContentControl = new n.ContentControl(this.context, e(this.context, this, "ParentContentControl", !1, !1))),
				this.m_parentContentControl
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "alignment", {
			get : function () {
				return t("alignment", this.m_alignment),
				this.m_alignment
			},
			set : function (n) {
				this.m_alignment = n;
				r(this.context, this, "Alignment", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "firstLineIndent", {
			get : function () {
				return t("firstLineIndent", this.m_firstLineIndent),
				this.m_firstLineIndent
			},
			set : function (n) {
				this.m_firstLineIndent = n;
				r(this.context, this, "FirstLineIndent", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "leftIndent", {
			get : function () {
				return t("leftIndent", this.m_leftIndent),
				this.m_leftIndent
			},
			set : function (n) {
				this.m_leftIndent = n;
				r(this.context, this, "LeftIndent", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "lineSpacing", {
			get : function () {
				return t("lineSpacing", this.m_lineSpacing),
				this.m_lineSpacing
			},
			set : function (n) {
				this.m_lineSpacing = n;
				r(this.context, this, "LineSpacing", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "lineUnitAfter", {
			get : function () {
				return t("lineUnitAfter", this.m_lineUnitAfter),
				this.m_lineUnitAfter
			},
			set : function (n) {
				this.m_lineUnitAfter = n;
				r(this.context, this, "LineUnitAfter", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "lineUnitBefore", {
			get : function () {
				return t("lineUnitBefore", this.m_lineUnitBefore),
				this.m_lineUnitBefore
			},
			set : function (n) {
				this.m_lineUnitBefore = n;
				r(this.context, this, "LineUnitBefore", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "outlineLevel", {
			get : function () {
				return t("outlineLevel", this.m_outlineLevel),
				this.m_outlineLevel
			},
			set : function (n) {
				this.m_outlineLevel = n;
				r(this.context, this, "OutlineLevel", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "rightIndent", {
			get : function () {
				return t("rightIndent", this.m_rightIndent),
				this.m_rightIndent
			},
			set : function (n) {
				this.m_rightIndent = n;
				r(this.context, this, "RightIndent", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "spaceAfter", {
			get : function () {
				return t("spaceAfter", this.m_spaceAfter),
				this.m_spaceAfter
			},
			set : function (n) {
				this.m_spaceAfter = n;
				r(this.context, this, "SpaceAfter", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "spaceBefore", {
			get : function () {
				return t("spaceBefore", this.m_spaceBefore),
				this.m_spaceBefore
			},
			set : function (n) {
				this.m_spaceBefore = n;
				r(this.context, this, "SpaceBefore", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "style", {
			get : function () {
				return t("style", this.m_style),
				this.m_style
			},
			set : function (n) {
				this.m_style = n;
				r(this.context, this, "Style", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "text", {
			get : function () {
				return t("text", this.m_text),
				this.m_text
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_Id", {
			get : function () {
				return t("_Id", this.m__Id),
				this.m__Id
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		v.prototype.clear = function () {
			f(this.context, this, "Clear", 0, [])
		},
		v.prototype.delete  = function () {
			f(this.context, this, "Delete", 0, [])
		},
		v.prototype.getHtml = function () {
			var t = f(this.context, this, "GetHtml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.getOoxml = function () {
			var t = f(this.context, this, "GetOoxml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.insertBreak = function (n, t) {
			f(this.context, this, "InsertBreak", 0, [n, t])
		},
		v.prototype.insertContentControl = function () {
			return new n.ContentControl(this.context, u(this.context, this, "InsertContentControl", 0, [], !1, !0))
		},
		v.prototype.insertFileFromBase64 = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertFileFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertHtml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertHtml", 0, [t, i], !1, !0))
		},
		v.prototype.insertInlinePictureFromBase64 = function (t, i) {
			return new n.InlinePicture(this.context, u(this.context, this, "InsertInlinePictureFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertOoxml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertOoxml", 0, [t, i], !1, !0))
		},
		v.prototype.insertParagraph = function (t, i) {
			return new n.Paragraph(this.context, u(this.context, this, "InsertParagraph", 0, [t, i], !1, !0))
		},
		v.prototype.insertText = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertText", 0, [t, i], !1, !0))
		},
		v.prototype.search = function (t, i) {
			return i = y(this.context, i),
			new n.SearchResultCollection(this.context, u(this.context, this, "Search", 1, [t, i], !0, !0))
		},
		v.prototype.select = function (n) {
			f(this.context, this, "Select", 1, [n])
		},
		v.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		v.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Alignment) || (this.m_alignment = t.Alignment);
				i(t.FirstLineIndent) || (this.m_firstLineIndent = t.FirstLineIndent);
				i(t.LeftIndent) || (this.m_leftIndent = t.LeftIndent);
				i(t.LineSpacing) || (this.m_lineSpacing = t.LineSpacing);
				i(t.LineUnitAfter) || (this.m_lineUnitAfter = t.LineUnitAfter);
				i(t.LineUnitBefore) || (this.m_lineUnitBefore = t.LineUnitBefore);
				i(t.OutlineLevel) || (this.m_outlineLevel = t.OutlineLevel);
				i(t.RightIndent) || (this.m_rightIndent = t.RightIndent);
				i(t.SpaceAfter) || (this.m_spaceAfter = t.SpaceAfter);
				i(t.SpaceBefore) || (this.m_spaceBefore = t.SpaceBefore);
				i(t.Style) || (this.m_style = t.Style);
				i(t.Text) || (this.m_text = t.Text);
				i(t._Id) || (this.m__Id = t._Id);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "parentContentControl", "ParentContentControl"])
			}
		},
		v.prototype.load = function (n) {
			return s(this, n),
			this
		},
		v.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		v
	}
	(OfficeExtension.ClientObject);
	n.Paragraph = nt;
	tt = function (r) {
		function u() {
			r.apply(this, arguments)
		}
		return __extends(u, r),
		Object.defineProperty(u.prototype, "items", {
			get : function () {
				return t("items", this.m__items),
				this.m__items
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		u.prototype._GetItem = function (t) {
			return new n.Paragraph(this.context, a(this.context, this, [t]))
		},
		u.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		u.prototype._handleResult = function (t) {
			var r,
			f,
			u,
			e;
			if (!o(t) && (r = t, h(this, r), i(r._ReferenceId) || (this.m__ReferenceId = r._ReferenceId), !o(r[OfficeExtension.Constants.items])))
				for (this.m__items = [], f = r[OfficeExtension.Constants.items], u = 0; u < f.length; u++)
					e = new n.Paragraph(this.context, v(this.context, this, f[u])), e._handleResult(f[u]), this.m__items.push(e)
		},
		u.prototype.load = function (n) {
			return s(this, n),
			this
		},
		u.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		u
	}
	(OfficeExtension.ClientObject);
	n.ParagraphCollection = tt;
	it = function (a) {
		function v() {
			a.apply(this, arguments)
		}
		return __extends(v, a),
		Object.defineProperty(v.prototype, "contentControls", {
			get : function () {
				return this.m_contentControls || (this.m_contentControls = new n.ContentControlCollection(this.context, e(this.context, this, "ContentControls", !0, !1))),
				this.m_contentControls
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "font", {
			get : function () {
				return this.m_font || (this.m_font = new n.Font(this.context, e(this.context, this, "Font", !1, !1))),
				this.m_font
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "inlinePictures", {
			get : function () {
				return this.m_inlinePictures || (this.m_inlinePictures = new n.InlinePictureCollection(this.context, e(this.context, this, "InlinePictures", !0, !1))),
				this.m_inlinePictures
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "paragraphs", {
			get : function () {
				return this.m_paragraphs || (this.m_paragraphs = new n.ParagraphCollection(this.context, e(this.context, this, "Paragraphs", !0, !1))),
				this.m_paragraphs
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "parentContentControl", {
			get : function () {
				return this.m_parentContentControl || (this.m_parentContentControl = new n.ContentControl(this.context, e(this.context, this, "ParentContentControl", !1, !1))),
				this.m_parentContentControl
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "style", {
			get : function () {
				return t("style", this.m_style),
				this.m_style
			},
			set : function (n) {
				this.m_style = n;
				r(this.context, this, "Style", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "text", {
			get : function () {
				return t("text", this.m_text),
				this.m_text
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_Id", {
			get : function () {
				return t("_Id", this.m__Id),
				this.m__Id
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(v.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		v.prototype.clear = function () {
			f(this.context, this, "Clear", 0, [])
		},
		v.prototype.delete  = function () {
			f(this.context, this, "Delete", 0, [])
		},
		v.prototype.getHtml = function () {
			var t = f(this.context, this, "GetHtml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.getOoxml = function () {
			var t = f(this.context, this, "GetOoxml", 1, []),
			n = new OfficeExtension.ClientResult;
			return c(this, t, n),
			n
		},
		v.prototype.insertBreak = function (n, t) {
			f(this.context, this, "InsertBreak", 0, [n, t])
		},
		v.prototype.insertContentControl = function () {
			return new n.ContentControl(this.context, u(this.context, this, "InsertContentControl", 0, [], !1, !0))
		},
		v.prototype.insertFileFromBase64 = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertFileFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertHtml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertHtml", 0, [t, i], !1, !0))
		},
		v.prototype.insertInlinePictureFromBase64 = function (t, i) {
			return new n.InlinePicture(this.context, u(this.context, this, "InsertInlinePictureFromBase64", 0, [t, i], !1, !0))
		},
		v.prototype.insertOoxml = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertOoxml", 0, [t, i], !1, !0))
		},
		v.prototype.insertParagraph = function (t, i) {
			return new n.Paragraph(this.context, u(this.context, this, "InsertParagraph", 0, [t, i], !1, !0))
		},
		v.prototype.insertText = function (t, i) {
			return new n.Range(this.context, u(this.context, this, "InsertText", 0, [t, i], !1, !0))
		},
		v.prototype.search = function (t, i) {
			return i = y(this.context, i),
			new n.SearchResultCollection(this.context, u(this.context, this, "Search", 1, [t, i], !0, !0))
		},
		v.prototype.select = function (n) {
			f(this.context, this, "Select", 1, [n])
		},
		v.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		v.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.Style) || (this.m_style = t.Style);
				i(t.Text) || (this.m_text = t.Text);
				i(t._Id) || (this.m__Id = t._Id);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "paragraphs", "Paragraphs", "parentContentControl", "ParentContentControl"])
			}
		},
		v.prototype.load = function (n) {
			return s(this, n),
			this
		},
		v.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		v
	}
	(OfficeExtension.ClientObject);
	n.Range = it;
	rt = function (u) {
		function f() {
			u.apply(this, arguments)
		}
		return __extends(f, u),
		Object.defineProperty(f.prototype, "matchWildCards", {
			get : function () {
				return t("matchWildcards", this.m_matchWildcards),
				this.m_matchWildcards
			},
			set : function (n) {
				this.m_matchWildcards = n;
				r(this.context, this, "MatchWildcards", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "ignorePunct", {
			get : function () {
				return t("ignorePunct", this.m_ignorePunct),
				this.m_ignorePunct
			},
			set : function (n) {
				this.m_ignorePunct = n;
				r(this.context, this, "IgnorePunct", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "ignoreSpace", {
			get : function () {
				return t("ignoreSpace", this.m_ignoreSpace),
				this.m_ignoreSpace
			},
			set : function (n) {
				this.m_ignoreSpace = n;
				r(this.context, this, "IgnoreSpace", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchCase", {
			get : function () {
				return t("matchCase", this.m_matchCase),
				this.m_matchCase
			},
			set : function (n) {
				this.m_matchCase = n;
				r(this.context, this, "MatchCase", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchPrefix", {
			get : function () {
				return t("matchPrefix", this.m_matchPrefix),
				this.m_matchPrefix
			},
			set : function (n) {
				this.m_matchPrefix = n;
				r(this.context, this, "MatchPrefix", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchSoundsLike", {
			get : function () {
				return t("matchSoundsLike", this.m_matchSoundsLike),
				this.m_matchSoundsLike
			},
			set : function (n) {
				this.m_matchSoundsLike = n;
				r(this.context, this, "MatchSoundsLike", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchSuffix", {
			get : function () {
				return t("matchSuffix", this.m_matchSuffix),
				this.m_matchSuffix
			},
			set : function (n) {
				this.m_matchSuffix = n;
				r(this.context, this, "MatchSuffix", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchWholeWord", {
			get : function () {
				return t("matchWholeWord", this.m_matchWholeWord),
				this.m_matchWholeWord
			},
			set : function (n) {
				this.m_matchWholeWord = n;
				r(this.context, this, "MatchWholeWord", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(f.prototype, "matchWildcards", {
			get : function () {
				return t("matchWildcards", this.m_matchWildcards),
				this.m_matchWildcards
			},
			set : function (n) {
				this.m_matchWildcards = n;
				r(this.context, this, "MatchWildcards", n)
			},
			enumerable : !0,
			configurable : !0
		}),
		f.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t.IgnorePunct) || (this.m_ignorePunct = t.IgnorePunct);
				i(t.IgnoreSpace) || (this.m_ignoreSpace = t.IgnoreSpace);
				i(t.MatchCase) || (this.m_matchCase = t.MatchCase);
				i(t.MatchPrefix) || (this.m_matchPrefix = t.MatchPrefix);
				i(t.MatchSoundsLike) || (this.m_matchSoundsLike = t.MatchSoundsLike);
				i(t.MatchSuffix) || (this.m_matchSuffix = t.MatchSuffix);
				i(t.MatchWholeWord) || (this.m_matchWholeWord = t.MatchWholeWord);
				i(t.MatchWildcards) || (this.m_matchWildcards = t.MatchWildcards)
			}
		},
		f.prototype.load = function (n) {
			return s(this, n),
			this
		},
		f.newObject = function (t) {
			return new n.SearchOptions(t, ot(t, "Microsoft.WordServices.SearchOptions", !1))
		},
		f
	}
	(OfficeExtension.ClientObject);
	n.SearchOptions = rt;
	ut = function (r) {
		function u() {
			r.apply(this, arguments)
		}
		return __extends(u, r),
		Object.defineProperty(u.prototype, "items", {
			get : function () {
				return t("items", this.m__items),
				this.m__items
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		u.prototype._GetItem = function (t) {
			return new n.Range(this.context, a(this.context, this, [t]))
		},
		u.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		u.prototype._handleResult = function (t) {
			var r,
			f,
			u,
			e;
			if (!o(t) && (r = t, h(this, r), i(r._ReferenceId) || (this.m__ReferenceId = r._ReferenceId), !o(r[OfficeExtension.Constants.items])))
				for (this.m__items = [], f = r[OfficeExtension.Constants.items], u = 0; u < f.length; u++)
					e = new n.Range(this.context, v(this.context, this, f[u])), e._handleResult(f[u]), this.m__items.push(e)
		},
		u.prototype.load = function (n) {
			return s(this, n),
			this
		},
		u.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		u
	}
	(OfficeExtension.ClientObject);
	n.SearchResultCollection = ut;
	ft = function (r) {
		function c() {
			r.apply(this, arguments)
		}
		return __extends(c, r),
		Object.defineProperty(c.prototype, "body", {
			get : function () {
				return this.m_body || (this.m_body = new n.Body(this.context, e(this.context, this, "Body", !1, !1))),
				this.m_body
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(c.prototype, "_Id", {
			get : function () {
				return t("_Id", this.m__Id),
				this.m__Id
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(c.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		c.prototype.getFooter = function (t) {
			return new n.Body(this.context, u(this.context, this, "GetFooter", 1, [t], !1, !0))
		},
		c.prototype.getHeader = function (t) {
			return new n.Body(this.context, u(this.context, this, "GetHeader", 1, [t], !1, !0))
		},
		c.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		c.prototype._handleResult = function (n) {
			if (!o(n)) {
				var t = n;
				h(this, t);
				i(t._Id) || (this.m__Id = t._Id);
				i(t._ReferenceId) || (this.m__ReferenceId = t._ReferenceId);
				l(this, t, ["body", "Body"])
			}
		},
		c.prototype.load = function (n) {
			return s(this, n),
			this
		},
		c.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		c
	}
	(OfficeExtension.ClientObject);
	n.Section = ft;
	et = function (r) {
		function u() {
			r.apply(this, arguments)
		}
		return __extends(u, r),
		Object.defineProperty(u.prototype, "items", {
			get : function () {
				return t("items", this.m__items),
				this.m__items
			},
			enumerable : !0,
			configurable : !0
		}),
		Object.defineProperty(u.prototype, "_ReferenceId", {
			get : function () {
				return t("_ReferenceId", this.m__ReferenceId),
				this.m__ReferenceId
			},
			enumerable : !0,
			configurable : !0
		}),
		u.prototype._GetItem = function (t) {
			return new n.Section(this.context, a(this.context, this, [t]))
		},
		u.prototype._KeepReference = function () {
			f(this.context, this, "_KeepReference", 1, [])
		},
		u.prototype._handleResult = function (t) {
			var r,
			f,
			u,
			e;
			if (!o(t) && (r = t, h(this, r), i(r._ReferenceId) || (this.m__ReferenceId = r._ReferenceId), !o(r[OfficeExtension.Constants.items])))
				for (this.m__items = [], f = r[OfficeExtension.Constants.items], u = 0; u < f.length; u++)
					e = new n.Section(this.context, v(this.context, this, f[u])), e._handleResult(f[u]), this.m__items.push(e)
		},
		u.prototype.load = function (n) {
			return s(this, n),
			this
		},
		u.prototype._initReferenceId = function (n) {
			this.m__ReferenceId = n
		},
		u
	}
	(OfficeExtension.ClientObject);
	n.SectionCollection = et,
	function (n) {
		n.richText = "RichText"
	}
	(ht = n.ContentControlType || (n.ContentControlType = {})),
	function (n) {
		n.boundingBox = "BoundingBox";
		n.tags = "Tags";
		n.hidden = "Hidden"
	}
	(ct = n.ContentControlAppearance || (n.ContentControlAppearance = {})),
	function (n) {
		n.none = "None";
		n.single = "Single";
		n.word = "Word";
		n.double = "Double";
		n.dotted = "Dotted";
		n.hidden = "Hidden";
		n.thick = "Thick";
		n.dashLine = "DashLine";
		n.dotLine = "DotLine";
		n.dotDashLine = "DotDashLine";
		n.twoDotDashLine = "TwoDotDashLine";
		n.wave = "Wave"
	}
	(lt = n.UnderlineType || (n.UnderlineType = {})),
	function (n) {
		n.page = "Page";
		n.column = "Column";
		n.next = "Next";
		n.sectionContinuous = "SectionContinuous";
		n.sectionEven = "SectionEven";
		n.sectionOdd = "SectionOdd";
		n.line = "Line";
		n.lineClearLeft = "LineClearLeft";
		n.lineClearRight = "LineClearRight";
		n.textWrapping = "TextWrapping"
	}
	(at = n.BreakType || (n.BreakType = {})),
	function (n) {
		n.before = "Before";
		n.after = "After";
		n.start = "Start";
		n.end = "End";
		n.replace = "Replace"
	}
	(vt = n.InsertLocation || (n.InsertLocation = {})),
	function (n) {
		n.unknown = "Unknown";
		n.left = "Left";
		n.centered = "Centered";
		n.right = "Right";
		n.justified = "Justified"
	}
	(yt = n.Alignment || (n.Alignment = {})),
	function (n) {
		n.primary = "Primary";
		n.firstPage = "FirstPage";
		n.evenPages = "EvenPages"
	}
	(pt = n.HeaderFooterType || (n.HeaderFooterType = {})),
	function (n) {
		n.select = "Select";
		n.start = "Start";
		n.end = "End"
	}
	(wt = n.SelectionMode || (n.SelectionMode = {})),
	function (n) {
		n.accessDenied = "AccessDenied";
		n.generalException = "GeneralException";
		n.invalidArgument = "InvalidArgument";
		n.itemNotFound = "ItemNotFound";
		n.notImplemented = "NotImplemented"
	}
	(bt = n.ErrorCodes || (n.ErrorCodes = {}))
}
(Word || (Word = {})), function (n) {
	function i(t) {
		return OfficeExtension.ClientRequestContext._run(function () {
			return new n.RequestContext
		}, t)
	}
	var t = function (t) {
		function i(i) {
			t.call(this, i);
			this.m_document = new n.Document(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
			this._rootObject = this.m_document
		}
		return __extends(i, t),
		Object.defineProperty(i.prototype, "document", {
			get : function () {
				return this.m_document
			},
			enumerable : !0,
			configurable : !0
		}),
		i
	}
	(OfficeExtension.ClientRequestContext);
	n.RequestContext = t;
	n.run = i
}
(Word || (Word = {}))