// jscriptowork bundled launcher
// Generated: 2026-08-03 17:13
//
// All libs are inlined. No separate libs/ folder is needed.
// open_hta() injects lib source inline into generated HTAs.
//
// Usage:  cscript launcher.js <yourscript.js>
//         launcher.bat <yourscript.js>

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

var _script = WScript;
function log(message) { _script.echo(message); }

var CURRENT_PATH   = _script.ScriptFullName;
var CURRENT_FOLDER = CURRENT_PATH.slice(0, CURRENT_PATH.lastIndexOf('\\') + 1);
var ROOT_FOLDER    = CURRENT_FOLDER;

// read_all_text_file: referenced by system.js (load_working_directory).
function read_all_text_file(path) {
    var _fso = new ActiveXObject('Scripting.FileSystemObject');
    try {
        var _f = _fso.OpenTextFile(path, 1);
        if (_f.AtEndOfStream) { _f.Close(); return ''; }
        var _s = _f.ReadAll(); _f.Close(); return _s;
    } catch(e) { return null; }
}

// load() is a no-op: all libs are already inlined below.
function load(modulename) {}

// Source of core + polyfills + system, pre-escaped for inline HTA injection.
// ui.js checks typeof _jsw_hta_inline_libs to decide whether to use this
// or fall back to <script src="file:///..."> references.
var _jsw_hta_inline_libs = '//modding jscript engine\nif (!Array.prototype.indexOf) {\n    \n    Array.prototype.indexOf = function(obj, start) {\n        var j = this.length;\n        var n = start || 0;\n        for (var i = n >= 0 ? n : Math.max(j + n, 0); i < j; i++) {\n            if (this[i] === obj) { return i; }\n        }\n        return -1;\n    }\n\n}\n\nif (!String.prototype.trim) {\n    (function() {\n        // Make sure we trim BOM and NBSP\n        var rtrim = /^[\\s\\uFEFF\\xA0]+|[\\s\\uFEFF\\xA0]+$/g;\n        String.prototype.trim = function() {\n            return this.replace(rtrim, \'\');\n        };\n    })();\n}\n\nif (!String.prototype.startsWith) {\n    String.prototype.startsWith = function(searchString, position){\n      position = position || 0;\n      return this.substr(position, searchString.length) === searchString;\n  };\n}\n\nif (!Array.prototype.filter) {\n  Array.prototype.filter = function(fun/*, thisArg*/) {\n    \'use strict\';\n\n    if (this === void 0 || this === null) {\n      throw new TypeError();\n    }\n\n    var t = Object(this);\n    var len = t.length >>> 0;\n    if (typeof fun !== \'function\') {\n      throw new TypeError();\n    }\n\n    var res = [];\n    var thisArg = arguments.length >= 2 ? arguments[1] : void 0;\n    for (var i = 0; i < len; i++) {\n      if (i in t) {\n        var val = t[i];\n\n        // NOTE: Technically this should Object.defineProperty at\n        //       the next index, as push can be affected by\n        //       properties on Object.prototype and Array.prototype.\n        //       But that method\'s new, and collisions should be\n        //       rare, so use the more-compatible alternative.\n        if (fun.call(thisArg, val, i, t)) {\n          res.push(val);\n        }\n      }\n    }\n\n    return res;\n  };\n}\n\n// Production steps of ECMA-262, Edition 5, 15.4.4.19\n// Reference: http://es5.github.io/#x15.4.4.19\nif (!Array.prototype.map) {\n\n  Array.prototype.map = function(callback, thisArg) {\n\n    var T, A, k;\n\n    if (this == null) {\n      throw new TypeError(\' this is null or not defined\');\n    }\n\n    // 1. Let O be the result of calling ToObject passing the |this| \n    //    value as the argument.\n    var O = Object(this);\n\n    // 2. Let lenValue be the result of calling the Get internal \n    //    method of O with the argument "length".\n    // 3. Let len be ToUint32(lenValue).\n    var len = O.length >>> 0;\n\n    // 4. If IsCallable(callback) is false, throw a TypeError exception.\n    // See: http://es5.github.com/#x9.11\n    if (typeof callback !== \'function\') {\n      throw new TypeError(callback + \' is not a function\');\n    }\n\n    // 5. If thisArg was supplied, let T be thisArg; else let T be undefined.\n    if (arguments.length > 1) {\n      T = thisArg;\n    }\n\n    // 6. Let A be a new array created as if by the expression new Array(len) \n    //    where Array is the standard built-in constructor with that name and \n    //    len is the value of len.\n    A = new Array(len);\n\n    // 7. Let k be 0\n    k = 0;\n\n    // 8. Repeat, while k < len\n    while (k < len) {\n\n      var kValue, mappedValue;\n\n      // a. Let Pk be ToString(k).\n      //   This is implicit for LHS operands of the in operator\n      // b. Let kPresent be the result of calling the HasProperty internal \n      //    method of O with argument Pk.\n      //   This step can be combined with c\n      // c. If kPresent is true, then\n      if (k in O) {\n\n        // i. Let kValue be the result of calling the Get internal \n        //    method of O with argument Pk.\n        kValue = O[k];\n\n        // ii. Let mappedValue be the result of calling the Call internal \n        //     method of callback with T as the this value and argument \n        //     list containing kValue, k, and O.\n        mappedValue = callback.call(T, kValue, k, O);\n\n        // iii. Call the DefineOwnProperty internal method of A with arguments\n        // Pk, Property Descriptor\n        // { Value: mappedValue,\n        //   Writable: true,\n        //   Enumerable: true,\n        //   Configurable: true },\n        // and false.\n\n        // In browsers that support Object.defineProperty, use the following:\n        // Object.defineProperty(A, k, {\n        //   value: mappedValue,\n        //   writable: true,\n        //   enumerable: true,\n        //   configurable: true\n        // });\n\n        // For best browser support, use the following:\n        A[k] = mappedValue;\n      }\n      // d. Increase k by 1.\n      k++;\n    }\n\n    // 9. return A\n    return A;\n  };\n}\n\n// Production steps of ECMA-262, Edition 5, 15.4.4.18\n// Reference: http://es5.github.io/#x15.4.4.18\nif (!Array.prototype.forEach) {\n\n  Array.prototype.forEach = function(callback, thisArg) {\n\n    var T, k;\n\n    if (this == null) {\n      throw new TypeError(\' this is null or not defined\');\n    }\n\n    // 1. Let O be the result of calling ToObject passing the |this| value as the argument.\n    var O = Object(this);\n\n    // 2. Let lenValue be the result of calling the Get internal method of O with the argument "length".\n    // 3. Let len be ToUint32(lenValue).\n    var len = O.length >>> 0;\n\n    // 4. If IsCallable(callback) is false, throw a TypeError exception.\n    // See: http://es5.github.com/#x9.11\n    if (typeof callback !== "function") {\n      throw new TypeError(callback + \' is not a function\');\n    }\n\n    // 5. If thisArg was supplied, let T be thisArg; else let T be undefined.\n    if (arguments.length > 1) {\n      T = thisArg;\n    }\n\n    // 6. Let k be 0\n    k = 0;\n\n    // 7. Repeat, while k < len\n    while (k < len) {\n\n      var kValue;\n\n      // a. Let Pk be ToString(k).\n      //   This is implicit for LHS operands of the in operator\n      // b. Let kPresent be the result of calling the HasProperty internal method of O with argument Pk.\n      //   This step can be combined with c\n      // c. If kPresent is true, then\n      if (k in O) {\n\n        // i. Let kValue be the result of calling the Get internal method of O with argument Pk.\n        kValue = O[k];\n\n        // ii. Call the Call internal method of callback with T as the this value and\n        // argument list containing kValue, k, and O.\n        callback.call(T, kValue, k, O);\n      }\n      // d. Increase k by 1.\n      k++;\n    }\n    // 8. return undefined\n  };\n}\n\nif (!Array.prototype.find) {\n  Array.prototype.find = function(predicate) {\n    if (this === null) {\n      throw new TypeError(\'Array.prototype.find called on null or undefined\');\n    }\n    if (typeof predicate !== \'function\') {\n      throw new TypeError(\'predicate must be a function\');\n    }\n    var list = Object(this);\n    var length = list.length >>> 0;\n    var thisArg = arguments[1];\n    var value;\n\n    for (var i = 0; i < length; i++) {\n      value = list[i];\n      if (predicate.call(thisArg, value, i, list)) {\n        return value;\n      }\n    }\n    return undefined;\n  };\n}\n\n// Production steps of ECMA-262, Edition 5, 15.4.4.21\n// Reference: http://es5.github.io/#x15.4.4.21\nif (!Array.prototype.reduce) {\n  Array.prototype.reduce = function(callback /*, initialValue*/) {\n    \'use strict\';\n    if (this == null) {\n      throw new TypeError(\'Array.prototype.reduce called on null or undefined\');\n    }\n    if (typeof callback !== \'function\') {\n      throw new TypeError(callback + \' is not a function\');\n    }\n    var t = Object(this), len = t.length >>> 0, k = 0, value;\n    if (arguments.length == 2) {\n      value = arguments[1];\n    } else {\n      while (k < len && !(k in t)) {\n        k++; \n      }\n      if (k >= len) {\n        throw new TypeError(\'Reduce of empty array with no initial value\');\n      }\n      value = t[k++];\n    }\n    for (; k < len; k++) {\n      if (k in t) {\n        value = callback(value, t[k], k, t);\n      }\n    }\n    return value;\n  };\n}\n\n\n\n\n\n/*\n    http://www.JSON.org/json2.js\n    2010-03-20\n\n    Public Domain.\n\n    NO WARRANTY EXPRESSED OR IMPLIED. USE AT YOUR OWN RISK.\n\n    See http://www.JSON.org/js.html\n\n\n    This code should be minified before deployment.\n    See http://javascript.crockford.com/jsmin.html\n\n    USE YOUR OWN COPY. IT IS EXTREMELY UNWISE TO LOAD CODE FROM SERVERS YOU DO\n    NOT CONTROL.\n\n\n    This file creates a global JSON object containing two methods: stringify\n    and parse.\n\n        JSON.stringify(value, replacer, space)\n            value       any JavaScript value, usually an object or array.\n\n            replacer    an optional parameter that determines how object\n                        values are stringified for objects. It can be a\n                        function or an array of strings.\n\n            space       an optional parameter that specifies the indentation\n                        of nested structures. If it is omitted, the text will\n                        be packed without extra whitespace. If it is a number,\n                        it will specify the number of spaces to indent at each\n                        level. If it is a string (such as \'\\t\' or \'&nbsp;\'),\n                        it contains the characters used to indent at each level.\n\n            This method produces a JSON text from a JavaScript value.\n\n            When an object value is found, if the object contains a toJSON\n            method, its toJSON method will be called and the result will be\n            stringified. A toJSON method does not serialize: it returns the\n            value represented by the name/value pair that should be serialized,\n            or undefined if nothing should be serialized. The toJSON method\n            will be passed the key associated with the value, and this will be\n            bound to the value\n\n            For example, this would serialize Dates as ISO strings.\n\n                Date.prototype.toJSON = function (key) {\n                    function f(n) {\n                        // Format integers to have at least two digits.\n                        return n < 10 ? \'0\' + n : n;\n                    }\n\n                    return this.getUTCFullYear()   + \'-\' +\n                         f(this.getUTCMonth() + 1) + \'-\' +\n                         f(this.getUTCDate())      + \'T\' +\n                         f(this.getUTCHours())     + \':\' +\n                         f(this.getUTCMinutes())   + \':\' +\n                         f(this.getUTCSeconds())   + \'Z\';\n                };\n\n            You can provide an optional replacer method. It will be passed the\n            key and value of each member, with this bound to the containing\n            object. The value that is returned from your method will be\n            serialized. If your method returns undefined, then the member will\n            be excluded from the serialization.\n\n            If the replacer parameter is an array of strings, then it will be\n            used to select the members to be serialized. It filters the results\n            such that only members with keys listed in the replacer array are\n            stringified.\n\n            Values that do not have JSON representations, such as undefined or\n            functions, will not be serialized. Such values in objects will be\n            dropped; in arrays they will be replaced with null. You can use\n            a replacer function to replace those with JSON values.\n            JSON.stringify(undefined) returns undefined.\n\n            The optional space parameter produces a stringification of the\n            value that is filled with line breaks and indentation to make it\n            easier to read.\n\n            If the space parameter is a non-empty string, then that string will\n            be used for indentation. If the space parameter is a number, then\n            the indentation will be that many spaces.\n\n            Example:\n\n            text = JSON.stringify([\'e\', {pluribus: \'unum\'}]);\n            // text is \'["e",{"pluribus":"unum"}]\'\n\n\n            text = JSON.stringify([\'e\', {pluribus: \'unum\'}], null, \'\\t\');\n            // text is \'[\\n\\t"e",\\n\\t{\\n\\t\\t"pluribus": "unum"\\n\\t}\\n]\'\n\n            text = JSON.stringify([new Date()], function (key, value) {\n                return this[key] instanceof Date ?\n                    \'Date(\' + this[key] + \')\' : value;\n            });\n            // text is \'["Date(---current time---)"]\'\n\n\n        JSON.parse(text, reviver)\n            This method parses a JSON text to produce an object or array.\n            It can throw a SyntaxError exception.\n\n            The optional reviver parameter is a function that can filter and\n            transform the results. It receives each of the keys and values,\n            and its return value is used instead of the original value.\n            If it returns what it received, then the structure is not modified.\n            If it returns undefined then the member is deleted.\n\n            Example:\n\n            // Parse the text. Values that look like ISO date strings will\n            // be converted to Date objects.\n\n            myData = JSON.parse(text, function (key, value) {\n                var a;\n                if (typeof value === \'string\') {\n                    a =\n/^(\\d{4})-(\\d{2})-(\\d{2})T(\\d{2}):(\\d{2}):(\\d{2}(?:\\.\\d*)?)Z$/.exec(value);\n                    if (a) {\n                        return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4],\n                            +a[5], +a[6]));\n                    }\n                }\n                return value;\n            });\n\n            myData = JSON.parse(\'["Date(09/09/2001)"]\', function (key, value) {\n                var d;\n                if (typeof value === \'string\' &&\n                        value.slice(0, 5) === \'Date(\' &&\n                        value.slice(-1) === \')\') {\n                    d = new Date(value.slice(5, -1));\n                    if (d) {\n                        return d;\n                    }\n                }\n                return value;\n            });\n\n\n    This is a reference implementation. You are free to copy, modify, or\n    redistribute.\n*/\n\n/*jslint evil: true, strict: false */\n\n/*members "", "\\b", "\\t", "\\n", "\\f", "\\r", "\\"", JSON, "\\\\", apply,\n    call, charCodeAt, getUTCDate, getUTCFullYear, getUTCHours,\n    getUTCMinutes, getUTCMonth, getUTCSeconds, hasOwnProperty, join,\n    lastIndex, length, parse, prototype, push, replace, slice, stringify,\n    test, toJSON, toString, valueOf\n*/\n\n\n// Create a JSON object only if one does not already exist. We create the\n// methods in a closure to avoid creating global variables.\n\nif (!this.JSON) {\n    this.JSON = {};\n}\n\n(function () {\n\n    function f(n) {\n        // Format integers to have at least two digits.\n        return n < 10 ? \'0\' + n : n;\n    }\n\n    if (typeof Date.prototype.toJSON !== \'function\') {\n\n        Date.prototype.toJSON = function (key) {\n\n            return isFinite(this.valueOf()) ?\n                   this.getUTCFullYear()   + \'-\' +\n                 f(this.getUTCMonth() + 1) + \'-\' +\n                 f(this.getUTCDate())      + \'T\' +\n                 f(this.getUTCHours())     + \':\' +\n                 f(this.getUTCMinutes())   + \':\' +\n                 f(this.getUTCSeconds())   + \'Z\' : null;\n        };\n\n        String.prototype.toJSON =\n        Number.prototype.toJSON =\n        Boolean.prototype.toJSON = function (key) {\n            return this.valueOf();\n        };\n    }\n\n    var cx = /[\\u0000\\u00ad\\u0600-\\u0604\\u070f\\u17b4\\u17b5\\u200c-\\u200f\\u2028-\\u202f\\u2060-\\u206f\\ufeff\\ufff0-\\uffff]/g,\n        escapable = /[\\\\\\"\\x00-\\x1f\\x7f-\\x9f\\u00ad\\u0600-\\u0604\\u070f\\u17b4\\u17b5\\u200c-\\u200f\\u2028-\\u202f\\u2060-\\u206f\\ufeff\\ufff0-\\uffff]/g,\n        gap,\n        indent,\n        meta = {    // table of character substitutions\n            \'\\b\': \'\\\\b\',\n            \'\\t\': \'\\\\t\',\n            \'\\n\': \'\\\\n\',\n            \'\\f\': \'\\\\f\',\n            \'\\r\': \'\\\\r\',\n            \'"\' : \'\\\\"\',\n            \'\\\\\': \'\\\\\\\\\'\n        },\n        rep;\n\n\n    function quote(string) {\n\n// If the string contains no control characters, no quote characters, and no\n// backslash characters, then we can safely slap some quotes around it.\n// Otherwise we must also replace the offending characters with safe escape\n// sequences.\n\n        escapable.lastIndex = 0;\n        return escapable.test(string) ?\n            \'"\' + string.replace(escapable, function (a) {\n                var c = meta[a];\n                return typeof c === \'string\' ? c :\n                    \'\\\\u\' + (\'0000\' + a.charCodeAt(0).toString(16)).slice(-4);\n            }) + \'"\' :\n            \'"\' + string + \'"\';\n    }\n\n\n    function str(key, holder) {\n\n// Produce a string from holder[key].\n\n        var i,          // The loop counter.\n            k,          // The member key.\n            v,          // The member value.\n            length,\n            mind = gap,\n            partial,\n            value = holder[key];\n\n// If the value has a toJSON method, call it to obtain a replacement value.\n\n        if (value && typeof value === \'object\' &&\n                typeof value.toJSON === \'function\') {\n            value = value.toJSON(key);\n        }\n\n// If we were called with a replacer function, then call the replacer to\n// obtain a replacement value.\n\n        if (typeof rep === \'function\') {\n            value = rep.call(holder, key, value);\n        }\n\n// What happens next depends on the value\'s type.\n\n        switch (typeof value) {\n        case \'string\':\n            return quote(value);\n\n        case \'number\':\n\n// JSON numbers must be finite. Encode non-finite numbers as null.\n\n            return isFinite(value) ? String(value) : \'null\';\n\n        case \'boolean\':\n        case \'null\':\n\n// If the value is a boolean or null, convert it to a string. Note:\n// typeof null does not produce \'null\'. The case is included here in\n// the remote chance that this gets fixed someday.\n\n            return String(value);\n\n// If the type is \'object\', we might be dealing with an object or an array or\n// null.\n\n        case \'object\':\n\n// Due to a specification blunder in ECMAScript, typeof null is \'object\',\n// so watch out for that case.\n\n            if (!value) {\n                return \'null\';\n            }\n\n// Make an array to hold the partial results of stringifying this object value.\n\n            gap += indent;\n            partial = [];\n\n// Is the value an array?\n\n            if (Object.prototype.toString.apply(value) === \'[object Array]\') {\n\n// The value is an array. Stringify every element. Use null as a placeholder\n// for non-JSON values.\n\n                length = value.length;\n                for (i = 0; i < length; i += 1) {\n                    partial[i] = str(i, value) || \'null\';\n                }\n\n// Join all of the elements together, separated with commas, and wrap them in\n// brackets.\n\n                v = partial.length === 0 ? \'[]\' :\n                    gap ? \'[\\n\' + gap +\n                            partial.join(\',\\n\' + gap) + \'\\n\' +\n                                mind + \']\' :\n                          \'[\' + partial.join(\',\') + \']\';\n                gap = mind;\n                return v;\n            }\n\n// If the replacer is an array, use it to select the members to be stringified.\n\n            if (rep && typeof rep === \'object\') {\n                length = rep.length;\n                for (i = 0; i < length; i += 1) {\n                    k = rep[i];\n                    if (typeof k === \'string\') {\n                        v = str(k, value);\n                        if (v) {\n                            partial.push(quote(k) + (gap ? \': \' : \':\') + v);\n                        }\n                    }\n                }\n            } else {\n\n// Otherwise, iterate through all of the keys in the object.\n\n                for (k in value) {\n                    if (Object.hasOwnProperty.call(value, k)) {\n                        v = str(k, value);\n                        if (v) {\n                            partial.push(quote(k) + (gap ? \': \' : \':\') + v);\n                        }\n                    }\n                }\n            }\n\n// Join all of the member texts together, separated with commas,\n// and wrap them in braces.\n\n            v = partial.length === 0 ? \'{}\' :\n                gap ? \'{\\n\' + gap + partial.join(\',\\n\' + gap) + \'\\n\' +\n                        mind + \'}\' : \'{\' + partial.join(\',\') + \'}\';\n            gap = mind;\n            return v;\n        }\n        return v;\n    }\n\n// If the JSON object does not yet have a stringify method, give it one.\n\n    if (typeof JSON.stringify !== \'function\') {\n        JSON.stringify = function (value, replacer, space) {\n\n// The stringify method takes a value and an optional replacer, and an optional\n// space parameter, and returns a JSON text. The replacer can be a function\n// that can replace values, or an array of strings that will select the keys.\n// A default replacer method can be provided. Use of the space parameter can\n// produce text that is more easily readable.\n\n            var i;\n            gap = \'\';\n            indent = \'\';\n\n// If the space parameter is a number, make an indent string containing that\n// many spaces.\n\n            if (typeof space === \'number\') {\n                for (i = 0; i < space; i += 1) {\n                    indent += \' \';\n                }\n\n// If the space parameter is a string, it will be used as the indent string.\n\n            } else if (typeof space === \'string\') {\n                indent = space;\n            }\n\n// If there is a replacer, it must be a function or an array.\n// Otherwise, throw an error.\n\n            rep = replacer;\n            if (replacer && typeof replacer !== \'function\' &&\n                    (typeof replacer !== \'object\' ||\n                     typeof replacer.length !== \'number\')) {\n                throw new Error(\'JSON.stringify\');\n            }\n\n// Make a fake root object containing our value under the key of \'\'.\n// Return the result of stringifying the value.\n\n            return str(\'\', {\'\': value});\n        };\n    }\n\n\n// If the JSON object does not yet have a parse method, give it one.\n\n    if (typeof JSON.parse !== \'function\') {\n        JSON.parse = function (text, reviver) {\n\n// The parse method takes a text and an optional reviver function, and returns\n// a JavaScript value if the text is a valid JSON text.\n\n            var j;\n\n            function walk(holder, key) {\n\n// The walk method is used to recursively walk the resulting structure so\n// that modifications can be made.\n\n                var k, v, value = holder[key];\n                if (value && typeof value === \'object\') {\n                    for (k in value) {\n                        if (Object.hasOwnProperty.call(value, k)) {\n                            v = walk(value, k);\n                            if (v !== undefined) {\n                                value[k] = v;\n                            } else {\n                                delete value[k];\n                            }\n                        }\n                    }\n                }\n                return reviver.call(holder, key, value);\n            }\n\n\n// Parsing happens in four stages. In the first stage, we replace certain\n// Unicode characters with escape sequences. JavaScript handles many characters\n// incorrectly, either silently deleting them, or treating them as line endings.\n\n            text = String(text);\n            cx.lastIndex = 0;\n            if (cx.test(text)) {\n                text = text.replace(cx, function (a) {\n                    return \'\\\\u\' +\n                        (\'0000\' + a.charCodeAt(0).toString(16)).slice(-4);\n                });\n            }\n\n// In the second stage, we run the text against regular expressions that look\n// for non-JSON patterns. We are especially concerned with \'()\' and \'new\'\n// because they can cause invocation, and \'=\' because it can cause mutation.\n// But just to be safe, we want to reject all unexpected forms.\n\n// We split the second stage into 4 regexp operations in order to work around\n// crippling inefficiencies in IE\'s and Safari\'s regexp engines. First we\n// replace the JSON backslash pairs with \'@\' (a non-JSON character). Second, we\n// replace all simple value tokens with \']\' characters. Third, we delete all\n// open brackets that follow a colon or comma or that begin the text. Finally,\n// we look to see that the remaining characters are only whitespace or \']\' or\n// \',\' or \':\' or \'{\' or \'}\'. If that is so, then the text is safe for eval.\n\n            if (/^[\\],:{}\\s]*$/.\ntest(text.replace(/\\\\(?:["\\\\\\/bfnrt]|u[0-9a-fA-F]{4})/g, \'@\').\nreplace(/"[^"\\\\\\n\\r]*"|true|false|null|-?\\d+(?:\\.\\d*)?(?:[eE][+\\-]?\\d+)?/g, \']\').\nreplace(/(?:^|:|,)(?:\\s*\\[)+/g, \'\'))) {\n\n// In the third stage we use the eval function to compile the text into a\n// JavaScript structure. The \'{\' operator is subject to a syntactic ambiguity\n// in JavaScript: it can begin a block or an object literal. We wrap the text\n// in parens to eliminate the ambiguity.\n\n                j = eval(\'(\' + text + \')\');\n\n// In the optional fourth stage, we recursively walk the new structure, passing\n// each name/value pair to a reviver function for possible transformation.\n\n                return typeof reviver === \'function\' ?\n                    walk({\'\': j}, \'\') : j;\n            }\n\n// If the text is not JSON parseable, then a SyntaxError is thrown.\n\n            throw new SyntaxError(\'JSON.parse\');\n        };\n    }\n}());\n\n\n// polyfills.js - Extended ECMAScript compatibility layer for JScript / CScript\n// All code is written in ES3-compatible syntax (no arrow functions, let/const,\n// template literals, destructuring, spread, classes, for..of, Promises).\n\n// ---------------------------------------------------------------------------\n// Array - static methods\n// ---------------------------------------------------------------------------\n\nif (!Array.isArray) {\n    Array.isArray = function(arg) {\n        return Object.prototype.toString.call(arg) === \'[object Array]\';\n    };\n}\n\nif (!Array.of) {\n    Array.of = function() {\n        return Array.prototype.slice.call(arguments);\n    };\n}\n\nif (!Array.from) {\n    Array.from = function(arrayLike, mapFn, thisArg) {\n        if (arrayLike == null) {\n            throw new TypeError(\'Array.from requires an array-like object\');\n        }\n        var arr = [];\n        var len = arrayLike.length >>> 0;\n        for (var i = 0; i < len; i++) {\n            var val = typeof arrayLike === \'string\'\n                ? arrayLike.charAt(i)\n                : arrayLike[i];\n            arr.push(mapFn ? mapFn.call(thisArg, val, i) : val);\n        }\n        return arr;\n    };\n}\n\n// ---------------------------------------------------------------------------\n// Array - prototype methods\n// ---------------------------------------------------------------------------\n\nif (!Array.prototype.every) {\n    Array.prototype.every = function(callback, thisArg) {\n        if (this == null) throw new TypeError(\'Array.prototype.every called on null or undefined\');\n        if (typeof callback !== \'function\') throw new TypeError(callback + \' is not a function\');\n        var O = Object(this);\n        var len = O.length >>> 0;\n        for (var i = 0; i < len; i++) {\n            if (i in O && !callback.call(thisArg, O[i], i, O)) return false;\n        }\n        return true;\n    };\n}\n\nif (!Array.prototype.some) {\n    Array.prototype.some = function(callback, thisArg) {\n        if (this == null) throw new TypeError(\'Array.prototype.some called on null or undefined\');\n        if (typeof callback !== \'function\') throw new TypeError(callback + \' is not a function\');\n        var O = Object(this);\n        var len = O.length >>> 0;\n        for (var i = 0; i < len; i++) {\n            if (i in O && callback.call(thisArg, O[i], i, O)) return true;\n        }\n        return false;\n    };\n}\n\nif (!Array.prototype.includes) {\n    // Uses SameValueZero: NaN equals NaN, unlike indexOf\n    Array.prototype.includes = function(searchElement, fromIndex) {\n        var O = Object(this);\n        var len = O.length >>> 0;\n        if (len === 0) return false;\n        var n = fromIndex | 0;\n        var k = n >= 0 ? n : Math.max(0, len + n);\n        for (; k < len; k++) {\n            var el = O[k];\n            if (el === searchElement || (el !== el && searchElement !== searchElement)) {\n                return true;\n            }\n        }\n        return false;\n    };\n}\n\nif (!Array.prototype.findIndex) {\n    Array.prototype.findIndex = function(predicate, thisArg) {\n        if (this == null) throw new TypeError(\'Array.prototype.findIndex called on null or undefined\');\n        if (typeof predicate !== \'function\') throw new TypeError(\'predicate must be a function\');\n        var O = Object(this);\n        var len = O.length >>> 0;\n        for (var i = 0; i < len; i++) {\n            if (predicate.call(thisArg, O[i], i, O)) return i;\n        }\n        return -1;\n    };\n}\n\nif (!Array.prototype.fill) {\n    Array.prototype.fill = function(value, start, end) {\n        var O = Object(this);\n        var len = O.length >>> 0;\n        var relStart = start === undefined ? 0 : (start | 0);\n        var k = relStart < 0 ? Math.max(len + relStart, 0) : Math.min(relStart, len);\n        var relEnd = end === undefined ? len : (end | 0);\n        var last = relEnd < 0 ? Math.max(len + relEnd, 0) : Math.min(relEnd, len);\n        while (k < last) {\n            O[k] = value;\n            k++;\n        }\n        return O;\n    };\n}\n\nif (!Array.prototype.flat) {\n    Array.prototype.flat = function(depth) {\n        depth = depth === undefined ? 1 : Math.floor(+depth);\n        var result = [];\n        var arr = this;\n        (function flatten(a, d) {\n            for (var i = 0; i < a.length; i++) {\n                if (Array.isArray(a[i]) && d > 0) {\n                    flatten(a[i], d - 1);\n                } else {\n                    result.push(a[i]);\n                }\n            }\n        }(arr, depth));\n        return result;\n    };\n}\n\nif (!Array.prototype.flatMap) {\n    Array.prototype.flatMap = function(callback, thisArg) {\n        return this.map(callback, thisArg).flat(1);\n    };\n}\n\n// ---------------------------------------------------------------------------\n// String - prototype methods\n// ---------------------------------------------------------------------------\n\nif (!String.prototype.endsWith) {\n    String.prototype.endsWith = function(searchString, endPosition) {\n        var str = String(this);\n        var pos = endPosition === undefined ? str.length : Math.min(endPosition | 0, str.length);\n        return str.slice(pos - searchString.length, pos) === searchString;\n    };\n}\n\nif (!String.prototype.includes) {\n    String.prototype.includes = function(searchString, position) {\n        return String(this).indexOf(searchString, position || 0) !== -1;\n    };\n}\n\nif (!String.prototype.repeat) {\n    String.prototype.repeat = function(count) {\n        if (this == null) throw new TypeError(\'String.prototype.repeat called on null or undefined\');\n        var str = String(this);\n        count = Math.floor(count);\n        if (count < 0 || count === Infinity) throw new RangeError(\'Invalid count value\');\n        var result = \'\';\n        for (var i = 0; i < count; i++) result += str;\n        return result;\n    };\n}\n\nif (!String.prototype.padStart) {\n    String.prototype.padStart = function(targetLength, padString) {\n        var str = String(this);\n        targetLength = targetLength >> 0;\n        padString = padString === undefined ? \' \' : String(padString);\n        if (str.length >= targetLength || padString.length === 0) return str;\n        var padLength = targetLength - str.length;\n        var pad = padString;\n        while (pad.length < padLength) pad += padString;\n        return pad.slice(0, padLength) + str;\n    };\n}\n\nif (!String.prototype.padEnd) {\n    String.prototype.padEnd = function(targetLength, padString) {\n        var str = String(this);\n        targetLength = targetLength >> 0;\n        padString = padString === undefined ? \' \' : String(padString);\n        if (str.length >= targetLength || padString.length === 0) return str;\n        var padLength = targetLength - str.length;\n        var pad = padString;\n        while (pad.length < padLength) pad += padString;\n        return str + pad.slice(0, padLength);\n    };\n}\n\nif (!String.prototype.trimStart) {\n    String.prototype.trimStart = function() {\n        return String(this).replace(/^[\\s\\uFEFF\\xA0]+/, \'\');\n    };\n    String.prototype.trimLeft = String.prototype.trimStart;\n}\n\nif (!String.prototype.trimEnd) {\n    String.prototype.trimEnd = function() {\n        return String(this).replace(/[\\s\\uFEFF\\xA0]+$/, \'\');\n    };\n    String.prototype.trimRight = String.prototype.trimEnd;\n}\n\n// ---------------------------------------------------------------------------\n// Object - static methods\n// ---------------------------------------------------------------------------\n\nif (!Object.keys) {\n    Object.keys = function(obj) {\n        if (typeof obj !== \'object\' && typeof obj !== \'function\' || obj === null) {\n            throw new TypeError(\'Object.keys called on non-object\');\n        }\n        var keys = [];\n        for (var k in obj) {\n            if (Object.prototype.hasOwnProperty.call(obj, k)) keys.push(k);\n        }\n        return keys;\n    };\n}\n\nif (!Object.values) {\n    Object.values = function(obj) {\n        var keys = Object.keys(obj);\n        var vals = [];\n        for (var i = 0; i < keys.length; i++) vals.push(obj[keys[i]]);\n        return vals;\n    };\n}\n\nif (!Object.entries) {\n    Object.entries = function(obj) {\n        var keys = Object.keys(obj);\n        var entries = [];\n        for (var i = 0; i < keys.length; i++) entries.push([keys[i], obj[keys[i]]]);\n        return entries;\n    };\n}\n\nif (!Object.assign) {\n    Object.assign = function(target) {\n        if (target == null) throw new TypeError(\'Cannot convert undefined or null to object\');\n        var to = Object(target);\n        for (var i = 1; i < arguments.length; i++) {\n            var src = arguments[i];\n            if (src == null) continue;\n            for (var k in src) {\n                if (Object.prototype.hasOwnProperty.call(src, k)) to[k] = src[k];\n            }\n        }\n        return to;\n    };\n}\n\nif (!Object.create) {\n    Object.create = function(proto) {\n        if (proto === null) throw new Error(\'Object.create: null prototype not supported in this polyfill\');\n        function F() {}\n        F.prototype = proto;\n        return new F();\n    };\n}\n\nif (!Object.freeze) {\n    // No-op: JScript cannot truly freeze objects\n    Object.freeze = function(obj) { return obj; };\n}\n\nif (!Object.isFrozen) {\n    Object.isFrozen = function(obj) {\n        return typeof obj !== \'object\' || obj === null;\n    };\n}\n\n// ---------------------------------------------------------------------------\n// Number - static methods and constants\n// ---------------------------------------------------------------------------\n\nif (Number.isNaN === undefined) {\n    Number.isNaN = function(value) {\n        return typeof value === \'number\' && value !== value;\n    };\n}\n\nif (Number.isFinite === undefined) {\n    Number.isFinite = function(value) {\n        return typeof value === \'number\' && isFinite(value);\n    };\n}\n\nif (Number.isInteger === undefined) {\n    Number.isInteger = function(value) {\n        return typeof value === \'number\' && isFinite(value) && Math.floor(value) === value;\n    };\n}\n\nif (Number.parseInt === undefined) {\n    Number.parseInt = parseInt;\n}\n\nif (Number.parseFloat === undefined) {\n    Number.parseFloat = parseFloat;\n}\n\nif (Number.EPSILON === undefined) {\n    Number.EPSILON = 2.220446049250313e-16;\n}\n\nif (Number.MAX_SAFE_INTEGER === undefined) {\n    Number.MAX_SAFE_INTEGER = 9007199254740991;\n}\n\nif (Number.MIN_SAFE_INTEGER === undefined) {\n    Number.MIN_SAFE_INTEGER = -9007199254740991;\n}\n\n// ---------------------------------------------------------------------------\n// Math - static methods\n// ---------------------------------------------------------------------------\n\nif (!Math.sign) {\n    Math.sign = function(x) {\n        x = +x;\n        if (x === 0 || x !== x) return x; // handles +0, -0, NaN\n        return x > 0 ? 1 : -1;\n    };\n}\n\nif (!Math.trunc) {\n    Math.trunc = function(x) {\n        return x < 0 ? Math.ceil(x) : Math.floor(x);\n    };\n}\n\nif (!Math.log2) {\n    Math.log2 = function(x) {\n        return Math.log(x) / Math.LN2;\n    };\n}\n\nif (!Math.log10) {\n    Math.log10 = function(x) {\n        return Math.log(x) / Math.LN10;\n    };\n}\n\nif (!Math.cbrt) {\n    Math.cbrt = function(x) {\n        var y = Math.pow(Math.abs(x), 1 / 3);\n        return x < 0 ? -y : y;\n    };\n}\n\nif (!Math.hypot) {\n    Math.hypot = function() {\n        var sum = 0;\n        for (var i = 0; i < arguments.length; i++) {\n            sum += arguments[i] * arguments[i];\n        }\n        return Math.sqrt(sum);\n    };\n}\n\nif (!Math.clz32) {\n    Math.clz32 = function(x) {\n        x = x >>> 0;\n        if (x === 0) return 32;\n        var n = 0;\n        if ((x & 0xFFFF0000) === 0) { n += 16; x <<= 16; }\n        if ((x & 0xFF000000) === 0) { n += 8;  x <<= 8;  }\n        if ((x & 0xF0000000) === 0) { n += 4;  x <<= 4;  }\n        if ((x & 0xC0000000) === 0) { n += 2;  x <<= 2;  }\n        if ((x & 0x80000000) === 0) { n += 1;            }\n        return n;\n    };\n}\n\n// ---------------------------------------------------------------------------\n// Date - static methods\n// ---------------------------------------------------------------------------\n\nif (!Date.now) {\n    Date.now = function() {\n        return new Date().getTime();\n    };\n}\n\nif (!Date.prototype.toISOString) {\n    Date.prototype.toISOString = function() {\n        if (!isFinite(this)) throw new RangeError(\'Invalid time value\');\n        function pad(n, w) {\n            var s = String(n);\n            while (s.length < (w || 2)) s = \'0\' + s;\n            return s;\n        }\n        return this.getUTCFullYear()        + \'-\' +\n               pad(this.getUTCMonth() + 1)  + \'-\' +\n               pad(this.getUTCDate())        + \'T\' +\n               pad(this.getUTCHours())       + \':\' +\n               pad(this.getUTCMinutes())     + \':\' +\n               pad(this.getUTCSeconds())     + \'.\' +\n               pad(this.getUTCMilliseconds(), 3) + \'Z\';\n    };\n}\n\n// ---------------------------------------------------------------------------\n// Function - prototype methods\n// ---------------------------------------------------------------------------\n\nif (!Function.prototype.bind) {\n    Function.prototype.bind = function(oThis) {\n        if (typeof this !== \'function\') {\n            throw new TypeError(\'Function.prototype.bind: target is not callable\');\n        }\n        var aArgs    = Array.prototype.slice.call(arguments, 1);\n        var fToBind  = this;\n        var fNOP     = function() {};\n        var fBound   = function() {\n            return fToBind.apply(\n                (this instanceof fNOP && oThis) ? this : oThis,\n                aArgs.concat(Array.prototype.slice.call(arguments))\n            );\n        };\n        fNOP.prototype = this.prototype;\n        fBound.prototype = new fNOP();\n        return fBound;\n    };\n}\n\n// ---------------------------------------------------------------------------\n// console shim  (maps to WScript output)\n// ---------------------------------------------------------------------------\n\nif (typeof console === \'undefined\') {\n    console = (function() {\n        function _print(prefix, args) {\n            var parts = [];\n            for (var i = 0; i < args.length; i++) {\n                var v = args[i];\n                parts.push((v !== null && typeof v === \'object\') ? JSON.stringify(v) : String(v));\n            }\n            WScript.Echo(prefix + parts.join(\' \'));\n        }\n        return {\n            log:   function() { _print(\'\',         arguments); },\n            info:  function() { _print(\'[INFO] \',  arguments); },\n            warn:  function() { _print(\'[WARN] \',  arguments); },\n            error: function() { _print(\'[ERROR] \', arguments); }\n        };\n    }());\n}\n\nload_working_directory = function (){\n	\n	var path = CURRENT_FOLDER+"/.workspace";\n    var workspace = read_all_text_file(path);\n	if(!workspace){\n\n		var input_folder = typeof INPUT_FOLDER !== "undefined" ? INPUT_FOLDER : "";\n		return CURRENT_FOLDER+input_folder+"\\\\";\n\n	}else{\n		\n		return workspace.trim();\n	}\n	\n}\n\nsave_working_directory = function(working_directory){\n	\n	var path = CURRENT_FOLDER+"/.workspace";\n	\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    try{\n		\n		var f = fso.CreateTextFile(path, true);\n		f.WriteLine(working_directory);\n		f.Close();\n	}\n	catch(exc){\n		log(exc.message);\n		return null;\n		\n	}\n	\n	\n	\n}\n\nload_properties = function (properties_file){\n    \n    \n    var path = CURRENT_FOLDER+"/"+properties_file;\n    var lib = read_all_text_file(path);\n    var rows = lib.split("\\n");\n    log("reading configuration...");\n    for(var i= 0;i < rows.length; i++){\n        var row = rows[i].trim();\n        var new_row = "";\n        if(row !== "" && row.charAt(0) !== "#"){\n            var eq = row.indexOf("=");\n            if(eq !== -1){\n                var key   = row.substring(0, eq).trim();\n                var value = row.substring(eq + 1)\n                    .replace(/\\\\/g, "\\\\\\\\")\n                    .replace(/"/g, "\\\\\\"");\n                new_row = key+"=\\""+value+"\\";";\n            }\n        }\n        rows[i] = new_row;\n    }\n    lib=rows.join(\'\\n\');\n    log("configuration loaded.");\n    eval(lib);\n    \n}\n\n\n\n\n\n// Writes text to a file. Pass unicode=true to write UTF-16LE (needed for\n// non-ASCII content); defaults to ASCII for backward compatibility.\nwrite_text_to_file = function (text,filepath,unicode){\n\n    var ForWriting = 2;\n    var TristateUnicode = -1;\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    var f = fso.OpenTextFile(filepath, ForWriting, true, unicode ? TristateUnicode : 0);\n    try{\n        f.Write(text);\n    } finally {\n        f.Close();\n    }\n\n}\n\nstdin   = _script.StdIn;\nstdout  = _script.StdOut;\n\n\nread_line = function (){\n    var str= "";\n\n    str += stdin.ReadLine();\n\n    return str;\n\n}\n\nread = function (n){\n    return stdin.Read(n);\n}\n\nread_all = function(){\n    try{\n\n		if (stdin.AtEndOfStream)\n			return ("");\n		else\n			return (stdin.ReadAll());\n\n\n	}\n	catch(exc){\n		log("can\'t read from stdin");\n		return null;\n\n	}\n}\n\nwrite_line = function (data){\n    \n     stdout.WriteLine(data);\n    \n}\n\nwrite = function (data){\n    \n    stdout.Write(data);\n    \n}\n\n// Returns the full paths of every file directly inside path (not folders,\n// despite the name list_folders kept below for backward compatibility).\nlist_files = function (path){\n\n    var files = [];\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    var folder = fso.GetFolder(path);\n    var enumerator = new Enumerator(folder.files);\n    for (; !enumerator.atEnd(); enumerator.moveNext()){\n\n        files.push(enumerator.item().path);\n\n    }\n    return files;\n}\n\n// Kept for backward compatibility: despite the name, this lists files (see TODO.md).\nlist_folders = list_files;\n\n// Returns the full paths of every subfolder directly inside path.\nlist_subfolders = function (path){\n\n    var folders = [];\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    var folder = fso.GetFolder(path);\n    var enumerator = new Enumerator(folder.subfolders);\n    for (; !enumerator.atEnd(); enumerator.moveNext()){\n\n        folders.push(enumerator.item().path);\n\n    }\n    return folders;\n}\n\n\nrandomString = function(len, charSet) {\n    charSet = charSet || \'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789\';\n    var randomString = \'\';\n    for (var i = 0; i < len; i++) {\n    	var randomPoz = Math.floor(Math.random() * charSet.length);\n    	randomString += charSet.substring(randomPoz,randomPoz+1);\n    }\n    return randomString;\n}\n\n\n//format date object, currently supports YYYY/MM/DD, YY/MM/DD, YYYYMMDD\nformat_date = function(d,format){\n\n	var yyyy = ""+d.getFullYear();\n	var mm   = ("0"+(d.getMonth()+1)).slice(-2);\n	var dd   = ("0"+(d.getDate())).slice(-2);\n\n	if(format == \'YYYY/MM/DD\' ){\n\n		return yyyy+"/"+mm+"/"+dd;\n\n	}else if(format == \'YY/MM/DD\' ){\n\n		return yyyy.slice(-2)+"/"+mm+"/"+dd;\n\n	}else if(format == \'YYYYMMDD\' ){\n\n		return yyyy+mm+dd;\n\n	}else{\n		log(\'format not recognized\')\n		return d.toString();\n	}\n\n}\n\nparse_date = function(ds,format){\n\n	var d = new Date();\n\n	if(format == \'DD/MM/YYYY\' ){\n\n		var day   = ds.substring(0,2);\n		var month = ds.substring(3,5);\n		var year  = ds.substring(6,10);\n\n		d.setFullYear(year);\n		d.setMonth(parseInt(month)-1);\n		d.setDate(parseInt(day))\n		return d;\n		//return (""+d.getFullYear())+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);\n\n	}else{\n		log(\'format not recognized\')\n		return ds;\n	}\n\n\n\n}\n\n\n// ---------------------------------------------------------------------------\n// File-system helpers\n// ---------------------------------------------------------------------------\n\n// Reads all text from a file. Returns "" for an empty file, throws on error.\nread_text_file = function(path) {\n    var ForReading = 1;\n    var TristateUseDefault = -2; // auto-detects a Unicode BOM; ASCII files are unaffected\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    try {\n        var f = fso.OpenTextFile(path, ForReading, false, TristateUseDefault);\n        if (f.AtEndOfStream) { f.Close(); return ""; }\n        var content = f.ReadAll();\n        f.Close();\n        return content;\n    } catch(exc) {\n        throw new Error("read_text_file: " + exc.message);\n    }\n};\n\n// Returns true if a file exists at path.\nfile_exists = function(path) {\n    return (new ActiveXObject("Scripting.FileSystemObject")).FileExists(path);\n};\n\n// Returns true if a folder exists at path.\nfolder_exists = function(path) {\n    return (new ActiveXObject("Scripting.FileSystemObject")).FolderExists(path);\n};\n\n// Deletes a file. Throws if the file does not exist.\ndelete_file = function(path) {\n    (new ActiveXObject("Scripting.FileSystemObject")).DeleteFile(path);\n};\n\n// Creates a single folder. No-op if it already exists.\ncreate_folder = function(path) {\n    var fso = new ActiveXObject("Scripting.FileSystemObject");\n    if (!fso.FolderExists(path)) { fso.CreateFolder(path); }\n};\n\n// Writes an array of integers (0-255) to a binary file using ADODB.Stream.\n// Each integer is stored as one raw byte (iso-8859-1, no BOM).\n// Note: byte values 128-159 may not round-trip on some locales (Windows-1252 overlap).\nwrite_binary_file = function(path, bytes) {\n    var str = "";\n    for (var i = 0; i < bytes.length; i++) {\n        str += String.fromCharCode(bytes[i] & 0xFF);\n    }\n    var stream = new ActiveXObject("ADODB.Stream");\n    stream.Type    = 2;           // adTypeText\n    stream.CharSet = "iso-8859-1";\n    stream.Open();\n    stream.WriteText(str);\n    stream.SaveToFile(path, 2);   // adSaveCreateOverWrite\n    stream.Close();\n};\n\n// Reads a binary file and returns an array of integers (0-255).\nread_binary_file = function(path) {\n    var stream = new ActiveXObject("ADODB.Stream");\n    stream.Type = 1; // adTypeBinary\n    stream.Open();\n    stream.LoadFromFile(path);\n    if (stream.Size === 0) { stream.Close(); return []; }\n    var bytes = new VBArray(stream.Read()).toArray();\n    stream.Close();\n    return bytes;\n};\n\n// ---------------------------------------------------------------------------\n// http_request(url, method, callback [, body [, headers [, timeout]]])\n// callback(responseText, statusCode, responseHeaders)\n// body            - optional string to send as request body (for POST/PUT)\n// headers         - optional plain object of request headers to set\n// timeout         - optional milliseconds; throws if any phase exceeds it\n// responseHeaders - raw header block as returned by getAllResponseHeaders()\n//\n// Still synchronous (open(..., false)) despite the callback-shaped API: true\n// async/streaming HTTP is tracked as a separate feature in TODO.md.\nhttp_request = function(url, method, reqListener, body, headers, timeout){\n	if([\'GET\',\'POST\',\'PUT\',\'DELETE\'].indexOf(method) == -1){\n		throw \'method not recognized:\' + method;\n	}\n	// ServerXMLHTTP, not plain XMLHTTP: XMLHTTP.6.0 does not implement\n	// setTimeouts() at all (it throws "Object doesn\'t support this property\n	// or method"), so a timeout could never actually take effect there.\n	var request = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");\n	request.open(method, url, false);\n	if (typeof timeout === "number") {\n		request.setTimeouts(timeout, timeout, timeout, timeout);\n	}\n	if (headers) {\n		for (var key in headers) {\n			if (Object.prototype.hasOwnProperty.call(headers, key)) {\n				request.setRequestHeader(key, headers[key]);\n			}\n		}\n	}\n	request.send(body || null);\n	reqListener(request.responseText, request.status, request.getAllResponseHeaders());\n}\n';

// ---------------------------------------------------------------------------
// core.js
// ---------------------------------------------------------------------------

//modding jscript engine
if (!Array.prototype.indexOf) {
    
    Array.prototype.indexOf = function(obj, start) {
        var j = this.length;
        var n = start || 0;
        for (var i = n >= 0 ? n : Math.max(j + n, 0); i < j; i++) {
            if (this[i] === obj) { return i; }
        }
        return -1;
    }

}

if (!String.prototype.trim) {
    (function() {
        // Make sure we trim BOM and NBSP
        var rtrim = /^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g;
        String.prototype.trim = function() {
            return this.replace(rtrim, '');
        };
    })();
}

if (!String.prototype.startsWith) {
    String.prototype.startsWith = function(searchString, position){
      position = position || 0;
      return this.substr(position, searchString.length) === searchString;
  };
}

if (!Array.prototype.filter) {
  Array.prototype.filter = function(fun/*, thisArg*/) {
    'use strict';

    if (this === void 0 || this === null) {
      throw new TypeError();
    }

    var t = Object(this);
    var len = t.length >>> 0;
    if (typeof fun !== 'function') {
      throw new TypeError();
    }

    var res = [];
    var thisArg = arguments.length >= 2 ? arguments[1] : void 0;
    for (var i = 0; i < len; i++) {
      if (i in t) {
        var val = t[i];

        // NOTE: Technically this should Object.defineProperty at
        //       the next index, as push can be affected by
        //       properties on Object.prototype and Array.prototype.
        //       But that method's new, and collisions should be
        //       rare, so use the more-compatible alternative.
        if (fun.call(thisArg, val, i, t)) {
          res.push(val);
        }
      }
    }

    return res;
  };
}

// Production steps of ECMA-262, Edition 5, 15.4.4.19
// Reference: http://es5.github.io/#x15.4.4.19
if (!Array.prototype.map) {

  Array.prototype.map = function(callback, thisArg) {

    var T, A, k;

    if (this == null) {
      throw new TypeError(' this is null or not defined');
    }

    // 1. Let O be the result of calling ToObject passing the |this| 
    //    value as the argument.
    var O = Object(this);

    // 2. Let lenValue be the result of calling the Get internal 
    //    method of O with the argument "length".
    // 3. Let len be ToUint32(lenValue).
    var len = O.length >>> 0;

    // 4. If IsCallable(callback) is false, throw a TypeError exception.
    // See: http://es5.github.com/#x9.11
    if (typeof callback !== 'function') {
      throw new TypeError(callback + ' is not a function');
    }

    // 5. If thisArg was supplied, let T be thisArg; else let T be undefined.
    if (arguments.length > 1) {
      T = thisArg;
    }

    // 6. Let A be a new array created as if by the expression new Array(len) 
    //    where Array is the standard built-in constructor with that name and 
    //    len is the value of len.
    A = new Array(len);

    // 7. Let k be 0
    k = 0;

    // 8. Repeat, while k < len
    while (k < len) {

      var kValue, mappedValue;

      // a. Let Pk be ToString(k).
      //   This is implicit for LHS operands of the in operator
      // b. Let kPresent be the result of calling the HasProperty internal 
      //    method of O with argument Pk.
      //   This step can be combined with c
      // c. If kPresent is true, then
      if (k in O) {

        // i. Let kValue be the result of calling the Get internal 
        //    method of O with argument Pk.
        kValue = O[k];

        // ii. Let mappedValue be the result of calling the Call internal 
        //     method of callback with T as the this value and argument 
        //     list containing kValue, k, and O.
        mappedValue = callback.call(T, kValue, k, O);

        // iii. Call the DefineOwnProperty internal method of A with arguments
        // Pk, Property Descriptor
        // { Value: mappedValue,
        //   Writable: true,
        //   Enumerable: true,
        //   Configurable: true },
        // and false.

        // In browsers that support Object.defineProperty, use the following:
        // Object.defineProperty(A, k, {
        //   value: mappedValue,
        //   writable: true,
        //   enumerable: true,
        //   configurable: true
        // });

        // For best browser support, use the following:
        A[k] = mappedValue;
      }
      // d. Increase k by 1.
      k++;
    }

    // 9. return A
    return A;
  };
}

// Production steps of ECMA-262, Edition 5, 15.4.4.18
// Reference: http://es5.github.io/#x15.4.4.18
if (!Array.prototype.forEach) {

  Array.prototype.forEach = function(callback, thisArg) {

    var T, k;

    if (this == null) {
      throw new TypeError(' this is null or not defined');
    }

    // 1. Let O be the result of calling ToObject passing the |this| value as the argument.
    var O = Object(this);

    // 2. Let lenValue be the result of calling the Get internal method of O with the argument "length".
    // 3. Let len be ToUint32(lenValue).
    var len = O.length >>> 0;

    // 4. If IsCallable(callback) is false, throw a TypeError exception.
    // See: http://es5.github.com/#x9.11
    if (typeof callback !== "function") {
      throw new TypeError(callback + ' is not a function');
    }

    // 5. If thisArg was supplied, let T be thisArg; else let T be undefined.
    if (arguments.length > 1) {
      T = thisArg;
    }

    // 6. Let k be 0
    k = 0;

    // 7. Repeat, while k < len
    while (k < len) {

      var kValue;

      // a. Let Pk be ToString(k).
      //   This is implicit for LHS operands of the in operator
      // b. Let kPresent be the result of calling the HasProperty internal method of O with argument Pk.
      //   This step can be combined with c
      // c. If kPresent is true, then
      if (k in O) {

        // i. Let kValue be the result of calling the Get internal method of O with argument Pk.
        kValue = O[k];

        // ii. Call the Call internal method of callback with T as the this value and
        // argument list containing kValue, k, and O.
        callback.call(T, kValue, k, O);
      }
      // d. Increase k by 1.
      k++;
    }
    // 8. return undefined
  };
}

if (!Array.prototype.find) {
  Array.prototype.find = function(predicate) {
    if (this === null) {
      throw new TypeError('Array.prototype.find called on null or undefined');
    }
    if (typeof predicate !== 'function') {
      throw new TypeError('predicate must be a function');
    }
    var list = Object(this);
    var length = list.length >>> 0;
    var thisArg = arguments[1];
    var value;

    for (var i = 0; i < length; i++) {
      value = list[i];
      if (predicate.call(thisArg, value, i, list)) {
        return value;
      }
    }
    return undefined;
  };
}

// Production steps of ECMA-262, Edition 5, 15.4.4.21
// Reference: http://es5.github.io/#x15.4.4.21
if (!Array.prototype.reduce) {
  Array.prototype.reduce = function(callback /*, initialValue*/) {
    'use strict';
    if (this == null) {
      throw new TypeError('Array.prototype.reduce called on null or undefined');
    }
    if (typeof callback !== 'function') {
      throw new TypeError(callback + ' is not a function');
    }
    var t = Object(this), len = t.length >>> 0, k = 0, value;
    if (arguments.length == 2) {
      value = arguments[1];
    } else {
      while (k < len && !(k in t)) {
        k++; 
      }
      if (k >= len) {
        throw new TypeError('Reduce of empty array with no initial value');
      }
      value = t[k++];
    }
    for (; k < len; k++) {
      if (k in t) {
        value = callback(value, t[k], k, t);
      }
    }
    return value;
  };
}





/*
    http://www.JSON.org/json2.js
    2010-03-20

    Public Domain.

    NO WARRANTY EXPRESSED OR IMPLIED. USE AT YOUR OWN RISK.

    See http://www.JSON.org/js.html


    This code should be minified before deployment.
    See http://javascript.crockford.com/jsmin.html

    USE YOUR OWN COPY. IT IS EXTREMELY UNWISE TO LOAD CODE FROM SERVERS YOU DO
    NOT CONTROL.


    This file creates a global JSON object containing two methods: stringify
    and parse.

        JSON.stringify(value, replacer, space)
            value       any JavaScript value, usually an object or array.

            replacer    an optional parameter that determines how object
                        values are stringified for objects. It can be a
                        function or an array of strings.

            space       an optional parameter that specifies the indentation
                        of nested structures. If it is omitted, the text will
                        be packed without extra whitespace. If it is a number,
                        it will specify the number of spaces to indent at each
                        level. If it is a string (such as '\t' or '&nbsp;'),
                        it contains the characters used to indent at each level.

            This method produces a JSON text from a JavaScript value.

            When an object value is found, if the object contains a toJSON
            method, its toJSON method will be called and the result will be
            stringified. A toJSON method does not serialize: it returns the
            value represented by the name/value pair that should be serialized,
            or undefined if nothing should be serialized. The toJSON method
            will be passed the key associated with the value, and this will be
            bound to the value

            For example, this would serialize Dates as ISO strings.

                Date.prototype.toJSON = function (key) {
                    function f(n) {
                        // Format integers to have at least two digits.
                        return n < 10 ? '0' + n : n;
                    }

                    return this.getUTCFullYear()   + '-' +
                         f(this.getUTCMonth() + 1) + '-' +
                         f(this.getUTCDate())      + 'T' +
                         f(this.getUTCHours())     + ':' +
                         f(this.getUTCMinutes())   + ':' +
                         f(this.getUTCSeconds())   + 'Z';
                };

            You can provide an optional replacer method. It will be passed the
            key and value of each member, with this bound to the containing
            object. The value that is returned from your method will be
            serialized. If your method returns undefined, then the member will
            be excluded from the serialization.

            If the replacer parameter is an array of strings, then it will be
            used to select the members to be serialized. It filters the results
            such that only members with keys listed in the replacer array are
            stringified.

            Values that do not have JSON representations, such as undefined or
            functions, will not be serialized. Such values in objects will be
            dropped; in arrays they will be replaced with null. You can use
            a replacer function to replace those with JSON values.
            JSON.stringify(undefined) returns undefined.

            The optional space parameter produces a stringification of the
            value that is filled with line breaks and indentation to make it
            easier to read.

            If the space parameter is a non-empty string, then that string will
            be used for indentation. If the space parameter is a number, then
            the indentation will be that many spaces.

            Example:

            text = JSON.stringify(['e', {pluribus: 'unum'}]);
            // text is '["e",{"pluribus":"unum"}]'


            text = JSON.stringify(['e', {pluribus: 'unum'}], null, '\t');
            // text is '[\n\t"e",\n\t{\n\t\t"pluribus": "unum"\n\t}\n]'

            text = JSON.stringify([new Date()], function (key, value) {
                return this[key] instanceof Date ?
                    'Date(' + this[key] + ')' : value;
            });
            // text is '["Date(---current time---)"]'


        JSON.parse(text, reviver)
            This method parses a JSON text to produce an object or array.
            It can throw a SyntaxError exception.

            The optional reviver parameter is a function that can filter and
            transform the results. It receives each of the keys and values,
            and its return value is used instead of the original value.
            If it returns what it received, then the structure is not modified.
            If it returns undefined then the member is deleted.

            Example:

            // Parse the text. Values that look like ISO date strings will
            // be converted to Date objects.

            myData = JSON.parse(text, function (key, value) {
                var a;
                if (typeof value === 'string') {
                    a =
/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/.exec(value);
                    if (a) {
                        return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4],
                            +a[5], +a[6]));
                    }
                }
                return value;
            });

            myData = JSON.parse('["Date(09/09/2001)"]', function (key, value) {
                var d;
                if (typeof value === 'string' &&
                        value.slice(0, 5) === 'Date(' &&
                        value.slice(-1) === ')') {
                    d = new Date(value.slice(5, -1));
                    if (d) {
                        return d;
                    }
                }
                return value;
            });


    This is a reference implementation. You are free to copy, modify, or
    redistribute.
*/

/*jslint evil: true, strict: false */

/*members "", "\b", "\t", "\n", "\f", "\r", "\"", JSON, "\\", apply,
    call, charCodeAt, getUTCDate, getUTCFullYear, getUTCHours,
    getUTCMinutes, getUTCMonth, getUTCSeconds, hasOwnProperty, join,
    lastIndex, length, parse, prototype, push, replace, slice, stringify,
    test, toJSON, toString, valueOf
*/


// Create a JSON object only if one does not already exist. We create the
// methods in a closure to avoid creating global variables.

if (!this.JSON) {
    this.JSON = {};
}

(function () {

    function f(n) {
        // Format integers to have at least two digits.
        return n < 10 ? '0' + n : n;
    }

    if (typeof Date.prototype.toJSON !== 'function') {

        Date.prototype.toJSON = function (key) {

            return isFinite(this.valueOf()) ?
                   this.getUTCFullYear()   + '-' +
                 f(this.getUTCMonth() + 1) + '-' +
                 f(this.getUTCDate())      + 'T' +
                 f(this.getUTCHours())     + ':' +
                 f(this.getUTCMinutes())   + ':' +
                 f(this.getUTCSeconds())   + 'Z' : null;
        };

        String.prototype.toJSON =
        Number.prototype.toJSON =
        Boolean.prototype.toJSON = function (key) {
            return this.valueOf();
        };
    }

    var cx = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
        escapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
        gap,
        indent,
        meta = {    // table of character substitutions
            '\b': '\\b',
            '\t': '\\t',
            '\n': '\\n',
            '\f': '\\f',
            '\r': '\\r',
            '"' : '\\"',
            '\\': '\\\\'
        },
        rep;


    function quote(string) {

// If the string contains no control characters, no quote characters, and no
// backslash characters, then we can safely slap some quotes around it.
// Otherwise we must also replace the offending characters with safe escape
// sequences.

        escapable.lastIndex = 0;
        return escapable.test(string) ?
            '"' + string.replace(escapable, function (a) {
                var c = meta[a];
                return typeof c === 'string' ? c :
                    '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
            }) + '"' :
            '"' + string + '"';
    }


    function str(key, holder) {

// Produce a string from holder[key].

        var i,          // The loop counter.
            k,          // The member key.
            v,          // The member value.
            length,
            mind = gap,
            partial,
            value = holder[key];

// If the value has a toJSON method, call it to obtain a replacement value.

        if (value && typeof value === 'object' &&
                typeof value.toJSON === 'function') {
            value = value.toJSON(key);
        }

// If we were called with a replacer function, then call the replacer to
// obtain a replacement value.

        if (typeof rep === 'function') {
            value = rep.call(holder, key, value);
        }

// What happens next depends on the value's type.

        switch (typeof value) {
        case 'string':
            return quote(value);

        case 'number':

// JSON numbers must be finite. Encode non-finite numbers as null.

            return isFinite(value) ? String(value) : 'null';

        case 'boolean':
        case 'null':

// If the value is a boolean or null, convert it to a string. Note:
// typeof null does not produce 'null'. The case is included here in
// the remote chance that this gets fixed someday.

            return String(value);

// If the type is 'object', we might be dealing with an object or an array or
// null.

        case 'object':

// Due to a specification blunder in ECMAScript, typeof null is 'object',
// so watch out for that case.

            if (!value) {
                return 'null';
            }

// Make an array to hold the partial results of stringifying this object value.

            gap += indent;
            partial = [];

// Is the value an array?

            if (Object.prototype.toString.apply(value) === '[object Array]') {

// The value is an array. Stringify every element. Use null as a placeholder
// for non-JSON values.

                length = value.length;
                for (i = 0; i < length; i += 1) {
                    partial[i] = str(i, value) || 'null';
                }

// Join all of the elements together, separated with commas, and wrap them in
// brackets.

                v = partial.length === 0 ? '[]' :
                    gap ? '[\n' + gap +
                            partial.join(',\n' + gap) + '\n' +
                                mind + ']' :
                          '[' + partial.join(',') + ']';
                gap = mind;
                return v;
            }

// If the replacer is an array, use it to select the members to be stringified.

            if (rep && typeof rep === 'object') {
                length = rep.length;
                for (i = 0; i < length; i += 1) {
                    k = rep[i];
                    if (typeof k === 'string') {
                        v = str(k, value);
                        if (v) {
                            partial.push(quote(k) + (gap ? ': ' : ':') + v);
                        }
                    }
                }
            } else {

// Otherwise, iterate through all of the keys in the object.

                for (k in value) {
                    if (Object.hasOwnProperty.call(value, k)) {
                        v = str(k, value);
                        if (v) {
                            partial.push(quote(k) + (gap ? ': ' : ':') + v);
                        }
                    }
                }
            }

// Join all of the member texts together, separated with commas,
// and wrap them in braces.

            v = partial.length === 0 ? '{}' :
                gap ? '{\n' + gap + partial.join(',\n' + gap) + '\n' +
                        mind + '}' : '{' + partial.join(',') + '}';
            gap = mind;
            return v;
        }
        return v;
    }

// If the JSON object does not yet have a stringify method, give it one.

    if (typeof JSON.stringify !== 'function') {
        JSON.stringify = function (value, replacer, space) {

// The stringify method takes a value and an optional replacer, and an optional
// space parameter, and returns a JSON text. The replacer can be a function
// that can replace values, or an array of strings that will select the keys.
// A default replacer method can be provided. Use of the space parameter can
// produce text that is more easily readable.

            var i;
            gap = '';
            indent = '';

// If the space parameter is a number, make an indent string containing that
// many spaces.

            if (typeof space === 'number') {
                for (i = 0; i < space; i += 1) {
                    indent += ' ';
                }

// If the space parameter is a string, it will be used as the indent string.

            } else if (typeof space === 'string') {
                indent = space;
            }

// If there is a replacer, it must be a function or an array.
// Otherwise, throw an error.

            rep = replacer;
            if (replacer && typeof replacer !== 'function' &&
                    (typeof replacer !== 'object' ||
                     typeof replacer.length !== 'number')) {
                throw new Error('JSON.stringify');
            }

// Make a fake root object containing our value under the key of ''.
// Return the result of stringifying the value.

            return str('', {'': value});
        };
    }


// If the JSON object does not yet have a parse method, give it one.

    if (typeof JSON.parse !== 'function') {
        JSON.parse = function (text, reviver) {

// The parse method takes a text and an optional reviver function, and returns
// a JavaScript value if the text is a valid JSON text.

            var j;

            function walk(holder, key) {

// The walk method is used to recursively walk the resulting structure so
// that modifications can be made.

                var k, v, value = holder[key];
                if (value && typeof value === 'object') {
                    for (k in value) {
                        if (Object.hasOwnProperty.call(value, k)) {
                            v = walk(value, k);
                            if (v !== undefined) {
                                value[k] = v;
                            } else {
                                delete value[k];
                            }
                        }
                    }
                }
                return reviver.call(holder, key, value);
            }


// Parsing happens in four stages. In the first stage, we replace certain
// Unicode characters with escape sequences. JavaScript handles many characters
// incorrectly, either silently deleting them, or treating them as line endings.

            text = String(text);
            cx.lastIndex = 0;
            if (cx.test(text)) {
                text = text.replace(cx, function (a) {
                    return '\\u' +
                        ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
                });
            }

// In the second stage, we run the text against regular expressions that look
// for non-JSON patterns. We are especially concerned with '()' and 'new'
// because they can cause invocation, and '=' because it can cause mutation.
// But just to be safe, we want to reject all unexpected forms.

// We split the second stage into 4 regexp operations in order to work around
// crippling inefficiencies in IE's and Safari's regexp engines. First we
// replace the JSON backslash pairs with '@' (a non-JSON character). Second, we
// replace all simple value tokens with ']' characters. Third, we delete all
// open brackets that follow a colon or comma or that begin the text. Finally,
// we look to see that the remaining characters are only whitespace or ']' or
// ',' or ':' or '{' or '}'. If that is so, then the text is safe for eval.

            if (/^[\],:{}\s]*$/.
test(text.replace(/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g, '@').
replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g, ']').
replace(/(?:^|:|,)(?:\s*\[)+/g, ''))) {

// In the third stage we use the eval function to compile the text into a
// JavaScript structure. The '{' operator is subject to a syntactic ambiguity
// in JavaScript: it can begin a block or an object literal. We wrap the text
// in parens to eliminate the ambiguity.

                j = eval('(' + text + ')');

// In the optional fourth stage, we recursively walk the new structure, passing
// each name/value pair to a reviver function for possible transformation.

                return typeof reviver === 'function' ?
                    walk({'': j}, '') : j;
            }

// If the text is not JSON parseable, then a SyntaxError is thrown.

            throw new SyntaxError('JSON.parse');
        };
    }
}());



// ---------------------------------------------------------------------------
// ext.js
// ---------------------------------------------------------------------------

// ext.js - Ext.util.JSON / Ext.encode / Ext.decode compatibility shim
//
// Nothing in this project uses these; they exist only for scripts written
// against the old ExtJS-style JSON API. Load explicitly with load("ext") if
// you need them - they are not pulled in by core.js.

if (!this.Ext) {

    Ext={};
    Ext.util = {};

    /**
     * @class Ext.util.JSON
     * Modified version of Douglas Crockford"s json.js that doesn"t
     * mess with the Object prototype
     * http://www.json.org/js.html
     * @singleton
     */
    Ext.util.JSON = {
        encode : function(o) {
            return JSON.stringify(o);
        },

        decode : function(s) {
            return JSON.parse(s);
        }
    };

    /**
     * Shorthand for {@link Ext.util.JSON#encode}
     * @param {Mixed} o The variable to encode
     * @return {String} The JSON string
     * @member Ext
     * @method encode
     */
    Ext.encode = Ext.util.JSON.encode;
    /**
     * Shorthand for {@link Ext.util.JSON#decode}
     * @param {String} json The JSON string
     * @param {Boolean} safe (optional) Whether to return null or throw an exception if the JSON is invalid.
     * @return {Object} The resulting object
     * @member Ext
     * @method decode
     */
    Ext.decode = Ext.util.JSON.decode;

}


// ---------------------------------------------------------------------------
// polyfills.js
// ---------------------------------------------------------------------------

// polyfills.js - Extended ECMAScript compatibility layer for JScript / CScript
// All code is written in ES3-compatible syntax (no arrow functions, let/const,
// template literals, destructuring, spread, classes, for..of, Promises).

// ---------------------------------------------------------------------------
// Array - static methods
// ---------------------------------------------------------------------------

if (!Array.isArray) {
    Array.isArray = function(arg) {
        return Object.prototype.toString.call(arg) === '[object Array]';
    };
}

if (!Array.of) {
    Array.of = function() {
        return Array.prototype.slice.call(arguments);
    };
}

if (!Array.from) {
    Array.from = function(arrayLike, mapFn, thisArg) {
        if (arrayLike == null) {
            throw new TypeError('Array.from requires an array-like object');
        }
        var arr = [];
        var len = arrayLike.length >>> 0;
        for (var i = 0; i < len; i++) {
            var val = typeof arrayLike === 'string'
                ? arrayLike.charAt(i)
                : arrayLike[i];
            arr.push(mapFn ? mapFn.call(thisArg, val, i) : val);
        }
        return arr;
    };
}

// ---------------------------------------------------------------------------
// Array - prototype methods
// ---------------------------------------------------------------------------

if (!Array.prototype.every) {
    Array.prototype.every = function(callback, thisArg) {
        if (this == null) throw new TypeError('Array.prototype.every called on null or undefined');
        if (typeof callback !== 'function') throw new TypeError(callback + ' is not a function');
        var O = Object(this);
        var len = O.length >>> 0;
        for (var i = 0; i < len; i++) {
            if (i in O && !callback.call(thisArg, O[i], i, O)) return false;
        }
        return true;
    };
}

if (!Array.prototype.some) {
    Array.prototype.some = function(callback, thisArg) {
        if (this == null) throw new TypeError('Array.prototype.some called on null or undefined');
        if (typeof callback !== 'function') throw new TypeError(callback + ' is not a function');
        var O = Object(this);
        var len = O.length >>> 0;
        for (var i = 0; i < len; i++) {
            if (i in O && callback.call(thisArg, O[i], i, O)) return true;
        }
        return false;
    };
}

if (!Array.prototype.includes) {
    // Uses SameValueZero: NaN equals NaN, unlike indexOf
    Array.prototype.includes = function(searchElement, fromIndex) {
        var O = Object(this);
        var len = O.length >>> 0;
        if (len === 0) return false;
        var n = fromIndex | 0;
        var k = n >= 0 ? n : Math.max(0, len + n);
        for (; k < len; k++) {
            var el = O[k];
            if (el === searchElement || (el !== el && searchElement !== searchElement)) {
                return true;
            }
        }
        return false;
    };
}

if (!Array.prototype.findIndex) {
    Array.prototype.findIndex = function(predicate, thisArg) {
        if (this == null) throw new TypeError('Array.prototype.findIndex called on null or undefined');
        if (typeof predicate !== 'function') throw new TypeError('predicate must be a function');
        var O = Object(this);
        var len = O.length >>> 0;
        for (var i = 0; i < len; i++) {
            if (predicate.call(thisArg, O[i], i, O)) return i;
        }
        return -1;
    };
}

if (!Array.prototype.fill) {
    Array.prototype.fill = function(value, start, end) {
        var O = Object(this);
        var len = O.length >>> 0;
        var relStart = start === undefined ? 0 : (start | 0);
        var k = relStart < 0 ? Math.max(len + relStart, 0) : Math.min(relStart, len);
        var relEnd = end === undefined ? len : (end | 0);
        var last = relEnd < 0 ? Math.max(len + relEnd, 0) : Math.min(relEnd, len);
        while (k < last) {
            O[k] = value;
            k++;
        }
        return O;
    };
}

if (!Array.prototype.flat) {
    Array.prototype.flat = function(depth) {
        depth = depth === undefined ? 1 : Math.floor(+depth);
        var result = [];
        var arr = this;
        (function flatten(a, d) {
            for (var i = 0; i < a.length; i++) {
                if (Array.isArray(a[i]) && d > 0) {
                    flatten(a[i], d - 1);
                } else {
                    result.push(a[i]);
                }
            }
        }(arr, depth));
        return result;
    };
}

if (!Array.prototype.flatMap) {
    Array.prototype.flatMap = function(callback, thisArg) {
        return this.map(callback, thisArg).flat(1);
    };
}

// ---------------------------------------------------------------------------
// String - prototype methods
// ---------------------------------------------------------------------------

if (!String.prototype.endsWith) {
    String.prototype.endsWith = function(searchString, endPosition) {
        var str = String(this);
        var pos = endPosition === undefined ? str.length : Math.min(endPosition | 0, str.length);
        return str.slice(pos - searchString.length, pos) === searchString;
    };
}

if (!String.prototype.includes) {
    String.prototype.includes = function(searchString, position) {
        return String(this).indexOf(searchString, position || 0) !== -1;
    };
}

if (!String.prototype.repeat) {
    String.prototype.repeat = function(count) {
        if (this == null) throw new TypeError('String.prototype.repeat called on null or undefined');
        var str = String(this);
        count = Math.floor(count);
        if (count < 0 || count === Infinity) throw new RangeError('Invalid count value');
        var result = '';
        for (var i = 0; i < count; i++) result += str;
        return result;
    };
}

if (!String.prototype.padStart) {
    String.prototype.padStart = function(targetLength, padString) {
        var str = String(this);
        targetLength = targetLength >> 0;
        padString = padString === undefined ? ' ' : String(padString);
        if (str.length >= targetLength || padString.length === 0) return str;
        var padLength = targetLength - str.length;
        var pad = padString;
        while (pad.length < padLength) pad += padString;
        return pad.slice(0, padLength) + str;
    };
}

if (!String.prototype.padEnd) {
    String.prototype.padEnd = function(targetLength, padString) {
        var str = String(this);
        targetLength = targetLength >> 0;
        padString = padString === undefined ? ' ' : String(padString);
        if (str.length >= targetLength || padString.length === 0) return str;
        var padLength = targetLength - str.length;
        var pad = padString;
        while (pad.length < padLength) pad += padString;
        return str + pad.slice(0, padLength);
    };
}

if (!String.prototype.trimStart) {
    String.prototype.trimStart = function() {
        return String(this).replace(/^[\s\uFEFF\xA0]+/, '');
    };
    String.prototype.trimLeft = String.prototype.trimStart;
}

if (!String.prototype.trimEnd) {
    String.prototype.trimEnd = function() {
        return String(this).replace(/[\s\uFEFF\xA0]+$/, '');
    };
    String.prototype.trimRight = String.prototype.trimEnd;
}

// ---------------------------------------------------------------------------
// Object - static methods
// ---------------------------------------------------------------------------

if (!Object.keys) {
    Object.keys = function(obj) {
        if (typeof obj !== 'object' && typeof obj !== 'function' || obj === null) {
            throw new TypeError('Object.keys called on non-object');
        }
        var keys = [];
        for (var k in obj) {
            if (Object.prototype.hasOwnProperty.call(obj, k)) keys.push(k);
        }
        return keys;
    };
}

if (!Object.values) {
    Object.values = function(obj) {
        var keys = Object.keys(obj);
        var vals = [];
        for (var i = 0; i < keys.length; i++) vals.push(obj[keys[i]]);
        return vals;
    };
}

if (!Object.entries) {
    Object.entries = function(obj) {
        var keys = Object.keys(obj);
        var entries = [];
        for (var i = 0; i < keys.length; i++) entries.push([keys[i], obj[keys[i]]]);
        return entries;
    };
}

if (!Object.assign) {
    Object.assign = function(target) {
        if (target == null) throw new TypeError('Cannot convert undefined or null to object');
        var to = Object(target);
        for (var i = 1; i < arguments.length; i++) {
            var src = arguments[i];
            if (src == null) continue;
            for (var k in src) {
                if (Object.prototype.hasOwnProperty.call(src, k)) to[k] = src[k];
            }
        }
        return to;
    };
}

if (!Object.create) {
    Object.create = function(proto) {
        if (proto === null) throw new Error('Object.create: null prototype not supported in this polyfill');
        function F() {}
        F.prototype = proto;
        return new F();
    };
}

if (!Object.freeze) {
    // No-op: JScript cannot truly freeze objects
    Object.freeze = function(obj) { return obj; };
}

if (!Object.isFrozen) {
    Object.isFrozen = function(obj) {
        return typeof obj !== 'object' || obj === null;
    };
}

// ---------------------------------------------------------------------------
// Number - static methods and constants
// ---------------------------------------------------------------------------

if (Number.isNaN === undefined) {
    Number.isNaN = function(value) {
        return typeof value === 'number' && value !== value;
    };
}

if (Number.isFinite === undefined) {
    Number.isFinite = function(value) {
        return typeof value === 'number' && isFinite(value);
    };
}

if (Number.isInteger === undefined) {
    Number.isInteger = function(value) {
        return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
    };
}

if (Number.parseInt === undefined) {
    Number.parseInt = parseInt;
}

if (Number.parseFloat === undefined) {
    Number.parseFloat = parseFloat;
}

if (Number.EPSILON === undefined) {
    Number.EPSILON = 2.220446049250313e-16;
}

if (Number.MAX_SAFE_INTEGER === undefined) {
    Number.MAX_SAFE_INTEGER = 9007199254740991;
}

if (Number.MIN_SAFE_INTEGER === undefined) {
    Number.MIN_SAFE_INTEGER = -9007199254740991;
}

// ---------------------------------------------------------------------------
// Math - static methods
// ---------------------------------------------------------------------------

if (!Math.sign) {
    Math.sign = function(x) {
        x = +x;
        if (x === 0 || x !== x) return x; // handles +0, -0, NaN
        return x > 0 ? 1 : -1;
    };
}

if (!Math.trunc) {
    Math.trunc = function(x) {
        return x < 0 ? Math.ceil(x) : Math.floor(x);
    };
}

if (!Math.log2) {
    Math.log2 = function(x) {
        return Math.log(x) / Math.LN2;
    };
}

if (!Math.log10) {
    Math.log10 = function(x) {
        return Math.log(x) / Math.LN10;
    };
}

if (!Math.cbrt) {
    Math.cbrt = function(x) {
        var y = Math.pow(Math.abs(x), 1 / 3);
        return x < 0 ? -y : y;
    };
}

if (!Math.hypot) {
    Math.hypot = function() {
        var sum = 0;
        for (var i = 0; i < arguments.length; i++) {
            sum += arguments[i] * arguments[i];
        }
        return Math.sqrt(sum);
    };
}

if (!Math.clz32) {
    Math.clz32 = function(x) {
        x = x >>> 0;
        if (x === 0) return 32;
        var n = 0;
        if ((x & 0xFFFF0000) === 0) { n += 16; x <<= 16; }
        if ((x & 0xFF000000) === 0) { n += 8;  x <<= 8;  }
        if ((x & 0xF0000000) === 0) { n += 4;  x <<= 4;  }
        if ((x & 0xC0000000) === 0) { n += 2;  x <<= 2;  }
        if ((x & 0x80000000) === 0) { n += 1;            }
        return n;
    };
}

// ---------------------------------------------------------------------------
// Date - static methods
// ---------------------------------------------------------------------------

if (!Date.now) {
    Date.now = function() {
        return new Date().getTime();
    };
}

if (!Date.prototype.toISOString) {
    Date.prototype.toISOString = function() {
        if (!isFinite(this)) throw new RangeError('Invalid time value');
        function pad(n, w) {
            var s = String(n);
            while (s.length < (w || 2)) s = '0' + s;
            return s;
        }
        return this.getUTCFullYear()        + '-' +
               pad(this.getUTCMonth() + 1)  + '-' +
               pad(this.getUTCDate())        + 'T' +
               pad(this.getUTCHours())       + ':' +
               pad(this.getUTCMinutes())     + ':' +
               pad(this.getUTCSeconds())     + '.' +
               pad(this.getUTCMilliseconds(), 3) + 'Z';
    };
}

// ---------------------------------------------------------------------------
// Function - prototype methods
// ---------------------------------------------------------------------------

if (!Function.prototype.bind) {
    Function.prototype.bind = function(oThis) {
        if (typeof this !== 'function') {
            throw new TypeError('Function.prototype.bind: target is not callable');
        }
        var aArgs    = Array.prototype.slice.call(arguments, 1);
        var fToBind  = this;
        var fNOP     = function() {};
        var fBound   = function() {
            return fToBind.apply(
                (this instanceof fNOP && oThis) ? this : oThis,
                aArgs.concat(Array.prototype.slice.call(arguments))
            );
        };
        fNOP.prototype = this.prototype;
        fBound.prototype = new fNOP();
        return fBound;
    };
}

// ---------------------------------------------------------------------------
// console shim  (maps to WScript output)
// ---------------------------------------------------------------------------

if (typeof console === 'undefined') {
    console = (function() {
        function _print(prefix, args) {
            var parts = [];
            for (var i = 0; i < args.length; i++) {
                var v = args[i];
                parts.push((v !== null && typeof v === 'object') ? JSON.stringify(v) : String(v));
            }
            WScript.Echo(prefix + parts.join(' '));
        }
        return {
            log:   function() { _print('',         arguments); },
            info:  function() { _print('[INFO] ',  arguments); },
            warn:  function() { _print('[WARN] ',  arguments); },
            error: function() { _print('[ERROR] ', arguments); }
        };
    }());
}


// ---------------------------------------------------------------------------
// system.js
// ---------------------------------------------------------------------------

load_working_directory = function (){
	
	var path = CURRENT_FOLDER+"/.workspace";
    var workspace = read_all_text_file(path);
	if(!workspace){

		var input_folder = typeof INPUT_FOLDER !== "undefined" ? INPUT_FOLDER : "";
		return CURRENT_FOLDER+input_folder+"\\";

	}else{
		
		return workspace.trim();
	}
	
}

save_working_directory = function(working_directory){
	
	var path = CURRENT_FOLDER+"/.workspace";
	
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try{
		
		var f = fso.CreateTextFile(path, true);
		f.WriteLine(working_directory);
		f.Close();
	}
	catch(exc){
		log(exc.message);
		return null;
		
	}
	
	
	
}

load_properties = function (properties_file){
    
    
    var path = CURRENT_FOLDER+"/"+properties_file;
    var lib = read_all_text_file(path);
    var rows = lib.split("\n");
    log("reading configuration...");
    for(var i= 0;i < rows.length; i++){
        var row = rows[i].trim();
        var new_row = "";
        if(row !== "" && row.charAt(0) !== "#"){
            var eq = row.indexOf("=");
            if(eq !== -1){
                var key   = row.substring(0, eq).trim();
                var value = row.substring(eq + 1)
                    .replace(/\\/g, "\\\\")
                    .replace(/"/g, "\\\"");
                new_row = key+"=\""+value+"\";";
            }
        }
        rows[i] = new_row;
    }
    lib=rows.join('\n');
    log("configuration loaded.");
    eval(lib);
    
}





// Writes text to a file. Pass unicode=true to write UTF-16LE (needed for
// non-ASCII content); defaults to ASCII for backward compatibility.
write_text_to_file = function (text,filepath,unicode){

    var ForWriting = 2;
    var TristateUnicode = -1;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var f = fso.OpenTextFile(filepath, ForWriting, true, unicode ? TristateUnicode : 0);
    try{
        f.Write(text);
    } finally {
        f.Close();
    }

}

stdin   = _script.StdIn;
stdout  = _script.StdOut;


read_line = function (){
    var str= "";

    str += stdin.ReadLine();

    return str;

}

read = function (n){
    return stdin.Read(n);
}

read_all = function(){
    try{

		if (stdin.AtEndOfStream)
			return ("");
		else
			return (stdin.ReadAll());


	}
	catch(exc){
		log("can't read from stdin");
		return null;

	}
}

write_line = function (data){
    
     stdout.WriteLine(data);
    
}

write = function (data){
    
    stdout.Write(data);
    
}

// Returns the full paths of every file directly inside path (not folders,
// despite the name list_folders kept below for backward compatibility).
list_files = function (path){

    var files = [];
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(path);
    var enumerator = new Enumerator(folder.files);
    for (; !enumerator.atEnd(); enumerator.moveNext()){

        files.push(enumerator.item().path);

    }
    return files;
}

// Kept for backward compatibility: despite the name, this lists files (see TODO.md).
list_folders = list_files;

// Returns the full paths of every subfolder directly inside path.
list_subfolders = function (path){

    var folders = [];
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(path);
    var enumerator = new Enumerator(folder.subfolders);
    for (; !enumerator.atEnd(); enumerator.moveNext()){

        folders.push(enumerator.item().path);

    }
    return folders;
}


randomString = function(len, charSet) {
    charSet = charSet || 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var randomString = '';
    for (var i = 0; i < len; i++) {
    	var randomPoz = Math.floor(Math.random() * charSet.length);
    	randomString += charSet.substring(randomPoz,randomPoz+1);
    }
    return randomString;
}


//format date object, currently supports YYYY/MM/DD, YY/MM/DD, YYYYMMDD
format_date = function(d,format){

	var yyyy = ""+d.getFullYear();
	var mm   = ("0"+(d.getMonth()+1)).slice(-2);
	var dd   = ("0"+(d.getDate())).slice(-2);

	if(format == 'YYYY/MM/DD' ){

		return yyyy+"/"+mm+"/"+dd;

	}else if(format == 'YY/MM/DD' ){

		return yyyy.slice(-2)+"/"+mm+"/"+dd;

	}else if(format == 'YYYYMMDD' ){

		return yyyy+mm+dd;

	}else{
		log('format not recognized')
		return d.toString();
	}

}

parse_date = function(ds,format){

	var d = new Date();

	if(format == 'DD/MM/YYYY' ){

		var day   = ds.substring(0,2);
		var month = ds.substring(3,5);
		var year  = ds.substring(6,10);

		d.setFullYear(year);
		d.setMonth(parseInt(month)-1);
		d.setDate(parseInt(day))
		return d;
		//return (""+d.getFullYear())+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+(d.getDate())).slice(-2);

	}else{
		log('format not recognized')
		return ds;
	}



}


// ---------------------------------------------------------------------------
// File-system helpers
// ---------------------------------------------------------------------------

// Reads all text from a file. Returns "" for an empty file, throws on error.
read_text_file = function(path) {
    var ForReading = 1;
    var TristateUseDefault = -2; // auto-detects a Unicode BOM; ASCII files are unaffected
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    try {
        var f = fso.OpenTextFile(path, ForReading, false, TristateUseDefault);
        if (f.AtEndOfStream) { f.Close(); return ""; }
        var content = f.ReadAll();
        f.Close();
        return content;
    } catch(exc) {
        throw new Error("read_text_file: " + exc.message);
    }
};

// Returns true if a file exists at path.
file_exists = function(path) {
    return (new ActiveXObject("Scripting.FileSystemObject")).FileExists(path);
};

// Returns true if a folder exists at path.
folder_exists = function(path) {
    return (new ActiveXObject("Scripting.FileSystemObject")).FolderExists(path);
};

// Deletes a file. Throws if the file does not exist.
delete_file = function(path) {
    (new ActiveXObject("Scripting.FileSystemObject")).DeleteFile(path);
};

// Creates a single folder. No-op if it already exists.
create_folder = function(path) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(path)) { fso.CreateFolder(path); }
};

// Writes an array of integers (0-255) to a binary file using ADODB.Stream.
// Each integer is stored as one raw byte (iso-8859-1, no BOM).
// Note: byte values 128-159 may not round-trip on some locales (Windows-1252 overlap).
write_binary_file = function(path, bytes) {
    var str = "";
    for (var i = 0; i < bytes.length; i++) {
        str += String.fromCharCode(bytes[i] & 0xFF);
    }
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type    = 2;           // adTypeText
    stream.CharSet = "iso-8859-1";
    stream.Open();
    stream.WriteText(str);
    stream.SaveToFile(path, 2);   // adSaveCreateOverWrite
    stream.Close();
};

// Reads a binary file and returns an array of integers (0-255).
read_binary_file = function(path) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 1; // adTypeBinary
    stream.Open();
    stream.LoadFromFile(path);
    if (stream.Size === 0) { stream.Close(); return []; }
    var bytes = new VBArray(stream.Read()).toArray();
    stream.Close();
    return bytes;
};

// ---------------------------------------------------------------------------
// http_request(url, method, callback [, body [, headers [, timeout]]])
// callback(responseText, statusCode, responseHeaders)
// body            - optional string to send as request body (for POST/PUT)
// headers         - optional plain object of request headers to set
// timeout         - optional milliseconds; throws if any phase exceeds it
// responseHeaders - raw header block as returned by getAllResponseHeaders()
//
// Still synchronous (open(..., false)) despite the callback-shaped API: true
// async/streaming HTTP is tracked as a separate feature in TODO.md.
http_request = function(url, method, reqListener, body, headers, timeout){
	if(['GET','POST','PUT','DELETE'].indexOf(method) == -1){
		throw 'method not recognized:' + method;
	}
	// ServerXMLHTTP, not plain XMLHTTP: XMLHTTP.6.0 does not implement
	// setTimeouts() at all (it throws "Object doesn't support this property
	// or method"), so a timeout could never actually take effect there.
	var request = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
	request.open(method, url, false);
	if (typeof timeout === "number") {
		request.setTimeouts(timeout, timeout, timeout, timeout);
	}
	if (headers) {
		for (var key in headers) {
			if (Object.prototype.hasOwnProperty.call(headers, key)) {
				request.setRequestHeader(key, headers[key]);
			}
		}
	}
	request.send(body || null);
	reqListener(request.responseText, request.status, request.getAllResponseHeaders());
}

// ---------------------------------------------------------------------------
// helpers.js
// ---------------------------------------------------------------------------

//version 2020-10-29
//@author Daniele Marsico

//file sstem

select_file_from_folder = function (folder, message, pattern){//return selected file path
    
    var file_list           = null; 
    var selected_file_index = 0;
    
    while(true){
                   
        write_line(message);
        read_line();
                    
        file_list = list_files(folder).filter(function(elem){
                   
            return elem.match(pattern);
       
        });
        if(file_list.length == 0){
            write_line("no files found");
        }else{
            break;
        }
    }
                
    if(file_list.length > 1){
                    
        selected_file_index = -1;
        while( selected_file_index >= file_list.length || selected_file_index < 0 ){
                    
            write_line("select file to process:");
            file_list.forEach(function(value,index){
        
                var path_parts = value.split("\\");
        
                write_line("("+(index+1)+") "+path_parts[path_parts.length-1]);
        
            });
            write(">:");
            selected_file_index = parseInt(read_line().trim(),10);
            if(isNaN(selected_file_index))
                selected_file_index = -1;
            else
                selected_file_index -= 1;
                    
                    
        }
                
    }
    var selected_file = file_list[selected_file_index]
    write_line("selected file: "+selected_file);
    return selected_file;                
    
}



read_choice_from_input = function (message,choices){//format YYYYMMDD
    
    while(true){
        
        write_line(message+" ("+choices.join("|")+") : ");       
        var input = read_line().trim().toUpperCase();
        var index = choices.indexOf(input);
        
        if(index!=-1){ 
            
            return input;
            
        }
        
    }
    
}



read_date_from_input = function (){//format YYYYMMDD
    
    while(true){
        
        write_line("enter delivery date format (YYYYMMDD): ");       
        var input = read_line().trim();
        var result = input.match(/^\d{8}$/);

        if(result!=null){

            //formatted_resolution_date = resolution_date = resolution_date.substring(0,4)+resolution_date.substring(5,7)+resolution_date.substring(8,10);
            var year  =  parseInt(input.substring(0,4),10);
            var month =  parseInt(input.substring(4,6),10)-1;
            var day   =  parseInt(input.substring(6,8),10);
            var delivery_date = new Date(year, month, day, 0, 0, 0, 0);

            if(delivery_date && !isNaN(delivery_date.getTime())){

                return delivery_date;

            }

        }
        
    }
    
}


select_files_from_folder = function (folder, message, pattern){//return file paths matching patterns
    
    var file_list = list_files(folder).filter(function(elem){
                   
            return elem.match(pattern);
       
    });
    
    return file_list;                
    
}



//launch Excel and and execute to_do function in its context
do_in_excel = function (to_do){
    
    var excel = new ActiveXObject("Excel.Application");
    
    
    excel.DisplayAlerts = false;

    excel.Visible = false;
    
    try{
    
        to_do(excel);
    
    }catch(ex){
        
        log(ex.message);
        
    }
    
    excel.DisplayAlerts = true;
    
    excel.Quit();
   
    
}


//launch Access and execute to_do function in its context
do_in_access = function (to_do,database_filename){
    
    var access = new ActiveXObject("Access.Application");
    access.UserControl = false;
	access.Visible = false;
    
    
	var database_path = typeof database_filename !== "undefined" ? database_filename
		: (typeof DATABASEPATH !== "undefined" ? DATABASEPATH : "");

    access.OpenCurrentDataBase(CURRENT_FOLDER+"/"+database_path);

    var db = access.CurrentDb();
  
    try{
        
        to_do(db);
        
        
    }catch(ex){
        
        log(ex.message);
        
    }

    
    
    access.CloseCurrentDatabase();
    access.Quit();
    
    
}




//launch Word and and execute to_do function in its context
do_in_word = function (to_do){
    
    var word = new ActiveXObject("Word.Application");
    
    
    word.DisplayAlerts = false;

    word.Visible = false;
    
    try{
    
        to_do(word);
    
    }catch(ex){
        
        log(ex.message);
        
    }
    
    word.DisplayAlerts = true;
    
    word.Quit();
   
    
}

alphabet = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";

// Converts a 1-based column number to its Excel column letter(s), so callers
// are not limited to the 26 columns alphabet.charAt() can address (A..Z).
_column_letter = function(n){
	var letter = "";
	while(n > 0){
		var rem = (n - 1) % 26;
		letter = String.fromCharCode(65 + rem) + letter;
		n = Math.floor((n - 1) / 26);
	}
	return letter;
}


//fill an Excel sheet with data form array of object, labels taken from the first element of the array
fill_sheet = function (sheet,tickets){
    
	
	if(tickets.length == 0){
		
		log("no tickets in the list");
		return;
		
	}
	
	
	
    var ticket = tickets[0];
	
	
    
    var ncolumns = 0;
    for (var k in ticket){
      
        if (ticket.hasOwnProperty(k)){
          
            ncolumns++;                
            
        } 
        
    } 
	
	
    
    
    var area = "B:"+_column_letter(ncolumns);
   
    
    
    sheet.Range("A:A").Copy(sheet.Range(area));
    
    var counter = 0;
    
    for(var key in ticket){
        
        sheet.Cells(1,++counter).Value = key;
        
    }
    for(var i=0; i < tickets.length; i++){
        
        var ticket = tickets[i];
        var counter = 0;
        var row = 2+i;
        for(var key in ticket){
            var value = ticket[key];
            sheet.Cells(row,++counter).Value = value;

        }
    }
    
    
}

/*read elements from excel shhet and returns array of object, keys are taken from first row

if ncolumns is defined takes ncolumns from the sheet. if the label is empty it will use the letter of thecolumnn

if ncolumsn is not defined takes at least 100 columns (override with params.max_columns)


max 3000 rows (override with params.max_rows)
*/
read_sheet_data = function (sheet,ncolumns,params){
	
	var MAX_ROWS 		= (params && params.max_rows)    || 3000;
	var MAX_COLUMNS		= (params && params.max_columns) ||  100;

    //cycling on first row to get labels
    var isempty = false;
    var counter = 1;
    var labels = [];


    var start_row = 1;
    if(params && params.start_row){

        start_row = params.start_row;
    }
    
    
	var check_empty = function(){return true}	
	
	if(ncolumns){
		
		//log('reading data until:' + ncolumns)
		
		check_empty = function(x){
			
			return (function(sheet,row){
			
				for(var i= 1; i<= x; i++){
				
					var v = sheet.Cells(row,i).Value
                    //log(v)
				
					if(typeof v != 'undefined' && v.toString().trim() != ""  ){
					
						return false;
					
					}
				}
			
				return true;
			});
			
		}(ncolumns)
		
		while(counter <= ncolumns){
			
			var cell = sheet.Cells(start_row,counter);
			var value = cell.Value;
			//log(counter+" : "+value)
			if(!value || value.trim() == "" || typeof(value) == 'undefined'){
				
				value =  alphabet.charAt(counter)
				//log('new value: '+value)
			}
			labels.push(value.trim());
			counter = counter +1;
		}
	}
	else{
		
		
		check_empty = function(sheet,row){
			
			var v = sheet.Cells(row,1).Value;

			return !v || String(v).trim() == "" || typeof(v)== 'undefined';
		}
		
		while(!isempty && counter < MAX_COLUMNS){
			var cell = sheet.Cells(start_row,counter);
			var value = cell.Value;
		   
			if(!value || value.trim() == "" ){
				isempty = true;
				break;
			}else{
			   labels.push(value.trim());
			}
			counter = counter +1;
		}
			
	}
    
    log("n� of columns: "+labels.length);
	log(labels);
                        
    var row_counter = start_row+1;
    
    
    
    
    //getting all problems
    isempty = false;
	
    var problems = [];
            
    while(!isempty && row_counter < MAX_ROWS){
       
        var problem = {};
        //log(row_counter+")");
		
		if(check_empty(sheet,row_counter)) { //end of rows
			
			//log('empty list')
			break;
        }
	
		for(var i=1; i<= labels.length; i++){
            
            var value = sheet.Cells(row_counter,i).Value;
			//log(labels[i-1]+":"+value+" - "+typeof value);
			
			if(typeof value == "number"){

				value = value.toString()

			}else if(value instanceof Date){
                
                
               try{
                    
                    var d = new Date("" +value);
                    value = "" + d.getFullYear()+"/"+("0"+(d.getMonth()+1)).slice(-2)+"/"+("0"+d.getDate()).slice(-2);
                    
                }catch(ex){
                    
                    log(ex.message);
                    value="-";
                    
                }
               
                
            }
            else if(typeof value == "undefined" || value==null || !value){
                    value= "";

            }
			else{
				value = String(value);
			}

			problem[labels[i-1]] = value.trim();
            
        }
        
        problems.push(problem);
        
        row_counter= row_counter +1;
    }
    
    
    return problems;
    
};

//save text file at a filepath replacing excel extension with txt extension in the filename
write_report_to_file = function (report_data,filepath){
    
    var d = new Date();
    var report = "";
    var output_folder = typeof OUTPUT_FOLDER !== "undefined" ? OUTPUT_FOLDER : "";

    if(filepath){

        var current_date = "-"+d.getTime();

        var tokens = filepath.split("\\");
        filename_radix = (tokens[tokens.length-1]).replace(/\.xls(x)?/,"");

        report = CURRENT_FOLDER.replace(/\\/g, "/")+output_folder+"/"+filename_radix+ current_date + ".txt";
    }else{

        var current_date = (""+d.getFullYear()).substring(2)+("0"+(d.getMonth()+1)).slice(-2)+("0"+(d.getDate())).slice(-2)+"-"+d.getTime();

        report = CURRENT_FOLDER.replace(/\\/g, "/")+output_folder+"/"+"consistency_analysis_report-"+ current_date + ".txt";

    }
    log("saving report to : " + report);
    write_text_to_file(report_data,report);
    
}



get_current_date_as_excel_text = function(){
	
	var d = new Date();
	return  "'"+d.getFullYear()+[d.getMonth()+1,d.getDate(),d.getHours(),d.getMinutes(),d.getSeconds()].map(function(x){return ("0"+x).slice(-2)}).join("");
	
	
}




//execute a saved query in Access
execute_query = function (saved_query,db){
    
    var dbOpenDynamic       = 16;
	var dbOpenDynaset       = 2;
    var dbOpenForwardOnly   = 8;
    var dbOpenSnapshot      = 4;
    var dbOpenTable         = 1;
    
    
    
    var rs = db.OpenRecordset(saved_query,dbOpenDynaset);
    var counter = 0;
    
    var records = [];
    while(!rs.EOF){
        
        var record = {};
        
        for(var  i= 0; i< rs.Fields.Count; i++){
            
			
			var v = rs.Fields(i).Value;
			
			if(v!=null){
				
				v = v.toString().trim();
				
			}else{
				
				v= "";
				
			}
			
			
			
            record[rs.Fields(i).Name]=v;
			
        }
        records.push(record);
        rs.MoveNext();
        
    }
    
	
	
	
    return records;
}

// ---------------------------------------------------------------------------
// minimist.js
// ---------------------------------------------------------------------------


minimist = function(args,opts){

    opts = opts || {};
    var flags = { bools : {}, strings : {}, unknownFn: null };
    
    


    if (typeof opts['unknown'] === 'function') {
        flags.unknownFn = opts['unknown'];
    }

    if (typeof opts['boolean'] === 'boolean' && opts['boolean']) {
      flags.allBools = true;
    } else {
      [].concat(opts['boolean']).filter(Boolean).forEach(function (key) {
          flags.bools[key] = true;
      });
    }
    
    

    var aliases = {};
    
    for(var key in (opts.alias || {}) ){
        
        aliases[key] = [].concat(opts.alias[key]);
        aliases[key].forEach(function (x) {
            aliases[x] = [key].concat(aliases[key].filter(function (y) {
                return x !== y;
            }));
        });
        
    };

    
    [].concat(opts.string).filter(Boolean).forEach(function (key) {
        flags.strings[key] = true;
        if (aliases[key]) {
            flags.strings[aliases[key]] = true;
        }
     });
     
     
     

    var defaults = opts['default'] || {};
    
    
   
    
    var argv = { _ : [] };
    for(var key in flags.bool){
        
        setArg(key, defaults[key] === undefined ? false : defaults[key]);
        
    }
    
    var notFlags = [];

    if (args.indexOf('--') !== -1) {
        notFlags = args.slice(args.indexOf('--')+1);
        args = args.slice(0, args.indexOf('--'));
    }

    function argDefined(key, arg) {
        return (flags.allBools && /^--[^=]+$/.test(arg)) ||
            flags.strings[key] || flags.bools[key] || aliases[key];
    }

    function setArg (key, val, arg) {
        if (arg && flags.unknownFn && !argDefined(key, arg)) {
            if (flags.unknownFn(arg) === false) return;
        }

        var value = !flags.strings[key] && isNumber(val)
            ? Number(val) : val
        ;
        setKey(argv, key.split('.'), value);
        
        (aliases[key] || []).forEach(function (x) {
            setKey(argv, x.split('.'), value);
        });
    }
    
     

    function setKey (obj, keys, value) {
        var o = obj;
        for (var i = 0; i < keys.length-1; i++) {
            var key = keys[i];
            if (key === '__proto__') return;
            if (o[key] === undefined) o[key] = {};
            if (o[key] === Object.prototype || o[key] === Number.prototype
                || o[key] === String.prototype) o[key] = {};
            if (o[key] === Array.prototype) o[key] = [];
            o = o[key];
        }

        var key = keys[keys.length - 1];
        if (key === '__proto__') return;
        if (o === Object.prototype || o === Number.prototype
            || o === String.prototype) o = {};
        if (o === Array.prototype) o = [];
        if (o[key] === undefined || flags.bools[key] || typeof o[key] === 'boolean') {
            o[key] = value;
        }
        else if (Array.isArray(o[key])) {
            o[key].push(value);
        }
        else {
            o[key] = [ o[key], value ];
        }
    }
    
    function aliasIsBoolean(key) {
      return aliases[key].some(function (x) {
          return flags.bools[x];
      });
    }

    
    for (var i = 0; i < args.length; i++) {
        var arg = args[i];
        if (/^--.+=/.test(arg)) {
            // Using [\s\S] instead of . because js doesn't support the
            // 'dotall' regex modifier. See:
            // http://stackoverflow.com/a/1068308/13216
            var m = arg.match(/^--([^=]+)=([\s\S]*)$/);
            var key = m[1];
            var value = m[2];
            if (flags.bools[key]) {
                value = value !== 'false';
            }
            setArg(key, value, arg);
        }
        else if (/^--no-.+/.test(arg)) {
            var key = arg.match(/^--no-(.+)/)[1];
            setArg(key, false, arg);
        }
        else if (/^--.+/.test(arg)) {
            var key = arg.match(/^--(.+)/)[1];
            var next = args[i + 1];
            if (next !== undefined && !/^-/.test(next)
            && !flags.bools[key]
            && !flags.allBools
            && (aliases[key] ? !aliasIsBoolean(key) : true)) {
                setArg(key, next, arg);
                i++;
            }
            else if (/^(true|false)$/.test(next)) {
                setArg(key, next === 'true', arg);
                i++;
            }
            else {
                setArg(key, flags.strings[key] ? '' : true, arg);
            }
        }
        else if (/^-[^-]+/.test(arg)) {
            var letters = arg.slice(1,-1).split('');
            
            var broken = false;
            for (var j = 0; j < letters.length; j++) {
                var next = arg.slice(j+2);
                
                if (next === '-') {
                    setArg(letters[j], next, arg)
                    continue;
                }
                
                if (/[A-Za-z]/.test(letters[j]) && /=/.test(next)) {
                    setArg(letters[j], next.split('=')[1], arg);
                    broken = true;
                    break;
                }
                
                if (/[A-Za-z]/.test(letters[j])
                && /-?\d+(\.\d*)?(e-?\d+)?$/.test(next)) {
                    setArg(letters[j], next, arg);
                    broken = true;
                    break;
                }
                
                if (letters[j+1] && letters[j+1].match(/\W/)) {
                    setArg(letters[j], arg.slice(j+2), arg);
                    broken = true;
                    break;
                }
                else {
                    setArg(letters[j], flags.strings[letters[j]] ? '' : true, arg);
                }
            }
            
            var key = arg.charAt(arg.length - 1);
            if (!broken && key !== '-') {
                if (args[i+1] && !/^(-|--)[^-]/.test(args[i+1])
                && !flags.bools[key]
                && (aliases[key] ? !aliasIsBoolean(key) : true)) {
                    setArg(key, args[i+1], arg);
                    i++;
                }
                else if (args[i+1] && /^(true|false)$/.test(args[i+1])) {
                    setArg(key, args[i+1] === 'true', arg);
                    i++;
                }
                else {
                    setArg(key, flags.strings[key] ? '' : true, arg);
                }
            }
        }
        else {
            if (!flags.unknownFn || flags.unknownFn(arg) !== false) {
                argv._.push(
                    flags.strings['_'] || !isNumber(arg) ? arg : Number(arg)
                );
            }
            if (opts.stopEarly) {
                argv._.push.apply(argv._, args.slice(i + 1));
                break;
            }
        }
    }
    
   
    
   
   
    for(var key in defaults){
        
        if (!hasKey(argv, key.split('.'))) {
            setKey(argv, key.split('.'), defaults[key]);
            
            (aliases[key] || []).forEach(function (x) {
                setKey(argv, x.split('.'), defaults[key]);
            });
        }
        
    }
    
   
    
    if (opts['--']) {
        argv['--'] = new Array();
        notFlags.forEach(function(key) {
            argv['--'].push(key);
        });
    }
    else {
        notFlags.forEach(function(key) {
            argv._.push(key);
        });
    }

    return argv;
    
    
    
    
};

function hasKey (obj, keys) {
    var o = obj;
    keys.slice(0,-1).forEach(function (key) {
        o = (o[key] || {});
    });

    var key = keys[keys.length - 1];
    return key in o;
}

function isNumber (x) {
    if (typeof x === 'number') return true;
    if (/^0x[0-9a-f]+$/i.test(x)) return true;
    return /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(e[-+]?\d+)?$/.test(x);
}

// ---------------------------------------------------------------------------
// ui.js
// ---------------------------------------------------------------------------

// ui.js  -  HTA (HTML Application) launcher for jscriptowork
//
// An HTA runs inside mshta.exe with a native Windows window and FULL trust:
//   - No browser chrome (no address bar, tabs, menus)
//   - All jscriptowork libs (core, polyfills, system) available
//   - Direct ActiveX access: filesystem, HTTP, COM objects
//   - jsw_return(value) sends a result back to the launching CScript
//
// IMPORTANT: system.js functions that require stdin/stdout (read_line,
// write_line, write) are NOT available inside the HTA because there is
// no WScript.StdIn/StdOut.  All other system.js helpers work normally.
//
// ---------------------------------------------------------------------------
// open_hta(options [, onClose])
// ---------------------------------------------------------------------------
//
// options — plain object (or a bare HTML string treated as {body:...}):
//
//   title   string   Window title bar text.           default: "jscriptowork"
//   width   number   Initial window width  (px).      default: 900
//   height  number   Initial window height (px).      default: 600
//   body    string   HTML to place inside <body>.     default: ""
//   script  string   Extra JS injected after libs.    default: ""
//   style   string   Extra CSS injected in <head>.    default: ""
//
// onClose(result) — optional callback.
//   When provided CScript BLOCKS until the HTA window closes, then calls
//   onClose with whatever value the HTA passed to jsw_return(value).
//   Returns null if the user closed the window without calling jsw_return.
//
// ---------------------------------------------------------------------------
// Inside the HTA the following globals are pre-wired:
//
//   jsw_return(value)           write result JSON and close window
//   http_request(...)           from system.js
//   write_text_to_file(...)     from system.js
//   read_text_file(...)         from system.js
//   file_exists(path)           from system.js
//   write_binary_file(...)      from system.js
//   read_binary_file(...)       from system.js
//   Array / String / Object / Number / Math polyfills (core + polyfills)
//   log(msg)                    appends a line to #_jsw_log (if present)
// ---------------------------------------------------------------------------

open_hta = function(options, onClose) {

    // ---- normalise options ----
    if (typeof options === 'string') { options = { body: options }; }
    options = options || {};

    var title  = options.title  || 'jscriptowork';
    var width  = options.width  || 900;
    var height = options.height || 600;
    var body   = options.body   || '';
    var style  = options.style  || '';
    var script = options.script || '';

    // ---- temp file paths ----
    var fso        = new ActiveXObject("Scripting.FileSystemObject");
    var tmpDir     = fso.GetSpecialFolder(2).Path;
    var ts         = (new Date()).getTime();
    var htaPath    = tmpDir + "\\jsw_hta_"    + ts + ".hta";
    var resultPath = tmpDir + "\\jsw_result_" + ts + ".json";

    // ---- lib URLs (file:/// with forward slashes) ----
    var libBase = "file:///" + ROOT_FOLDER.replace(/\\/g, '/').replace(/\/+$/, '') + "/libs/";

    // ---- escape a string to be safe inside a JS string literal in the HTA ----
    function jsStr(s) {
        return s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r/g, '').replace(/\n/g, '\\n');
    }

    // ---- build the HTA source ----
    var hta = [
        '<!DOCTYPE html>',
        '<html>',
        '<head>',
        '<meta http-equiv="X-UA-Compatible" content="IE=edge"/>',

        // HTA:APPLICATION — controls window chrome
        '<HTA:APPLICATION',
        '  APPLICATIONNAME="' + title + '"',
        '  CAPTION="' + title + '"',
        '  BORDER="thin"',
        '  BORDERSTYLE="normal"',
        '  MINIMIZEBUTTON="yes"',
        '  MAXIMIZEBUTTON="yes"',
        '  SCROLL="auto"',
        '  INNERBORDER="no"',
        '  SHOWINTASKBAR="yes"',
        '  SINGLEINSTANCE="no"',
        '/>',

        '<title>' + title + '</title>',

        // Default + user styles
        '<style>',
        'body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 12px; box-sizing: border-box; }',
        '#_jsw_log { font-family: Consolas, monospace; font-size: 11px; color: #555; border-top: 1px solid #ddd; margin-top: 8px; padding-top: 4px; }',
        style,
        '</style>',

        // 1. WScript compatibility shim — must come BEFORE system.js
        '<script type="text/javascript">',
        'var _script = { echo: function(m){}, StdIn: null, StdOut: null };',
        // log() appends to #_jsw_log if it exists, otherwise no-op
        'function log(m) {',
        '  var el = document.getElementById("_jsw_log");',
        '  if (el) { var d = document.createElement("div"); d.appendChild(document.createTextNode(String(m))); el.appendChild(d); }',
        '}',
        '</script>',

        // 2. jscriptowork libs
        // When _jsw_hta_inline_libs is defined (bundled launcher) the lib source is
        // injected inline so no separate libs/ folder is needed on the target machine.
        (typeof _jsw_hta_inline_libs !== 'undefined'
            ? '<script type="text/javascript">\n' + _jsw_hta_inline_libs + '\n</script>'
            : '<script type="text/javascript" src="' + libBase + 'core.js"></script>\n' +
              '<script type="text/javascript" src="' + libBase + 'polyfills.js"></script>\n' +
              '<script type="text/javascript" src="' + libBase + 'system.js"></script>'),

        // 3. Built-in result channel + window sizing
        '<script type="text/javascript">',
        // jsw_return: write result JSON to temp file then close
        'var _jsw_result_path = "' + jsStr(resultPath) + '";',
        'function jsw_return(value) {',
        '  try {',
        '    var fso = new ActiveXObject("Scripting.FileSystemObject");',
        '    var f = fso.CreateTextFile(_jsw_result_path, true);',
        '    f.Write(JSON.stringify(value !== undefined ? value : null));',
        '    f.Close();',
        '  } catch(e) {}',
        '  window.close();',
        '}',
        // size and centre the window once the DOM is ready
        'window.onload = function() {',
        '  window.resizeTo(' + width + ', ' + height + ');',
        '  window.moveTo(Math.max(0,(screen.availWidth-'  + width  + ')/2),',
        '                Math.max(0,(screen.availHeight-' + height + ')/2));',
        '};',
        '</script>',

        // 4. User-supplied script
        (script ? '<script type="text/javascript">\n' + script + '\n</script>' : ''),

        '</head>',
        '<body>',
        body,
        '</body>',
        '</html>'
    ].join('\n');

    // ---- write & launch ----
    write_text_to_file(hta, htaPath);

    var shell = new ActiveXObject("WScript.Shell");
    var wait  = !!onClose;
    shell.Run('mshta.exe "' + htaPath + '"', 1 /*SW_SHOWNORMAL*/, wait);

    if (onClose) {
        // mshta has exited — read result then clean up
        var result = null;
        if (file_exists(resultPath)) {
            try {
                result = JSON.parse(read_text_file(resultPath));
                delete_file(resultPath);
            } catch(e) {}
        }
        try { if (file_exists(htaPath)) delete_file(htaPath); } catch(e) {}
        onClose(result);
    }
    // non-blocking: temp files cleaned up by OS eventually
};


// ---------------------------------------------------------------------------
// minitest.js
// ---------------------------------------------------------------------------

// minitest.js - Minimal synchronous test framework for CScript / JScript
//
// Usage:
//   load("minitest");
//
//   describe("My feature", function() {
//       it("does something", function() {
//           assert.equal(1 + 1, 2);
//           assert.ok(true);
//       });
//       skip("does something else", "blocked on a known bug, see TODO.md");
//   });
//
//   _test.summary();   // always call at the end to print results

_test = (function() {

    var passed  = 0;
    var failed  = 0;
    var skipped = 0;

    function _print(msg) {
        if (typeof log === 'function') {
            log(msg);
        } else {
            WScript.Echo(msg);
        }
    }

    function _pad(n) {
        return n < 10 ? ' ' + n : '' + n;
    }

    // -----------------------------------------------------------------------
    // Public API
    // -----------------------------------------------------------------------

    var api = {};

    api.describe = function(name, fn) {
        _print('[' + name + ']');
        fn();
    };

    api.it = function(name, fn) {
        try {
            fn();
            passed++;
            _print('  PASS  ' + name);
        } catch (e) {
            failed++;
            _print('  FAIL  ' + name + ': ' + e.message);
        }
    };

    // Records a test as skipped without running it. Use for tests blocked on a
    // known bug (see TODO.md) or on a resource that is not always available
    // (Office, network, interactive desktop). Never fails the run.
    api.skip = function(name, reason) {
        skipped++;
        _print('  SKIP  ' + name + (reason ? ' (' + reason + ')' : ''));
    };

    api.summary = function() {
        var total = passed + failed;
        _print('');
        _print('=====================================');
        if (failed === 0) {
            _print('  ALL TESTS PASSED: ' + passed + ' / ' + total);
        } else {
            _print('  PASSED : ' + passed + ' / ' + total);
            _print('  FAILED : ' + failed + ' / ' + total);
        }
        if (skipped > 0) {
            _print('  SKIPPED: ' + skipped);
        }
        _print('=====================================');
    };

    // Counters, for suites that need to assert on the runner itself.
    api.counts = function() {
        return { passed: passed, failed: failed, skipped: skipped };
    };

    // -----------------------------------------------------------------------
    // assert helpers
    // -----------------------------------------------------------------------

    api.assert = {

        ok: function(val, msg) {
            if (!val) {
                throw new Error(msg || ('Expected truthy but got: ' + val));
            }
        },

        notOk: function(val, msg) {
            if (val) {
                throw new Error(msg || ('Expected falsy but got: ' + val));
            }
        },

        equal: function(actual, expected, msg) {
            if (actual !== expected) {
                throw new Error(msg || ('Expected ' + JSON.stringify(expected) + ' but got ' + JSON.stringify(actual)));
            }
        },

        notEqual: function(actual, expected, msg) {
            if (actual === expected) {
                throw new Error(msg || ('Expected value to differ from ' + JSON.stringify(expected)));
            }
        },

        deepEqual: function(actual, expected, msg) {
            var a = JSON.stringify(actual);
            var e = JSON.stringify(expected);
            if (a !== e) {
                throw new Error(msg || ('Expected ' + e + ' but got ' + a));
            }
        },

        throws: function(fn, msg) {
            var threw = false;
            try {
                fn();
            } catch (e) {
                threw = true;
            }
            if (!threw) {
                throw new Error(msg || 'Expected function to throw but it did not');
            }
        },

        doesNotThrow: function(fn, msg) {
            try {
                fn();
            } catch (e) {
                throw new Error(msg || ('Expected function not to throw but got: ' + e.message));
            }
        }
    };

    return api;
}());

// Expose top-level helpers as globals so test files stay concise
describe = function(name, fn) { _test.describe(name, fn); };
it       = function(name, fn) { _test.it(name, fn); };
skip     = function(name, reason) { _test.skip(name, reason); };
assert   = _test.assert;


// ---------------------------------------------------------------------------
// Script executor
// ---------------------------------------------------------------------------

(function() {
    if (_script.Arguments.Count() === 0) {
        log('Usage: cscript launcher.js <yourscript.js>');
        _script.Quit(1);
    }
    var scriptPath = _script.Arguments(0);
    log('executing:\t' + scriptPath);
    var src = read_all_text_file(scriptPath);
    if (src !== null) { eval(src); }
}());