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
// console shim
// ---------------------------------------------------------------------------
//
// The console shim now lives in its own lib so it can be loaded on its own:
//
//     load("console");
//
// polyfills.js deliberately no longer defines it - this file is the
// language-level compatibility layer (Array/String/Object/Number/Math/Date/
// Function) and nothing else.
