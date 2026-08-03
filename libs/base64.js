// base64.js - base64 encode/decode for JScript (no native btoa()/atob())
// ---------------------------------------------------------------------------
//   base64_encode_bytes(byteArray)   byte array (integers 0-255) -> base64 string
//   base64_decode_bytes(str)         base64 string -> byte array (integers 0-255)
//   base64_encode(str)               UTF-8 string -> base64 string
//   base64_decode(str)               base64 string -> UTF-8 string
//
// Pure ES3 syntax - no arrow functions, let/const, template literals, etc.
// All public functions are assigned without 'var' so they survive load()'s eval scope.
//
// NOTE: No IIFE wrapper - JScript's eval() does not reliably preserve closure
// scope for function declarations inside IIFEs (see the note at the top of
// crypto.js). Private helpers use the _base64_ prefix as a namespace instead.
// ---------------------------------------------------------------------------

_base64_alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

// Encode a JS string to a UTF-8 byte array (handles BMP + surrogate pairs).
_base64_to_utf8 = function(str) {
    var bytes = [], i = 0;
    while (i < str.length) {
        var c = str.charCodeAt(i++);
        if (c >= 0xD800 && c <= 0xDBFF && i < str.length) {
            var lo = str.charCodeAt(i);
            if (lo >= 0xDC00 && lo <= 0xDFFF) { c = ((c - 0xD800) << 10) + (lo - 0xDC00) + 0x10000; i++; }
        }
        if (c < 0x80) {
            bytes.push(c);
        } else if (c < 0x800) {
            bytes.push((c >> 6) | 0xC0);
            bytes.push((c & 0x3F) | 0x80);
        } else if (c < 0x10000) {
            bytes.push((c >> 12) | 0xE0);
            bytes.push(((c >> 6) & 0x3F) | 0x80);
            bytes.push((c & 0x3F) | 0x80);
        } else {
            bytes.push((c >> 18) | 0xF0);
            bytes.push(((c >> 12) & 0x3F) | 0x80);
            bytes.push(((c >> 6) & 0x3F) | 0x80);
            bytes.push((c & 0x3F) | 0x80);
        }
    }
    return bytes;
};

// Decode a UTF-8 byte array back to a JS string (handles multi-byte + surrogate pairs).
_base64_from_utf8 = function(bytes) {
    var out = "", i = 0;
    while (i < bytes.length) {
        var b0 = bytes[i++];
        var cp;
        if (b0 < 0x80) {
            cp = b0;
        } else if ((b0 & 0xE0) === 0xC0) {
            cp = ((b0 & 0x1F) << 6) | (bytes[i++] & 0x3F);
        } else if ((b0 & 0xF0) === 0xE0) {
            cp = ((b0 & 0x0F) << 12) | ((bytes[i++] & 0x3F) << 6) | (bytes[i++] & 0x3F);
        } else {
            cp = ((b0 & 0x07) << 18) | ((bytes[i++] & 0x3F) << 12) |
                 ((bytes[i++] & 0x3F) << 6) | (bytes[i++] & 0x3F);
        }
        if (cp > 0xFFFF) {
            cp -= 0x10000;
            out += String.fromCharCode(0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF));
        } else {
            out += String.fromCharCode(cp);
        }
    }
    return out;
};

base64_encode_bytes = function(bytes) {
    var out = "";
    var i = 0;

    for (; i + 2 < bytes.length; i += 3) {
        var n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        out += _base64_alphabet.charAt((n >>> 18) & 63) +
               _base64_alphabet.charAt((n >>> 12) & 63) +
               _base64_alphabet.charAt((n >>> 6)  & 63) +
               _base64_alphabet.charAt(n & 63);
    }

    var remaining = bytes.length - i;
    if (remaining === 1) {
        var n1 = bytes[i] << 16;
        out += _base64_alphabet.charAt((n1 >>> 18) & 63) +
               _base64_alphabet.charAt((n1 >>> 12) & 63) + "==";
    } else if (remaining === 2) {
        var n2 = (bytes[i] << 16) | (bytes[i + 1] << 8);
        out += _base64_alphabet.charAt((n2 >>> 18) & 63) +
               _base64_alphabet.charAt((n2 >>> 12) & 63) +
               _base64_alphabet.charAt((n2 >>> 6)  & 63) + "=";
    }

    return out;
};

base64_decode_bytes = function(str) {
    var lookup = {};
    for (var i = 0; i < _base64_alphabet.length; i++) {
        lookup[_base64_alphabet.charAt(i)] = i;
    }

    var bytes  = [];
    var buffer = 0;
    var bits   = 0;

    for (var j = 0; j < str.length; j++) {
        var c = str.charAt(j);
        if (c === "=" || !(c in lookup)) { continue; }
        buffer = (buffer << 6) | lookup[c];
        bits += 6;
        if (bits >= 8) {
            bits -= 8;
            bytes.push((buffer >>> bits) & 0xFF);
        }
    }

    return bytes;
};

// Convenience wrappers for plain text (UTF-8 safe both ways).
base64_encode = function(str) {
    return base64_encode_bytes(_base64_to_utf8(str));
};

base64_decode = function(str) {
    return _base64_from_utf8(base64_decode_bytes(str));
};
