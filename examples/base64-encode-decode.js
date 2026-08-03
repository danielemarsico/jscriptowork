// base64-encode-decode.js - base64 encode / decode from scratch
//
// JScript has no built-in btoa()/atob(), so this example implements standard
// base64 as plain ES3, working on byte arrays (0-255) - the same convention
// libs/system.js uses for write_binary_file() / read_binary_file().
//
// Part 1 encodes/decodes a short ASCII string.
// Part 2 writes a small file to disk, base64-encodes its contents, decodes
// them back, writes the result to a second file, and verifies the two files
// are byte-for-byte identical.
//
// Run via:
//   examples\run.bat base64-encode-decode.js
//   cscript.exe bin\launcher.js examples\base64-encode-decode.js
//   cscript.exe dist\launcher.js examples\base64-encode-decode.js

load("core");
load("polyfills");
load("system");

// ---------------------------------------------------------------------------
// base64 encode / decode over byte arrays
// ---------------------------------------------------------------------------

var BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function base64_encode_bytes(bytes) {
    var out = "";
    var i = 0;

    for (; i + 2 < bytes.length; i += 3) {
        var n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        out += BASE64_ALPHABET.charAt((n >>> 18) & 63) +
               BASE64_ALPHABET.charAt((n >>> 12) & 63) +
               BASE64_ALPHABET.charAt((n >>> 6)  & 63) +
               BASE64_ALPHABET.charAt(n & 63);
    }

    var remaining = bytes.length - i;
    if (remaining === 1) {
        var n1 = bytes[i] << 16;
        out += BASE64_ALPHABET.charAt((n1 >>> 18) & 63) +
               BASE64_ALPHABET.charAt((n1 >>> 12) & 63) + "==";
    } else if (remaining === 2) {
        var n2 = (bytes[i] << 16) | (bytes[i + 1] << 8);
        out += BASE64_ALPHABET.charAt((n2 >>> 18) & 63) +
               BASE64_ALPHABET.charAt((n2 >>> 12) & 63) +
               BASE64_ALPHABET.charAt((n2 >>> 6)  & 63) + "=";
    }

    return out;
}

function base64_decode_bytes(str) {
    var lookup = {};
    for (var i = 0; i < BASE64_ALPHABET.length; i++) {
        lookup[BASE64_ALPHABET.charAt(i)] = i;
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
}

function string_to_bytes(str) {
    var bytes = [];
    for (var i = 0; i < str.length; i++) { bytes.push(str.charCodeAt(i) & 0xFF); }
    return bytes;
}

function bytes_to_string(bytes) {
    var chars = [];
    for (var i = 0; i < bytes.length; i++) { chars.push(String.fromCharCode(bytes[i])); }
    return chars.join("");
}

// ---------------------------------------------------------------------------
// Part 1 - encode / decode a short string
// ---------------------------------------------------------------------------

log("--- Part 1: string encode/decode ---");

var message = "jscriptowork can run on plain CScript!";
var encoded = base64_encode_bytes(string_to_bytes(message));
var decoded = bytes_to_string(base64_decode_bytes(encoded));

log("Original:  " + message);
log("Base64:    " + encoded);
log("Decoded:   " + decoded);
log("Round-trip OK: " + (decoded === message));

// ---------------------------------------------------------------------------
// Part 2 - round-trip a file's contents through base64
// ---------------------------------------------------------------------------

log("");
log("--- Part 2: file encode/decode ---");

var fso        = new ActiveXObject("Scripting.FileSystemObject");
var tempFolder = fso.GetSpecialFolder(2).Path; // 2 = TemporaryFolder
var sourcePath = tempFolder + "\\jsw_base64_example_source.txt";
var outputPath = tempFolder + "\\jsw_base64_example_roundtrip.txt";

var sourceContent = "Some sample file content, round-tripped through base64.\n";
write_text_to_file(sourceContent, sourcePath);

var base64Text = base64_encode_bytes(string_to_bytes(read_text_file(sourcePath)));
log("File as base64: " + base64Text);

var decodedContent = bytes_to_string(base64_decode_bytes(base64Text));
write_text_to_file(decodedContent, outputPath);

log("File round-trip OK: " + (read_text_file(outputPath) === sourceContent));

delete_file(sourcePath);
delete_file(outputPath);
