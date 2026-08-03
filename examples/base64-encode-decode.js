// base64-encode-decode.js - base64 encode / decode via libs/base64.js
//
// JScript has no built-in btoa()/atob(); libs/base64.js provides a from-
// scratch ES3 implementation: base64_encode/decode for UTF-8 strings, and
// base64_encode_bytes/decode_bytes for raw byte arrays (0-255), the same
// convention libs/system.js uses for write_binary_file()/read_binary_file().
//
// Part 1 encodes/decodes a short string.
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
load("base64");

// ---------------------------------------------------------------------------
// Part 1 - encode / decode a short string
// ---------------------------------------------------------------------------

log("--- Part 1: string encode/decode ---");

var message = "jscriptowork can run on plain CScript!";
var encoded = base64_encode(message);
var decoded = base64_decode(encoded);

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

var base64Text = base64_encode(read_text_file(sourcePath));
log("File as base64: " + base64Text);

write_text_to_file(base64_decode(base64Text), outputPath);

log("File round-trip OK: " + (read_text_file(outputPath) === sourceContent));

delete_file(sourcePath);
delete_file(outputPath);
