// esp32-decimen-loader.js - generate an ESP32 HID sketch that types a file
// onto an air-gapped laptop, in verified chunks.
//
// The problem: an air-gapped laptop needs one file on it - here, decimen's
// standalone sender - but you won't plug in a USB stick (might be infected)
// and there is no network. An ESP32 in USB-HID (keyboard) mode can type the
// file in as keystrokes: it carries no filesystem the laptop could autorun,
// and every command it types is visible on screen. This tool turns the file
// into that sketch.
//
// How it works:
//   1. The file is zipped (tar.exe, built into Windows 10 1803+) and base64'd,
//      which cuts the character count to about a third.
//   2. The base64 is split into chunks. Each chunk is typed into its own file
//      and checked on the spot with certutil + findstr, so a mistyped chunk is
//      caught where it happened, not after the whole transfer.
//   3. The chunks are reassembled, decoded, unzipped, and the final file is
//      SHA-256 checked against the original.
//
// Everything the laptop runs uses only tools already on Windows: certutil,
// tar, copy, findstr. Nothing is pre-installed.
//
// The emitted commands contain NO double/single quotes, backticks or carets,
// so they type correctly under both the US and US-International layouts (whose
// only difference that matters here is that ' " ` ~ ^ are dead keys). base64
// itself is all A-Za-z0-9+/=, none of which are layout-sensitive on those two.
//
// Usage (on an internet-connected Windows PC, NOT the air-gapped one):
//   cscript bin\launcher.js examples\esp32-decimen-loader.js decimen-sender.html
//   cscript bin\launcher.js examples\esp32-decimen-loader.js sender.html --out C:\out --chunk 1800 --keyms 8
//
//   --out <dir>     where to write the sketch + manifest   (default: .\esp32-loader)
//   --chunk <n>     base64 characters per chunk            (default: 1800)
//   --keyms <n>     ESP32 delay between keys, ms            (default: 8)
//   --linems <n>    ESP32 delay between lines, ms           (default: 120)
//   --zip <path>    use this pre-made .zip instead of running tar (for testing
//                   off Windows; on Windows just pass the .html and tar runs)
//
// Output in <dir>:
//   esp32-decimen-loader.ino   flash this to an ESP32-S2/S3 (native USB)
//   manifest.txt               expected hashes + the run procedure
//
// Run through bin/launcher.js so load() is available.

load("core");
load("polyfills");
load("system");
load("win");
load("base64");
load("crypto");
load("minimist");

// --- arguments -------------------------------------------------------------

var argv = [];
for (var a = 0; a < WScript.Arguments.Count(); a++) { argv.push(WScript.Arguments(a)); }
// Under bin/launcher.js, Arguments(0) is this script's own path.
if (argv.length > 0 && /esp32-decimen-loader\.js$/i.test(argv[0])) { argv = argv.slice(1); }

var args = minimist(argv, {
    "string": ["out", "zip"],
    "default": { "out": ".\\esp32-loader", "chunk": 1800, "keyms": 8, "linems": 120 }
});

var fso = new ActiveXObject("Scripting.FileSystemObject");

var inputArg = args._ && args._.length > 0 ? String(args._[0]) : "";
if (inputArg === "") {
    log("Usage: cscript bin\\launcher.js examples\\esp32-decimen-loader.js <file.html> [--out dir] [--chunk n]");
    WScript.Quit(1);
}

var inputPath = fso.GetAbsolutePathName(inputArg);
if (!fso.FileExists(inputPath)) {
    log("No such file: " + inputPath);
    WScript.Quit(1);
}

var chunkChars = Number(args.chunk) || 1800;
var keyMs      = Number(args.keyms) || 8;
var lineMs     = Number(args.linems) || 120;
var outDir     = fso.GetAbsolutePathName(String(args.out));
var baseName   = inputPath.slice(inputPath.lastIndexOf("\\") + 1);

if (!fso.FolderExists(outDir)) { fso.CreateFolder(outDir); }

// --- 1. zip the file -------------------------------------------------------
//
// tar.exe on Windows is bsdtar/libarchive, so `-a -c -f out.zip` writes a real
// deflate zip whose single entry keeps the original name. GNU tar (Linux) does
// not do zip, which is why --zip exists for testing off Windows.

var zipPath = outDir + "\\_payload.zip";

if (args.zip) {
    var suppliedZip = fso.GetAbsolutePathName(String(args.zip));
    if (!fso.FileExists(suppliedZip)) { log("No such --zip file: " + suppliedZip); WScript.Quit(1); }
    if (fso.FileExists(zipPath)) { fso.DeleteFile(zipPath); }
    fso.CopyFile(suppliedZip, zipPath);
    log("Using supplied zip: " + suppliedZip);
} else {
    if (fso.FileExists(zipPath)) { fso.DeleteFile(zipPath); }
    var parent = inputPath.slice(0, inputPath.lastIndexOf("\\"));
    var cmd = 'tar.exe -a -c -f "' + zipPath + '" -C "' + parent + '" "' + baseName + '"';
    log("Zipping:  " + cmd);
    var r = exec_command(cmd);
    if (r.exit_code !== 0 || !fso.FileExists(zipPath)) {
        log("tar failed (exit " + r.exit_code + "): " + (r.stderr || r.stdout || "no output"));
        log("tar.exe needs Windows 10 1803 or later. On an older machine, zip the file");
        log("yourself and pass it with --zip.");
        WScript.Quit(1);
    }
}

// --- 2. base64 the zip -----------------------------------------------------

var zipBytes = read_binary_file(zipPath);
var htmlBytes = read_binary_file(inputPath);
var b64 = base64_encode_bytes(zipBytes);   // standard alphabet, no line wrapping

// --- 3. split into chunks and hash each ------------------------------------
//
// Each chunk is typed as `>pNN.b64 echo <chunk>` which writes the chunk plus a
// trailing CRLF. So the file certutil hashes is chunk + "\r\n" - the expected
// token must be computed over exactly those bytes.

function first(hex, n) { return hex.substring(0, n); }
var CHUNK_TOKEN_LEN = 16;   // 16 hex = 64 bits; plenty to catch a typo

var chunks = [];
for (var i = 0; i < b64.length; i += chunkChars) { chunks.push(b64.substr(i, chunkChars)); }
var n = chunks.length;
var width = String(n - 1).length; if (width < 2) { width = 2; }

function pad(num) {
    var s = String(num);
    while (s.length < width) { s = "0" + s; }
    return s;
}

var tokens = [];
for (var c = 0; c < n; c++) {
    tokens.push(first(sha256(chunks[c] + "\r\n"), CHUNK_TOKEN_LEN));
}

// canary: every base64 symbol, typed once, so a wrong layout is caught before
// the bulk is typed.
var ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
var canaryToken = first(sha256(ALPHABET + "\r\n"), CHUNK_TOKEN_LEN);

// final file: full SHA-256 of the original bytes, checked after unzip.
var fullHash = sha256_bytes(htmlBytes);

// --- self-check: the chunks really reconstruct the zip ---------------------

var rejoined = chunks.join("");
var decoded  = base64_decode_bytes(rejoined);
var ok = decoded.length === zipBytes.length;
if (ok) { for (var d = 0; d < decoded.length; d++) { if (decoded[d] !== zipBytes[d]) { ok = false; break; } } }
if (!ok) {
    log("SELF-CHECK FAILED: chunks do not decode back to the zip. Aborting.");
    WScript.Quit(1);
}

// --- 4. build the command lines --------------------------------------------
//
// Redirect-first (`>file echo x`) avoids echo's trailing-space quirk. No quotes
// anywhere, so the lines are dead-key-safe on US and US-International.

var stage1 = [];   // preamble + canary: typed on the first button press
var stage2 = [];   // payload + assembly + final check: second button press

stage1.push("cd /d %TEMP%");
stage1.push("rd /s /q dcm 2>nul");
stage1.push("mkdir dcm");
stage1.push("cd dcm");
stage1.push(">canary.txt echo " + ALPHABET);
stage1.push("certutil -hashfile canary.txt SHA256 | findstr /i " + canaryToken +
            " && echo CANARY OK || echo CANARY BAD");

for (var k = 0; k < n; k++) {
    var name = "p" + pad(k) + ".b64";
    stage2.push(">" + name + " echo " + chunks[k]);
    stage2.push("certutil -hashfile " + name + " SHA256 | findstr /i " + tokens[k] +
                " && echo P" + pad(k) + " OK || echo P" + pad(k) + " BAD");
}

var joinList = [];
for (var j = 0; j < n; j++) { joinList.push("p" + pad(j) + ".b64"); }
stage2.push("copy /b " + joinList.join("+") + " sender.b64");
stage2.push("certutil -decode sender.b64 sender.zip");
stage2.push("tar -xf sender.zip");
stage2.push("certutil -hashfile " + baseName + " SHA256 | findstr /i " + fullHash +
            " && echo FILE OK || echo FILE BAD");

// --- 5. emit the .ino ------------------------------------------------------

function cEscape(s) {
    return s.replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
}
function cArray(name, lines) {
    var out = ["const char* " + name + "[] = {"];
    for (var i = 0; i < lines.length; i++) {
        out.push('  "' + cEscape(lines[i]) + '"' + (i < lines.length - 1 ? "," : ""));
    }
    out.push("};");
    out.push("const int " + name + "_COUNT = " + lines.length + ";");
    return out.join("\r\n");
}

var ino = [];
ino.push("// esp32-decimen-loader.ino - generated by jscriptowork esp32-decimen-loader.js");
ino.push("// Types " + baseName + " onto an air-gapped laptop as verified chunks.");
ino.push("// Target: ESP32-S2 or ESP32-S3 (native USB). Board setting: USB Mode = 'USB-OTG',");
ino.push("//         or select Tools > USB CDC On Boot as your core requires.");
ino.push("//");
ino.push("// Procedure:");
ino.push("//   1. Flash this to the ESP32 on your ONLINE PC.");
ino.push("//   2. Plug the ESP32 into the air-gapped laptop. Open a Command Prompt and");
ino.push("//      click inside it so it has focus.");
ino.push("//   3. Press the BOOT button once. Stage 1 types the setup + a canary line.");
ino.push("//      Read the screen: it must say CANARY OK. If it says CANARY BAD, the");
ino.push("//      keyboard layout is wrong - fix it and press BOOT again.");
ino.push("//   4. Press BOOT a second time. Stage 2 types the payload and verifies each");
ino.push("//      chunk (Pnn OK / Pnn BAD), then assembles and prints FILE OK.");
ino.push("//   5. Any Pnn BAD: press the RESET button, then repeat from step 3 - every");
ino.push("//      chunk is rewritten from scratch, so a clean re-run fixes it.");
ino.push("");
ino.push("#include \"USB.h\"");
ino.push("#include \"USBHIDKeyboard.h\"");
ino.push("");
ino.push("USBHIDKeyboard Keyboard;");
ino.push("const int BOOT_BTN = 0;      // BOOT button on most ESP32-S2/S3 dev boards");
ino.push("const int KEY_MS  = " + keyMs + ";       // delay between keystrokes");
ino.push("const int LINE_MS = " + lineMs + ";      // delay after each Enter");
ino.push("int stage = 0;");
ino.push("");
ino.push(cArray("STAGE1", stage1));
ino.push("");
ino.push(cArray("STAGE2", stage2));
ino.push("");
ino.push("void typeLine(const char* s) {");
ino.push("  for (const char* p = s; *p; ++p) { Keyboard.write((uint8_t)*p); delay(KEY_MS); }");
ino.push("  Keyboard.write((uint8_t)'\\n');");
ino.push("  delay(LINE_MS);");
ino.push("}");
ino.push("");
ino.push("void typeAll(const char** lines, int count) {");
ino.push("  for (int i = 0; i < count; ++i) { typeLine(lines[i]); }");
ino.push("}");
ino.push("");
ino.push("bool pressed() {");
ino.push("  if (digitalRead(BOOT_BTN) == LOW) { delay(40);");
ino.push("    if (digitalRead(BOOT_BTN) == LOW) { while (digitalRead(BOOT_BTN) == LOW) delay(10); return true; } }");
ino.push("  return false;");
ino.push("}");
ino.push("");
ino.push("void setup() {");
ino.push("  pinMode(BOOT_BTN, INPUT_PULLUP);");
ino.push("  Keyboard.begin();");
ino.push("  USB.begin();");
ino.push("  delay(1500);");
ino.push("}");
ino.push("");
ino.push("void loop() {");
ino.push("  if (pressed()) {");
ino.push("    if (stage == 0)      { typeAll(STAGE1, STAGE1_COUNT); stage = 1; }");
ino.push("    else if (stage == 1) { typeAll(STAGE2, STAGE2_COUNT); stage = 2; }");
ino.push("  }");
ino.push("  delay(20);");
ino.push("}");
ino.push("");

var inoPath = outDir + "\\esp32-decimen-loader.ino";
write_text_to_file(ino.join("\r\n"), inoPath);

// --- 6. emit the manifest --------------------------------------------------

var man = [];
man.push("Decimen air-gap loader - manifest");
man.push("=================================");
man.push("");
man.push("source file        : " + baseName);
man.push("source SHA-256      : " + fullHash);
man.push("source size        : " + htmlBytes.length + " bytes");
man.push("zipped size        : " + zipBytes.length + " bytes");
man.push("base64 characters   : " + b64.length);
man.push("chunks             : " + n + " x " + chunkChars + " chars");
man.push("layout             : US / US-International (no dead-key characters emitted)");
man.push("");
man.push("Estimated typing time (payload only):");
man.push("  at " + Math.round(1000 / keyMs) + " keys/s : ~" +
         (Math.round(b64.length / (1000 / keyMs) / 6) / 10) + " min");
man.push("");
man.push("On the air-gapped laptop this ends with:  FILE OK");
man.push("If you see FILE BAD or any Pnn BAD, reset the ESP32 and re-run - each");
man.push("chunk file is overwritten, so a clean re-run is safe and idempotent.");
man.push("");
man.push("The received file lands in  %TEMP%\\dcm\\" + baseName);
man.push("Verify it yourself any time with:");
man.push("  certutil -hashfile %TEMP%\\dcm\\" + baseName + " SHA256");
man.push("and confirm it matches the source SHA-256 above.");
man.push("");
man.push("Canary expected token : " + canaryToken);

write_text_to_file(man.join("\r\n"), outDir + "\\manifest.txt");

// clean up the working zip
if (fso.FileExists(zipPath)) { fso.DeleteFile(zipPath); }

log("");
log("SELF-CHECK OK: chunks decode back to the zip byte-for-byte.");
log("Wrote " + inoPath);
log("Wrote " + outDir + "\\manifest.txt");
log("");
log("  source SHA-256 : " + fullHash);
log("  chunks         : " + n + " (" + chunkChars + " chars each)");
log("  base64 chars   : " + b64.length);
