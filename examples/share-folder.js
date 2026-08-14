// share-folder.js - zip a folder, upload it anonymously, show the link as a QR
//
// ===========================================================================
//  READ THIS FIRST: the upload is PUBLIC and TEMPORARY.
//
//  The zip is sent to https://0x0.st, an anonymous, no-signup file host. That
//  means:
//    * anyone who has (or guesses) the URL can download the file - there is no
//      password, no login, no access control of any kind;
//    * the file expires on its own, sooner the larger it is (roughly 30 days
//      for a small one, days for a large one), and can disappear at any time;
//    * you are handing your data to a third party you do not control.
//
//  Never point this at anything confidential. The script asks for explicit
//  confirmation before it uploads anything.
// ===========================================================================
//
// What it does:
//   1. Asks for a folder on the console.
//   2. Zips it with tar.exe - built into Windows 10 1803 and later - through
//      exec_command() from libs/win.js.
//   3. Uploads the zip as multipart/form-data via MSXML2.ServerXMLHTTP, with
//      the request body assembled in ADODB.Stream (raw bytes cannot live in a
//      JScript string safely).
//   4. Prints the returned URL, draws it as a QR code on the console, and
//      opens a window showing the same QR so a phone can scan it.
//
// The QR code is generated locally by libs/qrcode.js. No image is fetched, so
// only the upload itself needs the network.
//
// Requirements: Windows 10 1803+ (for tar.exe), network access, and a desktop
// for the window. There is deliberately no test suite for this - it needs all
// three, plus a human to scan the code.
//
// Run via:
//   examples\run.bat share-folder.js
//   cscript.exe bin\launcher.js examples\share-folder.js
//   cscript.exe dist\launcher.js examples\share-folder.js

load("core");
load("polyfills");
load("system");
load("win");
load("qrcode");
load("ui");

var UPLOAD_URL   = "https://0x0.st";        // anonymous, no signup, expiring
var FIELD_NAME   = "file";                  // 0x0.st's form field
var USER_AGENT   = "jscriptowork/1.0 (+https://github.com/danielemarsico/jscriptowork)";
var MAX_MEGABYTES = 100;                    // refuse anything sillier than this

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function fso() { return new ActiveXObject("Scripting.FileSystemObject"); }

function ask(prompt) {
    write_line(prompt);
    return read_line().trim();
}

function strip_trailing_slash(path) {
    return path.replace(/[\\\/]+$/, "");
}

function human_size(bytes) {
    if (bytes < 1024) { return bytes + " B"; }
    if (bytes < 1024 * 1024) { return (bytes / 1024).toFixed(1) + " KB"; }
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function html_escape(s) {
    return s.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
}

// tar.exe ships with Windows 10 1803 and later. Older machines need the
// Shell.Application "compressed folder" trick instead: create a file whose
// first 22 bytes are an empty-zip end-of-central-directory header, open it
// with Shell.Application.NameSpace(), call CopyHere(sourceFolder.Items()), and
// then poll the zip's item count because CopyHere returns immediately and
// copies in the background. That is a lot of fragile machinery, so this
// example requires tar.exe and says so rather than guessing.
function has_tar() {
    return exec_command("where tar.exe").exit_code === 0;
}

// Zips `folder` into `zipPath`. tar is run with -C so the archive contains the
// folder itself rather than the absolute path leading to it.
function zip_folder(folder, zipPath) {
    var f      = fso().GetFolder(folder);
    var parent = f.ParentFolder.Path;
    var name   = f.Name;

    var command = 'tar.exe -a -c -f "' + zipPath + '" -C "' + parent + '" "' + name + '"';
    log("Zipping:  " + command);

    var result = exec_command(command);
    if (result.exit_code !== 0) {
        throw new Error("tar failed (exit " + result.exit_code + "): " +
                        (result.stderr || result.stdout || "no output"));
    }
    if (!file_exists(zipPath)) {
        throw new Error("tar reported success but produced no file at " + zipPath);
    }
}

// ---------------------------------------------------------------------------
// multipart/form-data upload
//
// The body is part text, part raw file bytes, so it is assembled in
// ADODB.Stream rather than in a string: a JScript string is UTF-16 and would
// mangle the bytes on the way through. Text parts are written to a text stream
// as us-ascii, then that stream is switched to binary mode and copied into one
// binary stream along with the file's own bytes.
// ---------------------------------------------------------------------------

function text_as_binary_stream(text) {
    var s = new ActiveXObject("ADODB.Stream");
    s.Type    = 2;              // adTypeText
    s.CharSet = "us-ascii";     // boundaries and headers are ASCII by definition
    s.Open();
    s.WriteText(text);
    s.Position = 0;
    s.Type     = 1;             // adTypeBinary - allowed only at position 0
    return s;
}

function file_as_binary_stream(path) {
    var s = new ActiveXObject("ADODB.Stream");
    s.Type = 1;                 // adTypeBinary
    s.Open();
    s.LoadFromFile(path);
    s.Position = 0;
    return s;
}

function upload_file(url, path, fieldName) {
    var boundary = "----jscriptowork" + (new Date()).getTime();
    var name     = fso().GetFile(path).Name;

    var head = "--" + boundary + "\r\n" +
               'Content-Disposition: form-data; name="' + fieldName +
               '"; filename="' + name + '"' + "\r\n" +
               "Content-Type: application/zip\r\n\r\n";
    var tail = "\r\n--" + boundary + "--\r\n";

    var headStream = null, fileStream = null, tailStream = null, body = null;

    try {
        headStream = text_as_binary_stream(head);
        fileStream = file_as_binary_stream(path);
        tailStream = text_as_binary_stream(tail);

        body = new ActiveXObject("ADODB.Stream");
        body.Type = 1;
        body.Open();
        headStream.CopyTo(body);
        fileStream.CopyTo(body);
        tailStream.CopyTo(body);
        body.Position = 0;

        var request = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
        request.open("POST", url, false);
        request.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary);
        request.setRequestHeader("User-Agent", USER_AGENT);
        request.send(body.Read());

        return { status: request.status, text: String(request.responseText) };

    } finally {
        // Close every stream that was opened, on the error path too.
        if (headStream) { try { headStream.Close(); } catch (e) {} }
        if (fileStream) { try { fileStream.Close(); } catch (e) {} }
        if (tailStream) { try { tailStream.Close(); } catch (e) {} }
        if (body)       { try { body.Close();       } catch (e) {} }
    }
}

// ---------------------------------------------------------------------------
// 1. Pick a folder
// ---------------------------------------------------------------------------

log("=== share-folder ===");
log("");

if (!has_tar()) {
    log("tar.exe was not found on PATH.");
    log("It ships with Windows 10 1803 and later; on an older machine this");
    log("example cannot zip the folder. See the comment above has_tar().");
    WScript.Quit(1);
}

var folder = strip_trailing_slash(ask("Folder to share: "));
if (folder === "") {
    log("Nothing entered - stopping.");
    WScript.Quit(1);
}
if (!folder_exists(folder)) {
    log("No such folder: " + folder);
    WScript.Quit(1);
}

// ---------------------------------------------------------------------------
// 2. Zip it
// ---------------------------------------------------------------------------

var tempFolder = fso().GetSpecialFolder(2).Path;
var zipPath    = tempFolder + "\\jsw_share_" + (new Date()).getTime() + ".zip";

zip_folder(folder, zipPath);

var zipSize = fso().GetFile(zipPath).Size;
log("Created:  " + zipPath + " (" + human_size(zipSize) + ")");
log("");

if (zipSize > MAX_MEGABYTES * 1024 * 1024) {
    log("That is " + human_size(zipSize) + ", over this example's " +
        MAX_MEGABYTES + " MB limit. Stopping.");
    delete_file(zipPath);
    WScript.Quit(1);
}

// ---------------------------------------------------------------------------
// 3. Confirm, then upload
// ---------------------------------------------------------------------------

log("About to upload to " + UPLOAD_URL + ".");
log("");
log("  The file becomes PUBLIC: anyone with the link can download it.");
log("  There is no password and no way to un-share it.");
log("  It expires on its own, sooner the larger it is.");
log("");

var answer = ask("Type 'yes' to upload, anything else to stop: ");
if (answer.toLowerCase() !== "yes") {
    log("Not uploading. The zip is still at " + zipPath);
    WScript.Quit(0);
}

log("Uploading " + human_size(zipSize) + "...");

var response;
try {
    response = upload_file(UPLOAD_URL, zipPath, FIELD_NAME);
} catch (e) {
    log("Upload failed: " + e.message);
    delete_file(zipPath);
    WScript.Quit(1);
}

if (response.status < 200 || response.status >= 300) {
    log("Upload rejected: HTTP " + response.status);
    log(response.text);
    delete_file(zipPath);
    WScript.Quit(1);
}

var link = response.text.replace(/^\s+|\s+$/g, "");
if (!/^https?:\/\//.test(link)) {
    log("The host answered with something that is not a URL:");
    log(response.text);
    delete_file(zipPath);
    WScript.Quit(1);
}

delete_file(zipPath);

log("");
log("Uploaded: " + link);
log("");

// ---------------------------------------------------------------------------
// 4. Show the link as a QR code - encoded locally, no image fetched
// ---------------------------------------------------------------------------

var qr = qr_encode(link, { ec_level: "M" });
log(qr_to_ascii(qr));

open_hta(
    {
        title:  "Share folder",
        width:  400,
        height: 520,

        style:
            "body   { font-family: Segoe UI, Arial, sans-serif;" +
            "         display: flex; flex-direction: column;" +
            "         align-items: center; justify-content: center;" +
            "         background: #f0f4f8; margin: 0; }" +
            "table  { border: 1px solid #d0d7de; border-radius: 4px; }" +
            "p      { color: #2d3748; font-size: 12px; word-break: break-all;" +
            "         max-width: 340px; text-align: center; margin: 12px 0; }" +
            "small  { color: #718096; font-size: 11px; }" +
            "button { padding: 8px 28px; font-size: 14px; cursor: pointer;" +
            "         background: #0078d4; color: #fff;" +
            "         border: none; border-radius: 4px; }",

        body:
            qr_to_html(qr, { scale: 6, quiet_zone: 3 }) +
            "<p>" + html_escape(link) + "</p>" +
            "<p><small>public and temporary &middot; " + html_escape(human_size(zipSize)) +
            " &middot; QR generated offline</small></p>" +
            "<button onclick=\"jsw_return(true)\">Close</button>"
    },
    function(result) {
        log("Window closed.");
    }
);
