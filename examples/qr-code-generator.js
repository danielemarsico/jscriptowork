// qr-code-generator.js - display a QR code for a URL in a native window
//
// Prompts for a URL on the console, then opens an HTA window (via
// open_hta()) showing a QR code for it. The QR code image itself comes from
// api.qrserver.com (goqr.me's free, no-signup QR code API) - the <img> tag
// is fetched directly by the HTA's own IE rendering engine, so no image
// bytes ever need to be handled in JScript.
//
// Run via:
//   examples\run.bat qr-code-generator.js
//   cscript.exe bin\launcher.js examples\qr-code-generator.js
//   cscript.exe dist\launcher.js examples\qr-code-generator.js

load("core");
load("polyfills");
load("system");
load("ui");

var DEFAULT_URL = "https://example.com/";

function html_escape(s) {
    return s.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
}

write_line("Enter a URL to encode as a QR code (Enter for " + DEFAULT_URL + "): ");
var url = read_line().trim();
if (url === "") { url = DEFAULT_URL; }

var qrImageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
    encodeURIComponent(url);

log("Encoding: " + url);
log("QR image: " + qrImageUrl);

open_hta(
    {
        title:  "QR Code",
        width:  380,
        height: 460,

        style:
            "body   { font-family: Segoe UI, Arial, sans-serif;" +
            "         display: flex; flex-direction: column;" +
            "         align-items: center; justify-content: center;" +
            "         background: #f0f4f8; margin: 0; }" +
            "img    { border: 1px solid #d0d7de; border-radius: 4px; background: #fff; }" +
            "p      { color: #2d3748; font-size: 12px; word-break: break-all;" +
            "         max-width: 320px; text-align: center; margin: 12px 0; }" +
            "button { padding: 8px 28px; font-size: 14px; cursor: pointer;" +
            "         background: #0078d4; color: #fff;" +
            "         border: none; border-radius: 4px; }",

        body:
            "<img src=\"" + qrImageUrl + "\" width=\"300\" height=\"300\" alt=\"QR code\">" +
            "<p>" + html_escape(url) + "</p>" +
            "<button onclick=\"jsw_return(true)\">Close</button>"
    },
    function(result) {
        log("Window closed.");
    }
);
