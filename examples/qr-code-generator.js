// qr-code-generator.js - display a QR code for a URL in a native window
//
// Prompts for a URL on the console, prints the QR code as ASCII art, then
// opens an HTA window (via open_hta()) showing it.
//
// The QR code is generated locally by libs/qrcode.js - a from-scratch ES3
// encoder. Nothing is fetched: this example used to point an <img> at
// api.qrserver.com, which meant no network, no QR code. It now works with the
// network unplugged. The window renders the symbol as a <table> of coloured
// cells, which the old IE engine behind an HTA draws reliably (no SVG, no
// canvas, no image bytes to handle in JScript).
//
// Run via:
//   examples\run.bat qr-code-generator.js
//   cscript.exe bin\launcher.js examples\qr-code-generator.js
//   cscript.exe dist\launcher.js examples\qr-code-generator.js

load("core");
load("polyfills");
load("system");
load("qrcode");
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

// Level M corrects roughly 15% damage - the usual choice for a screen.
var qr = qr_encode(url, { ec_level: "M" });

log("Encoding: " + url);
log("Symbol:   version " + qr.version + ", " + qr.size + "x" + qr.size +
    " modules, level " + qr.ec_level + ", mask " + qr.mask + ", " + qr.mode + " mode");
log("");
log(qr_to_ascii(qr));

open_hta(
    {
        title:  "QR Code",
        width:  380,
        height: 480,

        style:
            "body   { font-family: Segoe UI, Arial, sans-serif;" +
            "         display: flex; flex-direction: column;" +
            "         align-items: center; justify-content: center;" +
            "         background: #f0f4f8; margin: 0; }" +
            "table  { border: 1px solid #d0d7de; border-radius: 4px; }" +
            "p      { color: #2d3748; font-size: 12px; word-break: break-all;" +
            "         max-width: 320px; text-align: center; margin: 12px 0; }" +
            "small  { color: #718096; font-size: 11px; }" +
            "button { padding: 8px 28px; font-size: 14px; cursor: pointer;" +
            "         background: #0078d4; color: #fff;" +
            "         border: none; border-radius: 4px; }",

        body:
            qr_to_html(qr, { scale: 6, quiet_zone: 3 }) +
            "<p>" + html_escape(url) + "</p>" +
            "<p><small>version " + qr.version + " &middot; level " + qr.ec_level +
            " &middot; generated offline</small></p>" +
            "<button onclick=\"jsw_return(true)\">Close</button>"
    },
    function(result) {
        log("Window closed.");
    }
);
