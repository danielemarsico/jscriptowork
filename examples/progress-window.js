// progress-window.js - stream progress out of an HTA while it is still open
//
// open_hta() has always been able to hand ONE value back to CScript, at the
// end, via jsw_return(). options.onProgress adds a live channel: the HTA calls
// jsw_progress(value) as often as it likes and CScript receives each value
// while the window is still on screen.
//
// This example shows a progress bar working through five steps. The window
// reports each step as it completes; CScript prints them as they arrive, then
// prints the final jsw_return value once the window closes.
//
// Under the hood: with onProgress the window is launched detached (rather than
// blocking inside shell.Run) and CScript polls a temp file the HTA appends to.
//
// Run via:
//   examples\run.bat progress-window.js
//   cscript.exe bin\launcher.js examples\progress-window.js
//   cscript.exe dist\launcher.js examples\progress-window.js

load("core");
load("polyfills");
load("system");
load("ui");

log("Opening progress window...");
log("");

var started = (new Date()).getTime();

open_hta(
    {
        title:  "Working...",
        width:  460,
        height: 220,

        style:
            "body   { font-family: Segoe UI, Arial, sans-serif; background: #f0f4f8;" +
            "         margin: 0; padding: 24px; }" +
            "h1     { color: #2d3748; font-size: 1.2em; margin: 0 0 16px; }" +
            "#bar   { height: 20px; background: #dde3ea; border-radius: 10px;" +
            "         overflow: hidden; }" +
            "#fill  { height: 100%; width: 0; background: #0078d4; }" +
            "#label { color: #4a5568; font-size: 0.9em; margin-top: 12px; }",

        body:
            "<h1>Processing records</h1>" +
            "<div id='bar'><div id='fill'></div></div>" +
            "<div id='label'>starting...</div>",

        // Inside the HTA. window.setTimeout exists here (this is the IE engine,
        // not CScript), so the steps can be spread out over time instead of
        // finishing instantly.
        script:
            "var step = 0;" +
            "var TOTAL = 5;" +
            "function next() {" +
            "  step++;" +
            "  document.getElementById('fill').style.width = ((step / TOTAL) * 100) + '%';" +
            "  document.getElementById('label').innerHTML = 'step ' + step + ' of ' + TOTAL;" +
            // Hand a structured value back to CScript, right now.
            "  jsw_progress({ step: step, total: TOTAL, label: 'finished step ' + step });" +
            "  if (step < TOTAL) {" +
            "    window.setTimeout(next, 600);" +
            "  } else {" +
            "    window.setTimeout(function() { jsw_return('all ' + TOTAL + ' steps done'); }, 400);" +
            "  }" +
            "}" +
            "window.onload = function() { window.setTimeout(next, 400); };",

        // Called once per jsw_progress, while the window is open.
        onProgress: function(value) {
            var elapsed = ((new Date()).getTime() - started) / 1000;
            log("  [" + elapsed.toFixed(1) + "s] " +
                value.step + "/" + value.total + " - " + value.label);
        },

        poll_ms: 100        // check for new progress this often (default 200)
    },

    // Called once, after the window closes, with the jsw_return value.
    function(result) {
        log("");
        log("Window closed. Result: " + JSON.stringify(result));
    }
);

log("");
log("Note the timestamps: each line printed while the window was still open,");
log("not in one burst at the end.");
