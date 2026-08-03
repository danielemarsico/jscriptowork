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
// onClose(result) - optional callback.
//   When provided CScript BLOCKS until the HTA window closes, then calls
//   onClose with whatever value the HTA passed to jsw_return(value).
//   Returns null if the user closed the window without calling jsw_return.
//
// options.onProgress(value) - optional callback.
//   Receives whatever the HTA passes to jsw_progress(value), WHILE the window
//   is still open, instead of having to wait for the window to close.
//
//   Progress changes how the window is launched. Without it, CScript blocks
//   inside shell.Run() and cannot do anything until mshta exits. With it, the
//   window is started detached and CScript polls a temp file the HTA appends
//   to, dispatching each new value as it appears. onClose, if given, still
//   fires once at the end with the jsw_return value.
//
//   options.poll_ms    how often to check for new progress.  default: 200
//   options.max_wait   give up after this many ms (0 = wait
//                      forever, which is the default).       default: 0
//
//     open_hta({
//         body:   '<h1>Working...</h1><div id="_jsw_log"></div>',
//         script: 'window.onload = function() {' +
//                 '  for (var i = 1; i <= 3; i++) { jsw_progress(i); }' +
//                 '  jsw_return("finished");' +
//                 '};',
//         onProgress: function(v) { log("progress: " + v); }
//     }, function(result) { log("done: " + result); });
//
// ---------------------------------------------------------------------------
// Inside the HTA the following globals are pre-wired:
//
//   jsw_return(value)           write result JSON and close window
//   jsw_progress(value)         send a value back to CScript right now,
//                               without closing the window (requires
//                               options.onProgress on the CScript side)
//   http_request(...)           from system.js
//   write_text_to_file(...)     from system.js
//   read_text_file(...)         from system.js
//   file_exists(path)           from system.js
//   write_binary_file(...)      from system.js
//   read_binary_file(...)       from system.js
//   Array / String / Object / Number / Math polyfills (core + polyfills)
//   log(msg)                    appends a line to #_jsw_log (if present)
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Progress polling (private; bare assignments so they survive load()'s eval)
// ---------------------------------------------------------------------------

// Reads the complete lines written so far. Returns null when the file cannot be
// read this instant - the HTA may be mid-write - so the caller can simply try
// again on the next tick rather than losing its place.
_hta_read_progress = function(path) {
    if (!file_exists(path)) { return []; }
    var text;
    try {
        text = read_text_file(path);
    } catch (e) {
        return null;
    }
    // Everything before the last newline is complete; anything after it is a
    // line still being written, so it waits for the next poll.
    var parts = text.split("\n");
    var lines = [];
    for (var i = 0; i < parts.length - 1; i++) {
        var line = parts[i];
        if (line.charAt(line.length - 1) === "\r") { line = line.substring(0, line.length - 1); }
        if (line.length > 0) { lines.push(line); }
    }
    return lines;
};

// Dispatches every line past `dispatched` and returns the new count.
_hta_dispatch_progress = function(path, dispatched, onProgress) {
    var lines = _hta_read_progress(path);
    if (lines === null) { return dispatched; }
    for (var i = dispatched; i < lines.length; i++) {
        var value;
        try {
            value = JSON.parse(lines[i]);
        } catch (e) {
            value = lines[i];   // not JSON: hand over the raw line
        }
        onProgress(value);
    }
    return lines.length;
};

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

    var onProgress = options.onProgress || null;
    var pollMs     = options.poll_ms  || 200;
    var maxWaitMs  = options.max_wait || 0;    // 0 = wait for as long as it takes

    // ---- temp file paths ----
    var fso        = new ActiveXObject("Scripting.FileSystemObject");
    var tmpDir     = fso.GetSpecialFolder(2).Path;
    var ts         = (new Date()).getTime();
    var htaPath      = tmpDir + "\\jsw_hta_"      + ts + ".hta";
    var resultPath   = tmpDir + "\\jsw_result_"   + ts + ".json";
    var progressPath = tmpDir + "\\jsw_progress_" + ts + ".txt";
    var donePath     = tmpDir + "\\jsw_done_"     + ts + ".txt";

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
              '<script type="text/javascript" src="' + libBase + 'console.js"></script>\n' +
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
        // jsw_progress: append one JSON line to the progress file. Opened and
        // closed per call so the CScript side can read it between writes.
        'var _jsw_progress_path = "' + jsStr(progressPath) + '";',
        'var _jsw_done_path = "' + jsStr(donePath) + '";',
        'function jsw_progress(value) {',
        '  try {',
        '    var fso = new ActiveXObject("Scripting.FileSystemObject");',
        '    var f = fso.OpenTextFile(_jsw_progress_path, 8, true);',
        '    f.WriteLine(JSON.stringify(value !== undefined ? value : null));',
        '    f.Close();',
        '  } catch(e) {}',
        '}',
        // Written however the window goes away - jsw_return, the close box, or
        // alt+F4 - so the poller always learns the window is gone. onunload,
        // not onload: a user script setting window.onload would clobber it.
        'window.onunload = function() {',
        '  try {',
        '    var fso = new ActiveXObject("Scripting.FileSystemObject");',
        '    var f = fso.CreateTextFile(_jsw_done_path, true);',
        '    f.Write("1");',
        '    f.Close();',
        '  } catch(e) {}',
        '};',
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

    // Reads the jsw_return value (null when the window was closed without one)
    // and removes every temp file this call created.
    var collect = function() {
        var result = null;
        if (file_exists(resultPath)) {
            try {
                result = JSON.parse(read_text_file(resultPath));
                delete_file(resultPath);
            } catch(e) {}
        }
        try { if (file_exists(htaPath))      delete_file(htaPath);      } catch(e) {}
        try { if (file_exists(progressPath)) delete_file(progressPath); } catch(e) {}
        try { if (file_exists(donePath))     delete_file(donePath);     } catch(e) {}
        return result;
    };

    // ---- streaming mode ----
    // shell.Run(wait=true) would park CScript inside mshta until the window
    // closed, which is exactly what progress reporting cannot afford. Launch
    // detached instead and poll the progress file until the HTA says it is
    // gone. This blocks until the window closes either way - the difference is
    // that here CScript gets to run code while it waits.
    if (onProgress) {
        shell.Run('mshta.exe "' + htaPath + '"', 1 /*SW_SHOWNORMAL*/, false);

        var dispatched = 0;
        var waited     = 0;
        while (true) {
            dispatched = _hta_dispatch_progress(progressPath, dispatched, onProgress);

            if (file_exists(donePath)) {
                // One last sweep: a value could have landed between the read
                // above and the window going away.
                _hta_dispatch_progress(progressPath, dispatched, onProgress);
                break;
            }
            if (maxWaitMs > 0 && waited >= maxWaitMs) { break; }

            sleep(pollMs);
            waited += pollMs;
        }

        var streamed = collect();
        if (onClose) { onClose(streamed); }
        return;
    }

    // ---- blocking / fire-and-forget mode (unchanged) ----
    shell.Run('mshta.exe "' + htaPath + '"', 1 /*SW_SHOWNORMAL*/, !!onClose);

    if (onClose) {
        // mshta has exited - read result then clean up
        onClose(collect());
    }
    // non-blocking: temp files cleaned up by OS eventually
};
