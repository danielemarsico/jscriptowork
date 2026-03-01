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
// onClose(result) — optional callback.
//   When provided CScript BLOCKS until the HTA window closes, then calls
//   onClose with whatever value the HTA passed to jsw_return(value).
//   Returns null if the user closed the window without calling jsw_return.
//
// ---------------------------------------------------------------------------
// Inside the HTA the following globals are pre-wired:
//
//   jsw_return(value)           write result JSON and close window
//   http_request(...)           from system.js
//   write_text_to_file(...)     from system.js
//   read_text_file(...)         from system.js
//   file_exists(path)           from system.js
//   write_binary_file(...)      from system.js
//   read_binary_file(...)       from system.js
//   Array / String / Object / Number / Math polyfills (core + polyfills)
//   log(msg)                    appends a line to #_jsw_log (if present)
// ---------------------------------------------------------------------------

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

    // ---- temp file paths ----
    var fso        = new ActiveXObject("Scripting.FileSystemObject");
    var tmpDir     = fso.GetSpecialFolder(2).Path;
    var ts         = (new Date()).getTime();
    var htaPath    = tmpDir + "\\jsw_hta_"    + ts + ".hta";
    var resultPath = tmpDir + "\\jsw_result_" + ts + ".json";

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
    var wait  = !!onClose;
    shell.Run('mshta.exe "' + htaPath + '"', 1 /*SW_SHOWNORMAL*/, wait);

    if (onClose) {
        // mshta has exited — read result then clean up
        var result = null;
        if (file_exists(resultPath)) {
            try {
                result = JSON.parse(read_text_file(resultPath));
                delete_file(resultPath);
            } catch(e) {}
        }
        try { if (file_exists(htaPath)) delete_file(htaPath); } catch(e) {}
        onClose(result);
    }
    // non-blocking: temp files cleaned up by OS eventually
};
