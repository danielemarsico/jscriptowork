// build.js - Creates the jscriptowork deployment package in dist/, and
// compiles single-file standalone bundles.
//
// Run from the project root:
//   cscript.exe build.js
//     Output (two files only):
//       dist/launcher.js   -- all libs inlined; open_hta() injects libs inline into HTAs
//       dist/launcher.bat  -- cscript.exe wrapper: launcher.bat <yourscript.js>
//
//   cscript.exe build.js --compile myscript.js [--out path.js] [--all-libs]
//     Output (one file):
//       myscript.bundled.js -- bootstrap + the libs the script load()s + the
//                              script itself, runnable as
//                              `cscript.exe myscript.bundled.js` with no
//                              libs/ folder and no launcher.

(function() {

    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var root = WScript.ScriptFullName;
    root = root.slice(0, root.lastIndexOf("\\"));  // project root (no trailing backslash)

    var distDir = root + "\\dist";
    var libsSrc = root + "\\libs";

    function echo(m) { WScript.Echo(m); }

    function readFile(path) {
        try {
            var f = fso.OpenTextFile(path, 1);
            if (f.AtEndOfStream) { f.Close(); return ""; }
            var s = f.ReadAll();
            f.Close();
            return s;
        } catch(e) {
            echo("ERROR reading " + path + " : " + e.message);
            return "";
        }
    }

    function writeFile(path, text) {
        var f = fso.CreateTextFile(path, true);
        f.Write(text);
        f.Close();
    }

    // Escapes a string for safe embedding as a JS single-quoted string literal.
    // Backslashes must be escaped first to avoid double-escaping.
    function escapeJsStr(s) {
        return s
            .replace(/\\/g,  '\\\\')
            .replace(/'/g,   "\\'")
            .replace(/\r/g,  '')
            .replace(/\n/g,  '\\n');
    }

    // ---- libs to bundle into launcher.js (CScript side) ----
    // Load order matters and is this array's order: core before polyfills,
    // console before log (log.js replaces the launcher's log()), and so on.
    var libNames = ["core", "ext", "polyfills", "console", "log", "system", "csv",
                    "helpers", "minimist", "ui", "win", "base64", "crypto", "minitest"];

    // ---- libs to embed inside HTAs (what ui.js previously loaded via <script src>) ----
    // Must match the files referenced in ui.js block 2.
    var htaLibNames = ["core", "polyfills", "console", "system"];

    function libPath(name) { return libsSrc + "\\" + name + ".js"; }

    function readHtaLibs() {
        var s = "";
        for (var h = 0; h < htaLibNames.length; h++) {
            var hp = libPath(htaLibNames[h]);
            if (fso.FileExists(hp)) { s += readFile(hp) + "\r\n"; }
        }
        return s;
    }

    // ---- shared emitters -------------------------------------------------
    //
    // Both outputs (dist/launcher.js and a --compile bundle) are the same
    // sandwich: bootstrap, then inlined libs, then something that runs a user
    // script. Only the last layer differs - dist/ eval()s a script named on the
    // command line, a compiled bundle carries its script inline.

    function bootstrapLines(target) {
        var L = [];
        L.push("var _script = WScript;");
        L.push("function log(message) { _script.echo(message); }");
        L.push("");
        L.push("var CURRENT_PATH   = _script.ScriptFullName;");
        L.push("var CURRENT_FOLDER = CURRENT_PATH.slice(0, CURRENT_PATH.lastIndexOf('\\\\') + 1);");
        L.push("var ROOT_FOLDER    = CURRENT_FOLDER;");
        L.push("");
        L.push("// read_all_text_file: referenced by system.js (load_working_directory).");
        L.push("function read_all_text_file(path) {");
        L.push("    var _fso = new ActiveXObject('Scripting.FileSystemObject');");
        L.push("    try {");
        L.push("        var _f = _fso.OpenTextFile(path, 1);");
        L.push("        if (_f.AtEndOfStream) { _f.Close(); return ''; }");
        L.push("        var _s = _f.ReadAll(); _f.Close(); return _s;");
        L.push("    } catch(e) { return null; }");
        L.push("}");
        L.push("");
        L.push("// load() is a no-op: all libs are already inlined below.");
        L.push("function load(modulename) {}");
        L.push("");
        return L;
    }

    function htaInlineLibsLines() {
        var L = [];
        L.push("// Source of core + polyfills + system, pre-escaped for inline HTA injection.");
        L.push("// ui.js checks typeof _jsw_hta_inline_libs to decide whether to use this");
        L.push("// or fall back to <script src=\"file:///...\"> references.");
        L.push("var _jsw_hta_inline_libs = '" + escapeJsStr(readHtaLibs()) + "';");
        L.push("");
        return L;
    }

    function inlinedLibLines(names) {
        var L = [];
        for (var i = 0; i < names.length; i++) {
            var p = libPath(names[i]);
            if (!fso.FileExists(p)) { continue; }
            L.push("// ---------------------------------------------------------------------------");
            L.push("// " + names[i] + ".js");
            L.push("// ---------------------------------------------------------------------------");
            L.push("");
            L.push(readFile(p));
            L.push("");
        }
        return L;
    }

    function concat(target, more) {
        for (var i = 0; i < more.length; i++) { target.push(more[i]); }
        return target;
    }

    // =======================================================================
    // Default build: dist/launcher.js + dist/launcher.bat
    // =======================================================================

    function buildDist() {

        // ---- clean and create dist ----
        if (fso.FolderExists(distDir)) {
            fso.DeleteFolder(distDir, true);
            echo("Removed previous dist/");
        }
        fso.CreateFolder(distDir);
        echo("Created dist/");

        var L = [];

        L.push("// jscriptowork bundled launcher");
        // Deliberately NOT a timestamp: dist/ is committed and CI rebuilds it and
        // runs `git diff --exit-code -- dist/` to catch a dist/ that was committed
        // without rebuilding from libs/. Any per-build varying content (a clock
        // reading, a random id) makes that check fail on every run regardless of
        // whether dist/ is actually stale. Keep this file byte-for-byte
        // reproducible from the contents of libs/ alone.
        L.push("// Generated by build.js from libs/ - do not edit by hand.");
        L.push("//");
        L.push("// All libs are inlined. No separate libs/ folder is needed.");
        L.push("// open_hta() injects lib source inline into generated HTAs.");
        L.push("//");
        L.push("// Usage:  cscript launcher.js <yourscript.js>");
        L.push("//         launcher.bat <yourscript.js>");
        L.push("");

        L.push("// ---------------------------------------------------------------------------");
        L.push("// Bootstrap");
        L.push("// ---------------------------------------------------------------------------");
        L.push("");
        concat(L, bootstrapLines());
        concat(L, htaInlineLibsLines());
        concat(L, inlinedLibLines(libNames));

        // ---- script executor ----
        L.push("// ---------------------------------------------------------------------------");
        L.push("// Script executor");
        L.push("// ---------------------------------------------------------------------------");
        L.push("");
        L.push("(function() {");
        L.push("    if (_script.Arguments.Count() === 0) {");
        L.push("        log('Usage: cscript launcher.js <yourscript.js>');");
        L.push("        _script.Quit(1);");
        L.push("    }");
        L.push("    var scriptPath = _script.Arguments(0);");
        L.push("    log('executing:\\t' + scriptPath);");
        L.push("    var src = read_all_text_file(scriptPath);");
        L.push("    if (src !== null) { eval(src); }");
        L.push("}());");

        writeFile(distDir + "\\launcher.js", L.join("\r\n"));
        echo("Written dist/launcher.js");

        // ---- dist/launcher.bat ----
        var bat = [
            "@echo off",
            "SET mypath=%~dp0",
            "cscript.exe \"%mypath%launcher.js\" %*"
        ].join("\r\n");
        writeFile(distDir + "\\launcher.bat", bat);
        echo("Written dist/launcher.bat");

        echo("");
        echo("=== Build complete ===");
        echo("  dist/launcher.bat <yourscript.js>");
        echo("  (two files total: launcher.js + launcher.bat)");
    }

    // =======================================================================
    // --compile: one standalone .js carrying its own libs
    // =======================================================================

    // Which libs does this script ask for? Finds load("name") / load('name')
    // calls. Over-inclusion is harmless (a lib nobody calls just sits there);
    // under-inclusion is a broken bundle, so anything unclear falls back to
    // every lib.
    function scanLoads(src) {
        var found  = {};
        var names  = [];
        var re     = /(^|[^\w$.])load\s*\(\s*(['"])([^'"]*)\2\s*\)/g;
        var m;
        while ((m = re.exec(src)) !== null) {
            var name = m[3];
            if (!found[name]) { found[name] = true; names.push(name); }
        }
        // A load() whose argument is not a plain string literal - load(name),
        // load("a" + b) - cannot be resolved by reading the source.
        var dynamic = /(^|[^\w$.])load\s*\(\s*[^'")\s]/.test(src);
        return { names: names, dynamic: dynamic };
    }

    // Returns the scan's names ordered the way libNames orders them, because
    // load order is load-bearing: core before polyfills, console before log.
    function orderLibs(names) {
        var wanted  = {};
        var unknown = [];
        var i;
        for (i = 0; i < names.length; i++) { wanted[names[i]] = true; }
        for (i = 0; i < names.length; i++) {
            var known = false;
            for (var j = 0; j < libNames.length; j++) {
                if (libNames[j] === names[i]) { known = true; break; }
            }
            if (!known) { unknown.push(names[i]); }
        }
        var ordered = [];
        for (i = 0; i < libNames.length; i++) {
            if (wanted[libNames[i]]) { ordered.push(libNames[i]); }
        }
        return { ordered: ordered, unknown: unknown };
    }

    function contains(arr, value) {
        for (var i = 0; i < arr.length; i++) { if (arr[i] === value) { return true; } }
        return false;
    }

    function defaultOutPath(scriptPath) {
        if (/\.js$/i.test(scriptPath)) {
            return scriptPath.replace(/\.js$/i, ".bundled.js");
        }
        return scriptPath + ".bundled.js";
    }

    function compileScript(scriptPath, outPath, forceAllLibs) {

        var absIn = fso.GetAbsolutePathName(scriptPath);
        if (!fso.FileExists(absIn)) {
            echo("ERROR: no such script: " + absIn);
            return 1;
        }

        var absOut = fso.GetAbsolutePathName(outPath ? outPath : defaultOutPath(absIn));
        if (absOut.toLowerCase() === absIn.toLowerCase()) {
            echo("ERROR: --out would overwrite the source script: " + absIn);
            return 1;
        }

        var src  = readFile(absIn);
        var scan = scanLoads(src);
        var libs, why;

        if (forceAllLibs) {
            libs = libNames;
            why  = "--all-libs";
        } else if (scan.dynamic) {
            libs = libNames;
            why  = "a load() call with a non-literal argument was found - cannot tell what it needs";
        } else {
            var ordered = orderLibs(scan.names);
            if (ordered.unknown.length > 0) {
                echo("ERROR: the script load()s libs that do not exist in libs/: " +
                     ordered.unknown.join(", "));
                return 1;
            }
            libs = ordered.ordered;
            why  = "scanned from the script's load() calls";
        }

        var name = absIn.slice(absIn.lastIndexOf("\\") + 1);

        var L = [];
        L.push("// " + name.replace(/\.js$/i, "") + " - compiled by jscriptowork build.js --compile");
        L.push("// Generated - do not edit by hand. Edit " + name + " and recompile.");
        L.push("//");
        L.push("// Standalone: no libs/ folder, no launcher. Run it with");
        L.push("//   cscript.exe " + absOut.slice(absOut.lastIndexOf("\\") + 1));
        L.push("//");
        L.push("// Inlined libs: " + (libs.length ? libs.join(", ") : "(none)"));
        L.push("//");
        L.push("// One difference from running through bin/launcher.js: there the");
        L.push("// launcher owns WScript.Arguments(0) (the script path) and the script's");
        L.push("// own arguments start at 1. Here nothing is in front of them, so they");
        L.push("// start at 0.");
        L.push("");

        L.push("// ---------------------------------------------------------------------------");
        L.push("// Bootstrap");
        L.push("// ---------------------------------------------------------------------------");
        L.push("");
        concat(L, bootstrapLines());

        // ui.js reads _jsw_hta_inline_libs to inject lib source into the HTAs it
        // opens; without it the HTA would <script src=...> a libs/ folder that a
        // standalone bundle is specifically meant not to need.
        if (contains(libs, "ui")) { concat(L, htaInlineLibsLines()); }

        concat(L, inlinedLibLines(libs));

        L.push("// ---------------------------------------------------------------------------");
        L.push("// " + name);
        L.push("// ---------------------------------------------------------------------------");
        L.push("");
        L.push(src);

        writeFile(absOut, L.join("\r\n"));

        echo("Compiled " + name + " -> " + absOut);
        echo("  libs inlined (" + why + "):");
        echo("    " + (libs.length ? libs.join(", ") : "(none)"));
        echo("");
        echo("Run it with:  cscript.exe \"" + absOut + "\"");
        return 0;
    }

    // =======================================================================
    // Arguments
    // =======================================================================

    function collectArgv() {
        var argv = [];
        for (var i = 0; i < WScript.Arguments.Count(); i++) {
            argv.push(WScript.Arguments(i));
        }
        return argv;
    }

    // libs/minimist.js is the project's own argument parser, so use it - but
    // build.js runs directly under cscript.exe, with no launcher and therefore
    // no load(). new Function(), not eval(): the source gets a clean function
    // scope of its own, so minimist's top-level `function hasKey()` helpers stay
    // reachable from the closure it exports, and its `minimist = function(){}`
    // bare assignment still lands on the global object. eval() inside this IIFE
    // is exactly the shape that loses inner function declarations in JScript
    // (see the note at the top of libs/crypto.js).
    //
    // Returns null if anything about that fails; parseArgv() then falls back to
    // a hand-rolled parser, because a broken lib must never take the build with
    // it.
    function tryLoadMinimist() {
        try {
            // FileExists first: readFile() echoes a loud ERROR line on a missing
            // file, and a missing lib here is handled, not fatal.
            if (!fso.FileExists(libPath("core")) || !fso.FileExists(libPath("minimist"))) {
                return null;
            }
            var coreSrc = readFile(libPath("core"));
            var miniSrc = readFile(libPath("minimist"));
            if (!coreSrc || !miniSrc) { return null; }
            new Function(coreSrc)();   // Array.prototype.forEach/filter, which minimist uses
            new Function(miniSrc)();
            return (typeof minimist === "function") ? minimist : null;
        } catch(e) {
            return null;
        }
    }

    // Long flags only, so this stays trivially equivalent to the minimist path:
    //   --compile <file>  --out <file>  --all-libs  --help
    function parseArgvFallback(argv) {
        var out = { compile: null, out: null, allLibs: false, help: false, unknown: [] };
        for (var i = 0; i < argv.length; i++) {
            var a = argv[i];
            if (a === "--compile") { out.compile = argv[++i]; }
            else if (a === "--out") { out.out = argv[++i]; }
            else if (a === "--all-libs") { out.allLibs = true; }
            else if (a === "--help" || a === "-h" || a === "/?") { out.help = true; }
            else { out.unknown.push(a); }
        }
        return out;
    }

    function parseArgv(argv) {
        var parse = tryLoadMinimist();
        if (!parse) { return parseArgvFallback(argv); }

        // Without an unknown handler minimist would silently accept
        // `--frobnicate` as a boolean flag and the build would carry on as if
        // nothing was wrong. Flags it does not recognise are collected and
        // rejected; anything positional stays in `_` and is rejected too, since
        // this CLI takes no positional arguments.
        var unknownFlags = [];
        var a = parse(argv, {
            'string':  ["compile", "out"],
            'boolean': ["all-libs", "help"],
            'alias':   { 'h': "help" },
            'unknown': function(arg) {
                if (arg.charAt(0) === "-") { unknownFlags.push(arg); return false; }
                return true;
            }
        });

        return {
            compile: (typeof a.compile === "string" && a.compile !== "") ? a.compile : null,
            out:     (typeof a.out === "string" && a.out !== "") ? a.out : null,
            allLibs: !!a["all-libs"],
            help:    !!a.help,
            unknown: unknownFlags.concat(a._ ? a._ : [])
        };
    }

    function usage() {
        echo("jscriptowork build.js");
        echo("");
        echo("  cscript.exe build.js");
        echo("      Rebuild dist/ (launcher.js with every lib inlined + launcher.bat).");
        echo("");
        echo("  cscript.exe build.js --compile <script.js> [--out <path.js>] [--all-libs]");
        echo("      Compile one script into a single standalone .js: bootstrap +");
        echo("      the libs it load()s + the script itself. Runs as");
        echo("      `cscript.exe script.bundled.js` with no libs/ and no launcher.");
        echo("");
        echo("      --out       output path (default: <script>.bundled.js)");
        echo("      --all-libs  inline every lib instead of only the load()ed ones");
    }

    // =======================================================================
    // Entry point
    // =======================================================================

    var args = parseArgv(collectArgv());

    if (args.help) {
        usage();
        WScript.Quit(0);
    }

    if (args.compile) {
        WScript.Quit(compileScript(args.compile, args.out, args.allLibs));
    }

    if (args.unknown && args.unknown.length > 0) {
        echo("Unrecognised argument: " + args.unknown[0]);
        echo("");
        usage();
        WScript.Quit(1);
    }

    buildDist();

}());
