// test-build.js - Tests for build.js (dist/ bundling)
//
// Runs `cscript.exe build.js` as a real subprocess (regenerating the
// project's dist/ folder exactly like the documented build step) and then
// inspects the resulting dist/launcher.js to confirm:
//   - every lib in build.js's libNames array got inlined - both the marker
//     comment AND the lib's actual source, not just the comment
//   - load() became a no-op, per the documented dist/ contract
//   - _jsw_hta_inline_libs round-trips (via eval) back to the exact
//     concatenated source of build.js's htaLibNames libs, proving the escaping
//     used for inline HTA injection is not corrupting anything
//
// libNames/htaLibNames are read out of build.js's own source rather than
// duplicated here, so this suite tracks build.js automatically if that list
// changes.
//
// Run via:  cscript.exe launcher.js test-build.js

load("core");
load("polyfills");
load("system");
load("minitest");

var fso   = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// Reads a file exactly the way build.js's own readFile() does (ForReading,
// default/ASCII format), so byte-for-byte comparisons against dist output
// are not thrown off by a differing read mode.
function readRaw(path) {
    var f = fso.OpenTextFile(path, 1); // 1 = ForReading
    var s = f.AtEndOfStream ? "" : f.ReadAll();
    f.Close();
    return s;
}

var buildJsSrc = readRaw(ROOT_FOLDER + "build.js");

function extractArrayLiteral(varName) {
    var re = new RegExp("var " + varName + "\\s*=\\s*(\\[[^\\]]*\\]);");
    var m  = buildJsSrc.match(re);
    if (!m) { throw new Error(varName + " not found in build.js"); }
    return eval(m[1]);
}

var libNames    = extractArrayLiteral("libNames");
var htaLibNames = extractArrayLiteral("htaLibNames");

// ---------------------------------------------------------------------------
// Run the real build
// ---------------------------------------------------------------------------

describe("build.js - running the build", function() {

    it("libNames and htaLibNames were found in build.js", function() {
        assert.ok(libNames.length > 0);
        assert.ok(htaLibNames.length > 0);
    });

    it("exits with code 0", function() {
        var cmd = 'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js"';
        var exitCode = shell.Run(cmd, 0, true);
        assert.equal(exitCode, 0);
    });

});

var distPath = ROOT_FOLDER + "dist/launcher.js";
var distSrc  = readRaw(distPath);

// ---------------------------------------------------------------------------
// libNames: every lib is actually inlined, not just marker-commented
// ---------------------------------------------------------------------------

describe("build.js - libs inlined into dist/launcher.js", function() {

    it("dist/launcher.js was regenerated and is non-empty", function() {
        assert.ok(distSrc.length > 0);
    });

    for (var i = 0; i < libNames.length; i++) {
        (function(name) {

            it("inlines " + name + ".js (marker comment present)", function() {
                assert.ok(distSrc.indexOf("// " + name + ".js") !== -1);
            });

            it("inlines " + name + ".js (actual source present verbatim)", function() {
                var raw = readRaw(ROOT_FOLDER + "libs/" + name + ".js");
                assert.ok(raw.length > 0, name + ".js is empty on disk");
                assert.ok(distSrc.indexOf(raw) !== -1,
                    "raw source of " + name + ".js not found verbatim in dist/launcher.js");
            });

        }(libNames[i]));
    }

});

// ---------------------------------------------------------------------------
// load() becomes a no-op
// ---------------------------------------------------------------------------

describe("build.js - load() is stubbed out", function() {

    it("dist/launcher.js defines load() as a no-op", function() {
        assert.ok(distSrc.indexOf("function load(modulename) {}") !== -1);
    });

});

// ---------------------------------------------------------------------------
// _jsw_hta_inline_libs: escaped source round-trips back to the original
// ---------------------------------------------------------------------------

describe("build.js - _jsw_hta_inline_libs is valid escaped source", function() {

    it("the assignment line is present", function() {
        assert.ok(distSrc.indexOf("var _jsw_hta_inline_libs = '") !== -1);
    });

    it("decodes (via eval) to the concatenated source of htaLibNames libs, modulo CR stripping", function() {
        var prefix = "var _jsw_hta_inline_libs = '";
        var start  = distSrc.indexOf(prefix) + prefix.length;
        // The escaped payload can never contain a raw CR or LF (both are
        // stripped/escaped by build.js's escapeJsStr), so the first literal
        // "';\r\n" after start is unambiguously the end of this statement.
        var end = distSrc.indexOf("';\r\n", start);
        assert.ok(end !== -1, "could not find end of _jsw_hta_inline_libs assignment");

        var escaped = distSrc.substring(start, end);

        var decoded;
        assert.doesNotThrow(function() {
            eval("decoded = '" + escaped + "';");
        }, "escaped _jsw_hta_inline_libs payload is not valid JS source");

        var expected = "";
        for (var h = 0; h < htaLibNames.length; h++) {
            expected += readRaw(ROOT_FOLDER + "libs/" + htaLibNames[h] + ".js") + "\r\n";
        }
        // escapeJsStr strips every raw \r before escaping \n, so CRLF source
        // decodes back as LF-only - normalise the same way here.
        expected = expected.replace(/\r/g, '');

        assert.equal(decoded, expected);
    });

});

// ---------------------------------------------------------------------------
// Reproducibility
//
// dist/ is committed, and CI rebuilds it and runs
// `git diff --exit-code -- dist/` to catch a dist/ committed without a rebuild.
// That check only means anything if the build is deterministic: build.js used
// to stamp `// Generated: <date time>` into the header, which made the check
// fail on every run no matter what. Guard against any such content returning.
// ---------------------------------------------------------------------------

describe("build.js - reproducible output", function() {

    it("a second build produces byte-identical output", function() {
        var cmd      = 'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js"';
        var exitCode = shell.Run(cmd, 0, true);
        assert.equal(exitCode, 0, "second build did not exit cleanly");

        var rebuilt = readRaw(distPath);
        assert.equal(rebuilt.length, distSrc.length,
            "dist/launcher.js changed length between two consecutive builds (" +
            distSrc.length + " then " + rebuilt.length + ") - the build is not " +
            "reproducible, so the CI drift check can never pass");
        assert.equal(rebuilt, distSrc,
            "dist/launcher.js differs between two consecutive builds - the build " +
            "is not reproducible, so the CI drift check can never pass");
    });

    it("the header carries no build timestamp", function() {
        assert.equal(distSrc.indexOf("// Generated: "), -1,
            "dist/launcher.js embeds a generation timestamp; that makes every " +
            "rebuild differ and breaks the CI dist/ drift check");
    });

});

// ---------------------------------------------------------------------------
// --compile: one standalone .js carrying its own libs
//
// Compiles a fixture script written to the temp folder and inspects the result:
// the libs it load()s are inlined (in libNames order), the ones it does not are
// not, load() is a no-op, and the script's own body is at the end. Then it runs
// the compiled file as a real subprocess and checks it produced what the
// fixture was written to produce.
// ---------------------------------------------------------------------------

var tempFolder  = fso.GetSpecialFolder(2).Path;
var fixturePath = tempFolder + "\\jsw_compile_fixture.js";
var bundlePath  = tempFolder + "\\jsw_compile_fixture.bundled.js";
var markerPath  = tempFolder + "\\jsw_compile_fixture_ran.txt";

// Deliberately loads polyfills before core: the compiler must emit them in
// libNames order (core first), not in the order the script happens to ask.
var fixtureSrc = [
    'load("polyfills");',
    'load("core");',
    'load("base64");',
    'var out = base64_encode("compiled") + "|" + [3,1,2].sort().join("");',
    'write_all_text_file(out, "' + markerPath.replace(/\\/g, "\\\\") + '");',
    'log("fixture ran: " + out);'
].join("\r\n");

// write_all_text_file lives in system.js, which the fixture does not load - so
// define it inside the fixture instead of pulling a whole lib in just for this.
fixtureSrc = 'write_all_text_file = function(text, path) {\r\n' +
             '    var f = new ActiveXObject("Scripting.FileSystemObject").CreateTextFile(path, true);\r\n' +
             '    f.Write(text); f.Close();\r\n' +
             '};\r\n' + fixtureSrc;

function writeRaw(path, text) {
    var f = fso.CreateTextFile(path, true);
    f.Write(text);
    f.Close();
}

function deleteIfPresent(path) {
    if (fso.FileExists(path)) { fso.DeleteFile(path); }
}

deleteIfPresent(fixturePath);
deleteIfPresent(bundlePath);
deleteIfPresent(markerPath);
writeRaw(fixturePath, fixtureSrc);

var compileExit = shell.Run(
    'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --compile "' + fixturePath + '"',
    0, true);

var bundleSrc = fso.FileExists(bundlePath) ? readRaw(bundlePath) : "";

describe("build.js --compile", function() {

    it("exits with code 0", function() {
        assert.equal(compileExit, 0);
    });

    it("writes <script>.bundled.js next to the script by default", function() {
        assert.ok(fso.FileExists(bundlePath), "expected " + bundlePath);
        assert.ok(bundleSrc.length > 0, "compiled bundle is empty");
    });

    it("inlines the libs the script load()s, source and all", function() {
        var wanted = ["core", "polyfills", "base64"];
        for (var i = 0; i < wanted.length; i++) {
            var raw = readRaw(ROOT_FOLDER + "libs/" + wanted[i] + ".js");
            assert.ok(bundleSrc.indexOf(raw) !== -1,
                wanted[i] + ".js was load()ed by the script but is not inlined verbatim");
        }
    });

    it("leaves out libs the script never load()s", function() {
        // helpers.js drags Excel/Access/Word COM in; nothing should pull it into
        // a bundle for a script that never asked for it.
        var helpers = readRaw(ROOT_FOLDER + "libs/helpers.js");
        assert.equal(bundleSrc.indexOf(helpers), -1,
            "helpers.js was inlined into a bundle whose script never load()s it");
    });

    it("emits libs in libNames order, not the order the script asks for them", function() {
        // The fixture loads polyfills before core; core must still come first,
        // because polyfills builds on what core installs.
        var corePos       = bundleSrc.indexOf("// core.js");
        var polyfillsPos  = bundleSrc.indexOf("// polyfills.js");
        assert.ok(corePos !== -1 && polyfillsPos !== -1, "marker comments missing");
        assert.ok(corePos < polyfillsPos,
            "core.js must be inlined before polyfills.js regardless of load() order");
    });

    it("stubs load() out, exactly as dist/ does", function() {
        assert.ok(bundleSrc.indexOf("function load(modulename) {}") !== -1);
    });

    it("carries the bootstrap the launcher would otherwise provide", function() {
        assert.ok(bundleSrc.indexOf("var _script = WScript;") !== -1);
        assert.ok(bundleSrc.indexOf("function read_all_text_file(path) {") !== -1);
        assert.ok(bundleSrc.indexOf("var ROOT_FOLDER    = CURRENT_FOLDER;") !== -1);
    });

    it("appends the script body itself, at the end", function() {
        var bodyPos = bundleSrc.indexOf('var out = base64_encode("compiled")');
        assert.ok(bodyPos !== -1, "the script's own body is missing from the bundle");
        assert.ok(bodyPos > bundleSrc.indexOf("// base64.js"),
            "the script body must come after the libs it depends on");
    });

    it("does not embed the HTA lib payload when ui is not loaded", function() {
        assert.equal(bundleSrc.indexOf("var _jsw_hta_inline_libs = '"), -1,
            "_jsw_hta_inline_libs is only needed when ui.js is inlined");
    });

    it("runs standalone under cscript.exe and does what the script says", function() {
        var exitCode = shell.Run('cscript.exe //nologo //B "' + bundlePath + '"', 0, true);
        assert.equal(exitCode, 0, "the compiled bundle did not exit cleanly");
        assert.ok(fso.FileExists(markerPath), "the compiled bundle did not run the script body");
        // base64_encode("compiled") from base64.js, .sort() from the engine:
        // proves an inlined lib and the script body are both live.
        assert.equal(readRaw(markerPath), "Y29tcGlsZWQ=|123");
    });

});

describe("build.js --compile - argument handling", function() {

    it("--all-libs inlines every lib", function() {
        var allPath = tempFolder + "\\jsw_compile_fixture.all.js";
        deleteIfPresent(allPath);
        var exitCode = shell.Run(
            'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --compile "' +
            fixturePath + '" --out "' + allPath + '" --all-libs', 0, true);
        assert.equal(exitCode, 0);

        var src = readRaw(allPath);
        for (var i = 0; i < libNames.length; i++) {
            assert.ok(src.indexOf("// " + libNames[i] + ".js") !== -1,
                "--all-libs left out " + libNames[i] + ".js");
        }
        // ui is in libNames, so the HTA payload has to come with it.
        assert.ok(src.indexOf("var _jsw_hta_inline_libs = '") !== -1,
            "--all-libs inlines ui.js, so the HTA lib payload must be present too");
        deleteIfPresent(allPath);
    });

    it("fails when the script load()s a lib that does not exist", function() {
        var badPath = tempFolder + "\\jsw_compile_bad.js";
        writeRaw(badPath, 'load("no_such_lib_at_all");\r\n');
        var exitCode = shell.Run(
            'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --compile "' + badPath + '"',
            0, true);
        assert.notEqual(exitCode, 0, "compiling a script with an unknown load() should fail");
        deleteIfPresent(badPath);
        deleteIfPresent(tempFolder + "\\jsw_compile_bad.bundled.js");
    });

    it("fails on a script that does not exist", function() {
        var exitCode = shell.Run(
            'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --compile "' +
            tempFolder + '\\jsw_definitely_not_here.js"', 0, true);
        assert.notEqual(exitCode, 0);
    });

    it("refuses to overwrite the source script", function() {
        var exitCode = shell.Run(
            'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --compile "' +
            fixturePath + '" --out "' + fixturePath + '"', 0, true);
        assert.notEqual(exitCode, 0);
        // and the fixture is still the fixture
        assert.equal(readRaw(fixturePath), fixtureSrc);
    });

    it("rejects an unrecognised argument instead of silently building dist/", function() {
        var exitCode = shell.Run(
            'cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --frobnicate', 0, true);
        assert.notEqual(exitCode, 0);
    });

    it("--help exits 0 without touching dist/", function() {
        var before   = readRaw(distPath);
        var exitCode = shell.Run('cscript.exe //nologo //B "' + ROOT_FOLDER + 'build.js" --help', 0, true);
        assert.equal(exitCode, 0);
        assert.equal(readRaw(distPath), before);
    });

});

deleteIfPresent(fixturePath);
deleteIfPresent(bundlePath);
deleteIfPresent(markerPath);

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
