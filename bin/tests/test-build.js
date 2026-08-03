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
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
