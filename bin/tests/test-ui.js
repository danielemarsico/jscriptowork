// test-ui.js - Tests for open_hta() in libs/ui.js
//
// Every blocking test opens an HTA whose window.onload immediately calls
// jsw_return() or window.close(), so the suite runs unattended.
// Each HTA window will flash briefly on screen — that is expected behaviour.
//
// Run via:  cscript.exe launcher.js test-ui.js

load("core");
load("polyfills");
load("system");
load("ui");
load("minitest");

// ---------------------------------------------------------------------------
// Desktop detection
//
// Opening an HTA needs an interactive desktop, which a CI runner does not have.
// Under CI those describes report as skipped instead of hanging the job; the
// progress-parsing tests below need no desktop and always run.
// ---------------------------------------------------------------------------

HEADLESS = (function() {
    var shellEnv = new ActiveXObject("WScript.Shell").Environment("Process");
    return shellEnv("CI") === "true" || shellEnv("JSW_TEST_NO_DESKTOP") !== "";
}());

if (HEADLESS) {
    log("ui tests running HEADLESS - no HTA window is opened.");
}

// ---------------------------------------------------------------------------
// Helper: launch a blocking HTA that evaluates jsExpr with jsw_return
// ---------------------------------------------------------------------------

function hta_eval(jsExpr) {
    var result;
    open_hta(
        { script: 'window.onload = function() { jsw_return(' + jsExpr + '); };' },
        function(r) { result = r; }
    );
    return result;
}

// ---------------------------------------------------------------------------
// Non-blocking launch
// ---------------------------------------------------------------------------

describe("open_hta - non-blocking", function() {

    if (HEADLESS) {
        var NO_DESKTOP = "needs an interactive desktop to open an HTA window";
        skip("does not throw when launched without a callback", NO_DESKTOP);
        return;
    }

    it("does not throw when launched without a callback", function() {
        assert.doesNotThrow(function() {
            open_hta({ script: 'window.onload = function(){ window.close(); };' });
        });
    });

});

// ---------------------------------------------------------------------------
// Blocking launch — closed without jsw_return
// ---------------------------------------------------------------------------

describe("open_hta - blocking, closed without jsw_return", function() {

    if (HEADLESS) {
        var NO_DESKTOP = "needs an interactive desktop to open an HTA window";
        skip("onClose receives null when HTA calls window.close() directly", NO_DESKTOP);
        return;
    }

    it("onClose receives null when HTA calls window.close() directly", function() {
        var received = '__unset__';
        open_hta(
            { script: 'window.onload = function(){ window.close(); };' },
            function(r) { received = r; }
        );
        assert.equal(received, null);
    });

});

// ---------------------------------------------------------------------------
// Blocking — jsw_return value round-trips
// ---------------------------------------------------------------------------

describe("open_hta - jsw_return value types", function() {

    if (HEADLESS) {
        var NO_DESKTOP = "needs an interactive desktop to open an HTA window";
        skip("returns a positive integer", NO_DESKTOP);
        skip("returns zero", NO_DESKTOP);
        skip("returns a negative number", NO_DESKTOP);
        skip("returns a string", NO_DESKTOP);
        skip("returns an empty string", NO_DESKTOP);
        skip("returns true", NO_DESKTOP);
        skip("returns false", NO_DESKTOP);
        skip("returns null explicitly", NO_DESKTOP);
        skip("returns a plain object", NO_DESKTOP);
        skip("returns an array", NO_DESKTOP);
        skip("returns a nested object", NO_DESKTOP);
        return;
    }

    it("returns a positive integer", function() {
        assert.equal(hta_eval('42'), 42);
    });

    it("returns zero", function() {
        assert.equal(hta_eval('0'), 0);
    });

    it("returns a negative number", function() {
        assert.equal(hta_eval('-7'), -7);
    });

    it("returns a string", function() {
        assert.equal(hta_eval('"hello"'), "hello");
    });

    it("returns an empty string", function() {
        assert.equal(hta_eval('""'), "");
    });

    it("returns true", function() {
        assert.equal(hta_eval('true'), true);
    });

    it("returns false", function() {
        assert.equal(hta_eval('false'), false);
    });

    it("returns null explicitly", function() {
        assert.equal(hta_eval('null'), null);
    });

    it("returns a plain object", function() {
        var result = hta_eval('({ key: "val", n: 7 })');
        assert.equal(result.key, "val");
        assert.equal(result.n,   7);
    });

    it("returns an array", function() {
        assert.deepEqual(hta_eval('[1, 2, 3]'), [1, 2, 3]);
    });

    it("returns a nested object", function() {
        var result = hta_eval('({ a: { b: 42 } })');
        assert.equal(result.a.b, 42);
    });

});

// ---------------------------------------------------------------------------
// Blocking — options
// ---------------------------------------------------------------------------

describe("open_hta - options", function() {

    if (HEADLESS) {
        var NO_DESKTOP = "needs an interactive desktop to open an HTA window";
        skip("accepts custom width and height without throwing", NO_DESKTOP);
        skip("accepts a title option without throwing", NO_DESKTOP);
        skip("accepts a style option without throwing", NO_DESKTOP);
        skip("body HTML is injected into the document", NO_DESKTOP);
        return;
    }

    it("accepts custom width and height without throwing", function() {
        var result;
        open_hta(
            { width: 400, height: 300, script: 'window.onload=function(){jsw_return(1);}' },
            function(r) { result = r; }
        );
        assert.equal(result, 1);
    });

    it("accepts a title option without throwing", function() {
        var result;
        open_hta(
            { title: 'Test Window', script: 'window.onload=function(){jsw_return(1);}' },
            function(r) { result = r; }
        );
        assert.equal(result, 1);
    });

    it("accepts a style option without throwing", function() {
        var result;
        open_hta(
            { style: 'body{background:#fff;}', script: 'window.onload=function(){jsw_return(1);}' },
            function(r) { result = r; }
        );
        assert.equal(result, 1);
    });

    it("body HTML is injected into the document", function() {
        var result;
        open_hta(
            {
                body:   '<p id="p">hi</p>',
                script: 'window.onload = function() { jsw_return(document.getElementById("p").innerText); };'
            },
            function(r) { result = r; }
        );
        assert.equal(result, "hi");
    });

});

// ---------------------------------------------------------------------------
// Progress streaming - line parsing
//
// _hta_read_progress / _hta_dispatch_progress turn the file the HTA appends to
// into callback calls. They only touch file_exists/read_text_file, so stubbing
// those two globals exercises every branch with no window and no desktop.
// ---------------------------------------------------------------------------

describe("open_hta - progress line parsing", function() {

    var _real_file_exists   = file_exists;
    var _real_read_text_file = read_text_file;

    // Pretends PATH holds `text`. Passing null makes the file unreadable, the
    // way a read racing an HTA mid-write behaves.
    function withFile(text, fn) {
        file_exists    = function() { return text !== undefined; };
        read_text_file = function() {
            if (text === null) { throw new Error("locked"); }
            return text;
        };
        try {
            return fn();
        } finally {
            file_exists    = _real_file_exists;
            read_text_file = _real_read_text_file;
        }
    }

    function dispatched(text, from) {
        var seen = [];
        var n = withFile(text, function() {
            return _hta_dispatch_progress("ignored", from || 0, function(v) { seen.push(v); });
        });
        return { values: seen, count: n };
    }

    it("returns no lines when the file does not exist yet", function() {
        assert.deepEqual(withFile(undefined, function() {
            return _hta_read_progress("ignored");
        }), []);
    });

    it("returns no lines for an empty file", function() {
        assert.deepEqual(withFile("", function() {
            return _hta_read_progress("ignored");
        }), []);
    });

    it("reads complete CRLF lines", function() {
        assert.deepEqual(withFile('1\r\n2\r\n', function() {
            return _hta_read_progress("ignored");
        }), ["1", "2"]);
    });

    it("reads complete LF lines", function() {
        assert.deepEqual(withFile('1\n2\n', function() {
            return _hta_read_progress("ignored");
        }), ["1", "2"]);
    });

    it("ignores a partial trailing line still being written", function() {
        assert.deepEqual(withFile('1\r\n"half', function() {
            return _hta_read_progress("ignored");
        }), ["1"]);
    });

    it("returns null when the file cannot be read right now", function() {
        assert.equal(withFile(null, function() {
            return _hta_read_progress("ignored");
        }), null);
    });

    it("dispatches each JSON value in order", function() {
        var out = dispatched('1\r\n2\r\n3\r\n');
        assert.deepEqual(out.values, [1, 2, 3]);
        assert.equal(out.count, 3);
    });

    it("decodes strings, objects and null", function() {
        var out = dispatched('"hello"\r\n{"a":1}\r\nnull\r\n');
        assert.equal(out.values[0], "hello");
        assert.equal(out.values[1].a, 1);
        assert.equal(out.values[2], null);
    });

    it("dispatches only lines past the given count", function() {
        var out = dispatched('1\r\n2\r\n3\r\n', 2);
        assert.deepEqual(out.values, [3]);
        assert.equal(out.count, 3);
    });

    it("dispatches nothing when everything was already seen", function() {
        var out = dispatched('1\r\n2\r\n', 2);
        assert.deepEqual(out.values, []);
    });

    it("hands over a non-JSON line unchanged rather than throwing", function() {
        var out = dispatched('not json\r\n');
        assert.deepEqual(out.values, ["not json"]);
    });

    it("keeps its place when a read fails", function() {
        var out = dispatched(null, 2);
        assert.deepEqual(out.values, []);
        assert.equal(out.count, 2, "a failed read must not rewind the dispatched count");
    });

    it("survives a value containing an escaped newline", function() {
        var out = dispatched('"line1\\nline2"\r\n');
        assert.deepEqual(out.values, ["line1\nline2"]);
    });

});

// ---------------------------------------------------------------------------
// Progress streaming - end to end (needs a desktop)
// ---------------------------------------------------------------------------

describe("open_hta - progress streaming", function() {

    if (HEADLESS) {
        var NO_DESKTOP_P = "needs an interactive desktop to open an HTA window";
        skip("delivers every progress value in order",            NO_DESKTOP_P);
        skip("delivers progress before the window closes",        NO_DESKTOP_P);
        skip("still delivers the jsw_return value to onClose",    NO_DESKTOP_P);
        skip("delivers object progress values",                   NO_DESKTOP_P);
        skip("works with no progress values at all",              NO_DESKTOP_P);
        skip("gives up after max_wait",                           NO_DESKTOP_P);
        return;
    }

    it("delivers every progress value in order", function() {
        var seen = [];
        open_hta({
            script: 'window.onload = function() {' +
                    '  for (var i = 1; i <= 3; i++) { jsw_progress(i); }' +
                    '  jsw_return("done");' +
                    '};',
            onProgress: function(v) { seen.push(v); }
        });
        assert.deepEqual(seen, [1, 2, 3]);
    });

    it("delivers progress before the window closes", function() {
        // The HTA reports, waits, then reports again and only then returns.
        // Receiving both values proves the poll loop is running while the
        // window is still open rather than draining the file at the end.
        var seen = [];
        var t0   = (new Date()).getTime();
        var elapsed_at_first = -1;
        open_hta({
            script: 'window.onload = function() {' +
                    '  jsw_progress("first");' +
                    '  window.setTimeout(function() {' +
                    '    jsw_progress("second");' +
                    '    jsw_return("done");' +
                    '  }, 1200);' +
                    '};',
            poll_ms: 100,
            onProgress: function(v) {
                if (elapsed_at_first < 0) { elapsed_at_first = (new Date()).getTime() - t0; }
                seen.push(v);
            }
        });
        assert.deepEqual(seen, ["first", "second"]);
        assert.ok(elapsed_at_first < 1200,
                  "first value arrived after the window closed (" + elapsed_at_first + "ms)");
    });

    it("still delivers the jsw_return value to onClose", function() {
        var result = "__unset__";
        open_hta({
            script: 'window.onload = function() { jsw_progress(1); jsw_return("the-result"); };',
            onProgress: function() {}
        }, function(r) { result = r; });
        assert.equal(result, "the-result");
    });

    it("delivers object progress values", function() {
        var seen = [];
        open_hta({
            script: 'window.onload = function() {' +
                    '  jsw_progress({ step: 1, label: "half" });' +
                    '  jsw_return(null);' +
                    '};',
            onProgress: function(v) { seen.push(v); }
        });
        assert.equal(seen.length, 1);
        assert.equal(seen[0].step,  1);
        assert.equal(seen[0].label, "half");
    });

    it("works with no progress values at all", function() {
        var seen   = [];
        var result = "__unset__";
        open_hta({
            script: 'window.onload = function() { jsw_return("nothing reported"); };',
            onProgress: function(v) { seen.push(v); }
        }, function(r) { result = r; });
        assert.deepEqual(seen, []);
        assert.equal(result, "nothing reported");
    });

    it("gives up after max_wait", function() {
        // The window never closes itself; max_wait must break the poll loop.
        var t0 = (new Date()).getTime();
        open_hta({
            title:  "max_wait probe (closes itself)",
            script: 'window.onload = function() {' +
                    '  jsw_progress("still here");' +
                    '  window.setTimeout(function(){ window.close(); }, 6000);' +
                    '};',
            poll_ms:  100,
            max_wait: 1500,
            onProgress: function() {}
        });
        var elapsed = (new Date()).getTime() - t0;
        assert.ok(elapsed < 5000, "max_wait did not break the poll loop (" + elapsed + "ms)");
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
