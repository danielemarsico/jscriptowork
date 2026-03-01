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
// Summary
// ---------------------------------------------------------------------------

_test.summary();
