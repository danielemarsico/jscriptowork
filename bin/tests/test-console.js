// test-console.js - Tests for libs/console.js
//
// The shim writes through log() when one exists, so these tests swap in a
// capturing log() to assert on what actually gets written, then put the real
// one back. log() is a global assigned without `var`, per the load()/eval
// scoping rule in CLAUDE.md.
//
// Run via:  cscript.exe launcher.js test-console.js

load("core");
load("polyfills");
load("console");
load("minitest");

// ---------------------------------------------------------------------------
// Capturing log()
// ---------------------------------------------------------------------------

var _real_log = log;
var _lines    = [];

function capture() {
    _lines = [];
    log = function(msg) { _lines.push(String(msg)); };
}

function release() {
    log = _real_log;
}

// Runs fn with log() captured and returns everything it wrote.
function written(fn) {
    capture();
    try {
        fn();
    } finally {
        release();
    }
    return _lines;
}

// ---------------------------------------------------------------------------
// Surface
// ---------------------------------------------------------------------------

describe("console - surface", function() {

    it("console exists", function() {
        assert.notEqual(typeof console, "undefined");
    });

    it("exposes log, info, warn, error and debug", function() {
        assert.equal(typeof console.log,   "function");
        assert.equal(typeof console.info,  "function");
        assert.equal(typeof console.warn,  "function");
        assert.equal(typeof console.error, "function");
        assert.equal(typeof console.debug, "function");
    });

    it("no method throws", function() {
        assert.doesNotThrow(function() {
            written(function() {
                console.log("test", 123, { a: 1 });
                console.info("i");
                console.warn("w");
                console.error("e");
                console.debug("d");
            });
        });
    });

});

// ---------------------------------------------------------------------------
// Output routing and formatting
// ---------------------------------------------------------------------------

describe("console - routing", function() {

    it("writes through log()", function() {
        assert.deepEqual(written(function() { console.log("hello"); }), ["hello"]);
    });

    it("writes one line per call", function() {
        var out = written(function() {
            console.log("one");
            console.log("two");
        });
        assert.deepEqual(out, ["one", "two"]);
    });

    it("resolves log() per call, so a later redefinition is honoured", function() {
        // console.js must not capture log() once at load time; libs/log.js
        // replaces the global log() and console must follow it there.
        var out = written(function() { console.log("late-bound"); });
        assert.deepEqual(out, ["late-bound"]);
    });

});

describe("console - formatting", function() {

    it("console.log has no prefix", function() {
        assert.deepEqual(written(function() { console.log("plain"); }), ["plain"]);
    });

    it("console.info prefixes [INFO]", function() {
        assert.deepEqual(written(function() { console.info("x"); }), ["[INFO] x"]);
    });

    it("console.warn prefixes [WARN]", function() {
        assert.deepEqual(written(function() { console.warn("x"); }), ["[WARN] x"]);
    });

    it("console.error prefixes [ERROR]", function() {
        assert.deepEqual(written(function() { console.error("x"); }), ["[ERROR] x"]);
    });

    it("console.debug prefixes [DEBUG]", function() {
        assert.deepEqual(written(function() { console.debug("x"); }), ["[DEBUG] x"]);
    });

    it("joins multiple arguments with a space", function() {
        assert.deepEqual(written(function() { console.log("a", "b", "c"); }), ["a b c"]);
    });

    it("stringifies numbers and booleans", function() {
        assert.deepEqual(written(function() { console.log(42, true, false); }), ["42 true false"]);
    });

    it("JSON-encodes objects and arrays", function() {
        assert.deepEqual(written(function() { console.log({ a: 1 }); }),      ['{"a":1}']);
        assert.deepEqual(written(function() { console.log([1, "two"]); }),    ['[1,"two"]']);
    });

    it("renders null and undefined without throwing", function() {
        assert.deepEqual(written(function() { console.log(null, undefined); }), ["null undefined"]);
    });

    it("writes an empty line for no arguments", function() {
        assert.deepEqual(written(function() { console.log(); }), [""]);
    });

    it("does not JSON-encode a string that looks like an object", function() {
        assert.deepEqual(written(function() { console.log('{"a":1}'); }), ['{"a":1}']);
    });

    it("falls back to String() for a value JSON cannot encode", function() {
        var circular = {};
        circular.self = circular;
        assert.doesNotThrow(function() {
            written(function() { console.log(circular); });
        });
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
