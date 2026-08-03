// test-minitest.js - Tests for libs/minitest.js (the test framework itself)
//
// Only the assertion helpers and the counters are exercised directly; the
// describe/it runner is exercised implicitly by every other suite. Nothing here
// deliberately fails a test, so a red result in this suite always means a real
// regression in minitest.
//
// Run via:  cscript.exe launcher.js test-minitest.js

load("core");
load("polyfills");
load("minitest");

// ---------------------------------------------------------------------------
// Public surface
// ---------------------------------------------------------------------------

describe("minitest - public surface", function() {

    it("exposes describe, it and skip as globals", function() {
        assert.equal(typeof describe, "function");
        assert.equal(typeof it,       "function");
        assert.equal(typeof skip,     "function");
    });

    it("exposes the _test object", function() {
        assert.equal(typeof _test.describe, "function");
        assert.equal(typeof _test.it,       "function");
        assert.equal(typeof _test.skip,     "function");
        assert.equal(typeof _test.summary,  "function");
        assert.equal(typeof _test.counts,   "function");
    });

    it("exposes every assert helper", function() {
        assert.equal(typeof assert.ok,           "function");
        assert.equal(typeof assert.notOk,        "function");
        assert.equal(typeof assert.equal,        "function");
        assert.equal(typeof assert.notEqual,     "function");
        assert.equal(typeof assert.deepEqual,    "function");
        assert.equal(typeof assert.throws,       "function");
        assert.equal(typeof assert.doesNotThrow, "function");
    });

    it("assert is the same object as _test.assert", function() {
        assert.equal(assert, _test.assert);
    });

});

// ---------------------------------------------------------------------------
// assert.ok / assert.notOk
// ---------------------------------------------------------------------------

describe("assert.ok", function() {

    it("accepts truthy values", function() {
        assert.doesNotThrow(function() { assert.ok(true); });
        assert.doesNotThrow(function() { assert.ok(1); });
        assert.doesNotThrow(function() { assert.ok("x"); });
        assert.doesNotThrow(function() { assert.ok([]); });
        assert.doesNotThrow(function() { assert.ok({}); });
    });

    it("throws on falsy values", function() {
        assert.throws(function() { assert.ok(false); });
        assert.throws(function() { assert.ok(0); });
        assert.throws(function() { assert.ok(""); });
        assert.throws(function() { assert.ok(null); });
        assert.throws(function() { assert.ok(undefined); });
        assert.throws(function() { assert.ok(NaN); });
    });

    it("uses the caller's message when given one", function() {
        var msg = null;
        try { assert.ok(false, "custom message"); } catch (e) { msg = e.message; }
        assert.equal(msg, "custom message");
    });

    it("builds a default message when none is given", function() {
        var msg = null;
        try { assert.ok(false); } catch (e) { msg = e.message; }
        assert.ok(msg && msg.length > 0);
    });

});

describe("assert.notOk", function() {

    it("accepts falsy values", function() {
        assert.doesNotThrow(function() { assert.notOk(false); });
        assert.doesNotThrow(function() { assert.notOk(0); });
        assert.doesNotThrow(function() { assert.notOk(""); });
        assert.doesNotThrow(function() { assert.notOk(null); });
        assert.doesNotThrow(function() { assert.notOk(undefined); });
    });

    it("throws on truthy values", function() {
        assert.throws(function() { assert.notOk(true); });
        assert.throws(function() { assert.notOk("x"); });
        assert.throws(function() { assert.notOk(1); });
    });

});

// ---------------------------------------------------------------------------
// assert.equal / assert.notEqual
// ---------------------------------------------------------------------------

describe("assert.equal", function() {

    it("accepts identical primitives", function() {
        assert.doesNotThrow(function() { assert.equal(1, 1); });
        assert.doesNotThrow(function() { assert.equal("a", "a"); });
        assert.doesNotThrow(function() { assert.equal(true, true); });
        assert.doesNotThrow(function() { assert.equal(null, null); });
        assert.doesNotThrow(function() { assert.equal(undefined, undefined); });
    });

    it("is strict - does not coerce types", function() {
        assert.throws(function() { assert.equal(1, "1"); });
        assert.throws(function() { assert.equal(0, false); });
        assert.throws(function() { assert.equal(null, undefined); });
        assert.throws(function() { assert.equal("", 0); });
    });

    it("compares objects by reference, not by value", function() {
        var o = { a: 1 };
        assert.doesNotThrow(function() { assert.equal(o, o); });
        assert.throws(function()       { assert.equal({ a: 1 }, { a: 1 }); });
    });

    it("never accepts NaN, even against itself", function() {
        assert.throws(function() { assert.equal(NaN, NaN); });
    });

    it("reports both values in the default message", function() {
        var msg = null;
        try { assert.equal(1, 2); } catch (e) { msg = e.message; }
        assert.ok(msg.indexOf("2") !== -1, "expected value missing from: " + msg);
        assert.ok(msg.indexOf("1") !== -1, "actual value missing from: "   + msg);
    });

});

describe("assert.notEqual", function() {

    it("accepts different values", function() {
        assert.doesNotThrow(function() { assert.notEqual(1, 2); });
        assert.doesNotThrow(function() { assert.notEqual("a", "b"); });
        assert.doesNotThrow(function() { assert.notEqual(1, "1"); });
    });

    it("throws on identical values", function() {
        assert.throws(function() { assert.notEqual(1, 1); });
        assert.throws(function() { assert.notEqual("x", "x"); });
    });

});

// ---------------------------------------------------------------------------
// assert.deepEqual
// ---------------------------------------------------------------------------

describe("assert.deepEqual", function() {

    it("accepts structurally identical arrays", function() {
        assert.doesNotThrow(function() { assert.deepEqual([1, 2, 3], [1, 2, 3]); });
        assert.doesNotThrow(function() { assert.deepEqual([], []); });
    });

    it("accepts structurally identical objects", function() {
        assert.doesNotThrow(function() { assert.deepEqual({ a: 1 }, { a: 1 }); });
        assert.doesNotThrow(function() { assert.deepEqual({}, {}); });
    });

    it("accepts nested structures", function() {
        assert.doesNotThrow(function() {
            assert.deepEqual({ a: [1, { b: 2 }] }, { a: [1, { b: 2 }] });
        });
    });

    it("throws when values differ", function() {
        assert.throws(function() { assert.deepEqual([1, 2], [1, 3]); });
        assert.throws(function() { assert.deepEqual({ a: 1 }, { a: 2 }); });
    });

    it("throws when lengths differ", function() {
        assert.throws(function() { assert.deepEqual([1], [1, 2]); });
    });

    it("is order sensitive for arrays", function() {
        assert.throws(function() { assert.deepEqual([1, 2], [2, 1]); });
    });

    it("is key-order sensitive for objects (JSON string comparison)", function() {
        // Documented limitation: deepEqual compares JSON.stringify output, so
        // {a:1,b:2} and {b:2,a:1} are NOT considered equal.
        assert.throws(function() { assert.deepEqual({ a: 1, b: 2 }, { b: 2, a: 1 }); });
    });

    it("distinguishes 1 from \"1\" (JSON encodes them differently)", function() {
        assert.throws(function() { assert.deepEqual([1], ["1"]); });
    });

});

// ---------------------------------------------------------------------------
// assert.throws / assert.doesNotThrow
// ---------------------------------------------------------------------------

describe("assert.throws", function() {

    it("passes when the function throws an Error", function() {
        assert.doesNotThrow(function() {
            assert.throws(function() { throw new Error("boom"); });
        });
    });

    it("passes when the function throws a bare value", function() {
        assert.doesNotThrow(function() {
            assert.throws(function() { throw "boom"; });
        });
    });

    it("passes on a runtime TypeError", function() {
        assert.doesNotThrow(function() {
            assert.throws(function() { var x = null; return x.y; });
        });
    });

    it("throws when the function completes normally", function() {
        var threw = false;
        try { assert.throws(function() { return 1; }); } catch (e) { threw = true; }
        assert.equal(threw, true);
    });

});

describe("assert.doesNotThrow", function() {

    it("passes when the function completes normally", function() {
        var ran = false;
        assert.doesNotThrow(function() { ran = true; });
        assert.equal(ran, true);
    });

    it("throws when the function throws", function() {
        var threw = false;
        try {
            assert.doesNotThrow(function() { throw new Error("inner"); });
        } catch (e) {
            threw = true;
        }
        assert.equal(threw, true);
    });

});

// ---------------------------------------------------------------------------
// Runner behaviour
// ---------------------------------------------------------------------------

describe("minitest - runner", function() {

    it("describe() runs its body synchronously", function() {
        var ran = false;
        _test.describe("(nested probe)", function() { ran = true; });
        assert.equal(ran, true);
    });

    it("counts() reports passed, failed and skipped", function() {
        var c = _test.counts();
        assert.equal(typeof c.passed,  "number");
        assert.equal(typeof c.failed,  "number");
        assert.equal(typeof c.skipped, "number");
    });

    it("passed count grows as tests run", function() {
        assert.ok(_test.counts().passed > 0);
    });

    it("skip() increments the skipped counter without failing", function() {
        var before = _test.counts().skipped;
        _test.skip("(skip counter probe)", "intentional, part of test-minitest");
        assert.equal(_test.counts().skipped, before + 1);
    });

    it("no test has failed in this suite so far", function() {
        assert.equal(_test.counts().failed, 0);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
