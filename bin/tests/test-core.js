// test-core.js - Tests for libs/core.js
//
// Covers the ES5 baseline that core.js guarantees on JScript:
//   Array.prototype  indexOf / filter / map / forEach / find / reduce
//   String.prototype trim / startsWith
//   JSON.stringify / JSON.parse (Crockford json2) and Date.prototype.toJSON
//   Ext.encode / Ext.decode
//
// These tests do not care whether the engine supplied the method natively or
// core.js polyfilled it — they assert the behaviour a caller can rely on.
//
// Run via:  cscript.exe launcher.js test-core.js

load("core");
load("minitest");

// ============================================================
// Array.prototype.indexOf
// ============================================================

describe("Array.prototype.indexOf", function() {

    it("finds an element and returns its index", function() {
        assert.equal([10, 20, 30].indexOf(20), 1);
    });

    it("returns -1 when the element is absent", function() {
        assert.equal([10, 20, 30].indexOf(99), -1);
    });

    it("returns -1 for an empty array", function() {
        assert.equal([].indexOf(1), -1);
    });

    it("returns the first index when the element repeats", function() {
        assert.equal(["a", "b", "a"].indexOf("a"), 0);
    });

    it("uses strict equality (no type coercion)", function() {
        assert.equal([1, 2, 3].indexOf("2"), -1);
    });

    it("honours a positive start index", function() {
        assert.equal(["a", "b", "a"].indexOf("a", 1), 2);
    });

    it("returns -1 when start is past the last match", function() {
        assert.equal(["a", "b"].indexOf("a", 1), -1);
    });

    it("never matches NaN", function() {
        assert.equal([NaN].indexOf(NaN), -1);
    });

});

// ============================================================
// Array.prototype.filter
// ============================================================

describe("Array.prototype.filter", function() {

    it("keeps only matching elements", function() {
        assert.deepEqual([1, 2, 3, 4].filter(function(x) { return x % 2 === 0; }), [2, 4]);
    });

    it("returns an empty array when nothing matches", function() {
        assert.deepEqual([1, 3].filter(function(x) { return x % 2 === 0; }), []);
    });

    it("returns all elements when everything matches", function() {
        assert.deepEqual([1, 2].filter(function() { return true; }), [1, 2]);
    });

    it("does not mutate the source array", function() {
        var src = [1, 2, 3];
        src.filter(function(x) { return x > 1; });
        assert.deepEqual(src, [1, 2, 3]);
    });

    it("passes value, index and array to the callback", function() {
        var seen = [];
        var src  = [7, 8];
        src.filter(function(v, i, a) { seen.push([v, i, a.length]); return true; });
        assert.deepEqual(seen, [[7, 0, 2], [8, 1, 2]]);
    });

    it("honours thisArg", function() {
        var ctx = { min: 2 };
        var out = [1, 2, 3].filter(function(x) { return x >= this.min; }, ctx);
        assert.deepEqual(out, [2, 3]);
    });

    it("throws when the callback is not a function", function() {
        assert.throws(function() { [1, 2].filter("nope"); });
    });

});

// ============================================================
// Array.prototype.map
// ============================================================

describe("Array.prototype.map", function() {

    it("transforms every element", function() {
        assert.deepEqual([1, 2, 3].map(function(x) { return x * 2; }), [2, 4, 6]);
    });

    it("returns an empty array for an empty source", function() {
        assert.deepEqual([].map(function(x) { return x; }), []);
    });

    it("preserves length", function() {
        assert.equal([1, 2, 3].map(function() { return 0; }).length, 3);
    });

    it("passes the index to the callback", function() {
        assert.deepEqual([9, 9].map(function(v, i) { return i; }), [0, 1]);
    });

    it("honours thisArg", function() {
        var ctx = { factor: 3 };
        assert.deepEqual([1, 2].map(function(x) { return x * this.factor; }, ctx), [3, 6]);
    });

    it("does not mutate the source array", function() {
        var src = [1, 2];
        src.map(function(x) { return x * 10; });
        assert.deepEqual(src, [1, 2]);
    });

    it("throws when the callback is not a function", function() {
        assert.throws(function() { [1].map(null); });
    });

});

// ============================================================
// Array.prototype.forEach
// ============================================================

describe("Array.prototype.forEach", function() {

    it("visits every element in order", function() {
        var out = [];
        [1, 2, 3].forEach(function(x) { out.push(x); });
        assert.deepEqual(out, [1, 2, 3]);
    });

    it("returns undefined", function() {
        assert.equal([1].forEach(function() {}), undefined);
    });

    it("does nothing for an empty array", function() {
        var calls = 0;
        [].forEach(function() { calls++; });
        assert.equal(calls, 0);
    });

    it("passes value, index and array", function() {
        var seen = [];
        ["a", "b"].forEach(function(v, i, a) { seen.push(v + i + a.length); });
        assert.deepEqual(seen, ["a02", "b12"]);
    });

    it("honours thisArg", function() {
        var ctx = { prefix: ">" };
        var out = [];
        ["x"].forEach(function(v) { out.push(this.prefix + v); }, ctx);
        assert.deepEqual(out, [">x"]);
    });

    it("throws when the callback is not a function", function() {
        assert.throws(function() { [1].forEach(undefined); });
    });

});

// ============================================================
// Array.prototype.find
// ============================================================

describe("Array.prototype.find", function() {

    it("returns the first matching element", function() {
        assert.equal([1, 5, 9].find(function(x) { return x > 3; }), 5);
    });

    it("returns undefined when nothing matches", function() {
        assert.equal([1, 2].find(function(x) { return x > 100; }), undefined);
    });

    it("returns undefined for an empty array", function() {
        assert.equal([].find(function() { return true; }), undefined);
    });

    it("finds objects by property", function() {
        var rows = [{ id: 1 }, { id: 2 }];
        assert.equal(rows.find(function(r) { return r.id === 2; }), rows[1]);
    });

    it("passes value, index and array to the predicate", function() {
        var seen = [];
        [4, 5].find(function(v, i, a) { seen.push(v + ":" + i + ":" + a.length); return false; });
        assert.deepEqual(seen, ["4:0:2", "5:1:2"]);
    });

    it("stops at the first match", function() {
        var calls = 0;
        [1, 2, 3].find(function(x) { calls++; return x === 2; });
        assert.equal(calls, 2);
    });

    it("throws when the predicate is not a function", function() {
        assert.throws(function() { [1].find("nope"); });
    });

});

// ============================================================
// Array.prototype.reduce
// ============================================================

describe("Array.prototype.reduce", function() {

    it("sums with an initial value", function() {
        assert.equal([1, 2, 3].reduce(function(a, b) { return a + b; }, 0), 6);
    });

    it("sums without an initial value", function() {
        assert.equal([1, 2, 3].reduce(function(a, b) { return a + b; }), 6);
    });

    it("returns the initial value for an empty array", function() {
        assert.equal([].reduce(function(a, b) { return a + b; }, 42), 42);
    });

    it("returns the only element when there is one and no initial value", function() {
        assert.equal([7].reduce(function(a, b) { return a + b; }), 7);
    });

    it("throws on an empty array with no initial value", function() {
        assert.throws(function() { [].reduce(function(a, b) { return a + b; }); });
    });

    it("can build an object accumulator", function() {
        var out = ["a", "b"].reduce(function(acc, k) { acc[k] = true; return acc; }, {});
        assert.deepEqual(out, { a: true, b: true });
    });

    it("can flatten arrays", function() {
        var out = [[1, 2], [3]].reduce(function(a, b) { return a.concat(b); }, []);
        assert.deepEqual(out, [1, 2, 3]);
    });

    it("passes index and source array", function() {
        var seen = [];
        [10, 20].reduce(function(acc, v, i, a) { seen.push(i + "/" + a.length); return acc; }, null);
        assert.deepEqual(seen, ["0/2", "1/2"]);
    });

    it("throws when the callback is not a function", function() {
        assert.throws(function() { [1].reduce("nope"); });
    });

});

// ============================================================
// String.prototype.trim
// ============================================================

describe("String.prototype.trim", function() {

    it("removes leading and trailing spaces", function() {
        assert.equal("  hello  ".trim(), "hello");
    });

    it("leaves inner whitespace alone", function() {
        assert.equal("  a b  ".trim(), "a b");
    });

    it("removes tabs and newlines", function() {
        assert.equal("\t\r\nvalue\n\t".trim(), "value");
    });

    it("removes a non-breaking space", function() {
        assert.equal("\xA0value\xA0".trim(), "value");
    });

    it("removes a BOM", function() {
        assert.equal("\uFEFFvalue\uFEFF".trim(), "value");
    });

    it("returns an empty string for whitespace only", function() {
        assert.equal("   ".trim(), "");
    });

    it("is a no-op on a clean string", function() {
        assert.equal("clean".trim(), "clean");
    });

});

// ============================================================
// String.prototype.startsWith
// ============================================================

describe("String.prototype.startsWith", function() {

    it("returns true for a matching prefix", function() {
        assert.equal("hello world".startsWith("hello"), true);
    });

    it("returns false for a non-matching prefix", function() {
        assert.equal("hello world".startsWith("world"), false);
    });

    it("is case sensitive", function() {
        assert.equal("Hello".startsWith("hello"), false);
    });

    it("returns true for an empty search string", function() {
        assert.equal("hello".startsWith(""), true);
    });

    it("returns true when the search string is the whole string", function() {
        assert.equal("hello".startsWith("hello"), true);
    });

    it("returns false when the search string is longer", function() {
        assert.equal("hi".startsWith("hiya"), false);
    });

    it("honours the position argument", function() {
        assert.equal("hello world".startsWith("world", 6), true);
        assert.equal("hello world".startsWith("hello", 6), false);
    });

});

// ============================================================
// JSON.stringify
// ============================================================

describe("JSON.stringify", function() {

    it("is available as a function", function() {
        assert.equal(typeof JSON.stringify, "function");
    });

    it("encodes primitives", function() {
        assert.equal(JSON.stringify(1),      "1");
        assert.equal(JSON.stringify("a"),    '"a"');
        assert.equal(JSON.stringify(true),   "true");
        assert.equal(JSON.stringify(null),   "null");
    });

    it("encodes an empty object and an empty array", function() {
        assert.equal(JSON.stringify({}), "{}");
        assert.equal(JSON.stringify([]), "[]");
    });

    it("encodes a flat object", function() {
        assert.equal(JSON.stringify({ a: 1, b: "x" }), '{"a":1,"b":"x"}');
    });

    it("encodes a nested structure", function() {
        assert.equal(JSON.stringify({ a: [1, { b: 2 }] }), '{"a":[1,{"b":2}]}');
    });

    it("escapes quotes and backslashes", function() {
        assert.equal(JSON.stringify('he said "hi"'), '"he said \\"hi\\""');
        assert.equal(JSON.stringify("a\\b"),         '"a\\\\b"');
    });

    it("escapes control characters", function() {
        assert.equal(JSON.stringify("a\nb"), '"a\\nb"');
        assert.equal(JSON.stringify("a\tb"), '"a\\tb"');
    });

    it("returns undefined for undefined", function() {
        assert.equal(JSON.stringify(undefined), undefined);
    });

    it("drops undefined and function members of an object", function() {
        assert.equal(JSON.stringify({ a: 1, b: undefined, c: function() {} }), '{"a":1}');
    });

    it("writes null for undefined array elements", function() {
        assert.equal(JSON.stringify([1, undefined, 2]), "[1,null,2]");
    });

    it("writes null for non-finite numbers", function() {
        assert.equal(JSON.stringify(1 / 0),      "null");
        assert.equal(JSON.stringify(0 / 0),      "null");
    });

    it("indents with a number of spaces", function() {
        assert.equal(JSON.stringify({ a: 1 }, null, 2), '{\n  "a": 1\n}');
    });

    it("indents with a string", function() {
        assert.equal(JSON.stringify({ a: 1 }, null, "\t"), '{\n\t"a": 1\n}');
    });

    it("applies a replacer function", function() {
        var out = JSON.stringify({ a: 1, b: 2 }, function(k, v) {
            return (k === "b") ? undefined : v;
        });
        assert.equal(out, '{"a":1}');
    });

    it("applies a replacer array as a key whitelist", function() {
        assert.equal(JSON.stringify({ a: 1, b: 2 }, ["a"]), '{"a":1}');
    });

});

// ============================================================
// JSON.parse
// ============================================================

describe("JSON.parse", function() {

    it("is available as a function", function() {
        assert.equal(typeof JSON.parse, "function");
    });

    it("parses primitives", function() {
        assert.equal(JSON.parse("1"),      1);
        assert.equal(JSON.parse('"a"'),    "a");
        assert.equal(JSON.parse("true"),   true);
        assert.equal(JSON.parse("null"),   null);
    });

    it("parses an object", function() {
        var o = JSON.parse('{"a":1,"b":"x"}');
        assert.equal(o.a, 1);
        assert.equal(o.b, "x");
    });

    it("parses an array", function() {
        assert.deepEqual(JSON.parse("[1,2,3]"), [1, 2, 3]);
    });

    it("parses a nested structure", function() {
        var o = JSON.parse('{"a":{"b":[1,2]}}');
        assert.equal(o.a.b[1], 2);
    });

    it("tolerates surrounding whitespace", function() {
        assert.deepEqual(JSON.parse('  {"a":1}  '), { a: 1 });
    });

    it("round-trips through stringify", function() {
        var src = { n: 1, s: "two", b: false, a: [1, 2], o: { x: null } };
        assert.deepEqual(JSON.parse(JSON.stringify(src)), src);
    });

    it("applies a reviver", function() {
        var o = JSON.parse('{"a":1,"b":2}', function(k, v) {
            return (typeof v === "number") ? v * 10 : v;
        });
        assert.equal(o.a, 10);
        assert.equal(o.b, 20);
    });

    it("throws on malformed input", function() {
        assert.throws(function() { JSON.parse("{a:1}"); });
        assert.throws(function() { JSON.parse("[1,"); });
    });

    it("rejects input that could execute code", function() {
        assert.throws(function() { JSON.parse('{"a":(function(){return 1})()}'); });
    });

});

// ============================================================
// Date.prototype.toJSON
// ============================================================

describe("Date.prototype.toJSON", function() {

    it("is available as a function", function() {
        assert.equal(typeof Date.prototype.toJSON, "function");
    });

    it("serialises a date as a UTC string", function() {
        var d   = new Date(Date.UTC(2020, 0, 2, 3, 4, 5));
        var out = JSON.stringify(d);
        assert.ok(out.indexOf('"2020-01-02T03:04:05') === 0,
                  "expected a 2020-01-02T03:04:05 prefix but got " + out);
    });

    it("serialises a date nested in an object", function() {
        var out = JSON.stringify({ when: new Date(Date.UTC(1999, 11, 31, 23, 59, 59)) });
        assert.ok(out.indexOf('{"when":"1999-12-31T23:59:59') === 0,
                  "unexpected output: " + out);
    });

    it("serialises an invalid date as null", function() {
        assert.equal(JSON.stringify(new Date("not a date")), "null");
    });

});

// ============================================================
// Ext.encode / Ext.decode
// ============================================================

describe("Ext.encode / Ext.decode", function() {

    it("exposes the Ext namespace", function() {
        assert.equal(typeof Ext,          "object");
        assert.equal(typeof Ext.util.JSON, "object");
    });

    it("Ext.encode matches JSON.stringify", function() {
        var src = { a: 1, b: [2, 3] };
        assert.equal(Ext.encode(src), JSON.stringify(src));
    });

    it("Ext.decode matches JSON.parse", function() {
        assert.deepEqual(Ext.decode('{"a":1}'), JSON.parse('{"a":1}'));
    });

    it("round-trips", function() {
        var src = { name: "test", values: [1, 2, 3] };
        assert.deepEqual(Ext.decode(Ext.encode(src)), src);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
