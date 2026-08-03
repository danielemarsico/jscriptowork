// test-polyfills.js - Test suite for libs/polyfills.js
// Run via:  cscript.exe launcher.js test-polyfills.js

load("core");
load("polyfills");
load("minitest");

// ============================================================
// Array - static methods
// ============================================================

describe("Array.isArray", function() {

    it("returns true for array literal", function() {
        assert.equal(Array.isArray([]), true);
    });

    it("returns true for new Array()", function() {
        assert.equal(Array.isArray(new Array(3)), true);
    });

    it("returns true for array with values", function() {
        assert.equal(Array.isArray([1, 2, 3]), true);
    });

    it("returns false for plain object", function() {
        assert.equal(Array.isArray({}), false);
    });

    it("returns false for string", function() {
        assert.equal(Array.isArray("hello"), false);
    });

    it("returns false for null", function() {
        assert.equal(Array.isArray(null), false);
    });

    it("returns false for undefined", function() {
        assert.equal(Array.isArray(undefined), false);
    });

    it("returns false for number", function() {
        assert.equal(Array.isArray(42), false);
    });

});

describe("Array.of", function() {

    it("creates array from arguments", function() {
        assert.deepEqual(Array.of(1, 2, 3), [1, 2, 3]);
    });

    it("creates single-element array (unlike new Array(n))", function() {
        assert.deepEqual(Array.of(7), [7]);
    });

    it("creates empty array with no arguments", function() {
        assert.deepEqual(Array.of(), []);
    });

    it("works with mixed types", function() {
        assert.deepEqual(Array.of(1, "two", null), [1, "two", null]);
    });

});

describe("Array.from", function() {

    it("converts string to char array", function() {
        assert.deepEqual(Array.from("hello"), ["h", "e", "l", "l", "o"]);
    });

    it("clones an array", function() {
        assert.deepEqual(Array.from([1, 2, 3]), [1, 2, 3]);
    });

    it("converts array-like object", function() {
        assert.deepEqual(Array.from({length: 3, 0: "a", 1: "b", 2: "c"}), ["a", "b", "c"]);
    });

    it("applies map function", function() {
        assert.deepEqual(Array.from([1, 2, 3], function(x) { return x * 2; }), [2, 4, 6]);
    });

    it("throws on null", function() {
        assert.throws(function() { Array.from(null); });
    });

});

// ============================================================
// Array - prototype methods
// ============================================================

describe("Array.prototype.every", function() {

    it("returns true when all elements match", function() {
        assert.equal([2, 4, 6].every(function(x) { return x % 2 === 0; }), true);
    });

    it("returns false when one element does not match", function() {
        assert.equal([2, 3, 6].every(function(x) { return x % 2 === 0; }), false);
    });

    it("returns true for empty array (vacuously)", function() {
        assert.equal([].every(function() { return false; }), true);
    });

    it("passes index and array to callback", function() {
        var indices = [];
        [10, 20, 30].every(function(v, i) { indices.push(i); return true; });
        assert.deepEqual(indices, [0, 1, 2]);
    });

});

describe("Array.prototype.some", function() {

    it("returns true when at least one element matches", function() {
        assert.equal([1, 2, 3].some(function(x) { return x > 2; }), true);
    });

    it("returns false when no element matches", function() {
        assert.equal([1, 2, 3].some(function(x) { return x > 5; }), false);
    });

    it("returns false for empty array", function() {
        assert.equal([].some(function() { return true; }), false);
    });

});

describe("Array.prototype.includes", function() {

    it("finds an existing value", function() {
        assert.equal([1, 2, 3].includes(2), true);
    });

    it("returns false for missing value", function() {
        assert.equal([1, 2, 3].includes(4), false);
    });

    it("finds NaN (SameValueZero - unlike indexOf)", function() {
        assert.equal([1, NaN, 3].includes(NaN), true);
    });

    it("respects fromIndex", function() {
        assert.equal([1, 2, 3, 2].includes(2, 2), true);
        assert.equal([1, 2, 3].includes(1, 1), false);
    });

    it("handles negative fromIndex", function() {
        assert.equal([1, 2, 3].includes(3, -1), true);
    });

});

describe("Array.prototype.findIndex", function() {

    it("returns index of first matching element", function() {
        assert.equal([1, 2, 3].findIndex(function(x) { return x > 1; }), 1);
    });

    it("returns -1 when no match", function() {
        assert.equal([1, 2, 3].findIndex(function(x) { return x > 5; }), -1);
    });

    it("returns 0 when first element matches", function() {
        assert.equal([5, 1, 2].findIndex(function(x) { return x === 5; }), 0);
    });

});

describe("Array.prototype.fill", function() {

    it("fills entire array", function() {
        assert.deepEqual([1, 2, 3, 4].fill(0), [0, 0, 0, 0]);
    });

    it("fills from start index", function() {
        assert.deepEqual([1, 2, 3, 4].fill(5, 2), [1, 2, 5, 5]);
    });

    it("fills between start and end index", function() {
        assert.deepEqual([1, 2, 3, 4].fill(5, 1, 3), [1, 5, 5, 4]);
    });

    it("returns the mutated array", function() {
        var a = [1, 2, 3];
        assert.ok(a.fill(0) === a);
    });

});

describe("Array.prototype.flat", function() {

    it("flattens one level by default", function() {
        assert.deepEqual([1, [2, 3]].flat(), [1, 2, 3]);
    });

    it("does not flatten deeper than specified depth", function() {
        assert.deepEqual([1, [2, [3]]].flat(), [1, 2, [3]]);
    });

    it("flattens to specified depth", function() {
        assert.deepEqual([1, [2, [3, [4]]]].flat(2), [1, 2, 3, [4]]);
    });

    it("flattens deeply with Infinity", function() {
        assert.deepEqual([1, [2, [3, [4]]]].flat(Infinity), [1, 2, 3, 4]);
    });

    it("handles empty arrays inside (outer [] vanishes, inner [] stays at depth 1)", function() {
        assert.deepEqual([1, [], [2, []]].flat(), [1, 2, []]);
    });

});

describe("Array.prototype.flatMap", function() {

    it("maps then flattens one level", function() {
        assert.deepEqual([1, 2, 3].flatMap(function(x) { return [x, x * 2]; }), [1, 2, 2, 4, 3, 6]);
    });

    it("works with identity (like map)", function() {
        assert.deepEqual([1, 2, 3].flatMap(function(x) { return x; }), [1, 2, 3]);
    });

    it("only flattens one level", function() {
        assert.deepEqual([[1], [2]].flatMap(function(x) { return [x]; }), [[1], [2]]);
    });

});

// ============================================================
// String - prototype methods
// ============================================================

describe("String.prototype.endsWith", function() {

    it("returns true when string ends with suffix", function() {
        assert.equal("hello world".endsWith("world"), true);
    });

    it("returns false when it does not", function() {
        assert.equal("hello world".endsWith("hello"), false);
    });

    it("works with endPosition parameter", function() {
        assert.equal("hello world".endsWith("hello", 5), true);
    });

    it("is case sensitive", function() {
        assert.equal("Hello".endsWith("hello"), false);
    });

    it("returns true for empty suffix", function() {
        assert.equal("hello".endsWith(""), true);
    });

});

describe("String.prototype.includes", function() {

    it("returns true when substring exists", function() {
        assert.equal("hello world".includes("world"), true);
    });

    it("returns false when substring is absent", function() {
        assert.equal("hello world".includes("xyz"), false);
    });

    it("is case sensitive", function() {
        assert.equal("Hello".includes("hello"), false);
    });

    it("works with position offset", function() {
        assert.equal("hello".includes("ell", 1), true);
        assert.equal("hello".includes("ell", 2), false);
    });

});

describe("String.prototype.repeat", function() {

    it("repeats string n times", function() {
        assert.equal("abc".repeat(3), "abcabcabc");
    });

    it("returns empty string for 0", function() {
        assert.equal("abc".repeat(0), "");
    });

    it("returns the string itself for 1", function() {
        assert.equal("abc".repeat(1), "abc");
    });

    it("throws RangeError for negative count", function() {
        assert.throws(function() { "x".repeat(-1); });
    });

    it("throws RangeError for Infinity", function() {
        assert.throws(function() { "x".repeat(Infinity); });
    });

});

describe("String.prototype.padStart", function() {

    it("pads to target length with default space", function() {
        assert.equal("5".padStart(3), "  5");
    });

    it("pads with custom fill string", function() {
        assert.equal("5".padStart(3, "0"), "005");
    });

    it("does not truncate when string is longer than target", function() {
        assert.equal("hello".padStart(3), "hello");
    });

    it("repeats fill string if needed", function() {
        assert.equal("1".padStart(5, "ab"), "abab1");
    });

});

describe("String.prototype.padEnd", function() {

    it("pads to target length with default space", function() {
        assert.equal("5".padEnd(3), "5  ");
    });

    it("pads with custom fill string", function() {
        assert.equal("5".padEnd(3, "0"), "500");
    });

    it("does not truncate when string is longer than target", function() {
        assert.equal("hello".padEnd(3), "hello");
    });

    it("repeats fill string if needed", function() {
        assert.equal("1".padEnd(5, "ab"), "1abab");
    });

});

describe("String.prototype.trimStart", function() {

    it("removes leading whitespace", function() {
        assert.equal("  hello  ".trimStart(), "hello  ");
    });

    it("leaves trailing whitespace intact", function() {
        assert.equal("  hi  ".trimStart(), "hi  ");
    });

    it("trimLeft is an alias", function() {
        assert.equal("  hi".trimLeft(), "hi");
    });

    it("does nothing when no leading whitespace", function() {
        assert.equal("hello".trimStart(), "hello");
    });

});

describe("String.prototype.trimEnd", function() {

    it("removes trailing whitespace", function() {
        assert.equal("  hello  ".trimEnd(), "  hello");
    });

    it("leaves leading whitespace intact", function() {
        assert.equal("  hi  ".trimEnd(), "  hi");
    });

    it("trimRight is an alias", function() {
        assert.equal("hi  ".trimRight(), "hi");
    });

});

// ============================================================
// Object - static methods
// ============================================================

describe("Object.keys", function() {

    it("returns own enumerable keys", function() {
        assert.deepEqual(Object.keys({a: 1, b: 2, c: 3}), ["a", "b", "c"]);
    });

    it("returns empty array for empty object", function() {
        assert.deepEqual(Object.keys({}), []);
    });

    it("does not include inherited keys", function() {
        function Parent() { this.x = 1; }
        Parent.prototype.y = 2;
        var child = new Parent();
        assert.deepEqual(Object.keys(child), ["x"]);
    });

    it("throws on null", function() {
        assert.throws(function() { Object.keys(null); });
    });

});

describe("Object.values", function() {

    it("returns own enumerable values", function() {
        assert.deepEqual(Object.values({a: 1, b: 2}), [1, 2]);
    });

    it("returns empty array for empty object", function() {
        assert.deepEqual(Object.values({}), []);
    });

});

describe("Object.entries", function() {

    it("returns [key, value] pairs", function() {
        assert.deepEqual(Object.entries({a: 1, b: 2}), [["a", 1], ["b", 2]]);
    });

    it("returns empty array for empty object", function() {
        assert.deepEqual(Object.entries({}), []);
    });

});

describe("Object.assign", function() {

    it("copies properties from one source to target", function() {
        var result = Object.assign({}, {a: 1, b: 2});
        assert.deepEqual(result, {a: 1, b: 2});
    });

    it("merges multiple sources", function() {
        var result = Object.assign({}, {a: 1}, {b: 2}, {c: 3});
        assert.deepEqual(result, {a: 1, b: 2, c: 3});
    });

    it("later sources overwrite earlier ones", function() {
        var result = Object.assign({}, {a: 1}, {a: 99});
        assert.equal(result.a, 99);
    });

    it("returns the target object", function() {
        var target = {};
        var result = Object.assign(target, {a: 1});
        assert.ok(result === target);
    });

    it("skips null/undefined sources", function() {
        assert.doesNotThrow(function() {
            Object.assign({}, null, undefined, {a: 1});
        });
    });

    it("throws when target is null", function() {
        assert.throws(function() { Object.assign(null, {a: 1}); });
    });

});

// ============================================================
// Number - static methods and constants
// ============================================================

describe("Number.isNaN", function() {

    it("returns true for NaN", function() {
        assert.equal(Number.isNaN(NaN), true);
    });

    it("returns false for number (unlike global isNaN)", function() {
        assert.equal(Number.isNaN(1), false);
    });

    it("returns false for string NaN (unlike global isNaN)", function() {
        assert.equal(Number.isNaN("NaN"), false);
    });

    it("returns false for undefined", function() {
        assert.equal(Number.isNaN(undefined), false);
    });

});

describe("Number.isFinite", function() {

    it("returns true for finite number", function() {
        assert.equal(Number.isFinite(42), true);
    });

    it("returns false for Infinity", function() {
        assert.equal(Number.isFinite(Infinity), false);
    });

    it("returns false for -Infinity", function() {
        assert.equal(Number.isFinite(-Infinity), false);
    });

    it("returns false for NaN", function() {
        assert.equal(Number.isFinite(NaN), false);
    });

    it("returns false for string (unlike global isFinite)", function() {
        assert.equal(Number.isFinite("1"), false);
    });

});

describe("Number.isInteger", function() {

    it("returns true for whole number", function() {
        assert.equal(Number.isInteger(4), true);
    });

    it("returns true for negative whole number", function() {
        assert.equal(Number.isInteger(-4), true);
    });

    it("returns false for float", function() {
        assert.equal(Number.isInteger(1.5), false);
    });

    it("returns false for Infinity", function() {
        assert.equal(Number.isInteger(Infinity), false);
    });

    it("returns false for string", function() {
        assert.equal(Number.isInteger("1"), false);
    });

});

describe("Number constants", function() {

    it("EPSILON is a small positive number", function() {
        assert.ok(Number.EPSILON > 0 && Number.EPSILON < 0.001);
    });

    it("MAX_SAFE_INTEGER is 2^53 - 1", function() {
        assert.equal(Number.MAX_SAFE_INTEGER, 9007199254740991);
    });

    it("MIN_SAFE_INTEGER is -(2^53 - 1)", function() {
        assert.equal(Number.MIN_SAFE_INTEGER, -9007199254740991);
    });

});

// ============================================================
// Math - static methods
// ============================================================

describe("Math.sign", function() {

    it("returns 1 for positive number", function() {
        assert.equal(Math.sign(5), 1);
    });

    it("returns -1 for negative number", function() {
        assert.equal(Math.sign(-5), -1);
    });

    it("returns 0 for zero", function() {
        assert.equal(Math.sign(0), 0);
    });

    it("returns NaN for NaN", function() {
        assert.ok(Number.isNaN(Math.sign(NaN)));
    });

});

describe("Math.trunc", function() {

    it("truncates positive float", function() {
        assert.equal(Math.trunc(4.9), 4);
    });

    it("truncates negative float toward zero", function() {
        assert.equal(Math.trunc(-4.9), -4);
    });

    it("leaves integer unchanged", function() {
        assert.equal(Math.trunc(5), 5);
    });

});

describe("Math.log2", function() {

    it("computes log base 2 of 8", function() {
        assert.equal(Math.round(Math.log2(8)), 3);
    });

    it("returns 0 for log2(1)", function() {
        assert.equal(Math.log2(1), 0);
    });

});

describe("Math.log10", function() {

    it("computes log base 10 of 1000", function() {
        assert.equal(Math.round(Math.log10(1000)), 3);
    });

    it("returns 0 for log10(1)", function() {
        assert.equal(Math.log10(1), 0);
    });

});

describe("Math.cbrt", function() {

    it("returns cube root of 27", function() {
        assert.equal(Math.round(Math.cbrt(27)), 3);
    });

    it("handles negative values", function() {
        assert.equal(Math.round(Math.cbrt(-27)), -3);
    });

});

describe("Math.hypot", function() {

    it("computes hypotenuse of 3-4-5 triangle", function() {
        assert.equal(Math.hypot(3, 4), 5);
    });

    it("works with single argument", function() {
        assert.equal(Math.hypot(5), 5);
    });

});

describe("Math.clz32", function() {

    it("returns 32 for 0", function() {
        assert.equal(Math.clz32(0), 32);
    });

    it("returns 31 for 1", function() {
        assert.equal(Math.clz32(1), 31);
    });

    it("returns 0 for 2^31", function() {
        assert.equal(Math.clz32(Math.pow(2, 31)), 0);
    });

});

// ============================================================
// Date - static methods
// ============================================================

describe("Date.now", function() {

    it("returns a number", function() {
        assert.equal(typeof Date.now(), "number");
    });

    it("returns a positive integer", function() {
        assert.ok(Date.now() > 0);
    });

    it("is roughly equal to new Date().getTime()", function() {
        var a = Date.now();
        var b = new Date().getTime();
        assert.ok(Math.abs(b - a) < 100);
    });

});

describe("Date.prototype.toISOString", function() {

    it("returns a string", function() {
        assert.equal(typeof new Date().toISOString(), "string");
    });

    it("matches the ISO 8601 pattern", function() {
        var s = new Date().toISOString();
        assert.ok(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(s));
    });

    it("correctly encodes a known UTC timestamp", function() {
        // 2001-09-09T01:46:40.000Z  (Unix epoch 1000000000000 ms)
        assert.equal(new Date(1000000000000).toISOString(), "2001-09-09T01:46:40.000Z");
    });

    it("zero-pads month, day, hours, minutes, seconds", function() {
        // 2000-01-02T03:04:05.006Z
        assert.equal(new Date(Date.UTC(2000, 0, 2, 3, 4, 5, 6)).toISOString(),
                     "2000-01-02T03:04:05.006Z");
    });

    it("pads milliseconds to three digits", function() {
        var s = new Date(Date.UTC(2020, 0, 1, 0, 0, 0, 7)).toISOString();
        assert.ok(s.indexOf(".007Z") !== -1);
    });

    it("handles midnight UTC (all time fields zero)", function() {
        assert.equal(new Date(Date.UTC(2024, 5, 15, 0, 0, 0, 0)).toISOString(),
                     "2024-06-15T00:00:00.000Z");
    });

    it("throws RangeError for an invalid date", function() {
        assert.throws(function() { new Date(NaN).toISOString(); });
    });

});

// ============================================================
// Function - prototype methods
// ============================================================

describe("Function.prototype.bind", function() {

    it("binds this context", function() {
        var obj = {value: 42};
        var fn = function() { return this.value; };
        assert.equal(fn.bind(obj)(), 42);
    });

    it("partially applies arguments", function() {
        var add = function(a, b) { return a + b; };
        var add5 = add.bind(null, 5);
        assert.equal(add5(3), 8);
    });

    it("bound function has correct length", function() {
        var fn = function(a, b, c) { return a + b + c; };
        assert.doesNotThrow(function() { fn.bind(null, 1, 2)(3); });
    });

    it("curries arguments across bind and call", function() {
        var fn = function(a, b, c) { return a + b + c; };
        assert.equal(fn.bind(null, 1)(2, 3), 6);
    });

});


// ============================================================
// Print summary
// ============================================================

_test.summary({ exit: true });
