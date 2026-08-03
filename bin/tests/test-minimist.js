// test-minimist.js - Tests for libs/minimist.js
//
// Run via:  cscript.exe launcher.js test-minimist.js

load("core");
load("polyfills");
load("minimist");
load("minitest");

// ---------------------------------------------------------------------------
// Basics
// ---------------------------------------------------------------------------

describe("minimist - basics", function() {

    it("returns only the positional bucket for empty input", function() {
        assert.deepEqual(minimist([], {}), { _: [] });
    });

    it("collects bare arguments as positionals", function() {
        var argv = minimist(["a", "b"], {});
        assert.deepEqual(argv._, ["a", "b"]);
    });

    it("converts numeric positionals to numbers", function() {
        var argv = minimist(["3", "4.5"], {});
        assert.deepEqual(argv._, [3, 4.5]);
    });

    it("keeps positionals as strings when '_' is declared a string option", function() {
        var argv = minimist(["3"], { string: "_" });
        assert.deepEqual(argv._, ["3"]);
    });

});

// ---------------------------------------------------------------------------
// Long options
// ---------------------------------------------------------------------------

describe("minimist - long options", function() {

    it("parses --flag value", function() {
        var argv = minimist(["--foo", "bar"], {});
        assert.equal(argv.foo, "bar");
        assert.deepEqual(argv._, []);
    });

    it("parses --flag=value", function() {
        assert.equal(minimist(["--foo=bar"], {}).foo, "bar");
    });

    it("keeps '=' characters inside the value of --flag=value", function() {
        assert.equal(minimist(["--foo=a=b"], {}).foo, "a=b");
    });

    it("parses an empty value in --flag=", function() {
        assert.equal(minimist(["--foo="], {}).foo, "");
    });

    it("converts numeric values to numbers", function() {
        var argv = minimist(["--count", "5"], {});
        assert.equal(argv.count, 5);
        assert.equal(typeof argv.count, "number");
    });

    it("treats a trailing --flag with no value as true", function() {
        assert.equal(minimist(["--verbose"], {}).verbose, true);
    });

    it("treats --flag followed by another flag as true", function() {
        var argv = minimist(["--a", "--b"], {});
        assert.equal(argv.a, true);
        assert.equal(argv.b, true);
    });

    it("parses --no-flag as false", function() {
        assert.equal(minimist(["--no-cache"], {}).cache, false);
    });

    it("collects repeated flags into an array", function() {
        var argv = minimist(["--x", "1", "--x", "2", "--x", "3"], {});
        assert.deepEqual(argv.x, [1, 2, 3]);
    });

    it("mixes flags and positionals", function() {
        var argv = minimist(["cmd", "--foo", "bar", "extra"], {});
        assert.equal(argv.foo, "bar");
        assert.deepEqual(argv._, ["cmd", "extra"]);
    });

    it("builds nested objects from dotted keys", function() {
        var argv = minimist(["--db.host=localhost", "--db.port=5432"], {});
        assert.equal(argv.db.host, "localhost");
        assert.equal(argv.db.port, 5432);
    });

    it("ignores a __proto__ key", function() {
        var argv = minimist(["--__proto__.polluted=yes"], {});
        assert.equal(argv.polluted, undefined);
        assert.equal(({}).polluted,  undefined);
    });

});

// ---------------------------------------------------------------------------
// opts.string
// ---------------------------------------------------------------------------

describe("minimist - opts.string", function() {

    it("keeps a numeric value as a string", function() {
        var argv = minimist(["--count", "5"], { string: "count" });
        assert.equal(argv.count, "5");
        assert.equal(typeof argv.count, "string");
    });

    it("accepts an array of string option names", function() {
        var argv = minimist(["--a", "1", "--b", "2"], { string: ["a", "b"] });
        assert.equal(argv.a, "1");
        assert.equal(argv.b, "2");
    });

    it("yields an empty string for a valueless string flag", function() {
        assert.equal(minimist(["--name"], { string: "name" }).name, "");
    });

});

// ---------------------------------------------------------------------------
// opts.boolean
// ---------------------------------------------------------------------------

describe("minimist - opts.boolean", function() {

    it("does not consume the next argument for a boolean flag", function() {
        var argv = minimist(["--verbose", "file.txt"], { boolean: "verbose" });
        assert.equal(argv.verbose, true);
        assert.deepEqual(argv._, ["file.txt"]);
    });

    it("accepts an explicit true/false value", function() {
        assert.equal(minimist(["--verbose", "false"], { boolean: "verbose" }).verbose, false);
        assert.equal(minimist(["--verbose", "true"],  { boolean: "verbose" }).verbose, true);
    });

    it("parses --flag=false as false", function() {
        assert.equal(minimist(["--verbose=false"], { boolean: "verbose" }).verbose, false);
    });

    it("accepts an array of boolean option names", function() {
        var argv = minimist(["--a", "--b"], { boolean: ["a", "b"] });
        assert.equal(argv.a, true);
        assert.equal(argv.b, true);
    });

    it("treats every long flag as boolean when boolean:true", function() {
        var argv = minimist(["--x", "--y"], { boolean: true });
        assert.equal(argv.x, true);
        assert.equal(argv.y, true);
    });

});

// ---------------------------------------------------------------------------
// opts.alias
// ---------------------------------------------------------------------------

describe("minimist - opts.alias", function() {

    it("mirrors a value onto its alias", function() {
        var argv = minimist(["--foo", "bar"], { alias: { foo: "f" } });
        assert.equal(argv.foo, "bar");
        assert.equal(argv.f,   "bar");
    });

    it("supports multiple aliases for one key", function() {
        var argv = minimist(["--foo", "bar"], { alias: { foo: ["f", "fu"] } });
        assert.equal(argv.f,  "bar");
        assert.equal(argv.fu, "bar");
    });

    it("mirrors defaults onto aliases too", function() {
        var argv = minimist([], { alias: { foo: "f" }, "default": { foo: "d" } });
        assert.equal(argv.foo, "d");
        assert.equal(argv.f,   "d");
    });

});

// ---------------------------------------------------------------------------
// opts.default
// ---------------------------------------------------------------------------

describe("minimist - opts.default", function() {

    it("applies a default when the flag is absent", function() {
        assert.equal(minimist([], { "default": { port: 8080 } }).port, 8080);
    });

    it("does not override a value that was supplied", function() {
        assert.equal(minimist(["--port", "9"], { "default": { port: 8080 } }).port, 9);
    });

    it("applies several defaults at once", function() {
        var argv = minimist([], { "default": { a: 1, b: "two" } });
        assert.equal(argv.a, 1);
        assert.equal(argv.b, "two");
    });

    it("applies a false default", function() {
        assert.equal(minimist([], { "default": { flag: false } }).flag, false);
    });

});

// ---------------------------------------------------------------------------
// The -- separator
// ---------------------------------------------------------------------------

describe("minimist - '--' separator", function() {

    it("pushes everything after -- into positionals by default", function() {
        var argv = minimist(["a", "--", "b", "--c"], {});
        assert.deepEqual(argv._, ["a", "b", "--c"]);
    });

    it("collects them under '--' when opts['--'] is set", function() {
        var argv = minimist(["a", "--", "b", "--c"], { "--": true });
        assert.deepEqual(argv._,      ["a"]);
        assert.deepEqual(argv["--"], ["b", "--c"]);
    });

    it("does not parse flags that appear after --", function() {
        var argv = minimist(["--", "--foo", "bar"], { "--": true });
        assert.equal(argv.foo, undefined);
    });

});

// ---------------------------------------------------------------------------
// opts.stopEarly and opts.unknown
// ---------------------------------------------------------------------------

describe("minimist - stopEarly", function() {

    it("stops parsing at the first positional", function() {
        var argv = minimist(["--a", "1", "b", "--c"], { stopEarly: true });
        assert.equal(argv.a, 1);
        assert.equal(argv.c, undefined);
        assert.deepEqual(argv._, ["b", "--c"]);
    });

});

describe("minimist - unknown callback", function() {

    it("sees every undeclared argument", function() {
        var seen = [];
        minimist(["--foo", "bar", "baz"], {
            unknown: function(arg) { seen.push(arg); return false; }
        });
        assert.deepEqual(seen, ["--foo", "baz"]);
    });

    it("drops arguments the callback rejects", function() {
        var argv = minimist(["--foo", "bar", "baz"], {
            unknown: function() { return false; }
        });
        assert.equal(argv.foo, undefined);
        assert.deepEqual(argv._, []);
    });

    it("keeps arguments the callback accepts", function() {
        var argv = minimist(["--foo", "bar", "baz"], {
            unknown: function() { return true; }
        });
        assert.equal(argv.foo, "bar");
        assert.deepEqual(argv._, ["baz"]);
    });

    it("is not consulted for declared options", function() {
        var seen = [];
        var argv = minimist(["--foo", "bar"], {
            string:  "foo",
            unknown: function(arg) { seen.push(arg); return false; }
        });
        assert.equal(argv.foo, "bar");
        assert.deepEqual(seen, []);
    });

});

// ---------------------------------------------------------------------------
// Known gaps — see TODO.md
// ---------------------------------------------------------------------------

describe("minimist - short options", function() {

    it("parses -f value", function() {
        var argv = minimist(["-f", "bar"], {});
        assert.equal(argv.f, "bar");
        assert.deepEqual(argv._, []);
    });

    it("parses a boolean short flag -v", function() {
        var argv = minimist(["-v"], { boolean: "v" });
        assert.equal(argv.v, true);
    });

    it("expands a cluster -abc to three flags", function() {
        var argv = minimist(["-abc"], {});
        assert.equal(argv.a, true);
        assert.equal(argv.b, true);
        assert.equal(argv.c, true);
    });

    it("parses -n5 as n=5", function() {
        var argv = minimist(["-n5"], {});
        assert.equal(argv.n, 5);
    });

    it("parses -x=1 as x=1", function() {
        var argv = minimist(["-x=1"], {});
        assert.equal(argv.x, 1);
    });

    it("mirrors a short flag onto its alias", function() {
        var argv = minimist(["-f", "bar"], { alias: { f: "file" } });
        assert.equal(argv.f, "bar");
        assert.equal(argv.file, "bar");
    });

});

describe("minimist - options object", function() {

    it("treats the options argument as optional", function() {
        var argv = minimist(["--foo", "bar", "positional"]);
        assert.equal(argv.foo, "bar");
        assert.deepEqual(argv._, ["positional"]);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
