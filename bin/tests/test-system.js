// test-system.js - Tests for the non-file-system parts of libs/system.js
//
// File-system helpers live in test-filesystem.js and http_request in
// test-http.js; this suite covers dates, random strings, the properties loader,
// the workspace file, and the stdio surface.
//
// It writes two temporary files into the launcher's own folder
// (CURRENT_FOLDER), because load_properties() and save_working_directory()
// resolve their paths relative to it. Both are removed at the end, and any
// pre-existing .workspace file is saved and restored.
//
// Run via:  cscript.exe launcher.js test-system.js

load("core");
load("polyfills");
load("system");
load("minitest");

// ---------------------------------------------------------------------------
// format_date
// ---------------------------------------------------------------------------

describe("format_date", function() {

    it("formats a date as YYYY/MM/DD", function() {
        assert.equal(format_date(new Date(2020, 0, 2), "YYYY/MM/DD"), "2020/01/02");
    });

    it("zero-pads month and day", function() {
        assert.equal(format_date(new Date(2021, 8, 5), "YYYY/MM/DD"), "2021/09/05");
    });

    it("does not pad a two-digit month or day", function() {
        assert.equal(format_date(new Date(2021, 10, 25), "YYYY/MM/DD"), "2021/11/25");
    });

    it("handles the last day of the year", function() {
        assert.equal(format_date(new Date(1999, 11, 31), "YYYY/MM/DD"), "1999/12/31");
    });

    it("handles a leap day", function() {
        assert.equal(format_date(new Date(2024, 1, 29), "YYYY/MM/DD"), "2024/02/29");
    });

    it("falls back to Date#toString for an unknown format", function() {
        var d = new Date(2020, 0, 2);
        assert.equal(format_date(d, "DD-MM-YYYY"), d.toString());
    });

    skip("formats a date as YYYYMMDD",
         "the comment on libs/system.js:146 promises it but only YYYY/MM/DD is implemented (TODO.md)");
    skip("formats a date as YY/MM/DD",
         "the comment on libs/system.js:146 promises it but only YYYY/MM/DD is implemented (TODO.md)");

});

// ---------------------------------------------------------------------------
// parse_date
// ---------------------------------------------------------------------------

describe("parse_date", function() {

    var BUG = "libs/system.js:161 reads d.substring() before d is assigned and ignores its ds parameter, so every call throws (TODO.md)";

    skip("parses DD/MM/YYYY into a Date",                BUG);
    skip("returns a Date whose day/month/year match",    BUG);
    skip("round-trips with format_date",                 BUG);
    skip("falls back to the input for an unknown format", BUG);

});

// ---------------------------------------------------------------------------
// randomString
// ---------------------------------------------------------------------------

describe("randomString", function() {

    // randomString is a `function` declaration inside system.js, so it is
    // scoped to load() and unreachable when running through bin/launcher.js.
    // It IS reachable from the bundled dist/launcher.js, where libs are inlined
    // at top level — so run the tests when it is there, skip when it is not.
    if (typeof randomString !== "function") {
        var GONE = "randomString is not a global under bin/launcher.js; it is declared with `function` instead of an assignment (TODO.md)";
        skip("returns a string of the requested length", GONE);
        skip("returns an empty string for length 0",     GONE);
        skip("uses only characters from the charset",    GONE);
        skip("uses the default alphanumeric charset",    GONE);
        skip("produces different values on repeat calls", GONE);
        return;
    }

    it("returns a string of the requested length", function() {
        assert.equal(typeof randomString(10), "string");
        assert.equal(randomString(10).length,  10);
        assert.equal(randomString(1).length,   1);
    });

    it("returns an empty string for length 0", function() {
        assert.equal(randomString(0), "");
    });

    it("uses only characters from the charset", function() {
        assert.ok(/^[AB]{20}$/.test(randomString(20, "AB")));
    });

    it("uses the default alphanumeric charset", function() {
        assert.ok(/^[A-Za-z0-9]{50}$/.test(randomString(50)));
    });

    it("produces different values on repeat calls", function() {
        // 32 chars from a 62-char alphabet: a collision means it is not random.
        assert.notEqual(randomString(32), randomString(32));
    });

});

// ---------------------------------------------------------------------------
// load_properties
// ---------------------------------------------------------------------------

describe("load_properties", function() {

    // load_properties resolves its argument against CURRENT_FOLDER and eval()s
    // the result, defining one global per key=value line.
    var PROPS_NAME = "_test_system.properties";
    var PROPS_PATH = CURRENT_FOLDER + "/" + PROPS_NAME;

    write_text_to_file("greeting=hello\nnumber=42\nempty_line_follows=yes\n\n", PROPS_PATH);
    load_properties(PROPS_NAME);

    it("defines a global for each key", function() {
        assert.equal(greeting, "hello");
    });

    it("defines every value as a string", function() {
        assert.equal(number, "42");
        assert.equal(typeof number, "string");
    });

    it("ignores blank lines", function() {
        assert.equal(empty_line_follows, "yes");
    });

    it("does not define anything for a key that was not in the file", function() {
        assert.equal(typeof not_in_the_properties_file, "undefined");
    });

    skip("keeps '=' characters inside a value",
         "load_properties requires exactly two split('=') parts, so 'k=a=b' is silently dropped (TODO.md)");
    skip("ignores # comment lines",
         "comment lines are not recognised; '#x=y' would define a global named '#x' (TODO.md)");

    delete_file(PROPS_PATH);

});

// ---------------------------------------------------------------------------
// save_working_directory / load_working_directory
// ---------------------------------------------------------------------------

describe("working directory", function() {

    var WS_PATH = CURRENT_FOLDER + "/.workspace";
    var backup  = file_exists(WS_PATH) ? read_text_file(WS_PATH) : null;

    it("round-trips a path through the .workspace file", function() {
        save_working_directory("C:\\some\\folder");
        assert.equal(load_working_directory(), "C:\\some\\folder");
    });

    it("creates the .workspace file", function() {
        save_working_directory("C:\\another");
        assert.ok(file_exists(WS_PATH));
    });

    it("overwrites a previously saved path", function() {
        save_working_directory("C:\\first");
        save_working_directory("C:\\second");
        assert.equal(load_working_directory(), "C:\\second");
    });

    it("strips the trailing newline written by save_working_directory", function() {
        save_working_directory("C:\\trimmed");
        assert.equal(load_working_directory().indexOf("\n"), -1);
        assert.equal(load_working_directory().indexOf("\r"), -1);
    });

    skip("falls back to CURRENT_FOLDER + INPUT_FOLDER when no .workspace exists",
         "INPUT_FOLDER is never defined by either launcher, so the fallback throws ReferenceError (TODO.md)");

    // restore whatever was there before
    if (backup === null) {
        if (file_exists(WS_PATH)) { delete_file(WS_PATH); }
    } else {
        write_text_to_file(backup, WS_PATH);
    }

});

// ---------------------------------------------------------------------------
// stdio surface
// ---------------------------------------------------------------------------

describe("stdio helpers", function() {

    // Reading from stdin would block an unattended run, so only the surface is
    // checked here; the interactive prompts are exercised in test-helpers.js
    // against a stubbed read_line.

    it("exposes the stdin/stdout stream objects", function() {
        assert.notEqual(typeof stdin,  "undefined");
        assert.notEqual(typeof stdout, "undefined");
    });

    it("exposes read_line, read, read_all, write_line and write", function() {
        assert.equal(typeof read_line,  "function");
        assert.equal(typeof read,       "function");
        assert.equal(typeof read_all,   "function");
        assert.equal(typeof write_line, "function");
        assert.equal(typeof write,      "function");
    });

    it("write and write_line do not throw", function() {
        assert.doesNotThrow(function() { write(""); });
        assert.doesNotThrow(function() { write_line(""); });
    });

    skip("read(n) returns n characters",
         "read() ignores its argument and always reads a single character (TODO.md)");
    skip("read_all() returns null at end of stream",
         "read_all() returns the literal string 'end of stream' as a sentinel (TODO.md)");

});

// ---------------------------------------------------------------------------
// http_request — offline checks only (see test-http.js for the network suite)
// ---------------------------------------------------------------------------

describe("http_request - offline", function() {

    it("is defined", function() {
        assert.equal(typeof http_request, "function");
    });

    it("rejects an unsupported method before touching the network", function() {
        assert.throws(function() { http_request("http://example.invalid/", "PATCH", function() {}); });
        assert.throws(function() { http_request("http://example.invalid/", "get",   function() {}); });
        assert.throws(function() { http_request("http://example.invalid/", "",      function() {}); });
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
