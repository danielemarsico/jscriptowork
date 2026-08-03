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

    it("formats a date as YYYYMMDD", function() {
        assert.equal(format_date(new Date(2020, 0, 2), "YYYYMMDD"), "20200102");
    });

    it("formats a date as YY/MM/DD", function() {
        assert.equal(format_date(new Date(2020, 0, 2), "YY/MM/DD"), "20/01/02");
    });

});

// ---------------------------------------------------------------------------
// parse_date
// ---------------------------------------------------------------------------

describe("parse_date", function() {

    it("parses DD/MM/YYYY into a Date", function() {
        var d = parse_date("02/01/2020", "DD/MM/YYYY");
        assert.ok(d instanceof Date);
    });

    it("returns a Date whose day/month/year match", function() {
        var d = parse_date("25/11/2021", "DD/MM/YYYY");
        assert.equal(d.getDate(), 25);
        assert.equal(d.getMonth(), 10);
        assert.equal(d.getFullYear(), 2021);
    });

    it("round-trips with format_date", function() {
        var d = parse_date("02/01/2020", "DD/MM/YYYY");
        assert.equal(format_date(d, "YYYY/MM/DD"), "2020/01/02");
    });

    it("falls back to the input for an unknown format", function() {
        assert.equal(parse_date("02/01/2020", "DD-MM-YYYY"), "02/01/2020");
    });

});

// ---------------------------------------------------------------------------
// randomString
// ---------------------------------------------------------------------------

describe("randomString", function() {

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

    it("keeps '=' characters inside a value", function() {
        var EQ_NAME = "_test_system_eq.properties";
        var EQ_PATH = CURRENT_FOLDER + "/" + EQ_NAME;
        write_text_to_file("k=a=b\n", EQ_PATH);
        load_properties(EQ_NAME);
        assert.equal(k, "a=b");
        delete_file(EQ_PATH);
    });

    it("ignores # comment lines", function() {
        var CMT_NAME = "_test_system_comments.properties";
        var CMT_PATH = CURRENT_FOLDER + "/" + CMT_NAME;
        write_text_to_file("# a comment\nalpha=1\n", CMT_PATH);
        load_properties(CMT_NAME);
        assert.equal(alpha, "1");
        delete_file(CMT_PATH);
    });

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

    it("falls back to CURRENT_FOLDER + INPUT_FOLDER when no .workspace exists", function() {
        if (file_exists(WS_PATH)) { delete_file(WS_PATH); }
        assert.equal(load_working_directory(), CURRENT_FOLDER + "\\");

        // assigned without `var` so it becomes a real global, per the
        // load()/eval scoping rule documented in CLAUDE.md.
        INPUT_FOLDER = "in";
        assert.equal(load_working_directory(), CURRENT_FOLDER + "in" + "\\");
    });

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

    // Reading from the real stdin would block an unattended run, so read(n)
    // and read_all() are exercised against a stubbed `stdin` object (assigned
    // without `var`, per the load()/eval scoping rule in CLAUDE.md).
    var _real_stdin = stdin;

    function stub_stdin(text) {
        var pos = 0;
        var obj = {};
        obj.AtEndOfStream = (pos >= text.length);
        obj.Read = function(n) {
            var chunk = text.substr(pos, n);
            pos += chunk.length;
            obj.AtEndOfStream = (pos >= text.length);
            return chunk;
        };
        obj.ReadAll = function() {
            var rest = text.substr(pos);
            pos = text.length;
            obj.AtEndOfStream = true;
            return rest;
        };
        stdin = obj;
    }

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

    it("read(n) returns n characters", function() {
        stub_stdin("hello world");
        assert.equal(read(5), "hello");
        assert.equal(read(1), " ");
        assert.equal(read(5), "world");
        stdin = _real_stdin;
    });

    it("read_all() returns an empty string at end of stream", function() {
        stub_stdin("");
        assert.equal(read_all(), "");
        stdin = _real_stdin;
    });

    it("read_all() reads everything left in the stream", function() {
        stub_stdin("rest of the input");
        assert.equal(read_all(), "rest of the input");
        stdin = _real_stdin;
    });

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

_test.summary({ exit: true });
