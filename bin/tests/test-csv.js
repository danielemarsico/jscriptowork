// test-csv.js - Tests for libs/csv.js
//
// csv_parse/csv_format are pure string work and need nothing but a JS engine.
// read_csv_file/write_csv_file round-trip through the system temp folder.
//
// Run via:  cscript.exe launcher.js test-csv.js

load("core");
load("polyfills");
load("system");
load("csv");
load("minitest");

// ---------------------------------------------------------------------------
// Temp folder
// ---------------------------------------------------------------------------

var TEMP = (function() {
    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var base = fso.GetSpecialFolder(2).Path; // 2 = TemporaryFolder
    return base + "\\jscriptowork_csv_test";
}());

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
    fso.CreateFolder(TEMP);
}());

function tmp(name) { return TEMP + "\\" + name; }

// ===========================================================================
// csv_parse
// ===========================================================================

describe("csv_parse - basics", function() {

    it("parses a single row", function() {
        assert.deepEqual(csv_parse("a,b,c"), [["a", "b", "c"]]);
    });

    it("parses several rows", function() {
        assert.deepEqual(csv_parse("a,b\nc,d"), [["a", "b"], ["c", "d"]]);
    });

    it("parses a single field", function() {
        assert.deepEqual(csv_parse("hello"), [["hello"]]);
    });

    it("returns an empty array for empty input", function() {
        assert.deepEqual(csv_parse(""), []);
    });

    it("returns an empty array for null or undefined", function() {
        assert.deepEqual(csv_parse(null), []);
        assert.deepEqual(csv_parse(undefined), []);
    });

    it("keeps empty fields", function() {
        assert.deepEqual(csv_parse("a,,c"), [["a", "", "c"]]);
    });

    it("keeps a leading empty field", function() {
        assert.deepEqual(csv_parse(",b"), [["", "b"]]);
    });

    it("keeps a trailing empty field", function() {
        assert.deepEqual(csv_parse("a,"), [["a", ""]]);
    });

    it("keeps every value as a string", function() {
        var rows = csv_parse("1,2.5,true");
        assert.equal(typeof rows[0][0], "string");
        assert.equal(rows[0][0], "1");
        assert.equal(rows[0][2], "true");
    });

    it("preserves surrounding whitespace by default", function() {
        assert.deepEqual(csv_parse("  a  ,b"), [["  a  ", "b"]]);
    });

    it("trims fields when options.trim is set", function() {
        assert.deepEqual(csv_parse("  a  ,  b  ", { trim: true }), [["a", "b"]]);
    });

});

describe("csv_parse - line endings", function() {

    it("handles LF", function() {
        assert.deepEqual(csv_parse("a\nb"), [["a"], ["b"]]);
    });

    it("handles CRLF", function() {
        assert.deepEqual(csv_parse("a\r\nb"), [["a"], ["b"]]);
    });

    it("handles a lone CR", function() {
        assert.deepEqual(csv_parse("a\rb"), [["a"], ["b"]]);
    });

    it("handles mixed line endings in one document", function() {
        assert.deepEqual(csv_parse("a\r\nb\nc\rd"), [["a"], ["b"], ["c"], ["d"]]);
    });

    it("ignores a single trailing newline", function() {
        assert.deepEqual(csv_parse("a,b\n"),   [["a", "b"]]);
        assert.deepEqual(csv_parse("a,b\r\n"), [["a", "b"]]);
    });

    it("treats a blank line in the middle as a one-empty-field row", function() {
        assert.deepEqual(csv_parse("a\n\nb"), [["a"], [""], ["b"]]);
    });

});

describe("csv_parse - quoting", function() {

    it("unwraps a quoted field", function() {
        assert.deepEqual(csv_parse('"a",b'), [["a", "b"]]);
    });

    it("keeps a delimiter inside quotes", function() {
        assert.deepEqual(csv_parse('"a,b",c'), [["a,b", "c"]]);
    });

    it("keeps a newline inside quotes", function() {
        assert.deepEqual(csv_parse('"line1\nline2",b'), [["line1\nline2", "b"]]);
    });

    it("keeps a CRLF inside quotes", function() {
        assert.deepEqual(csv_parse('"a\r\nb"'), [["a\r\nb"]]);
    });

    it("unescapes a doubled quote", function() {
        assert.deepEqual(csv_parse('"say ""hi"""'), [['say "hi"']]);
    });

    it("handles a field that is only a doubled quote", function() {
        assert.deepEqual(csv_parse('""""'), [['"']]);
    });

    it("handles an empty quoted field", function() {
        assert.deepEqual(csv_parse('"",a'), [["", "a"]]);
    });

    it("handles every field quoted", function() {
        assert.deepEqual(csv_parse('"a","b"\n"c","d"'), [["a", "b"], ["c", "d"]]);
    });

    it("keeps whitespace inside quotes even with trim", function() {
        assert.deepEqual(csv_parse('"  a  "', { trim: true }), [["  a  "]]);
    });

    it("handles a quoted field followed by a row break", function() {
        assert.deepEqual(csv_parse('"a"\n"b"'), [["a"], ["b"]]);
    });

});

describe("csv_parse - delimiter", function() {

    it("parses semicolon-separated values", function() {
        assert.deepEqual(csv_parse("a;b;c", { delimiter: ";" }), [["a", "b", "c"]]);
    });

    it("parses tab-separated values", function() {
        assert.deepEqual(csv_parse("a\tb", { delimiter: "\t" }), [["a", "b"]]);
    });

    it("treats a comma as ordinary text when the delimiter is a semicolon", function() {
        assert.deepEqual(csv_parse("a,b;c", { delimiter: ";" }), [["a,b", "c"]]);
    });

    it("still honours quoting with a custom delimiter", function() {
        assert.deepEqual(csv_parse('"a;b";c', { delimiter: ";" }), [["a;b", "c"]]);
    });

});

describe("csv_parse - headers", function() {

    it("returns objects keyed by the first row", function() {
        assert.deepEqual(
            csv_parse("name,age\nalice,30", { headers: true }),
            [{ name: "alice", age: "30" }]
        );
    });

    it("returns one object per data row", function() {
        var out = csv_parse("a,b\n1,2\n3,4", { headers: true });
        assert.equal(out.length, 2);
        assert.equal(out[1].a, "3");
    });

    it("returns an empty array when there are no data rows", function() {
        assert.deepEqual(csv_parse("name,age", { headers: true }), []);
    });

    it("returns an empty array for empty input", function() {
        assert.deepEqual(csv_parse("", { headers: true }), []);
    });

    it("pads a short row with empty strings", function() {
        assert.deepEqual(
            csv_parse("a,b,c\n1,2", { headers: true }),
            [{ a: "1", b: "2", c: "" }]
        );
    });

    it("ignores extra fields beyond the header count", function() {
        assert.deepEqual(
            csv_parse("a,b\n1,2,3", { headers: true }),
            [{ a: "1", b: "2" }]
        );
    });

    it("handles quoted headers", function() {
        assert.deepEqual(
            csv_parse('"first name",age\nalice,30', { headers: true }),
            [{ "first name": "alice", age: "30" }]
        );
    });

});

describe("csv_parse - BOM", function() {

    it("strips a U+FEFF BOM", function() {
        assert.deepEqual(csv_parse("\uFEFFa,b"), [["a", "b"]]);
    });

    it("strips raw UTF-8 BOM bytes read as ASCII", function() {
        assert.deepEqual(csv_parse("\u00EF\u00BB\u00BFa,b"), [["a", "b"]]);
    });

    it("does not strip anything from a normal document", function() {
        assert.deepEqual(csv_parse("a,b"), [["a", "b"]]);
    });

});

// ===========================================================================
// csv_format
// ===========================================================================

describe("csv_format - arrays", function() {

    it("formats a single row", function() {
        assert.equal(csv_format([["a", "b"]]), "a,b\r\n");
    });

    it("formats several rows", function() {
        assert.equal(csv_format([["a"], ["b"]]), "a\r\nb\r\n");
    });

    it("returns an empty string for an empty array", function() {
        assert.equal(csv_format([]), "");
    });

    it("returns an empty string for null", function() {
        assert.equal(csv_format(null), "");
    });

    it("honours a custom end-of-line", function() {
        assert.equal(csv_format([["a"], ["b"]], { eol: "\n" }), "a\nb\n");
    });

    it("honours a custom delimiter", function() {
        assert.equal(csv_format([["a", "b"]], { delimiter: ";" }), "a;b\r\n");
    });

    it("writes numbers and booleans as text", function() {
        assert.equal(csv_format([[1, true]]), "1,true\r\n");
    });

    it("writes null and undefined as empty fields", function() {
        assert.equal(csv_format([[null, undefined]]), ",\r\n");
    });

});

describe("csv_format - quoting", function() {

    it("does not quote a plain field", function() {
        assert.equal(csv_format([["plain"]]), "plain\r\n");
    });

    it("quotes a field containing the delimiter", function() {
        assert.equal(csv_format([["a,b"]]), '"a,b"\r\n');
    });

    it("quotes and doubles an embedded quote", function() {
        assert.equal(csv_format([['say "hi"']]), '"say ""hi"""\r\n');
    });

    it("quotes a field containing a newline", function() {
        assert.equal(csv_format([["a\nb"]]), '"a\nb"\r\n');
    });

    it("quotes a field containing a carriage return", function() {
        assert.equal(csv_format([["a\rb"]]), '"a\rb"\r\n');
    });

    it("quotes every field when quote_all is set", function() {
        assert.equal(csv_format([["a", "b"]], { quote_all: true }), '"a","b"\r\n');
    });

    it("quotes on the custom delimiter, not on a comma", function() {
        assert.equal(csv_format([["a,b"]], { delimiter: ";" }), "a,b\r\n");
        assert.equal(csv_format([["a;b"]], { delimiter: ";" }), '"a;b"\r\n');
    });

});

describe("csv_format - objects", function() {

    it("derives a header row from the object keys", function() {
        assert.equal(csv_format([{ a: 1, b: 2 }]), "a,b\r\n1,2\r\n");
    });

    it("writes one row per object", function() {
        assert.equal(csv_format([{ a: 1 }, { a: 2 }]), "a\r\n1\r\n2\r\n");
    });

    it("uses an explicit header array as the column order", function() {
        assert.equal(csv_format([{ a: 1, b: 2 }], { headers: ["b", "a"] }), "b,a\r\n2,1\r\n");
    });

    it("selects a subset of columns with an explicit header array", function() {
        assert.equal(csv_format([{ a: 1, b: 2 }], { headers: ["a"] }), "a\r\n1\r\n");
    });

    it("omits the header row when headers is false", function() {
        assert.equal(csv_format([{ a: 1, b: 2 }], { headers: false }), "1,2\r\n");
    });

    it("includes a key that only a later object introduces", function() {
        assert.equal(csv_format([{ a: 1 }, { a: 2, b: 3 }]), "a,b\r\n1,\r\n2,3\r\n");
    });

    it("writes an empty field for a missing key", function() {
        assert.equal(csv_format([{ a: 1, b: 2 }, { a: 3 }]), "a,b\r\n1,2\r\n3,\r\n");
    });

    it("quotes header names that need it", function() {
        assert.equal(csv_format([{ "a,b": 1 }]), '"a,b"\r\n1\r\n');
    });

});

// ===========================================================================
// Round trips
// ===========================================================================

describe("csv round trip", function() {

    it("survives quotes, delimiters and newlines", function() {
        var rows = [["plain", 'has "quotes"'], ["has,comma", "has\nnewline"]];
        assert.deepEqual(csv_parse(csv_format(rows)), rows);
    });

    it("survives empty fields", function() {
        var rows = [["", "a", ""], ["b", "", "c"]];
        assert.deepEqual(csv_parse(csv_format(rows)), rows);
    });

    it("survives a custom delimiter", function() {
        var rows = [["a;b", "c"], ["d", "e"]];
        var text = csv_format(rows, { delimiter: ";" });
        assert.deepEqual(csv_parse(text, { delimiter: ";" }), rows);
    });

    it("survives objects through headers", function() {
        var records = [{ name: "alice", note: "likes, commas" }, { name: "bob", note: "" }];
        assert.deepEqual(csv_parse(csv_format(records), { headers: true }), records);
    });

    it("survives a field that is only whitespace", function() {
        var rows = [["   ", "a"]];
        assert.deepEqual(csv_parse(csv_format(rows)), rows);
    });

});

// ===========================================================================
// File helpers
// ===========================================================================

describe("write_csv_file / read_csv_file", function() {

    it("writes a file that can be read back", function() {
        var p = tmp("basic.csv");
        write_csv_file(p, [["a", "b"], ["c", "d"]]);
        assert.ok(file_exists(p));
        assert.deepEqual(read_csv_file(p), [["a", "b"], ["c", "d"]]);
    });

    it("round-trips objects with a header row", function() {
        var p = tmp("objects.csv");
        var records = [{ id: "1", name: "alice" }, { id: "2", name: "bob" }];
        write_csv_file(p, records);
        assert.deepEqual(read_csv_file(p, { headers: true }), records);
    });

    it("round-trips fields containing delimiters and newlines", function() {
        var p = tmp("tricky.csv");
        var rows = [["a,b", "c\nd"], ['e"f', "g"]];
        write_csv_file(p, rows);
        assert.deepEqual(read_csv_file(p), rows);
    });

    it("honours a custom delimiter on both sides", function() {
        var p = tmp("semicolon.csv");
        write_csv_file(p, [["a", "b"]], { delimiter: ";" });
        assert.equal(read_text_file(p), "a;b\r\n");
        assert.deepEqual(read_csv_file(p, { delimiter: ";" }), [["a", "b"]]);
    });

    it("overwrites an existing file", function() {
        var p = tmp("overwrite.csv");
        write_csv_file(p, [["first"]]);
        write_csv_file(p, [["second"]]);
        assert.deepEqual(read_csv_file(p), [["second"]]);
    });

    it("writes an empty file for no rows", function() {
        var p = tmp("empty.csv");
        write_csv_file(p, []);
        assert.deepEqual(read_csv_file(p), []);
    });

    it("round-trips a unicode file", function() {
        var p = tmp("unicode.csv");
        var rows = [["caf\u00E9", "\u65E5\u672C"]];
        write_csv_file(p, rows, { unicode: true });
        assert.deepEqual(read_csv_file(p), rows);
    });

    it("reads a file written by hand with LF endings", function() {
        var p = tmp("handwritten.csv");
        write_text_to_file("name,age\nalice,30\nbob,25\n", p);
        assert.deepEqual(read_csv_file(p, { headers: true }), [
            { name: "alice", age: "30" },
            { name: "bob",   age: "25" }
        ]);
    });

});

// ---------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
}());

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
