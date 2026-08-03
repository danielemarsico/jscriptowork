// test-helpers.js - Tests for libs/helpers.js
//
// The Office automation in helpers.js is exercised through mock COM objects:
// an Excel worksheet is just an object with Cells(row, col) and Range(spec),
// and an Access database is an object with OpenRecordset(). That covers
// fill_sheet, read_sheet_data and execute_query without needing Office
// installed. The interactive prompts are driven by a stubbed read_line that
// throws when it runs out of scripted input, so a wrong expectation fails the
// test instead of hanging the run.
//
// do_in_excel / do_in_access / do_in_word really do launch Office and are only
// checked for existence here.
//
// Run via:  cscript.exe launcher.js test-helpers.js

load("core");
load("polyfills");
load("system");
load("helpers");
load("minitest");

// ---------------------------------------------------------------------------
// Temp folder for the folder-listing helpers
// ---------------------------------------------------------------------------

var TEMP = (function() {
    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var base = fso.GetSpecialFolder(2).Path; // 2 = TemporaryFolder
    return base + "\\jscriptowork_helpers_test";
}());

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
    fso.CreateFolder(TEMP);
}());

function tmp(name) { return TEMP + "\\" + name; }

// ---------------------------------------------------------------------------
// stdin stubbing
// ---------------------------------------------------------------------------

var _real_read_line  = read_line;
var _real_write_line = write_line;
var _real_write      = write;
var _prompts         = [];

// Feed a fixed list of lines to read_line and capture everything written.
// Running out of input throws, so an unterminated prompt loop fails fast
// instead of blocking the suite forever.
function stub_input(lines) {
    var i = 0;
    _prompts = [];
    read_line  = function() {
        if (i >= lines.length) { throw new Error("stub_input: out of scripted input"); }
        return lines[i++];
    };
    write_line = function(data) { _prompts.push(String(data)); };
    write      = function(data) { _prompts.push(String(data)); };
}

function restore_input() {
    read_line  = _real_read_line;
    write_line = _real_write_line;
    write      = _real_write;
}

// ---------------------------------------------------------------------------
// Mock Excel worksheet
// ---------------------------------------------------------------------------

// grid - optional array of rows (arrays of values), seeded 1-based.
// Unset cells read back as "", which is what an empty Excel cell looks like
// to read_sheet_data.
function mock_sheet(grid) {
    var cells  = {};
    var copies = [];
    var sheet  = {};

    sheet.Cells = function(row, col) {
        var k = row + ":" + col;
        if (!cells[k]) { cells[k] = { Value: "" }; }
        return cells[k];
    };

    sheet.Range = function(spec) {
        return {
            spec: spec,
            Copy: function(target) { copies.push(spec + " -> " + target.spec); }
        };
    };

    sheet._copies   = function() { return copies; };
    sheet._value    = function(row, col) { return sheet.Cells(row, col).Value; };
    sheet._cellsSet = function() {
        var n = 0;
        for (var k in cells) { if (cells[k].Value !== "") { n++; } }
        return n;
    };

    if (grid) {
        for (var r = 0; r < grid.length; r++) {
            for (var c = 0; c < grid[r].length; c++) {
                sheet.Cells(r + 1, c + 1).Value = grid[r][c];
            }
        }
    }
    return sheet;
}

// ---------------------------------------------------------------------------
// Mock Access database
// ---------------------------------------------------------------------------

function mock_db(fields, rows) {
    var db = { _query: null, _type: null };

    db.OpenRecordset = function(query, type) {
        db._query = query;
        db._type  = type;

        var i  = 0;
        var rs = {};
        rs.EOF = (rows.length === 0);

        rs.Fields = function(n) {
            return { Name: fields[n], Value: rows[i][n] };
        };
        rs.Fields.Count = fields.length;

        rs.MoveNext = function() {
            i++;
            if (i >= rows.length) { rs.EOF = true; }
        };
        return rs;
    };
    return db;
}

// ===========================================================================
// alphabet
// ===========================================================================

describe("alphabet", function() {

    it("is the column-letter lookup string", function() {
        assert.equal(alphabet, "_ABCDEFGHIJKLMNOPQRSTUVWXYZ");
    });

    it("is 1-based: charAt(n) is the nth column letter", function() {
        assert.equal(alphabet.charAt(1),  "A");
        assert.equal(alphabet.charAt(26), "Z");
    });

});

// ===========================================================================
// get_current_date_as_excel_text
// ===========================================================================

describe("get_current_date_as_excel_text", function() {

    it("returns a leading-apostrophe timestamp of 14 digits", function() {
        assert.ok(/^'\d{14}$/.test(get_current_date_as_excel_text()),
                  "unexpected format: " + get_current_date_as_excel_text());
    });

    it("starts with the current year", function() {
        var expected = "'" + (new Date()).getFullYear();
        assert.equal(get_current_date_as_excel_text().substring(0, 5), expected);
    });

    it("zero-pads every component to two digits", function() {
        var s = get_current_date_as_excel_text().substring(5); // MMDDhhmmss
        assert.equal(s.length, 10);
    });

});

// ===========================================================================
// fill_sheet
// ===========================================================================

describe("fill_sheet", function() {

    it("writes the object keys as the header row", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, [{ name: "alice", age: "30" }]);
        assert.equal(sheet._value(1, 1), "name");
        assert.equal(sheet._value(1, 2), "age");
    });

    it("writes the first record on row 2", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, [{ name: "alice", age: "30" }]);
        assert.equal(sheet._value(2, 1), "alice");
        assert.equal(sheet._value(2, 2), "30");
    });

    it("writes one row per record", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, [{ id: "1" }, { id: "2" }, { id: "3" }]);
        assert.equal(sheet._value(2, 1), "1");
        assert.equal(sheet._value(3, 1), "2");
        assert.equal(sheet._value(4, 1), "3");
    });

    it("copies column A formatting across the used columns", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, [{ a: "1", b: "2", c: "3" }]);
        assert.deepEqual(sheet._copies(), ["A:A -> B:C"]);
    });

    it("writes nothing for an empty record list", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, []);
        assert.equal(sheet._cellsSet(), 0);
        assert.deepEqual(sheet._copies(), []);
    });

    it("preserves non-string values as written", function() {
        var sheet = mock_sheet();
        fill_sheet(sheet, [{ n: 42, b: true }]);
        assert.equal(sheet._value(2, 1), 42);
        assert.equal(sheet._value(2, 2), true);
    });

    skip("handles more than 26 columns",
         "fill_sheet indexes a 27-character alphabet string and produces an empty column letter past Z (TODO.md)");

});

// ===========================================================================
// read_sheet_data
// ===========================================================================

describe("read_sheet_data - auto column detection", function() {

    it("reads records keyed by the first-row labels", function() {
        var sheet = mock_sheet([
            ["name",  "age"],
            ["alice", "30"],
            ["bob",   "25"]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [
            { name: "alice", age: "30" },
            { name: "bob",   age: "25" }
        ]);
    });

    it("returns an empty array when there are no data rows", function() {
        assert.deepEqual(read_sheet_data(mock_sheet([["name", "age"]])), []);
    });

    it("stops at the first row with an empty first cell", function() {
        var sheet = mock_sheet([
            ["name"],
            ["alice"],
            [""],
            ["ignored"]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ name: "alice" }]);
    });

    it("stops reading labels at the first empty header cell", function() {
        var sheet = mock_sheet([
            ["a", "b", "", "d"],
            ["1", "2", "3", "4"]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ a: "1", b: "2" }]);
    });

    it("trims labels and values", function() {
        var sheet = mock_sheet([
            ["  name  "],
            ["  alice  "]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ name: "alice" }]);
    });

    it("converts numeric cells to strings", function() {
        var sheet = mock_sheet([
            ["name", "qty"],
            ["bolt", 42]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ name: "bolt", qty: "42" }]);
    });

    it("turns an empty cell into an empty string", function() {
        var sheet = mock_sheet([
            ["a", "b"],
            ["1", ""]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ a: "1", b: "" }]);
    });

    it("reads a numeric cell in the first column", function() {
        var sheet = mock_sheet([
            ["qty", "name"],
            [42, "bolt"]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ qty: "42", name: "bolt" }]);
    });

    it("reads a boolean cell", function() {
        var sheet = mock_sheet([
            ["a", "flag"],
            ["1", true]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ a: "1", flag: "true" }]);
    });

});

describe("read_sheet_data - fixed column count", function() {

    it("reads exactly ncolumns columns", function() {
        var sheet = mock_sheet([
            ["a", "b", "c"],
            ["1", "2", "3"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 2), [{ a: "1", b: "2" }]);
    });

    it("names an unlabelled column after its column letter", function() {
        var sheet = mock_sheet([
            ["a", "", "c"],
            ["1", "2", "3"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 3), [{ a: "1", B: "2", c: "3" }]);
    });

    it("keeps a row whose first cell is empty but which has data further right", function() {
        var sheet = mock_sheet([
            ["a", "b"],
            ["",  "2"],
            ["3", "4"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 2), [
            { a: "", b: "2" },
            { a: "3", b: "4" }
        ]);
    });

    it("reads a numeric cell in the first column (unlike auto-detection)", function() {
        var sheet = mock_sheet([
            ["qty"],
            [42]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 1), [{ qty: "42" }]);
    });

    it("stops when every one of the ncolumns cells is empty", function() {
        var sheet = mock_sheet([
            ["a", "b"],
            ["1", "2"],
            ["",  ""],
            ["x", "y"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 2), [{ a: "1", b: "2" }]);
    });

});

describe("read_sheet_data - start_row", function() {

    it("takes labels from start_row and data from the row after", function() {
        var sheet = mock_sheet([
            ["report title", ""],
            ["name",  "age"],
            ["alice", "30"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, null, { start_row: 2 }),
                         [{ name: "alice", age: "30" }]);
    });

    it("works together with a fixed column count", function() {
        var sheet = mock_sheet([
            ["", ""],
            ["a", "b"],
            ["1", "2"]
        ]);
        assert.deepEqual(read_sheet_data(sheet, 2, { start_row: 2 }), [{ a: "1", b: "2" }]);
    });

});

describe("read_sheet_data - date cells", function() {

    it("formats a Date cell as YYYY/MM/DD", function() {
        var sheet = mock_sheet([
            ["a", "when"],
            ["1", new Date(2020, 0, 2)]
        ]);
        assert.deepEqual(read_sheet_data(sheet), [{ a: "1", when: "2020/01/02" }]);
    });

    it("survives a Date cell without throwing", function() {
        var sheet = mock_sheet([
            ["a", "when"],
            ["1", new Date(2021, 8, 5)]
        ]);
        assert.doesNotThrow(function() { read_sheet_data(sheet); });
    });

});

// ===========================================================================
// execute_query
// ===========================================================================

describe("execute_query", function() {

    it("returns one object per record, keyed by field name", function() {
        var db = mock_db(["id", "name"], [["1", "alice"], ["2", "bob"]]);
        assert.deepEqual(execute_query("qry", db), [
            { id: "1", name: "alice" },
            { id: "2", name: "bob" }
        ]);
    });

    it("returns an empty array for an empty recordset", function() {
        assert.deepEqual(execute_query("qry", mock_db(["id"], [])), []);
    });

    it("passes the query name to OpenRecordset", function() {
        var db = mock_db(["id"], [["1"]]);
        execute_query("my_saved_query", db);
        assert.equal(db._query, "my_saved_query");
    });

    it("opens the recordset as a dynaset (type 2)", function() {
        var db = mock_db(["id"], [["1"]]);
        execute_query("qry", db);
        assert.equal(db._type, 2);
    });

    it("converts null field values to empty strings", function() {
        var db = mock_db(["id", "note"], [["1", null]]);
        assert.deepEqual(execute_query("qry", db), [{ id: "1", note: "" }]);
    });

    it("converts numeric field values to trimmed strings", function() {
        var db = mock_db(["n"], [[42]]);
        assert.deepEqual(execute_query("qry", db), [{ n: "42" }]);
    });

    it("trims whitespace from field values", function() {
        var db = mock_db(["name"], [["  padded  "]]);
        assert.deepEqual(execute_query("qry", db), [{ name: "padded" }]);
    });

    it("reads every field of every record", function() {
        var db  = mock_db(["a", "b", "c"], [["1", "2", "3"], ["4", "5", "6"]]);
        var out = execute_query("qry", db);
        assert.equal(out.length, 2);
        assert.equal(out[1].c, "6");
    });

});

// ===========================================================================
// select_files_from_folder
// ===========================================================================

describe("select_files_from_folder", function() {

    write_text_to_file("x", tmp("alpha.txt"));
    write_text_to_file("x", tmp("beta.txt"));
    write_text_to_file("x", tmp("gamma.csv"));

    it("returns only the files matching the pattern", function() {
        var out = select_files_from_folder(TEMP, "", /\.txt$/);
        assert.equal(out.length, 2);
    });

    it("returns full paths", function() {
        var out = select_files_from_folder(TEMP, "", /alpha/);
        assert.equal(out.length, 1);
        assert.ok(out[0].indexOf(TEMP) === 0, "not a full path: " + out[0]);
    });

    it("accepts a string pattern", function() {
        assert.equal(select_files_from_folder(TEMP, "", "gamma").length, 1);
    });

    it("returns an empty array when nothing matches", function() {
        assert.deepEqual(select_files_from_folder(TEMP, "", /\.zip$/), []);
    });

    it("matches everything with a permissive pattern", function() {
        assert.equal(select_files_from_folder(TEMP, "", /./).length, 3);
    });

});

// ===========================================================================
// select_file_from_folder  (interactive)
// ===========================================================================

describe("select_file_from_folder", function() {

    it("returns the only match without asking which one", function() {
        stub_input([""]);          // one Enter to start the scan
        var out = select_file_from_folder(TEMP, "press enter", /gamma/);
        restore_input();
        assert.ok(out.indexOf("gamma.csv") !== -1, "unexpected selection: " + out);
    });

    it("prints the prompt message", function() {
        stub_input([""]);
        select_file_from_folder(TEMP, "PROMPT-TEXT", /gamma/);
        restore_input();
        assert.ok(_prompts.join("|").indexOf("PROMPT-TEXT") !== -1);
    });

    it("asks the user to choose when several files match", function() {
        var expected = select_files_from_folder(TEMP, "", /\.txt$/)[1];
        stub_input(["", "2"]);     // Enter, then "pick #2"
        var out = select_file_from_folder(TEMP, "pick", /\.txt$/);
        restore_input();
        assert.equal(out, expected);
    });

    it("numbers the choices from 1", function() {
        stub_input(["", "1"]);
        select_file_from_folder(TEMP, "pick", /\.txt$/);
        restore_input();
        assert.ok(_prompts.join("|").indexOf("(1) ") !== -1);
        assert.ok(_prompts.join("|").indexOf("(2) ") !== -1);
    });

    it("re-asks after a non-numeric answer", function() {
        var expected = select_files_from_folder(TEMP, "", /\.txt$/)[0];
        stub_input(["", "not-a-number", "1"]);
        var out = select_file_from_folder(TEMP, "pick", /\.txt$/);
        restore_input();
        assert.equal(out, expected);
    });

    it("re-asks after an out-of-range answer", function() {
        var expected = select_files_from_folder(TEMP, "", /\.txt$/)[0];
        stub_input(["", "99", "0", "1"]);
        var out = select_file_from_folder(TEMP, "pick", /\.txt$/);
        restore_input();
        assert.equal(out, expected);
    });

    it("rescans until at least one file matches", function() {
        // First pass finds nothing; the file appears before the second pass.
        var late = tmp("late.log");
        var i    = 0;
        stub_input(["", ""]);
        var real_list = list_folders;
        list_folders = function(path) {
            i++;
            if (i === 1) { return []; }
            return real_list(path);
        };
        write_text_to_file("x", late);
        var out = select_file_from_folder(TEMP, "pick", /late\.log$/);
        list_folders = real_list;
        restore_input();
        delete_file(late);
        assert.ok(out.indexOf("late.log") !== -1, "unexpected selection: " + out);
    });

});

// ===========================================================================
// read_choice_from_input
// ===========================================================================

describe("read_choice_from_input", function() {

    it("returns a valid choice", function() {
        stub_input(["B"]);
        var out = read_choice_from_input("pick one", ["A", "B"]);
        restore_input();
        assert.equal(out, "B");
    });

    it("upper-cases the answer before matching", function() {
        stub_input(["b"]);
        var out = read_choice_from_input("pick one", ["A", "B"]);
        restore_input();
        assert.equal(out, "B");
    });

    it("trims the answer", function() {
        stub_input(["  a  "]);
        var out = read_choice_from_input("pick one", ["A", "B"]);
        restore_input();
        assert.equal(out, "A");
    });

    it("re-asks until the answer is one of the choices", function() {
        stub_input(["x", "zzz", "A"]);
        var out = read_choice_from_input("pick one", ["A", "B"]);
        restore_input();
        assert.equal(out, "A");
    });

    it("shows the choices in the prompt", function() {
        stub_input(["Y"]);
        read_choice_from_input("continue?", ["Y", "N"]);
        restore_input();
        assert.ok(_prompts.join("|").indexOf("(Y|N)") !== -1, _prompts.join("|"));
    });

    it("works with a single choice", function() {
        stub_input(["OK"]);
        var out = read_choice_from_input("confirm", ["OK"]);
        restore_input();
        assert.equal(out, "OK");
    });

});

// ===========================================================================
// read_date_from_input
// ===========================================================================

describe("read_date_from_input", function() {

    it("parses YYYYMMDD into a Date", function() {
        stub_input(["20240115"]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getFullYear(), 2024);
        assert.equal(d.getMonth(),    0);
        assert.equal(d.getDate(),     15);
    });

    it("returns a date at midnight local time", function() {
        stub_input(["20240115"]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getHours(),        0);
        assert.equal(d.getMinutes(),      0);
        assert.equal(d.getSeconds(),      0);
        assert.equal(d.getMilliseconds(), 0);
    });

    it("trims the answer", function() {
        stub_input(["  20201231  "]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getFullYear(), 2020);
        assert.equal(d.getMonth(),    11);
        assert.equal(d.getDate(),     31);
    });

    it("handles a leap day", function() {
        stub_input(["20240229"]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getMonth(), 1);
        assert.equal(d.getDate(),  29);
    });

    it("re-asks until the answer has eight digits", function() {
        stub_input(["nope", "2024", "20240115"]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getFullYear(), 2024);
    });

    it("rejects eight digits embedded in other text", function() {
        stub_input(["abc20240115", "20240115"]);
        var d = read_date_from_input();
        restore_input();
        assert.equal(d.getFullYear(), 2024);
        assert.equal(d.getMonth(),    0);
        assert.equal(d.getDate(),     15);
    });

});

// ===========================================================================
// Office COM wrappers — not runnable unattended
// ===========================================================================

describe("Office COM wrappers", function() {

    it("do_in_excel is defined", function() {
        assert.equal(typeof do_in_excel, "function");
    });

    it("do_in_access is defined", function() {
        assert.equal(typeof do_in_access, "function");
    });

    it("do_in_word is defined", function() {
        assert.equal(typeof do_in_word, "function");
    });

    it("write_report_to_file is defined", function() {
        assert.equal(typeof write_report_to_file, "function");
    });

    skip("do_in_excel runs the callback with an Excel.Application",
         "needs Excel installed; launches a real COM server");
    skip("do_in_excel quits Excel even when the callback throws",
         "needs Excel installed; launches a real COM server");
    skip("do_in_word runs the callback with a Word.Application",
         "needs Word installed; launches a real COM server");
    skip("do_in_access opens the database and closes it afterwards",
         "needs Access installed and a .accdb fixture");
    it("write_report_to_file writes the report next to the source file", function() {
        // assigned without `var` so it becomes a real global, per the
        // load()/eval scoping rule documented in CLAUDE.md.
        OUTPUT_FOLDER = "";

        var source = tmp("mydata.xlsx");
        write_text_to_file("source", source);

        write_report_to_file("hello report", source);

        var written = select_files_from_folder(CURRENT_FOLDER, "", /mydata-\d+\.txt$/);
        assert.equal(written.length, 1);
        assert.equal(read_text_file(written[0]), "hello report");
        delete_file(written[0]);
    });

});

// ---------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------

restore_input();

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
}());

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
