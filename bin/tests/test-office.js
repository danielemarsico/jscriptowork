// test-office.js - Tests for the real Office COM wrappers in libs/helpers.js:
//                  do_in_excel, do_in_access, do_in_word.
//
// **Opt-in.** These launch actual Office applications, so they are off unless
// you ask for them:
//
//     set JSW_TEST_OFFICE=1
//     cscript.exe bin\launcher.js bin\tests\test-office.js
//
// Without that variable every test reports as skipped and nothing is launched,
// which is why `run-tests.bat` and CI can run this file unconditionally.
//
// With it, each application is probed once, and only the ones actually
// installed run - Excel, Word and Access are checked independently, so a
// machine with Excel but no Access still gets useful coverage.
//
// What is covered, per application:
//   * the wrapper hands the callback a live COM object
//   * work done inside the callback really lands on disk
//   * the wrapper quits the application afterwards, so nothing is left running
//   * a callback that throws does not take the wrapper down, and the
//     application is still quit
//
// That last one documents behaviour rather than endorsing it: all three
// wrappers swallow the exception and log it, so a caller cannot tell a failed
// run from a successful one. Changing that is an API break; the tests pin what
// the code does today.
//
// Run via:  cscript.exe launcher.js test-office.js

load("core");
load("polyfills");
load("system");
load("win");
load("helpers");
load("minitest");

// ---------------------------------------------------------------------------
// Opt-in and availability probing
// ---------------------------------------------------------------------------

OFFICE_OPT_IN = (function() {
    var value = new ActiveXObject("WScript.Shell").Environment("Process")("JSW_TEST_OFFICE");
    return value === "1" || String(value).toLowerCase() === "true";
}());

var OPT_OUT_REASON = "opt-in: set JSW_TEST_OFFICE=1 on a machine with Office";

// Creates the COM server once to see whether it exists, then quits it again.
// Only ever called when opted in.
function office_available(progId) {
    var app = null;
    try {
        app = new ActiveXObject(progId);
    } catch (e) {
        return false;
    }
    try { app.Quit(); } catch (e2) {}
    return true;
}

var HAS_EXCEL  = OFFICE_OPT_IN && office_available("Excel.Application");
var HAS_WORD   = OFFICE_OPT_IN && office_available("Word.Application");
var HAS_ACCESS = OFFICE_OPT_IN && office_available("Access.Application");

if (!OFFICE_OPT_IN) {
    log("office tests skipped - " + OPT_OUT_REASON);
} else {
    log("office tests: Excel " + (HAS_EXCEL ? "yes" : "no") +
        ", Word " + (HAS_WORD ? "yes" : "no") +
        ", Access " + (HAS_ACCESS ? "yes" : "no"));
}

function missing(app) { return app + " is not installed on this machine"; }

// ---------------------------------------------------------------------------
// Scratch space
// ---------------------------------------------------------------------------

var TEMP = (function() {
    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var base = fso.GetSpecialFolder(2).Path;   // 2 = TemporaryFolder
    return base + "\\jscriptowork_office_test";
}());

if (OFFICE_OPT_IN) {
    (function() {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
        fso.CreateFolder(TEMP);
    }());
}

function tmp(name) { return TEMP + "\\" + name; }

function remove_if_present(path) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FileExists(path)) {
        try { fso.DeleteFile(path, true); } catch (e) {}
    }
}

// How many copies of a process are running right now. The machine may already
// have Excel open for its own reasons, so the tests compare before and after
// rather than expecting zero.
function running(name) {
    try { return list_processes(name).length; } catch (e) { return -1; }
}

// ---------------------------------------------------------------------------
// Excel
// ---------------------------------------------------------------------------

describe("do_in_excel", function() {

    if (!HAS_EXCEL) {
        var reason = OFFICE_OPT_IN ? missing("Excel") : OPT_OUT_REASON;
        skip("hands the callback a live Excel.Application", reason);
        skip("writes a workbook that survives to disk", reason);
        skip("quits Excel afterwards", reason);
        skip("survives a callback that throws, and still quits Excel", reason);
        return;
    }

    it("hands the callback a live Excel.Application", function() {
        var version = null;
        do_in_excel(function(excel) {
            version = String(excel.Version);
        });
        assert.ok(version !== null, "the callback never ran");
        assert.ok(version.length > 0, "Excel.Version was empty");
    });

    it("writes a workbook that survives to disk", function() {
        var path = tmp("jsw_excel.xlsx");
        remove_if_present(path);

        do_in_excel(function(excel) {
            var book  = excel.Workbooks.Add();
            var sheet = book.Worksheets(1);
            sheet.Cells(1, 1).Value = "jscriptowork";
            sheet.Cells(2, 1).Value = 42;
            book.SaveAs(path, 51);          // 51 = xlOpenXMLWorkbook (.xlsx)
            book.Close(false);
        });

        assert.ok(file_exists(path), "SaveAs produced no file at " + path);

        var readBack = [];
        do_in_excel(function(excel) {
            var book = excel.Workbooks.Open(path);
            readBack.push(String(book.Worksheets(1).Cells(1, 1).Value));
            readBack.push(Number(book.Worksheets(1).Cells(2, 1).Value));
            book.Close(false);
        });

        assert.equal(readBack[0], "jscriptowork");
        assert.equal(readBack[1], 42);
        remove_if_present(path);
    });

    it("quits Excel afterwards", function() {
        var before = running("excel");
        do_in_excel(function(excel) { var v = excel.Version; });
        sleep(1500);                        // Quit() returns before the process is gone
        var after = running("excel");
        assert.ok(after <= before,
            "Excel processes went from " + before + " to " + after +
            " - do_in_excel left one running");
    });

    it("survives a callback that throws, and still quits Excel", function() {
        var before = running("excel");

        // Documented (if unfortunate) behaviour: the wrapper catches, logs and
        // returns normally, so the caller sees success either way.
        assert.doesNotThrow(function() {
            do_in_excel(function(excel) { throw new Error("deliberate"); });
        });

        sleep(1500);
        var after = running("excel");
        assert.ok(after <= before,
            "a throwing callback left Excel running (" + before + " -> " + after + ")");
    });

});

// ---------------------------------------------------------------------------
// Word
// ---------------------------------------------------------------------------

describe("do_in_word", function() {

    if (!HAS_WORD) {
        var reason = OFFICE_OPT_IN ? missing("Word") : OPT_OUT_REASON;
        skip("hands the callback a live Word.Application", reason);
        skip("writes a document that survives to disk", reason);
        skip("quits Word afterwards", reason);
        skip("survives a callback that throws, and still quits Word", reason);
        return;
    }

    it("hands the callback a live Word.Application", function() {
        var version = null;
        do_in_word(function(word) { version = String(word.Version); });
        assert.ok(version !== null, "the callback never ran");
        assert.ok(version.length > 0, "Word.Version was empty");
    });

    it("writes a document that survives to disk", function() {
        var path = tmp("jsw_word.docx");
        remove_if_present(path);

        do_in_word(function(word) {
            var doc = word.Documents.Add();
            doc.Content.Text = "jscriptowork wrote this";
            doc.SaveAs2(path, 16);          // 16 = wdFormatDocumentDefault (.docx)
            doc.Close(false);
        });

        assert.ok(file_exists(path), "SaveAs2 produced no file at " + path);

        var text = null;
        do_in_word(function(word) {
            var doc = word.Documents.Open(path);
            text = String(doc.Content.Text);
            doc.Close(false);
        });

        assert.ok(text !== null, "the document never opened");
        assert.ok(text.indexOf("jscriptowork wrote this") !== -1,
            "the document did not contain what was written: " + text);
        remove_if_present(path);
    });

    it("quits Word afterwards", function() {
        var before = running("winword");
        do_in_word(function(word) { var v = word.Version; });
        sleep(1500);
        var after = running("winword");
        assert.ok(after <= before,
            "Word processes went from " + before + " to " + after +
            " - do_in_word left one running");
    });

    it("survives a callback that throws, and still quits Word", function() {
        var before = running("winword");

        assert.doesNotThrow(function() {
            do_in_word(function(word) { throw new Error("deliberate"); });
        });

        sleep(1500);
        var after = running("winword");
        assert.ok(after <= before,
            "a throwing callback left Word running (" + before + " -> " + after + ")");
    });

});

// ---------------------------------------------------------------------------
// Access
//
// do_in_access opens `CURRENT_FOLDER + "/" + database_filename`, so the file it
// is pointed at must live next to the launcher - an absolute path cannot be
// passed. The fixture is therefore created in CURRENT_FOLDER and deleted again,
// and the limitation is recorded in TODO.md.
// ---------------------------------------------------------------------------

var ACCESS_DB_NAME = "jsw_office_test.accdb";
var ACCESS_DB_PATH = CURRENT_FOLDER + ACCESS_DB_NAME;

// Builds an .accdb with one table, using a throwaway Access instance.
function make_access_fixture() {
    remove_if_present(ACCESS_DB_PATH);
    var access = new ActiveXObject("Access.Application");
    try {
        access.NewCurrentDatabase(ACCESS_DB_PATH);
        var db = access.CurrentDb();
        db.Execute("CREATE TABLE widgets (id INTEGER, name TEXT(50))");
        db.Execute("INSERT INTO widgets (id, name) VALUES (1, 'sprocket')");
        db.Execute("INSERT INTO widgets (id, name) VALUES (2, 'flange')");
        access.CloseCurrentDatabase();
    } finally {
        try { access.Quit(); } catch (e) {}
    }
}

describe("do_in_access", function() {

    if (!HAS_ACCESS) {
        var reason = OFFICE_OPT_IN ? missing("Access") : OPT_OUT_REASON;
        skip("hands the callback the opened database", reason);
        skip("reads rows written by a previous session", reason);
        skip("quits Access afterwards", reason);
        skip("survives a callback that throws, and still quits Access", reason);
        skip("opens a database given as an absolute path", reason);
        return;
    }

    make_access_fixture();

    it("hands the callback the opened database", function() {
        var name = null;
        do_in_access(function(db) { name = String(db.Name); }, ACCESS_DB_NAME);
        assert.ok(name !== null, "the callback never ran");
        assert.ok(name.length > 0, "CurrentDb().Name was empty");
    });

    it("reads rows written by a previous session", function() {
        var rows = [];
        do_in_access(function(db) {
            var rs = db.OpenRecordset("SELECT id, name FROM widgets ORDER BY id");
            while (!rs.EOF) {
                rows.push(String(rs.Fields("name").Value));
                rs.MoveNext();
            }
            rs.Close();
        }, ACCESS_DB_NAME);

        assert.deepEqual(rows, ["sprocket", "flange"]);
    });

    it("quits Access afterwards", function() {
        var before = running("msaccess");
        do_in_access(function(db) { var n = db.Name; }, ACCESS_DB_NAME);
        sleep(1500);
        var after = running("msaccess");
        assert.ok(after <= before,
            "Access processes went from " + before + " to " + after +
            " - do_in_access left one running");
    });

    it("survives a callback that throws, and still quits Access", function() {
        var before = running("msaccess");

        assert.doesNotThrow(function() {
            do_in_access(function(db) { throw new Error("deliberate"); }, ACCESS_DB_NAME);
        });

        sleep(1500);
        var after = running("msaccess");
        assert.ok(after <= before,
            "a throwing callback left Access running (" + before + " -> " + after + ")");
    });

    // Known limitation, tracked in TODO.md: do_in_access always prefixes
    // CURRENT_FOLDER, so a caller cannot point it at a database anywhere else.
    skip("opens a database given as an absolute path",
         "do_in_access prefixes CURRENT_FOLDER to its argument - see TODO.md");

    remove_if_present(ACCESS_DB_PATH);

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
