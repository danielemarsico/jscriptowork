// test-filesystem.js - Tests for file-system helpers in system.js
// All files are written to the system temp folder and cleaned up at the end.
// Run via:  cscript.exe launcher.js test-filesystem.js

load("core");
load("polyfills");
load("system");
load("minitest");

// ---------------------------------------------------------------------------
// Temp folder setup — isolated directory under %TEMP% for all test artefacts
// ---------------------------------------------------------------------------

var TEMP = (function() {
    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var base = fso.GetSpecialFolder(2).Path; // 2 = TemporaryFolder
    return base + "\\jscriptowork_fs_test";
}());

// Clean up any leftover from a previous run, then create fresh
(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
    fso.CreateFolder(TEMP);
}());

function tmp(name) { return TEMP + "\\" + name; }

// ---------------------------------------------------------------------------
// Text file — write_text_to_file / read_text_file
// ---------------------------------------------------------------------------

describe("write_text_to_file", function() {

    it("creates a file that did not exist before", function() {
        var p = tmp("create.txt");
        write_text_to_file("hello", p);
        assert.ok(file_exists(p));
    });

    it("round-trips a simple string", function() {
        var p = tmp("simple.txt");
        write_text_to_file("hello world", p);
        assert.equal(read_text_file(p), "hello world");
    });

    it("round-trips multiline text", function() {
        var p   = tmp("multiline.txt");
        var src = "line one\nline two\nline three";
        write_text_to_file(src, p);
        assert.equal(read_text_file(p), src);
    });

    it("round-trips text with tabs and spaces", function() {
        var p   = tmp("whitespace.txt");
        var src = "col1\tcol2\t col3";
        write_text_to_file(src, p);
        assert.equal(read_text_file(p), src);
    });

    it("overwrites an existing file", function() {
        var p = tmp("overwrite.txt");
        write_text_to_file("first", p);
        write_text_to_file("second", p);
        assert.equal(read_text_file(p), "second");
    });

    it("can write an empty string", function() {
        var p = tmp("empty.txt");
        write_text_to_file("", p);
        assert.equal(read_text_file(p), "");
    });

    it("round-trips numeric content", function() {
        var p = tmp("numbers.txt");
        write_text_to_file("3.14159", p);
        assert.equal(read_text_file(p), "3.14159");
    });

    it("round-trips JSON text", function() {
        var p   = tmp("data.json");
        var src = JSON.stringify({ key: "value", n: 42 });
        write_text_to_file(src, p);
        var parsed = JSON.parse(read_text_file(p));
        assert.equal(parsed.key, "value");
        assert.equal(parsed.n,   42);
    });

});

// ---------------------------------------------------------------------------
// read_text_file
// ---------------------------------------------------------------------------

describe("read_text_file", function() {

    it("throws when file does not exist", function() {
        assert.throws(function() {
            read_text_file(tmp("nonexistent_" + Date.now() + ".txt"));
        });
    });

    it("returns empty string for an empty file", function() {
        var p = tmp("readempty.txt");
        write_text_to_file("", p);
        assert.equal(read_text_file(p), "");
    });

});

// ---------------------------------------------------------------------------
// file_exists
// ---------------------------------------------------------------------------

describe("file_exists", function() {

    it("returns true for a file that exists", function() {
        var p = tmp("exists.txt");
        write_text_to_file("x", p);
        assert.equal(file_exists(p), true);
    });

    it("returns false for a file that does not exist", function() {
        assert.equal(file_exists(tmp("ghost_" + Date.now() + ".txt")), false);
    });

    it("returns false for a folder path", function() {
        assert.equal(file_exists(TEMP), false);
    });

});

// ---------------------------------------------------------------------------
// delete_file
// ---------------------------------------------------------------------------

describe("delete_file", function() {

    it("removes the file from disk", function() {
        var p = tmp("todelete.txt");
        write_text_to_file("bye", p);
        delete_file(p);
        assert.equal(file_exists(p), false);
    });

    it("throws when deleting a non-existent file", function() {
        assert.throws(function() {
            delete_file(tmp("nodelet_" + Date.now() + ".txt"));
        });
    });

});

// ---------------------------------------------------------------------------
// folder_exists / create_folder
// ---------------------------------------------------------------------------

describe("folder_exists", function() {

    it("returns true for the temp test folder", function() {
        assert.equal(folder_exists(TEMP), true);
    });

    it("returns false for a non-existent folder", function() {
        assert.equal(folder_exists(tmp("nofolder_" + Date.now())), false);
    });

    it("returns false for a file path", function() {
        var p = tmp("notafolder.txt");
        write_text_to_file("x", p);
        assert.equal(folder_exists(p), false);
    });

});

describe("create_folder", function() {

    it("creates a new folder", function() {
        var p = tmp("newfolder_" + Date.now());
        create_folder(p);
        assert.equal(folder_exists(p), true);
    });

    it("does not throw when folder already exists", function() {
        assert.doesNotThrow(function() {
            create_folder(TEMP); // already exists
        });
    });

});

// ---------------------------------------------------------------------------
// list_folders (lists files in a directory)
// ---------------------------------------------------------------------------

describe("list_folders", function() {

    it("returns an array", function() {
        assert.ok(Array.isArray(list_folders(TEMP)));
    });

    it("includes a file that was just written", function() {
        var p    = tmp("listed.txt");
        write_text_to_file("x", p);
        var list = list_folders(TEMP);
        var found = list.some(function(entry) {
            return entry.indexOf("listed.txt") !== -1;
        });
        assert.ok(found);
    });

    it("does not include a file that was deleted", function() {
        var p = tmp("unlisted.txt");
        write_text_to_file("x", p);
        delete_file(p);
        var list = list_folders(TEMP);
        var found = list.some(function(entry) {
            return entry.indexOf("unlisted.txt") !== -1;
        });
        assert.notOk(found);
    });

});

// ---------------------------------------------------------------------------
// write_binary_file / read_binary_file
// ---------------------------------------------------------------------------

describe("write_binary_file", function() {

    it("creates a file that did not exist", function() {
        var p = tmp("binary.bin");
        write_binary_file(p, [1, 2, 3]);
        assert.ok(file_exists(p));
    });

    it("can write an empty byte array", function() {
        var p = tmp("empty.bin");
        assert.doesNotThrow(function() { write_binary_file(p, []); });
        assert.ok(file_exists(p));
    });

});

describe("read_binary_file", function() {

    it("returns an array", function() {
        var p = tmp("readbin.bin");
        write_binary_file(p, [10, 20, 30]);
        assert.ok(Array.isArray(read_binary_file(p)));
    });

    it("returns empty array for an empty binary file", function() {
        var p = tmp("emptyread.bin");
        write_binary_file(p, []);
        assert.deepEqual(read_binary_file(p), []);
    });

    it("round-trips ASCII byte values (0-127)", function() {
        var bytes = [];
        for (var i = 1; i <= 127; i++) { bytes.push(i); } // skip 0 (null)
        var p = tmp("ascii.bin");
        write_binary_file(p, bytes);
        assert.deepEqual(read_binary_file(p), bytes);
    });

    it("round-trips high byte values (160-255)", function() {
        // 128-159 overlap with Windows-1252 control chars; use 160-255 which
        // are identical in both iso-8859-1 and Windows-1252
        var bytes = [];
        for (var i = 160; i <= 255; i++) { bytes.push(i); }
        var p = tmp("high.bin");
        write_binary_file(p, bytes);
        assert.deepEqual(read_binary_file(p), bytes);
    });

    it("round-trips a short known sequence", function() {
        var src = [72, 101, 108, 108, 111]; // ASCII "Hello"
        var p   = tmp("hello.bin");
        write_binary_file(p, src);
        assert.deepEqual(read_binary_file(p), src);
    });

    it("length of read matches bytes written", function() {
        var src = [1, 2, 3, 4, 5, 6, 7, 8];
        var p   = tmp("len.bin");
        write_binary_file(p, src);
        assert.equal(read_binary_file(p).length, src.length);
    });

    it("each byte value is preserved individually", function() {
        var src = [0x41, 0x42, 0xFF, 0xA0, 0x01];
        var p   = tmp("vals.bin");
        write_binary_file(p, src);
        var got = read_binary_file(p);
        for (var i = 0; i < src.length; i++) {
            assert.equal(got[i], src[i]);
        }
    });

    it("overwrites a binary file on second write", function() {
        var p = tmp("overwrite.bin");
        write_binary_file(p, [1, 2, 3]);
        write_binary_file(p, [9, 8]);
        assert.deepEqual(read_binary_file(p), [9, 8]);
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

_test.summary();
