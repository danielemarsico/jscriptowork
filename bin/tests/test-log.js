// test-log.js - Tests for libs/log.js
//
// Level filtering and formatting are asserted through a captured sink, so no
// console inspection is needed. The file tee writes into the system temp
// folder.
//
// This suite loads log.js, which REPLACES the global log() that minitest
// prints through. Every test restores the default level and sink afterwards,
// so a failed assertion cannot silence the rest of the run.
//
// Run via:  cscript.exe launcher.js test-log.js

load("core");
load("polyfills");
load("system");
load("log");
load("minitest");

// ---------------------------------------------------------------------------
// Temp folder
// ---------------------------------------------------------------------------

var TEMP = (function() {
    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var base = fso.GetSpecialFolder(2).Path; // 2 = TemporaryFolder
    return base + "\\jscriptowork_log_test";
}());

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
    fso.CreateFolder(TEMP);
}());

function tmp(name) { return TEMP + "\\" + name; }

// ---------------------------------------------------------------------------
// Capture helper: runs fn with the sink captured, always restoring afterwards
// ---------------------------------------------------------------------------

function captured(fn) {
    var lines = [];
    set_log_sink(function(text) { lines.push(text); });
    try {
        fn();
    } finally {
        set_log_sink(null);
        set_log_level("info");
        set_log_timestamps(false);
        log_to_file(null);
    }
    return lines;
}

// ===========================================================================
// Surface
// ===========================================================================

describe("log - surface", function() {

    it("exposes the level functions", function() {
        assert.equal(typeof log_debug, "function");
        assert.equal(typeof log_info,  "function");
        assert.equal(typeof log_warn,  "function");
        assert.equal(typeof log_error, "function");
    });

    it("exposes the configuration functions", function() {
        assert.equal(typeof set_log_level,      "function");
        assert.equal(typeof get_log_level,      "function");
        assert.equal(typeof get_log_level_name, "function");
        assert.equal(typeof set_log_timestamps, "function");
        assert.equal(typeof set_log_sink,       "function");
        assert.equal(typeof log_to_file,        "function");
        assert.equal(typeof get_log_file,       "function");
    });

    it("still exposes log()", function() {
        assert.equal(typeof log, "function");
    });

    it("defines the level constants in ascending order", function() {
        assert.ok(LOG_LEVELS.debug < LOG_LEVELS.info);
        assert.ok(LOG_LEVELS.info  < LOG_LEVELS.warn);
        assert.ok(LOG_LEVELS.warn  < LOG_LEVELS.error);
        assert.ok(LOG_LEVELS.error < LOG_LEVELS.none);
    });

    it("defaults to the info level", function() {
        assert.equal(get_log_level(),      LOG_LEVELS.info);
        assert.equal(get_log_level_name(), "info");
    });

});

// ===========================================================================
// Formatting
// ===========================================================================

describe("log - formatting", function() {

    it("log() writes the message with no prefix, as before", function() {
        assert.deepEqual(captured(function() { log("plain"); }), ["plain"]);
    });

    it("prefixes each level", function() {
        var out = captured(function() {
            set_log_level("debug");
            log_debug("d");
            log_info("i");
            log_warn("w");
            log_error("e");
        });
        assert.deepEqual(out, ["[DEBUG] d", "[INFO] i", "[WARN] w", "[ERROR] e"]);
    });

    it("joins multiple arguments with a space", function() {
        assert.deepEqual(captured(function() { log_warn("a", "b", 1); }), ["[WARN] a b 1"]);
    });

    it("JSON-encodes objects and arrays", function() {
        assert.deepEqual(captured(function() { log({ a: 1 }); }),   ['{"a":1}']);
        assert.deepEqual(captured(function() { log([1, 2]); }),     ['[1,2]']);
    });

    it("renders null and undefined without throwing", function() {
        assert.deepEqual(captured(function() { log(null, undefined); }), ["null undefined"]);
    });

    it("writes an empty line for no arguments", function() {
        assert.deepEqual(captured(function() { log(); }), [""]);
    });

    it("survives a value JSON cannot encode", function() {
        var circular = {};
        circular.self = circular;
        assert.doesNotThrow(function() {
            captured(function() { log(circular); });
        });
    });

});

// ===========================================================================
// Level filtering
// ===========================================================================

describe("log - level filtering", function() {

    it("drops debug at the default info level", function() {
        assert.deepEqual(captured(function() { log_debug("hidden"); }), []);
    });

    it("keeps info, warn and error at the default level", function() {
        var out = captured(function() {
            log_info("i");
            log_warn("w");
            log_error("e");
        });
        assert.equal(out.length, 3);
    });

    it("shows debug once the level is lowered", function() {
        var out = captured(function() {
            set_log_level("debug");
            log_debug("now visible");
        });
        assert.deepEqual(out, ["[DEBUG] now visible"]);
    });

    it("drops info and debug at the warn level", function() {
        var out = captured(function() {
            set_log_level("warn");
            log_debug("no");
            log_info("no");
            log_warn("yes");
            log_error("yes");
        });
        assert.deepEqual(out, ["[WARN] yes", "[ERROR] yes"]);
    });

    it("keeps only error at the error level", function() {
        var out = captured(function() {
            set_log_level("error");
            log_warn("no");
            log_error("yes");
        });
        assert.deepEqual(out, ["[ERROR] yes"]);
    });

    it("drops everything at the none level", function() {
        var out = captured(function() {
            set_log_level("none");
            log("no");
            log_error("no");
        });
        assert.deepEqual(out, []);
    });

    it("filters plain log() too, since it logs at info", function() {
        var out = captured(function() {
            set_log_level("warn");
            log("suppressed");
        });
        assert.deepEqual(out, []);
    });

    it("accepts a level name in any case", function() {
        var out = captured(function() {
            set_log_level("WARN");
            log_info("no");
            log_warn("yes");
        });
        assert.deepEqual(out, ["[WARN] yes"]);
    });

    it("accepts a numeric level", function() {
        var out = captured(function() {
            set_log_level(LOG_LEVELS.error);
            log_warn("no");
            log_error("yes");
        });
        assert.deepEqual(out, ["[ERROR] yes"]);
    });

    it("throws on an unknown level name", function() {
        assert.throws(function() { set_log_level("chatty"); });
        assert.throws(function() { set_log_level(null); });
    });

    it("reports the level back", function() {
        captured(function() {
            set_log_level("error");
            assert.equal(get_log_level(),      LOG_LEVELS.error);
            assert.equal(get_log_level_name(), "error");
        });
    });

    it("restores the default level after each capture", function() {
        assert.equal(get_log_level(), LOG_LEVELS.info);
    });

});

// ===========================================================================
// Timestamps
// ===========================================================================

describe("log - timestamps", function() {

    it("adds no timestamp by default", function() {
        assert.deepEqual(captured(function() { log("bare"); }), ["bare"]);
    });

    it("prefixes a timestamp when enabled", function() {
        var out = captured(function() {
            set_log_timestamps(true);
            log("stamped");
        });
        assert.equal(out.length, 1);
        assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} stamped$/.test(out[0]),
                  "unexpected line: " + out[0]);
    });

    it("puts the timestamp before the level prefix", function() {
        var out = captured(function() {
            set_log_timestamps(true);
            log_error("boom");
        });
        assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[ERROR\] boom$/.test(out[0]),
                  "unexpected line: " + out[0]);
    });

    it("can be turned back off", function() {
        var out = captured(function() {
            set_log_timestamps(true);
            set_log_timestamps(false);
            log("bare again");
        });
        assert.deepEqual(out, ["bare again"]);
    });

});

// ===========================================================================
// Custom sink
// ===========================================================================

describe("log - sink", function() {

    it("routes every line through the sink", function() {
        assert.deepEqual(captured(function() { log("via sink"); }), ["via sink"]);
    });

    it("restores the default sink when passed null", function() {
        set_log_sink(function() {});
        set_log_sink(null);
        assert.doesNotThrow(function() { log(""); });
    });

});

// ===========================================================================
// File tee
// ===========================================================================

describe("log - file tee", function() {

    it("is off by default", function() {
        assert.equal(get_log_file(), null);
    });

    it("reports the path it is teeing to", function() {
        var p = tmp("reported.log");
        captured(function() {
            log_to_file(p);
            assert.equal(get_log_file(), p);
        });
    });

    it("creates the file and writes the line", function() {
        var p = tmp("basic.log");
        captured(function() {
            log_to_file(p);
            log("to file");
        });
        assert.ok(file_exists(p));
        assert.ok(read_text_file(p).indexOf("to file") !== -1);
    });

    it("appends rather than overwriting", function() {
        var p = tmp("append.log");
        captured(function() {
            log_to_file(p);
            log("first");
            log("second");
        });
        var text = read_text_file(p);
        assert.ok(text.indexOf("first")  !== -1, text);
        assert.ok(text.indexOf("second") !== -1, text);
    });

    it("writes one line per call", function() {
        var p = tmp("lines.log");
        captured(function() {
            log_to_file(p);
            log("a");
            log("b");
            log("c");
        });
        // WriteLine terminates each entry, so a trailing empty split part is
        // expected; count only the non-empty ones.
        var parts = read_text_file(p).split("\r\n");
        var count = 0;
        for (var i = 0; i < parts.length; i++) { if (parts[i].length > 0) { count++; } }
        assert.equal(count, 3);
    });

    it("writes the same text the sink gets, prefixes included", function() {
        var p = tmp("prefix.log");
        var out = captured(function() {
            log_to_file(p);
            log_error("with prefix");
        });
        assert.deepEqual(out, ["[ERROR] with prefix"]);
        assert.ok(read_text_file(p).indexOf("[ERROR] with prefix") !== -1);
    });

    it("does not write a line the level filter dropped", function() {
        var p = tmp("filtered.log");
        captured(function() {
            log_to_file(p);
            set_log_level("error");
            log_info("dropped");
        });
        assert.notOk(file_exists(p) && read_text_file(p).indexOf("dropped") !== -1);
    });

    it("stops teeing when passed null", function() {
        var p = tmp("stopped.log");
        captured(function() {
            log_to_file(p);
            log("kept");
            log_to_file(null);
            log("not kept");
        });
        var text = read_text_file(p);
        assert.ok(text.indexOf("kept")     !== -1);
        assert.ok(text.indexOf("not kept") === -1);
    });

    it("keeps going when the path cannot be written", function() {
        // A path under a folder that does not exist: the tee must fail quietly
        // rather than take down the caller.
        var out = captured(function() {
            log_to_file(TEMP + "\\no_such_folder\\out.log");
            log("still reaches the sink");
        });
        assert.deepEqual(out, ["still reaches the sink"]);
    });

});

// ---------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------

set_log_sink(null);
set_log_level("info");
set_log_timestamps(false);
log_to_file(null);

(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(TEMP)) { fso.DeleteFolder(TEMP, true); }
}());

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
