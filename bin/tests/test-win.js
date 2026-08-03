// test-win.js - Tests for libs/win.js
//
// Everything here runs against the real machine, but nothing here needs admin
// rights or an interactive desktop:
//   - registry tests confine themselves to HKCU\Software\jscriptowork_test and
//     delete the key again at the end;
//   - process tests run cmd.exe one-liners and look for cscript.exe, which is
//     necessarily running (it is running this suite);
//   - kill_process is exercised against a process this suite starts itself.
//
// Run via:  cscript.exe launcher.js test-win.js

load("core");
load("polyfills");
load("system");
load("win");
load("minitest");

var REG_BASE = "HKCU\\Software\\jscriptowork_test\\";

// Clean up anything a previous interrupted run left behind.
(function() {
    try { reg_delete(REG_BASE + "value"); } catch (e) {}
    try { reg_delete(REG_BASE);           } catch (e) {}
}());

// ===========================================================================
// Registry
// ===========================================================================

describe("registry - read and write", function() {

    it("writes and reads a string value", function() {
        reg_write(REG_BASE + "str", "hello");
        assert.equal(reg_read(REG_BASE + "str"), "hello");
    });

    it("overwrites an existing value", function() {
        reg_write(REG_BASE + "str", "first");
        reg_write(REG_BASE + "str", "second");
        assert.equal(reg_read(REG_BASE + "str"), "second");
    });

    it("writes and reads a DWORD", function() {
        reg_write(REG_BASE + "num", 42, "REG_DWORD");
        assert.equal(reg_read(REG_BASE + "num"), 42);
    });

    it("defaults a number to REG_DWORD", function() {
        reg_write(REG_BASE + "auto_num", 7);
        assert.equal(reg_read(REG_BASE + "auto_num"), 7);
    });

    it("writes and reads an expandable string", function() {
        reg_write(REG_BASE + "expand", "%TEMP%\\x", "REG_EXPAND_SZ");
        // RegRead expands REG_EXPAND_SZ on the way out.
        assert.ok(String(reg_read(REG_BASE + "expand")).indexOf("%TEMP%") === -1);
    });

    it("writes and reads the key's default value", function() {
        reg_write(REG_BASE, "default value");
        assert.equal(reg_read(REG_BASE), "default value");
    });

    it("round-trips an empty string", function() {
        reg_write(REG_BASE + "empty", "");
        assert.equal(reg_read(REG_BASE + "empty"), "");
    });

});

describe("registry - missing values", function() {

    it("reg_read throws for a value that does not exist", function() {
        assert.throws(function() { reg_read(REG_BASE + "no_such_value"); });
    });

    it("reg_read returns the fallback instead of throwing", function() {
        assert.equal(reg_read(REG_BASE + "no_such_value", "fallback"), "fallback");
    });

    it("reg_read treats an explicit undefined as a fallback", function() {
        assert.equal(reg_read(REG_BASE + "no_such_value", undefined), undefined);
    });

    it("reg_read returns a falsy fallback rather than throwing", function() {
        assert.equal(reg_read(REG_BASE + "no_such_value", 0), 0);
    });

    it("reg_exists is false for a missing value", function() {
        assert.equal(reg_exists(REG_BASE + "no_such_value"), false);
    });

    it("reg_exists is true for a value that was written", function() {
        reg_write(REG_BASE + "present", "x");
        assert.equal(reg_exists(REG_BASE + "present"), true);
    });

    it("reg_exists is false for a nonsense root", function() {
        assert.equal(reg_exists("HKCU\\Software\\jscriptowork_test_absent\\x"), false);
    });

});

describe("registry - delete", function() {

    it("deletes a value", function() {
        reg_write(REG_BASE + "doomed", "x");
        reg_delete(REG_BASE + "doomed");
        assert.equal(reg_exists(REG_BASE + "doomed"), false);
    });

    it("throws when deleting something that is not there", function() {
        assert.throws(function() { reg_delete(REG_BASE + "never_existed"); });
    });

});

// ===========================================================================
// Environment
// ===========================================================================

describe("environment", function() {

    it("reads a variable that always exists", function() {
        assert.ok(env("SystemRoot").length > 0, "SystemRoot came back empty");
    });

    it("returns an empty string for a variable that is not set", function() {
        assert.equal(env("JSW_DEFINITELY_NOT_SET_12345"), "");
    });

    it("expands a variable inside a string", function() {
        var expanded = expand_env("%SystemRoot%\\notepad.exe");
        assert.equal(expanded.indexOf("%SystemRoot%"), -1, expanded);
        assert.ok(expanded.length > "\\notepad.exe".length);
    });

    it("leaves text with no variables alone", function() {
        assert.equal(expand_env("plain text"), "plain text");
    });

    it("agrees with env() for the same variable", function() {
        assert.equal(expand_env("%SystemRoot%"), env("SystemRoot"));
    });

});

// ===========================================================================
// run_command / exec_command
// ===========================================================================

describe("run_command", function() {

    it("returns the exit code of a command", function() {
        assert.equal(run_command('cmd.exe /c exit 0'), 0);
    });

    it("reports a non-zero exit code", function() {
        assert.equal(run_command('cmd.exe /c exit 3'), 3);
    });

    it("returns immediately with wait:false", function() {
        var t0 = (new Date()).getTime();
        run_command('cmd.exe /c ping -n 3 127.0.0.1', { wait: false });
        assert.ok((new Date()).getTime() - t0 < 1500, "wait:false blocked anyway");
    });

});

describe("exec_command", function() {

    it("captures stdout", function() {
        var r = exec_command('echo hello');
        assert.ok(r.stdout.indexOf("hello") !== -1, "stdout was: " + r.stdout);
    });

    it("reports the exit code", function() {
        assert.equal(exec_command('exit 0').exit_code, 0);
        assert.equal(exec_command('exit 5').exit_code, 5);
    });

    it("captures stderr separately from stdout", function() {
        var r = exec_command('echo to-out & echo to-err 1>&2');
        assert.ok(r.stdout.indexOf("to-out") !== -1, "stdout was: " + r.stdout);
        assert.ok(r.stderr.indexOf("to-err") !== -1, "stderr was: " + r.stderr);
        assert.equal(r.stdout.indexOf("to-err"), -1, "stderr leaked into stdout");
    });

    it("folds stderr into stdout with merge", function() {
        var r = exec_command('echo to-out & echo to-err 1>&2', { merge: true });
        assert.ok(r.stdout.indexOf("to-out") !== -1, r.stdout);
        assert.ok(r.stdout.indexOf("to-err") !== -1, r.stdout);
        assert.equal(r.stderr, "");
    });

    it("returns empty output for a command that prints nothing", function() {
        var r = exec_command('exit 0');
        assert.equal(r.stdout, "");
        assert.equal(r.stderr, "");
    });

    it("captures output larger than a pipe buffer", function() {
        // ~64KB, well past the ~4KB pipe buffer that makes shell.Exec()
        // deadlock; redirection has to handle it without blocking.
        var r = exec_command('for /L %i in (1,1,2000) do @echo 0123456789012345678901234567890');
        assert.ok(r.stdout.length > 60000, "only got " + r.stdout.length + " bytes");
    });

    it("keeps quoted paths in the command intact", function() {
        var r = exec_command('cmd /c echo "C:\\Program Files\\x"');
        assert.ok(r.stdout.indexOf("Program Files") !== -1, r.stdout);
    });

    it("leaves no temp files behind", function() {
        var fso    = new ActiveXObject("Scripting.FileSystemObject");
        var tmpDir = fso.GetSpecialFolder(2).Path;
        exec_command('echo x');
        var leftover = 0;
        var files = new Enumerator(fso.GetFolder(tmpDir).files);
        for (; !files.atEnd(); files.moveNext()) {
            if (String(files.item().Name).indexOf("jsw_exec_") === 0) { leftover++; }
        }
        assert.equal(leftover, 0, leftover + " jsw_exec_ temp files left behind");
    });

});

// ===========================================================================
// Process listing
// ===========================================================================

describe("list_processes", function() {

    it("returns an array", function() {
        assert.ok(Array.isArray(list_processes()));
    });

    it("finds more than one process running", function() {
        assert.ok(list_processes().length > 1);
    });

    it("gives each entry a pid, name, command line and parent pid", function() {
        var p = list_processes()[0];
        assert.equal(typeof p.pid,          "number");
        assert.equal(typeof p.name,         "string");
        assert.equal(typeof p.command_line, "string");
        assert.equal(typeof p.parent_pid,   "number");
    });

    it("finds cscript.exe, which is running this suite", function() {
        assert.ok(list_processes("cscript.exe").length > 0);
    });

    it("matches a name with no .exe suffix", function() {
        assert.ok(list_processes("cscript").length > 0);
    });

    it("matches case-insensitively", function() {
        assert.ok(list_processes("CSCRIPT.EXE").length > 0);
    });

    it("returns an empty array for a name nothing matches", function() {
        assert.deepEqual(list_processes("jsw_no_such_process_12345.exe"), []);
    });

    it("filters to only the requested process", function() {
        var found = list_processes("cscript.exe");
        var all_match = true;
        for (var i = 0; i < found.length; i++) {
            if (String(found[i].name).toLowerCase() !== "cscript.exe") { all_match = false; }
        }
        assert.ok(all_match, "filter returned a process with another name");
    });

});

describe("process_exists", function() {

    it("is true for cscript.exe", function() {
        assert.equal(process_exists("cscript.exe"), true);
    });

    it("is false for a process that is not running", function() {
        assert.equal(process_exists("jsw_no_such_process_12345.exe"), false);
    });

});

describe("kill_process", function() {

    it("returns 0 when nothing matches", function() {
        assert.equal(kill_process("jsw_no_such_process_12345.exe"), 0);
    });

    it("returns 0 for a pid that is not running", function() {
        // Pids are recycled, but 0x7FFFFFFF is out of range for a real one.
        assert.equal(kill_process(2147483647), 0);
    });

    it("terminates a process this test started", function() {
        // A long ping keeps a distinctly-named copy of cmd.exe alive; start it
        // detached, confirm it is there, kill it by pid, confirm it is gone.
        var before = list_processes("ping.exe").length;
        run_command('cmd.exe /c ping -n 30 127.0.0.1', { wait: false });
        sleep(700);

        var running = list_processes("ping.exe");
        assert.ok(running.length > before,
                  "the ping process under test never started (before=" + before +
                  ", now=" + running.length + ")");

        var target = running[running.length - 1].pid;
        assert.equal(kill_process(target), 1);
        sleep(500);

        var still_there = false;
        var after = list_processes("ping.exe");
        for (var i = 0; i < after.length; i++) {
            if (after[i].pid === target) { still_there = true; }
        }
        assert.notOk(still_there, "pid " + target + " survived kill_process");
    });

});

// ---------------------------------------------------------------------------
// Cleanup - remove every value written above, then the key itself
// ---------------------------------------------------------------------------

(function() {
    var values = ["str", "num", "auto_num", "expand", "empty", "present"];
    for (var i = 0; i < values.length; i++) {
        try { reg_delete(REG_BASE + values[i]); } catch (e) {}
    }
    try { reg_delete(REG_BASE); } catch (e) {}
}());

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
