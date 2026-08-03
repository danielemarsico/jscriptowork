// win.js - Windows registry and process helpers for JScript / CScript
//
//   load("core"); load("system"); load("win");
//
// Registry (WScript.Shell):
//   reg_read(key [, fallback])        read a value
//   reg_write(key, value [, type])    write a value
//   reg_delete(key)                   delete a value or an empty key
//   reg_exists(key)                   true when the value/key can be read
//
// Processes:
//   run_command(command [, options])  run something, get its exit code
//   exec_command(command [, opts])    run something, capture stdout/stderr
//   list_processes([filter])          running processes via WMI
//   process_exists(name)              true when a process by that name runs
//   kill_process(name_or_pid)         terminate matching processes
//
// Environment:
//   env(name)                         one process environment variable
//   expand_env(text)                  expand %VARS% in a string
//
// Registry key syntax is WScript.Shell's: a key path ending in a backslash
// addresses the key's default value, anything else addresses a named value.
//
//     reg_read("HKCU\\Software\\MyApp\\Setting")
//     reg_write("HKCU\\Software\\MyApp\\Count", 3, "REG_DWORD")
//
// Roots may be written as HKCU/HKLM/HKCR/HKU/HKCC or spelled out
// (HKEY_CURRENT_USER and friends).
//
// Depends on libs/system.js for read_text_file / file_exists / delete_file.

// ---------------------------------------------------------------------------
// Private helpers (_win_ prefix, per the load()/eval scoping note in CLAUDE.md)
// ---------------------------------------------------------------------------

_win_shell = function() {
    return new ActiveXObject("WScript.Shell");
};

// Connects to WMI on the local machine. Kept in one place so a future
// remote-machine option only has to change here.
_win_wmi = function() {
    return GetObject("winmgmts:\\\\.\\root\\cimv2");
};

// Case-insensitive compare that also lets "notepad" match "notepad.exe", so
// callers do not have to remember the extension.
_win_name_matches = function(processName, wanted) {
    var a = String(processName).toLowerCase();
    var b = String(wanted).toLowerCase();
    if (a === b) { return true; }
    if (b.length > 0 && a === b + ".exe") { return true; }
    return false;
};

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

// Reads a registry value. Throws when the value does not exist, unless a
// fallback argument is supplied - in which case that is returned instead.
// Passing an explicit undefined counts as supplying one.
reg_read = function(key, fallback) {
    try {
        return _win_shell().RegRead(key);
    } catch (e) {
        if (arguments.length > 1) { return fallback; }
        throw new Error("reg_read: cannot read " + key + ": " + e.message);
    }
};

// Writes a registry value, creating the key if needed. type defaults to
// REG_DWORD for numbers and REG_SZ for everything else; pass it explicitly for
// REG_EXPAND_SZ or REG_BINARY.
reg_write = function(key, value, type) {
    if (!type) { type = (typeof value === "number") ? "REG_DWORD" : "REG_SZ"; }
    try {
        _win_shell().RegWrite(key, value, type);
    } catch (e) {
        throw new Error("reg_write: cannot write " + key + ": " + e.message);
    }
};

// Deletes a value, or a key when the path ends in a backslash. Windows only
// deletes an empty key - remove its values and subkeys first.
reg_delete = function(key) {
    try {
        _win_shell().RegDelete(key);
    } catch (e) {
        throw new Error("reg_delete: cannot delete " + key + ": " + e.message);
    }
};

// True when the value (or key) can be read. Never throws.
reg_exists = function(key) {
    try {
        _win_shell().RegRead(key);
        return true;
    } catch (e) {
        return false;
    }
};

// ---------------------------------------------------------------------------
// Environment
// ---------------------------------------------------------------------------

// One process environment variable, or "" when it is not set.
env = function(name) {
    return String(_win_shell().Environment("Process")(name));
};

// Expands %VARIABLES% in a string: expand_env("%TEMP%\\out.txt").
expand_env = function(text) {
    return String(_win_shell().ExpandEnvironmentStrings(text));
};

// ---------------------------------------------------------------------------
// Processes
// ---------------------------------------------------------------------------

// Runs a command and returns its exit code.
//
// options:
//   window   0 hidden, 1 normal.        default: 0
//   wait     wait for it to finish.     default: true
//
// With wait:false the process is left running and 0 is returned - there is no
// exit code to report yet.
run_command = function(command, options) {
    options = options || {};
    var window_style = (typeof options.window === "number") ? options.window : 0;
    var wait         = (typeof options.wait   === "boolean") ? options.wait : true;
    return _win_shell().Run(command, window_style, wait);
};

// Runs a command through cmd.exe and captures what it printed.
//
// Returns { exit_code: number, stdout: string, stderr: string }.
//
// Output is redirected to temp files rather than read from shell.Exec()'s
// pipes. Exec() looks simpler but deadlocks as soon as the child fills a pipe
// buffer while the script is not draining that exact stream - a real failure
// mode for anything chatty. Redirection has no such limit.
//
// options:
//   window   0 hidden, 1 normal.               default: 0
//   merge    fold stderr into stdout (2>&1).   default: false
exec_command = function(command, options) {
    options = options || {};
    var window_style = (typeof options.window === "number") ? options.window : 0;

    var fso    = new ActiveXObject("Scripting.FileSystemObject");
    var tmpDir = fso.GetSpecialFolder(2).Path;
    var stamp  = (new Date()).getTime() + "_" + Math.floor(Math.random() * 1000000);
    var outPath = tmpDir + "\\jsw_exec_out_" + stamp + ".txt";
    var errPath = tmpDir + "\\jsw_exec_err_" + stamp + ".txt";

    // cmd.exe /s /c "..." strips exactly the outer pair of quotes and runs the
    // rest verbatim, which is what keeps quoted paths inside `command` intact.
    var redirect = options.merge
        ? ' > "' + outPath + '" 2>&1'
        : ' > "' + outPath + '" 2> "' + errPath + '"';
    var full = 'cmd.exe /s /c "' + command + redirect + '"';

    var exit_code = _win_shell().Run(full, window_style, true);

    var readIfPresent = function(path) {
        if (!file_exists(path)) { return ""; }
        var text = "";
        try { text = read_text_file(path); } catch (e) { text = ""; }
        try { delete_file(path); } catch (e) {}
        return text;
    };

    return {
        exit_code: exit_code,
        stdout:    readIfPresent(outPath),
        stderr:    options.merge ? "" : readIfPresent(errPath)
    };
};

// Lists running processes as objects:
//   { pid: number, name: string, command_line: string, parent_pid: number }
//
// filter, when given, keeps only processes whose name matches it ("notepad"
// and "notepad.exe" both match notepad.exe, case-insensitively).
list_processes = function(filter) {
    var items = _win_wmi().ExecQuery(
        "SELECT ProcessId, Name, CommandLine, ParentProcessId FROM Win32_Process");
    var out = [];
    var e   = new Enumerator(items);
    for (; !e.atEnd(); e.moveNext()) {
        var p = e.item();
        var name = String(p.Name);
        if (filter && !_win_name_matches(name, filter)) { continue; }
        out.push({
            pid:          Number(p.ProcessId),
            name:         name,
            // CommandLine is null for processes this account may not inspect.
            command_line: (p.CommandLine === null) ? "" : String(p.CommandLine),
            parent_pid:   Number(p.ParentProcessId)
        });
    }
    return out;
};

// True when at least one process with that name is running.
process_exists = function(name) {
    return list_processes(name).length > 0;
};

// Terminates processes by name ("notepad.exe") or by pid (a number).
// Returns how many were terminated - 0 when nothing matched.
//
// Processes the current account may not terminate are skipped rather than
// throwing, so one protected process does not abandon the rest of the batch.
kill_process = function(name_or_pid) {
    var by_pid = (typeof name_or_pid === "number");
    var items  = _win_wmi().ExecQuery("SELECT ProcessId, Name FROM Win32_Process");
    var killed = 0;
    var e      = new Enumerator(items);
    for (; !e.atEnd(); e.moveNext()) {
        var p = e.item();
        var match = by_pid
            ? (Number(p.ProcessId) === name_or_pid)
            : _win_name_matches(String(p.Name), name_or_pid);
        if (!match) { continue; }
        try {
            p.Terminate();
            killed++;
        } catch (ex) {}
    }
    return killed;
};
