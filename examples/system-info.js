// system-info.js - registry, environment and process helpers via libs/win.js
//
// Reads a few well-known registry values, expands environment variables, runs
// a command and captures its output, and lists running processes - the kind of
// machine inspection automation scripts usually shell out to reg.exe and
// tasklist.exe for.
//
// Everything here is read-only except one value written into (and deleted
// from) HKCU\Software\jscriptowork_example, which needs no admin rights.
//
// Run via:
//   examples\run.bat system-info.js
//   cscript.exe bin\launcher.js examples\system-info.js
//   cscript.exe dist\launcher.js examples\system-info.js

load("core");
load("polyfills");
load("system");
load("win");

// ---------------------------------------------------------------------------
// 1. Registry - read
// ---------------------------------------------------------------------------

log("--- Windows ---");

var CV = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\";

// Second argument = fallback: returned instead of throwing when the value is
// missing, which keeps this working across Windows versions that renamed it.
log("  product   : " + reg_read(CV + "ProductName", "(unknown)"));
log("  build     : " + reg_read(CV + "CurrentBuild", "(unknown)"));
log("  installed : " + reg_read(CV + "RegisteredOwner", "(not set)"));

// ---------------------------------------------------------------------------
// 2. Registry - write, read back, delete
// ---------------------------------------------------------------------------

log("");
log("--- Registry round trip (HKCU) ---");

var KEY = "HKCU\\Software\\jscriptowork_example\\";

log("  exists before write : " + reg_exists(KEY + "last_run"));
reg_write(KEY + "last_run", format_date(new Date(), "YYYY/MM/DD"));
reg_write(KEY + "run_count", 1, "REG_DWORD");

log("  last_run            : " + reg_read(KEY + "last_run"));
log("  run_count           : " + reg_read(KEY + "run_count"));

reg_delete(KEY + "last_run");
reg_delete(KEY + "run_count");
reg_delete(KEY);
log("  exists after delete : " + reg_exists(KEY + "last_run"));

// ---------------------------------------------------------------------------
// 3. Environment
// ---------------------------------------------------------------------------

log("");
log("--- Environment ---");
log("  user      : " + env("USERNAME"));
log("  computer  : " + env("COMPUTERNAME"));
log("  temp      : " + expand_env("%TEMP%"));
log("  not set   : '" + env("JSW_NOT_SET") + "'  (missing variables read as empty)");

// ---------------------------------------------------------------------------
// 4. Running a command and capturing its output
// ---------------------------------------------------------------------------

log("");
log("--- Commands ---");

// run_command: exit code only, nothing captured.
log("  'exit 3' exit code   : " + run_command('cmd.exe /c exit 3'));

// exec_command: stdout, stderr and exit code, via temp-file redirection (so it
// cannot deadlock on a full pipe the way shell.Exec() does).
var whoami = exec_command('whoami');
log("  whoami stdout        : " + whoami.stdout.trim());
log("  whoami exit code     : " + whoami.exit_code);

var failing = exec_command('dir "C:\\no_such_folder_12345"');
log("  failing exit code    : " + failing.exit_code);
log("  failing stderr       : " + failing.stderr.trim());

// ---------------------------------------------------------------------------
// 5. Processes
// ---------------------------------------------------------------------------

log("");
log("--- Processes ---");

var all = list_processes();
log("  running processes    : " + all.length);

// This script is running under cscript.exe, so this can never be empty.
var mine = list_processes("cscript.exe");
log("  cscript.exe instances: " + mine.length);
for (var i = 0; i < mine.length; i++) {
    log("    pid " + mine[i].pid + "  parent " + mine[i].parent_pid);
}

log("  explorer running     : " + process_exists("explorer.exe"));
log("  nonsense.exe running : " + process_exists("jsw_nonsense_12345.exe"));

// The five processes with the most instances, just to show the data is real.
var counts = {};
for (var j = 0; j < all.length; j++) {
    var n = all[j].name.toLowerCase();
    counts[n] = (counts[n] || 0) + 1;
}
var ranked = Object.keys(counts).map(function(name) {
    return { name: name, count: counts[name] };
});
ranked.sort(function(a, b) { return b.count - a.count; });

log("");
log("  most instances:");
for (var k = 0; k < Math.min(5, ranked.length); k++) {
    log("    " + ranked[k].count + " x " + ranked[k].name);
}

// kill_process(name_or_pid) terminates matching processes and returns how many
// it got. Not demonstrated here for obvious reasons:
//
//     kill_process("notepad.exe");   // by name
//     kill_process(12345);           // by pid
