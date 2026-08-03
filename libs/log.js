// log.js - logging levels and file tee-ing for JScript / CScript
//
//   load("log");
//
//   log("still works exactly as before");   // INFO, no prefix
//   log_debug("noisy detail");              // [DEBUG] noisy detail
//   log_info("something happened");         // [INFO] something happened
//   log_warn("something looks wrong");      // [WARN] something looks wrong
//   log_error("something broke");           // [ERROR] something broke
//
//   set_log_level("warn");        // drop everything below WARN
//   log_to_file("C:\\run.log");   // also append every line to a file
//   set_log_timestamps(true);     // prefix each line with an ISO-ish stamp
//
// Loading this lib REPLACES the global log() defined by the launcher. It stays
// backward compatible on purpose:
//   - log("x") writes exactly "x" - no level prefix - as it always did;
//   - the default level is INFO, so nothing that used to print stops printing.
// Everything minitest and the existing suites emit goes through log(), so the
// test output is unchanged unless a caller opts in to a higher level.
//
// The original launcher log() is kept as _log_console, and all console output
// goes through the replaceable _log_sink, which is what makes level filtering
// testable without a console to inspect.

// ---------------------------------------------------------------------------
// Levels
// ---------------------------------------------------------------------------

// Numeric so comparisons are cheap; gaps left for future levels.
LOG_LEVELS = { debug: 10, info: 20, warn: 30, error: 40, none: 100 };

_log_level      = LOG_LEVELS.info;
_log_timestamps = false;
_log_file_path  = null;

// The original launcher log(), captured before it is replaced below. Under
// bin/launcher.js this is the global function declaration; under
// dist/launcher.js it is the same function, inlined at the top of the bundle.
_log_console = (typeof log === 'function') ? log : null;

// Everything that reaches the console goes through here. Tests replace it to
// capture output; set_log_sink(null) restores the default.
_log_sink = function(text) {
    if (_log_console) {
        _log_console(text);
    } else if (typeof WScript !== 'undefined') {
        WScript.Echo(text);
    }
};

// ---------------------------------------------------------------------------
// Private helpers (_log_ prefix, per the load()/eval scoping note in CLAUDE.md)
// ---------------------------------------------------------------------------

// Accepts a name ("warn", case-insensitive) or a number. Returns null when the
// value names no known level.
_log_resolve_level = function(level) {
    if (typeof level === 'number') { return level; }
    if (typeof level === 'string') {
        var key = level.toLowerCase();
        if (typeof LOG_LEVELS[key] === 'number') { return LOG_LEVELS[key]; }
    }
    return null;
};

// "2026-08-03 15:52:48"
_log_timestamp = function() {
    var d = new Date();
    var p = function(n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
           ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
};

// Renders one call's arguments the way console.js does: objects as JSON, the
// rest via String(), joined with spaces.
_log_format_args = function(args) {
    var parts = [];
    for (var i = 0; i < args.length; i++) {
        var v = args[i];
        if (v !== null && typeof v === 'object' && typeof JSON !== 'undefined') {
            try {
                parts.push(JSON.stringify(v));
            } catch (e) {
                parts.push(String(v));
            }
        } else {
            parts.push(String(v));
        }
    }
    return parts.join(' ');
};

// Appends one line to the tee file. A failure here must never take down the
// caller - losing a log line is better than losing the run - so the write is
// swallowed and tee-ing is switched off to avoid failing once per line.
_log_write_file = function(line) {
    if (!_log_file_path) { return; }
    var ForAppending = 8;
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var f   = fso.OpenTextFile(_log_file_path, ForAppending, true);
        try {
            f.WriteLine(line);
        } finally {
            f.Close();
        }
    } catch (e) {
        _log_file_path = null;
    }
};

// The one path every level function funnels through.
_log_emit = function(levelValue, prefix, args) {
    if (levelValue < _log_level) { return; }
    var line = prefix + _log_format_args(args);
    if (_log_timestamps) { line = _log_timestamp() + ' ' + line; }
    _log_sink(line);
    _log_write_file(line);
};

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

// Sets the minimum level that gets written. Accepts "debug"/"info"/"warn"/
// "error"/"none" or a number. Throws on an unknown name, rather than silently
// muting or unmuting the whole run.
set_log_level = function(level) {
    var resolved = _log_resolve_level(level);
    if (resolved === null) {
        throw new Error("set_log_level: unknown level: " + level);
    }
    _log_level = resolved;
};

// Returns the current level as a number (compare against LOG_LEVELS).
get_log_level = function() {
    return _log_level;
};

// Returns the current level's name, or "" for a numeric level with no name.
get_log_level_name = function() {
    for (var name in LOG_LEVELS) {
        if (Object.prototype.hasOwnProperty.call(LOG_LEVELS, name) &&
            LOG_LEVELS[name] === _log_level) {
            return name;
        }
    }
    return "";
};

// Starts appending every emitted line to path (created if missing).
// Pass null to stop tee-ing. Returns the path now in use, or null.
log_to_file = function(path) {
    _log_file_path = path ? path : null;
    return _log_file_path;
};

// Returns the current tee file path, or null.
get_log_file = function() {
    return _log_file_path;
};

// Turns the leading timestamp on each line on or off.
set_log_timestamps = function(on) {
    _log_timestamps = !!on;
};

// Replaces the console sink; pass null to restore the default. Used by tests to
// capture output, and usable to route logging somewhere else entirely.
set_log_sink = function(fn) {
    if (fn) {
        _log_sink = fn;
    } else {
        _log_sink = function(text) {
            if (_log_console) {
                _log_console(text);
            } else if (typeof WScript !== 'undefined') {
                WScript.Echo(text);
            }
        };
    }
};

// ---------------------------------------------------------------------------
// Level functions
// ---------------------------------------------------------------------------

log_debug = function() { _log_emit(LOG_LEVELS.debug, '[DEBUG] ', arguments); };
log_info  = function() { _log_emit(LOG_LEVELS.info,  '[INFO] ',  arguments); };
log_warn  = function() { _log_emit(LOG_LEVELS.warn,  '[WARN] ',  arguments); };
log_error = function() { _log_emit(LOG_LEVELS.error, '[ERROR] ', arguments); };

// Replaces the launcher's log(). Logs at INFO with no prefix, so existing
// output is byte-identical until a caller changes the level.
log = function() { _log_emit(LOG_LEVELS.info, '', arguments); };
