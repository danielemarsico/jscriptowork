// console.js - console shim for JScript / CScript
//
// JScript has no console object. This provides the familiar
// console.log/info/warn/error/debug surface on top of whatever output channel
// the host offers.
//
//   load("console");
//   console.log("hello", 42, {a: 1});   ->  hello 42 {"a":1}
//   console.warn("careful");            ->  [WARN] careful
//
// Output routing, in order of preference:
//   1. log()          - defined by bin/launcher.js, dist/launcher.js AND by the
//                       HTA bootstrap in ui.js, so this is the portable path
//                       and the one that honours libs/log.js levels and file
//                       tee-ing when that lib is loaded.
//   2. WScript.Echo() - plain WSH with no launcher.
//   3. no-op          - nothing to write to; never throws.
//
// This shim used to live in polyfills.js. It was split out so a script can pull
// in console on its own, and so polyfills.js stays purely a language-level
// compatibility layer.
//
// Objects are rendered with JSON.stringify, so load("core") first if you want
// anything better than [object Object] on a host without native JSON.

if (typeof console === 'undefined') {
    console = (function() {

        function _write(text) {
            // Resolved per call, not captured once: log() may be redefined
            // later (libs/log.js does exactly that).
            if (typeof log === 'function') {
                log(text);
            } else if (typeof WScript !== 'undefined') {
                WScript.Echo(text);
            }
        }

        function _format(args) {
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
        }

        function _print(prefix, args) {
            _write(prefix + _format(args));
        }

        return {
            log:   function() { _print('',         arguments); },
            info:  function() { _print('[INFO] ',  arguments); },
            warn:  function() { _print('[WARN] ',  arguments); },
            error: function() { _print('[ERROR] ', arguments); },
            debug: function() { _print('[DEBUG] ', arguments); }
        };
    }());
}
