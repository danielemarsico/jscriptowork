// minitest.js - Minimal synchronous test framework for CScript / JScript
//
// Usage:
//   load("minitest");
//
//   describe("My feature", function() {
//       it("does something", function() {
//           assert.equal(1 + 1, 2);
//           assert.ok(true);
//       });
//       skip("does something else", "blocked on a known bug, see TODO.md");
//   });
//
//   _test.summary();   // always call at the end to print results

_test = (function() {

    var passed  = 0;
    var failed  = 0;
    var skipped = 0;

    function _print(msg) {
        if (typeof log === 'function') {
            log(msg);
        } else {
            WScript.Echo(msg);
        }
    }

    function _pad(n) {
        return n < 10 ? ' ' + n : '' + n;
    }

    // -----------------------------------------------------------------------
    // Public API
    // -----------------------------------------------------------------------

    var api = {};

    api.describe = function(name, fn) {
        _print('[' + name + ']');
        fn();
    };

    api.it = function(name, fn) {
        try {
            fn();
            passed++;
            _print('  PASS  ' + name);
        } catch (e) {
            failed++;
            _print('  FAIL  ' + name + ': ' + e.message);
        }
    };

    // Records a test as skipped without running it. Use for tests blocked on a
    // known bug (see TODO.md) or on a resource that is not always available
    // (Office, network, interactive desktop). Never fails the run.
    api.skip = function(name, reason) {
        skipped++;
        _print('  SKIP  ' + name + (reason ? ' (' + reason + ')' : ''));
    };

    api.summary = function() {
        var total = passed + failed;
        _print('');
        _print('=====================================');
        if (failed === 0) {
            _print('  ALL TESTS PASSED: ' + passed + ' / ' + total);
        } else {
            _print('  PASSED : ' + passed + ' / ' + total);
            _print('  FAILED : ' + failed + ' / ' + total);
        }
        if (skipped > 0) {
            _print('  SKIPPED: ' + skipped);
        }
        _print('=====================================');
    };

    // Counters, for suites that need to assert on the runner itself.
    api.counts = function() {
        return { passed: passed, failed: failed, skipped: skipped };
    };

    // -----------------------------------------------------------------------
    // assert helpers
    // -----------------------------------------------------------------------

    api.assert = {

        ok: function(val, msg) {
            if (!val) {
                throw new Error(msg || ('Expected truthy but got: ' + val));
            }
        },

        notOk: function(val, msg) {
            if (val) {
                throw new Error(msg || ('Expected falsy but got: ' + val));
            }
        },

        equal: function(actual, expected, msg) {
            if (actual !== expected) {
                throw new Error(msg || ('Expected ' + JSON.stringify(expected) + ' but got ' + JSON.stringify(actual)));
            }
        },

        notEqual: function(actual, expected, msg) {
            if (actual === expected) {
                throw new Error(msg || ('Expected value to differ from ' + JSON.stringify(expected)));
            }
        },

        deepEqual: function(actual, expected, msg) {
            var a = JSON.stringify(actual);
            var e = JSON.stringify(expected);
            if (a !== e) {
                throw new Error(msg || ('Expected ' + e + ' but got ' + a));
            }
        },

        throws: function(fn, msg) {
            var threw = false;
            try {
                fn();
            } catch (e) {
                threw = true;
            }
            if (!threw) {
                throw new Error(msg || 'Expected function to throw but it did not');
            }
        },

        doesNotThrow: function(fn, msg) {
            try {
                fn();
            } catch (e) {
                throw new Error(msg || ('Expected function not to throw but got: ' + e.message));
            }
        }
    };

    return api;
}());

// Expose top-level helpers as globals so test files stay concise
describe = function(name, fn) { _test.describe(name, fn); };
it       = function(name, fn) { _test.it(name, fn); };
skip     = function(name, reason) { _test.skip(name, reason); };
assert   = _test.assert;
