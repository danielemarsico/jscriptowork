# CLAUDE.md

Guidance for Claude Code (and any other agent) working in this repository.

## What this project is

`jscriptowork` is a small framework that makes Windows Script Host (`cscript.exe`)
usable for real automation work. WSH runs **JScript 5.8**, an ES3-era engine, so
the project ships a polyfill layer, a set of helpers (file system, HTTP, Office
COM automation, SHA-256, HTA-based GUIs), a minimal test framework, and a build
step that bundles everything into a single distributable launcher.

There is **no npm, no package.json, no node_modules**. Nothing here runs under
Node.js. The runtime is `cscript.exe` on Windows.

## Repository layout

```
bin/
  launcher.js         entry point: defines log(), load(), read_all_text_file(), then evals the target script
  launcher.bat        batch wrapper
  tests/
    test-*.js         test suites (see "Testing"); run through bin/launcher.js
    run-tests.bat     runs every suite in order
libs/
  core.js             ES5 baseline: Array/String basics + Crockford json2
  ext.js              Ext.util.JSON / Ext.encode / Ext.decode compatibility shim (unused elsewhere)
  polyfills.js        extended layer: Array/String/Object/Number/Math/Date/Function
  console.js          console.log/info/warn/error/debug, routed through log()
  log.js              log() levels + tee to a file; REPLACES the launcher's log()
  csv.js              csv_parse/csv_format/read_csv_file/write_csv_file
  win.js              registry, processes (WMI), command execution with captured output
  system.js           stdin/stdout, file system, binary files, HTTP, date formatting
  helpers.js          interactive prompts + Excel / Access / Word COM automation
  crypto.js           sha256, sha256_bytes, hmac_sha256
  base64.js           base64_encode/decode, base64_encode_bytes/decode_bytes (no native btoa/atob)
  ui.js               open_hta(): native Windows GUI windows via mshta.exe
  minimist.js         command-line argument parser (vendored)
  minitest.js         describe / it / assert / skip test framework
build.js              bundles libs + launcher into dist/
dist/                 generated: launcher.js (all libs inlined) + launcher.bat
examples/             runnable examples
```

## How code loading works — read this before editing any lib

`bin/launcher.js` defines:

```js
function load(modulename){
    var path = ROOT_FOLDER + "/libs/" + modulename + ".js";
    var lib  = read_all_text_file(path);
    eval(lib);            // <-- eval happens INSIDE load()'s function scope
}
```

The `eval()` runs inside `load()`'s scope, not at global scope. Consequences that
bite every time:

1. **Declare public functions as bare assignments, never as declarations.**

   ```js
   my_helper = function(a, b) { ... };   // GOOD — becomes a global
   function my_helper(a, b) { ... }      // BAD  — dies with load()'s scope
   ```

   `libs/system.js` still has `function randomString(...)`, which is why that
   helper is unreachable from scripts loaded through `bin/launcher.js` (it *is*
   reachable from `dist/launcher.js`, where libs are inlined at top level). Do
   not copy that pattern.

2. **Private helpers must be prefix-namespaced, not IIFE-wrapped.** JScript's
   `eval` does not reliably preserve closure scope for function declarations
   inside an IIFE — inner helpers silently return `undefined`. See the note at
   the top of `libs/crypto.js`; it uses `_sha256_*` prefixes instead.
   `libs/minitest.js` gets away with an IIFE because it only assigns the
   *result* to a global.

3. **User scripts are also `eval`'d** (inside the launcher's bootstrap IIFE), so
   a top-level `var x` in a test file is local to that IIFE. To override a
   library global from a test (e.g. stubbing `read_line`), assign **without**
   `var`.

## Hard language constraints (JScript 5.8)

Anything in the "cannot be polyfilled" table in `README.md` is a syntax error and
will break the whole file at parse time. In practice, never write:

`let` / `const` · arrow functions · template literals · destructuring · spread or
rest · `class` · `for...of` · default parameters · `async` / `await` / `Promise`
· generators · `import` / `export` · `Symbol` · `WeakMap` / `WeakSet` · `?.` ·
`??` · trailing commas in object/array literals · getters/setters in object
literals · reserved words as bare property names (use `obj['default']`).

Additional runtime traps:

- **`str[i]` does not work.** Use `str.charAt(i)`. This is a real bug source —
  `libs/minimist.js` uses `arg.slice(-1)[0]` and short-flag parsing throws
  because of it (tracked in `TODO.md`).
- `typeof someDate` is `"object"`, never `"date"` (another live bug in
  `read_sheet_data`, tracked in `TODO.md`).
- No `setTimeout`/`setInterval` — everything is synchronous.
- `'use strict'` parses but is not enforced.
- Bit operations are 32-bit signed; `crypto.js` relies on `| 0` and `>>> 0`.

## Writing library code

- ES3 syntax only, 4-space indent, `snake_case` for public helpers
  (`write_text_to_file`, `read_sheet_data`), `_prefixed_snake_case` for private
  ones.
- Guard every polyfill with a feature check (`if (!Array.prototype.x) { ... }`).
- COM objects come from `new ActiveXObject("...")`. Always `Close()` / `Quit()`
  what you open, including on the error path.
- Adding a new lib means adding its name to `libNames` in `build.js` (and to
  `htaLibNames` too if HTAs need it).

## Testing

Tests use `libs/minitest.js`:

```js
load("core"); load("polyfills"); load("minitest");

describe("thing", function() {
    it("does something", function() { assert.equal(1 + 1, 2); });
    skip("does something else", "reason it is skipped");
});

_test.summary();               // required at the end of every suite
_test.summary({ exit: true }); // also sets the process exit code (0/1) — every
                                // suite in bin/tests/ uses this form so a
                                // runner (run-tests.bat, CI) can detect failure
```

Available assertions: `assert.ok`, `notOk`, `equal` (strict `!==`), `notEqual`,
`deepEqual` (JSON-string comparison), `throws`, `doesNotThrow`.
`skip(name, reason)` records a test as skipped without failing the run — use it
for known bugs and for anything that needs Office, a network, or a human.

Run everything from a Windows shell:

```bat
bin\tests\run-tests.bat
```

Or a single suite:

```bat
cscript.exe bin\launcher.js bin\tests\test-core.js
```

**Tests cannot be run from Linux/macOS** — there is no `cscript.exe` outside
Windows. When you change library code on a non-Windows machine, write or
update the matching suite and state plainly that the suite was not executed;
do not claim tests pass. CI (`.github/workflows/tests.yml`) runs on a
`windows-latest` GitHub Actions runner, which does have `cscript.exe`, and
invokes `bin\tests\run-tests.bat` directly; it sets `CI=true`, which
`run-tests.bat` uses to skip the final `pause`, `test-http.js` uses to switch
to its offline stub (so the job isn't at the mercy of httpbin.org's uptime),
and `test-ui.js` uses to skip the describes that open a window while still
running its parsing tests.

Suites and what they need:

| Suite | Needs |
|---|---|
| `test-core.js` | nothing |
| `test-ext.js` | nothing |
| `test-polyfills.js` | nothing |
| `test-minitest.js` | nothing |
| `test-crypto.js` | nothing |
| `test-base64.js` | nothing |
| `test-minimist.js` | nothing |
| `test-console.js` | nothing |
| `test-csv.js` | temp folder write access (parsing tests need nothing) |
| `test-system.js` | temp folder write access |
| `test-filesystem.js` | temp folder write access |
| `test-log.js` | temp folder write access (level tests use a captured sink) |
| `test-helpers.js` | temp folder write access (Office parts are skipped) |
| `test-win.js` | writes under `HKCU\Software\jscriptowork_test`, spawns `cmd.exe`; no admin rights needed |
| `test-build.js` | spawns `cscript.exe build.js` as a subprocess (regenerates `dist/`) |
| `test-http.js` | network access to httpbin.org — or set `JSW_TEST_HTTP_OFFLINE=1` (or `CI=true`) to use a local stub instead. The async/binary tests have no stub and skip themselves offline |
| `test-ui.js` | interactive desktop for the window tests; they skip themselves when `CI=true` (or `JSW_TEST_NO_DESKTOP` is set), and the progress-parsing tests always run |

## Build

```bat
build.bat            :: or: cscript.exe build.js
```

Regenerates `dist/launcher.js` (every lib inlined, `load()` becomes a no-op, HTA
libs embedded as an escaped string) and `dist/launcher.bat`. `dist/` is
committed, so regenerate and commit it whenever `libs/` changes.

## Conventions for changes

- Keep `README.md`, `CHANGELOG.md`, and `TODO.md` in step with the code:
  new feature → CHANGELOG entry under `Unreleased`; bug found but not fixed →
  `TODO.md` entry, and a `skip()`-ed test documenting the intended behaviour.
- Prefer fixing a lib over working around it in a caller.
- Do not add Node.js tooling, transpilers, or dependencies without being asked —
  the point of the project is that a target machine needs nothing but Windows.
