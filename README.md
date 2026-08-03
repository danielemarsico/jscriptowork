# jscriptowork

Bring the power of modern JavaScript to Windows CScript.

CScript (Windows Script Host) uses the JScript engine, which is roughly ES3/ES5 compatible and lacks many modern ECMAScript features. This project provides a polyfill layer and utilities to close that gap, letting you write more expressive JavaScript for automation tasks on Windows.

## Usage

Run a script via the launcher:

```bat
cscript.exe bin\launcher.js path\to\your\script.js
```

Inside your script, load libraries with `load()`:

```js
load("core");      // base polyfills (Array, String, JSON)
load("polyfills"); // extended polyfill layer (Array, String, Object, Number, Math, Date, Function)
load("system");    // file I/O, stdin/stdout, HTTP
load("helpers");   // Excel, Access, Word COM automation
```

## Project structure

```
bin/
  launcher.js       entry point: defines log(), load(), read_all_text_file(), then evals the target script
  launcher.bat      convenience batch wrapper
  tests/            test suites (run through bin/launcher.js) + run-tests.bat
libs/
  core.js           ES5 baseline: Array/String basics + Crockford json2
  ext.js            Ext.util.JSON / Ext.encode / Ext.decode compatibility shim (opt-in, unused elsewhere)
  polyfills.js      extended layer: Array/String/Object/Number/Math/Date/Function
  console.js        console.log/info/warn/error/debug shim
  log.js            log() levels (debug/info/warn/error) + tee to a file
  system.js         stdin/stdout, file system, binary files, HTTP (sync, async, binary), date formatting
  csv.js            csv_parse/csv_format, read_csv_file/write_csv_file (no Excel needed)
  helpers.js        interactive prompts + Excel / Access / Word COM automation
  win.js            registry, process listing/killing, command execution with captured output
  crypto.js         sha256, sha256_bytes, hmac_sha256, hmac_sha256_bytes
  base64.js         base64_encode/decode, base64_encode_bytes/decode_bytes (no native btoa/atob)
  ui.js             open_hta(): native Windows GUI windows via mshta.exe, with live progress
  minimist.js       command-line argument parser (vendored)
  minitest.js       describe / it / assert / skip test framework
build.js            bundles libs + launcher into dist/
dist/               generated: launcher.js (all libs inlined) + launcher.bat
examples/           runnable examples
```

Run `build.bat` (or `cscript.exe build.js`) to regenerate `dist/`, a
self-contained two-file distributable (`launcher.js` + `launcher.bat`) that
needs no separate `libs/` folder — every lib is inlined and `load()` becomes a
no-op.

## How `load()` works

`bin/launcher.js` reads a lib's source and `eval()`s it **inside `load()`'s own
function scope**, not at global scope:

```js
function load(modulename){
    var path = ROOT_FOLDER + "/libs/" + modulename + ".js";
    var lib  = read_all_text_file(path);
    eval(lib);            // <-- runs inside load()'s scope
}
```

This has two consequences that matter if you write your own lib or script:

1. **Public functions must be bare assignments, not declarations.**

   ```js
   my_helper = function(a, b) { ... };   // GOOD — becomes a global
   function my_helper(a, b) { ... }      // BAD  — dies with load()'s scope
   ```

2. **A top-level `var` in your script is local to the bootstrap, not global.**
   To override a library global (e.g. stubbing `read_line` for a test), assign
   **without** `var`.

`dist/launcher.js` does not have this quirk — `load()` is a no-op there
because every lib is already inlined at the top level.

## Polyfilled features (core.js)

These are missing from JScript and provided by the framework:

| Feature | Status |
|---|---|
| `Array.prototype.indexOf` | polyfilled |
| `Array.prototype.filter` | polyfilled |
| `Array.prototype.map` | polyfilled |
| `Array.prototype.forEach` | polyfilled |
| `Array.prototype.find` | polyfilled |
| `Array.prototype.reduce` | polyfilled |
| `String.prototype.trim` | polyfilled |
| `String.prototype.startsWith` | polyfilled |
| `JSON.stringify` / `JSON.parse` | polyfilled |

## CScript / JScript limitations

### Features that can be polyfilled (runtime)

The following are absent from JScript but can be added via the polyfill layer:

**Array**
- `Array.isArray`, `Array.from`, `Array.of`
- `Array.prototype.every`, `some`, `includes`, `findIndex`
- `Array.prototype.flat`, `flatMap`
- `Array.prototype.fill`, `copyWithin`, `keys`, `values`, `entries`

**String**
- `String.prototype.endsWith`, `includes`, `repeat`
- `String.prototype.padStart`, `padEnd`
- `String.prototype.trimStart` / `trimLeft`, `trimEnd` / `trimRight`

**Object**
- `Object.keys`, `Object.values`, `Object.entries`
- `Object.assign`, `Object.create` (partial)
- `Object.freeze`, `Object.isFrozen` — JScript objects cannot truly be frozen,
  so `freeze` is a documented no-op (returns its argument unchanged) and
  `isFrozen` reports `true` for any non-object (including `null`), matching
  the "frozen" contract only in the narrow sense that primitives can't be
  mutated. Treat both as advisory, not enforced.

**Number**
- `Number.isNaN`, `Number.isFinite`, `Number.isInteger`
- `Number.parseInt`, `Number.parseFloat`
- `Number.EPSILON`, `Number.MAX_SAFE_INTEGER`, `Number.MIN_SAFE_INTEGER`

**Math**
- `Math.sign`, `Math.trunc`, `Math.cbrt`
- `Math.log2`, `Math.log10`
- `Math.hypot`, `Math.clz32`

**Other**
- `Date.now()`
- `Function.prototype.bind`
- `console` (shim to `WScript.Echo` / `StdOut`) — in `libs/console.js`

### Features that CANNOT be polyfilled (syntax-level)

These require a transpiler (e.g. Babel) and **cannot** be used directly in CScript scripts:

| Feature | Workaround |
|---|---|
| `let` / `const` | use `var` |
| Arrow functions `() => {}` | use `function() {}` |
| Template literals `` `Hello ${name}` `` | use string concatenation |
| Destructuring `const {a, b} = obj` | assign properties manually |
| Spread / rest `...args` | use `arguments` or loops |
| `class` syntax | use prototype-based patterns |
| `for...of` loops | use `for` or `forEach` |
| Default parameters `fn(x = 1)` | check `arguments` inside the function |
| `Promise` / `async` / `await` | use callbacks |
| Generators / `yield` | not available |
| `import` / `export` | use `load()` helper |
| `Symbol` | not available |
| `WeakMap` / `WeakSet` | not available |
| Optional chaining `?.` | use explicit null checks |
| Nullish coalescing `??` | use `||` with care |

### Other runtime constraints

- No `setTimeout` / `setInterval` / `clearTimeout` — CScript is synchronous
- No DOM, `window`, `document`, `navigator`
- No `fetch` — use `MSXML2.XMLHTTP` via `http_request()` helper
- No module system — files are loaded via `eval()` through `load()`
- No `require` / CommonJS / ESM
- Regex lookbehind assertions may not work
- `typeof` on undeclared variables works, but behaviour may differ in edge cases
- `'use strict'` is accepted but not fully enforced by JScript

## Examples

`examples/` holds runnable scripts — `examples\run.bat <name>.js`:

| Example | Shows |
|---|---|
| `hello-world.js` | a native window via `open_hta()` |
| `progress-window.js` | streaming progress out of an HTA while it is still open |
| `csv-report.js` | reading/filtering/writing CSV, with log levels and a log file |
| `system-info.js` | registry, environment, subprocess output, process listing |
| `json-encode-decode.js` | `JSON.stringify`/`parse` plus a real HTTP GET |
| `base64-encode-decode.js` | base64 over strings and over file bytes |
| `qr-code-generator.js` | prompt for a URL, render its QR code in a window |

## Roadmap

The polyfill layer, test framework, full test-suite runner, and the feature
backlog that was outstanding — the `console` shim split into `libs/console.js`,
CSV helpers, log levels, async and binary HTTP, HTA progress streaming, and
registry/process helpers — are all done. Remaining known bugs and ideas are
tracked in [TODO.md](TODO.md).

## License

MIT — see [LICENSE](LICENSE)
