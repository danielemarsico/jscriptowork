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
load("system");    // file I/O, stdin/stdout, HTTP
load("helpers");   // Excel, Access, Word COM automation
```

## Project structure

```
bin/
  launcher.js       entry point for running scripts
  launcher.bat      convenience batch wrapper
  tests/            test scripts
libs/
  core.js           polyfills and JSON shim
  system.js         file system, stdio, HTTP helpers
  helpers.js        Excel / Access / Word automation
  minimist.js       argument parsing
templates/          report templates
```

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
- `Object.assign`, `Object.create` (partial), `Object.freeze`, `Object.isFrozen`

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
- `console` (shim to `WScript.Echo` / `StdOut`)

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

## Roadmap

- [ ] `libs/polyfills.js` — extended polyfill layer (Array, String, Object, Number, Math, Date, Function)
- [ ] `libs/minitest.js` — minimal test framework (describe / it / assert)
- [ ] `bin/tests/test-polyfills.js` — test suite for all polyfills
- [ ] `bin/tests/run-tests.bat` — batch runner for the full test suite
- [ ] `libs/console.js` — `console.log/warn/error` shim

## License

MIT — see [LICENSE](LICENSE)
