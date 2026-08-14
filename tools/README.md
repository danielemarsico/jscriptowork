# tools/

Maintainer-only build tooling. **Nothing in this folder is needed to run
jscriptowork.** A target machine still needs nothing but Windows — the shipped
artifact is plain JScript either way.

This is the one place in the project where Node.js and npm are allowed, and only
at *build* time. `build.js` (run under `cscript.exe`) remains the canonical
bundler; `build.bat` never calls anything in here.

## Setup

```bash
cd tools
npm install
```

## minify.mjs

Minifies a built bundle with [terser](https://terser.org/):

```bash
cscript.exe build.js            # on Windows: produces dist/launcher.js
node tools/minify.mjs           # produces dist/launcher.min.js
```

Options: `--in <file>` (default `dist/launcher.js`), `--out <file>` (default:
the input with `.min.js`), `--quiet`.

The terser settings are constrained by JScript 5.8 and by how the bundle is
loaded — the comment block at the top of `minify.mjs` explains each one. In
short:

- **ES5 output only.** ES6 syntax is a *parse* error in JScript, which kills the
  whole file rather than one line.
- **Top-level names are never mangled.** The bundle's public API *is* its set of
  top-level globals; user scripts are `eval`'d against them. Locals inside
  functions are mangled normally.
- **Property access is left alone** (`obj["default"]` is not collapsed to
  `obj.default`, object keys stay quoted): ES5 permits reserved words as
  property names, ES3 does not.

Before writing anything, the tool checks that every public top-level name found
in the input still appears in the output, and that the output parses as ES5
(via acorn). Both are smoke tests, not proof — only `cscript.exe` can prove
JScript accepts the file, which is why CI runs a test suite and an example
through the minified bundle on a `windows-latest` runner.

`dist/launcher.min.js` is deliberately not committed: `build.js` deletes and
recreates `dist/` on every run, so the file is transient by construction. It is
generated on demand and attached to GitHub releases.

## changelog-notes.mjs

Extracts one version's section from `CHANGELOG.md`, for use as GitHub release
notes:

```bash
node tools/changelog-notes.mjs v1.2.3               # to stdout
node tools/changelog-notes.mjs 1.2.3 --out notes.md
```

It exits non-zero when the version has no section, or the section is empty.
`.github/workflows/release.yml` runs it before creating the release, so a tag
whose CHANGELOG entry was never written fails the release instead of shipping
with empty notes.

