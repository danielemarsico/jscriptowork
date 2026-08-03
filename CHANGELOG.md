# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
This project does not yet publish versioned releases; dated sections below are
reconstructed from the git history.

## [Unreleased]

### Added

- `CLAUDE.md` — working notes for agents: the `load()`/`eval` scoping rules,
  JScript syntax constraints, testing and build instructions.
- `TODO.md` — known bugs, gaps, and planned work.
- `CHANGELOG.md` — this file.
- `libs/minitest.js`: `skip(name, reason)` / `_test.skip()` for tests that are
  blocked on a known bug or on an unavailable resource (Office, network,
  desktop). Skipped tests are counted separately in the summary and never fail
  the run.
- New test suites:
  - `bin/tests/test-core.js` — Array/String polyfills, `JSON.stringify`,
    `JSON.parse`, `Date.prototype.toJSON`, `Ext.encode` / `Ext.decode`.
  - `bin/tests/test-minitest.js` — the assertion helpers and the runner itself.
  - `bin/tests/test-minimist.js` — long options, `=` form, `--no-` form,
    booleans, strings, aliases, defaults, `--` separator, `stopEarly`.
  - `bin/tests/test-system.js` — `format_date`, `parse_date`, `randomString`,
    `load_properties`, `save_working_directory` / `load_working_directory`,
    `read`/`write` surface checks.
  - `bin/tests/test-helpers.js` — `read_choice_from_input`,
    `read_date_from_input`, `select_file_from_folder`,
    `select_files_from_folder`, `fill_sheet`, `read_sheet_data`,
    `execute_query`, `get_current_date_as_excel_text`, using stubbed stdin and
    mock Excel/Access COM objects.
- `bin/tests/test-crypto.js` gained SHA-256 message-padding boundary vectors
  (55/56/57/63/64/65/1000 bytes), UTF-8 encoding vectors (2-, 3- and 4-byte
  characters), byte-array vectors (`0xFF`, all 256 byte values), RFC 4231 test
  case 1, and HMAC key-length vectors covering the 64-byte block-size boundary
  and the oversized-key hashing path.
- `bin/tests/run-tests.bat` now runs every suite, grouped by what each one
  requires (offline / disk / network / desktop).

### Fixed

- `parse_date` (`libs/system.js`) read `d.substring(...)` before `d` was
  assigned and ignored its own `ds` parameter, so every call threw
  `TypeError: 'd' is undefined`. Now uses `ds` for the substrings, declares `d`
  first, and falls back to returning the input string for an unrecognized
  format.
- `minimist` short options (`libs/minimist.js`) were unusable: `arg.slice(-1)[0]`
  relies on string bracket indexing, which JScript 5.8 does not support, so
  `key` was always `undefined`. Changed to `arg.charAt(arg.length - 1)`; `-f
  value`, `-v`, `-abc` clustering, `-n5`, `-x=1`, and short-flag aliasing all
  work now.
- `read_sheet_data` (`libs/helpers.js`) branched on `typeof value == "date"`,
  which is never true (`typeof` a `Date` is `"object"`), so date cells fell
  through to `value.trim()` and threw. Now checks `value instanceof Date`.
- `read_sheet_data` threw on a numeric cell in the first column under
  auto-detected column count: the end-of-rows check called `v.trim()` before
  any type conversion. Now converts with `String(v)` first.
- `read_sheet_data` threw on a boolean cell, since booleans matched none of the
  `typeof` branches and fell through to `value.trim()`. Added a `String(value)`
  fallback for any unhandled type.
- `read_date_from_input` (`libs/helpers.js`) matched with an unanchored
  `/\d{8}/`, so `abc20240115` was accepted, `parseInt("abc2")` produced `NaN`,
  and the function returned a (truthy) Invalid Date instead of retrying. The
  pattern is now anchored (`/^\d{8}$/`) and the resulting date is validated
  with `isNaN(d.getTime())` before returning.
- `randomString` (`libs/system.js`) was a `function` declaration inside a lib,
  so it was scoped to `load()` and unreachable from scripts run through
  `bin/launcher.js` (only reachable via the bundled `dist/launcher.js`).
  Converted to `randomString = function(...)` for consistent behaviour.
- `INPUT_FOLDER`, `OUTPUT_FOLDER`, and `DATABASEPATH` are read but never
  defined by either launcher, so `load_working_directory`, `write_report_to_file`,
  and `do_in_access` (without an explicit filename) threw `ReferenceError`
  whenever a caller hadn't defined them. All three now guard the reference with
  `typeof ... !== "undefined"` and fall back to `""`, so they only throw if a
  caller genuinely needs a value neither the global nor an explicit parameter
  supplied.
- `bin/tests/test-crypto.js` had two incorrect expected digests, so the suite
  reported failures against a correct implementation. Both now match the
  published vectors (verified independently):
  `sha256("abc")` is `ba7816bf…f20015ad`, not `ba7816bf…df54f6b8`; RFC 4231 test
  case 2 is `5bdcc146…64ec3843`, not `5bdcc146…64a37827`. No change to
  `libs/crypto.js` was needed.

## [2026-03-01] — UI, crypto, and the test framework

### Added

- `libs/ui.js` — `open_hta(options, onClose)`: builds and launches a native
  Windows HTA window through `mshta.exe`, with `jsw_return(value)` to hand a
  result back to the calling CScript process.
- `libs/polyfills.js` — extended compatibility layer: `Array.isArray/of/from`,
  `every`, `some`, `includes`, `findIndex`, `fill`, `flat`, `flatMap`;
  `String.prototype.endsWith/includes/repeat/padStart/padEnd/trimStart/trimEnd`;
  `Object.keys/values/entries/assign/create/freeze/isFrozen`; `Number.isNaN`,
  `isFinite`, `isInteger`, `parseInt`, `parseFloat`, `EPSILON`,
  `MAX_SAFE_INTEGER`, `MIN_SAFE_INTEGER`; `Math.sign/trunc/log2/log10/cbrt/
  hypot/clz32`; `Date.now`, `Date.prototype.toISOString`;
  `Function.prototype.bind`; and a `console` shim over `WScript.Echo`.
- `libs/crypto.js` — `sha256`, `sha256_bytes`, `hmac_sha256`, pure ES3 with no
  external dependencies.
- `libs/minitest.js` — minimal synchronous test framework (`describe`, `it`,
  `assert`, `_test.summary()`).
- `libs/system.js` file-system helpers: `read_text_file`, `file_exists`,
  `folder_exists`, `delete_file`, `create_folder`, `write_binary_file`,
  `read_binary_file`.
- `build.js` / `build.bat` — bundles every lib plus the bootstrap into
  `dist/launcher.js` and `dist/launcher.bat`, so a target machine needs only two
  files.
- `examples/hello-world.js` and `examples/run.bat`.
- Test suites `test-polyfills.js`, `test-filesystem.js`, `test-http.js`,
  `test-ui.js`, `test-crypto.js`, and the `run-tests.bat` runner.
- `README.md` rewritten: usage, project structure, polyfill table, and an
  explicit list of JScript features that can and cannot be polyfilled.

### Changed

- `http_request(url, method, callback, body, headers)` now accepts an optional
  request body and an optional headers object, and validates the method against
  `GET`/`POST`/`PUT`/`DELETE`.
- `bin/tests/launcher.js` reports the failing path (`cannot load file: <path>`)
  instead of a generic message.

### Removed

- `bin/Nuova cartella/` scratch folder (`cllist.js`, `test-winsock.*`, old
  `test.*`, `test.properties`).
- `index.html`, `README.txt`, `bin/README.txt`, `templates/REPORT_YYMMDD.xlsx`.

## [2022-02-24] — Project restructure

### Added

- `bin/launcher.js` and `bin/launcher.bat`: the `load()` + `eval()` module
  loader and the `log()` / `read_all_text_file()` bootstrap.
- `bin/tests/` with its own launcher copy, `test.js`, `test-httpconnect.js`,
  and batch wrappers.
- `libs/system.js`: `load_working_directory`, `save_working_directory`,
  `load_properties`, `write_text_to_file`, stdin/stdout helpers
  (`read_line`, `read`, `read_all`, `write_line`, `write`), `list_folders`,
  `randomString`, `format_date`, `parse_date`, `http_request`.

### Changed

- Everything moved out of the repository root into `bin/`, `libs/`, and
  `templates/`.

## [2021-10-07] — First draft

### Added

- `libs/core.js` — Array (`indexOf`, `filter`, `map`, `forEach`, `find`,
  `reduce`) and String (`trim`, `startsWith`) polyfills, Crockford's `json2`
  `JSON.stringify` / `JSON.parse`, and the `Ext.encode` / `Ext.decode`
  shorthands.
- `libs/helpers.js` — interactive prompts (`select_file_from_folder`,
  `select_files_from_folder`, `read_choice_from_input`,
  `read_date_from_input`) and COM automation (`do_in_excel`, `do_in_access`,
  `do_in_word`, `fill_sheet`, `read_sheet_data`, `execute_query`,
  `write_report_to_file`, `get_current_date_as_excel_text`).
- `libs/minimist.js` — vendored command-line argument parser.
- `cllist.js`, `test-winsock.*`, `test.*`, `index.html`, and the initial
  `templates/REPORT_YYMMDD.xlsx` report template.
- MIT `LICENSE`.
