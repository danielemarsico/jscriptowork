# TODO

Known bugs, gaps, and planned work. Items marked **[test]** have a `skip()`-ed
test in `bin/tests/` that documents the intended behaviour — un-skip it once the
bug is fixed.

## Bugs

### High — these throw at runtime

- [ ] **[test]** `parse_date(ds, format)` (`libs/system.js:161`) is broken. The
      body reads `d.substring(...)` before `var d = new Date()` is assigned, and
      ignores its own `ds` parameter, so *every* call throws
      `TypeError: 'd' is undefined` — including the `format not recognized`
      branch, which also ends with `return d.toString()`.
      Fix: use `ds` for the substrings and declare `d` first.
      → `bin/tests/test-system.js`, `describe("parse_date")`
- [ ] **[test]** `minimist` short options are unusable
      (`libs/minimist.js:195`). `var key = arg.slice(-1)[0]` uses string bracket
      indexing, which JScript 5.8 does not support (see the note at the top of
      `libs/crypto.js`), so `key` is `undefined` and `key.split('.')` throws for
      any `-x`-style argument. Fix: `arg.charAt(arg.length - 1)`.
      → `bin/tests/test-minimist.js`, `describe("minimist - short options")`
- [ ] **[test]** `read_sheet_data` (`libs/helpers.js:402`) branches on
      `typeof value == "date"`, which is never true — `typeof` a `Date` is
      `"object"`. A real date cell therefore falls through to `value.trim()` and
      throws. Fix: test with `value instanceof Date` (or
      `Object.prototype.toString.call(value) === '[object Date]'`).
      → `bin/tests/test-helpers.js`, `describe("read_sheet_data - date cells")`
- [ ] **[test]** `read_sheet_data` throws on a numeric cell in the **first**
      column when column count is auto-detected: the end-of-rows check
      (`libs/helpers.js:349`) calls `v.trim()` before any type conversion, and
      numbers have no `trim`. Passing `ncolumns` avoids it, because that branch
      uses `v.toString().trim()`. Fix: convert with `String(v)` in both checks.
      → `bin/tests/test-helpers.js`, `describe("read_sheet_data - auto column detection")`
- [ ] **[test]** `read_sheet_data` throws on a boolean cell — booleans match none
      of the `typeof` branches (`libs/helpers.js:398`) and fall through to
      `value.trim()`. Fix: default to `String(value)` for anything unhandled.
      → `bin/tests/test-helpers.js`, `describe("read_sheet_data - auto column detection")`
- [ ] **[test]** `read_date_from_input` matches with an unanchored `/\d{8}/`
      (`libs/helpers.js:86`), so `abc20240115` is accepted, `parseInt("abc2")`
      yields `NaN`, and the function returns an **Invalid Date** — which is
      truthy, so the retry loop does not catch it. Fix: anchor the pattern with
      `/^\d{8}$/` and validate the resulting date.
      → `bin/tests/test-helpers.js`, `describe("read_date_from_input")`
- [ ] `INPUT_FOLDER`, `OUTPUT_FOLDER`, and `DATABASEPATH` are read but never
      defined by either launcher, so `load_working_directory`
      (`libs/system.js:7`), `write_report_to_file` (`libs/helpers.js:451`), and
      `do_in_access` without an explicit filename (`libs/helpers.js:159`) throw
      `ReferenceError`. Decide whether the launchers should define defaults or
      the helpers should take them as parameters.
- [ ] `randomString` (`libs/system.js:135`) is a `function` declaration inside a
      lib, so it is scoped to `load()` and is **not** reachable from scripts run
      through `bin/launcher.js` — while it *is* reachable from the bundled
      `dist/launcher.js`. Convert it to `randomString = function(...)` for
      consistent behaviour. Same class of bug to watch for in any new lib.
      → `bin/tests/test-system.js`, `describe("randomString")`

### Medium

- [ ] `read_all_text_file` in `bin/launcher.js` and `bin/tests/launcher.js` has
      `f.Close()` after the `return`, so the handle is never closed. It also
      swallows every error as `"file doesn't exist"` (the `bin/tests` copy at
      least reports the path).
- [ ] `read(n)` (`libs/system.js:87`) ignores its `n` argument and always reads a
      single character.
- [ ] `read_all()` (`libs/system.js:91`) returns the literal string
      `"end of stream"` as a sentinel, which callers cannot distinguish from
      real input. Return `null` or `""`.
- [ ] `list_folders` (`libs/system.js:120`) enumerates `folder.files`, so it
      returns *files*, not folders. It also assigns `fso` without `var`, leaking
      a global. Rename to `list_files` (keeping an alias) and add a real
      folder-listing helper.
- [ ] `load_properties` (`libs/system.js:37`) drops any line whose value
      contains `=` (it requires exactly two `split("=")` parts), has no comment
      or blank-line handling, and builds JS source for `eval()` without escaping
      quotes in values.
      → `bin/tests/test-system.js`, `describe("load_properties")`
- [ ] `format_date` (`libs/system.js:147`) documents `YY/MM/DD` and `YYYYMMDD`
      in its comment but only implements `YYYY/MM/DD`.
      → `bin/tests/test-system.js`, `describe("format_date")`
- [ ] `write_text_to_file` (`libs/system.js:62`) opens the file in ASCII mode and
      never closes it if `Write` throws. Non-ASCII content is mangled; add the
      Unicode flag and a `try/finally`.
- [ ] `fill_sheet` (`libs/helpers.js:248`) indexes a 27-character `alphabet`
      string, so it silently produces an empty column letter past 26 columns.
- [ ] `do_in_access` (`libs/helpers.js:159`) declares `var db` twice — once for
      the database filename, once for `access.CurrentDb()`. Rename the first.
- [ ] `minimist` logs `'procssing args'` (typo included) on every argument
      (`libs/minimist.js:125`). Remove the debug line.
- [ ] `minimist` throws if called with a single argument, because it reads
      `opts['unknown']` unguarded. Upstream minimist treats `opts` as optional.
- [ ] `http_request` (`libs/system.js:260`) is synchronous (`open(..., false)`)
      despite its callback-shaped API, has no timeout, and exposes no response
      headers.

### Low

- [ ] `core.js`'s `Array.prototype.indexOf` polyfill ignores a negative `start`
      argument (`start || 0`), unlike the spec.
- [ ] `core.js` defines the globals `Ext`, `Ext.util.JSON`, `Ext.encode`, and
      `Ext.decode` unconditionally; they are a compatibility shim nothing in the
      project uses. Consider dropping or moving them to their own lib.
- [ ] `Object.freeze` in `polyfills.js` is a documented no-op, and `Object.isFrozen`
      reports non-objects as frozen. Both are unavoidable on JScript — keep, but
      make the limitation visible in the README polyfill table.
- [ ] `minitest.describe` tracks a `current` block name that is never used when
      reporting an `it` result.
- [ ] `read_sheet_data` hardcodes `MAX_ROWS = 3000` / `MAX_COLUMNS = 100`; make
      them overridable via `params`.
- [ ] `bin/launcher.js` and `bin/tests/launcher.js` are near-duplicates that have
      already drifted (different error messages). Generate one from the other, or
      make the tests use `bin/launcher.js`.
- [ ] `bin/launcher.bat` and the older `bin/tests/test*.bat` wrappers print a
      joke banner and call `cscript.exe launcher.js` with a relative path, so
      they only work from inside `bin/`.
- [ ] `bin/tests/test.js` and `bin/tests/test-httpconnect.js` predate `minitest`
      and are manual scripts, not suites. Port them or delete them.

## Documentation

- [ ] `README.md` "Project structure" lists a `templates/` folder that no longer
      exists (the report template was removed in the 2026-03-01 restructure).
- [ ] `README.md` "Roadmap" is stale: `libs/polyfills.js`, `libs/minitest.js`,
      `bin/tests/test-polyfills.js`, and `bin/tests/run-tests.bat` all exist.
      Only `libs/console.js` is outstanding — and the `console` shim now lives in
      `polyfills.js`, so the item should either be dropped or the shim split out.
- [ ] `README.md` does not mention `libs/crypto.js`, `libs/ui.js`,
      `libs/minitest.js`, `build.js`/`dist/`, or `examples/`.
- [ ] Document the `load()`/`eval` scoping rule in `README.md` too — it is the
      single most surprising thing about writing a lib here (currently only in
      `CLAUDE.md`).

## Features

- [ ] `libs/console.js` — split the `console` shim out of `polyfills.js` so it
      can be loaded on its own (original roadmap item).
- [ ] Async/streaming HTTP: `MSXML2.ServerXMLHTTP` with a timeout, plus access to
      response headers and non-text responses.
- [ ] CSV read/write helpers in `system.js` — currently every caller goes through
      Excel COM for tabular data.
- [ ] Logging levels for `log()` (`debug`/`info`/`warn`/`error`) and an option to
      tee output to a file.
- [ ] `open_hta`: a way to stream progress back from the HTA to CScript while the
      window is open (today only the final `jsw_return` value crosses the
      boundary).
- [ ] Registry and process helpers (`WScript.Shell` / WMI) — common in the kind
      of automation this framework targets.

## Testing and tooling

- [ ] No CI. `cscript.exe` only exists on Windows, so this needs a
      `windows-latest` GitHub Actions runner invoking `bin\tests\run-tests.bat`
      and failing the job on a non-zero failure count.
- [ ] `minitest` never sets a process exit code — `_test.summary()` should call
      `WScript.Quit(failed ? 1 : 0)` (behind an option, so it does not kill
      interactive sessions) before CI is worth wiring up.
- [ ] No test for `build.js`: nothing verifies that `dist/launcher.js` contains
      every lib in `libNames`, that `load()` is stubbed out, or that
      `_jsw_hta_inline_libs` is valid escaped source.
- [ ] `test-http.js` depends on `httpbin.org`; a local `ServerXMLHTTP`-free stub
      or a recorded-response mode would make the suite runnable offline.
- [ ] `hmac_sha256` only accepts UTF-8 strings, so RFC 4231 test cases 3, 4, 6
      and 7 (keys and data made of raw `0xaa`/`0xdd` bytes) cannot be expressed.
      An `hmac_sha256_bytes(keyBytes, msgBytes)` overload would make the full
      RFC vector set testable.
      → `bin/tests/test-crypto.js`, `describe("hmac_sha256 - RFC 4231 known-answer vectors")`
- [ ] `do_in_excel` / `do_in_access` / `do_in_word` are only smoke-checked for
      existence; they need an opt-in suite that runs on a machine with Office.
      → `bin/tests/test-helpers.js`, `describe("Office COM wrappers")`
- [ ] `dist/launcher.js` is committed but nothing checks it is in sync with
      `libs/`; a drift check (rebuild + compare) belongs in CI.
