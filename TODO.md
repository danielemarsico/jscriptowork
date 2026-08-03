# TODO

Known bugs, gaps, and planned work. Items marked **[test]** have a `skip()`-ed
test in `bin/tests/` that documents the intended behaviour — un-skip it once the
bug is fixed.

## Features

- [ ] `libs/console.js` — split the `console` shim out of `polyfills.js` so it
      can be loaded on its own (original roadmap item).
- [ ] Async/streaming HTTP: `MSXML2.ServerXMLHTTP`, plus non-text responses.
      `http_request` already has a `timeout` parameter and exposes response
      headers to the callback (see CHANGELOG); it is still synchronous
      (`open(..., false)`) despite the callback-shaped API — a real async
      rewrite is what remains here.
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

## New feature and improvements

- [ ] create github pages to advertise the project
- [ ] add button ko-fi to github pages 
- [ ] add a feture to minimize and possible obfuscate the library js
- [ ] add the minimized and obfuscated (if possible) library as pacakge for every release
- [ ] add a tool to "compile" a single file from library and the script devveloped, so it can be luacnhed just with cscript.exe myscript.exe
- [ ] create an example whihc allows the user to select a folder than the script zip the files contained in the folder, upload them on pastebin or equivalent web site, or discord, than take the urlk and create an QR code out of it and display to the user.
- [ ] study the possibility to have the compiled library + script inbase64 format and run cscript on it, like base64 decoding and run eval