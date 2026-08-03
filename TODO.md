# TODO

Known bugs, gaps, and planned work. Items marked **[test]** have a `skip()`-ed
test in `bin/tests/` that documents the intended behaviour — un-skip it once the
bug is fixed.

## Features

- [ ] `http_request()` itself is still synchronous (`open(..., false)`), on
      purpose: its signature and callback shape are what every existing caller
      and the test suite's offline stub are built on. Real async lives alongside
      it as `http_request_async()` / `http_wait_all()`. Revisit only if the
      synchronous entry point becomes the bottleneck.

## Testing and tooling

- [ ] `do_in_excel` / `do_in_access` / `do_in_word` are only smoke-checked for
      existence; they need an opt-in suite that runs on a machine with Office.
      → `bin/tests/test-helpers.js`, `describe("Office COM wrappers")`

## New feature and improvements

- [ ] create github pages to advertise the project
- [ ] add button ko-fi to github pages 
- [ ] add a feture to minimize and possible obfuscate the library js
- [ ] add the minimized and obfuscated (if possible) library as pacakge for every release
- [ ] add a tool to "compile" a single file from library and the script devveloped, so it can be luacnhed just with cscript.exe myscript.exe
- [ ] create an example whihc allows the user to select a folder than the script zip the files contained in the folder, upload them on pastebin or equivalent web site, or discord, than take the urlk and create an QR code out of it and display to the user.
- [ ] study the possibility to have the compiled library + script inbase64 format and run cscript on it, like base64 decoding and run eval
- [ ] `examples/qr-code-generator.js` currently renders the QR code via
      api.qrserver.com (network required). Implement offline QR code
      generation (a real ES3 encoder — data encoding modes, Reed-Solomon
      error correction, mask selection — or an embedded/vendored generator)
      so the example works with no network access.