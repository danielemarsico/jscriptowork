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

- [x] `do_in_excel` / `do_in_access` / `do_in_word` — done:
      `bin/tests/test-office.js`, opt-in through `JSW_TEST_OFFICE=1`. Everything
      skips without it, so the runner and CI can include the file
      unconditionally; with it set, each application is probed separately.
      **Never executed:** writing this suite needs no Office, running it does.
      It has been checked only in its skip-everything state.

- [ ] **`do_in_access` cannot open a database outside `CURRENT_FOLDER`.** It
      builds its path as `CURRENT_FOLDER + "/" + database_filename`, so an
      absolute path becomes nonsense and a database anywhere else is
      unreachable. Fix: use the argument as-is when it is already absolute
      (`^[A-Za-z]:\\`, `^\\\\`), and keep the `CURRENT_FOLDER` prefix only for a
      bare filename — that keeps every existing caller working.
      **[test]** → `bin/tests/test-office.js`,
      `skip("opens a database given as an absolute path")`

## New feature and improvements

Design decisions already taken are recorded inline so the work can start
without re-litigating them. The recurring constraint across all of these is the
project's core promise (README / CLAUDE.md): **a target machine needs nothing
but Windows** — no Node.js, no npm, at *run* time. Where a task relaxes that, it
says so and confines the relaxation to build/maintainer time.

### Website

- [x] **GitHub Pages: a hand-written landing page.** Done — `docs/index.html`.
      Single file, all CSS inline, no fonts/CDN/scripts, favicon as a `data:`
      URI; the only external URLs are anchor `href`s. Responsive, light/dark via
      `prefers-color-scheme`.
      **Remaining manual step (repo owner):** Settings → Pages → "Deploy from a
      branch" → `main` / `/docs`. Nothing in the repo can enable that.

- [ ] **Ko-fi donate button on the landing page.**
      **Blocked: needs the Ko-fi handle from the repo owner.** The markup is
      already in `docs/index.html` — a plain styled `.kofi` link in the header
      (no widget script, so the page stays self-contained), carrying `hidden`
      and a `<!-- TODO: ko-fi handle -->` comment. To finish: set the `href` to
      `https://ko-fi.com/<HANDLE>` and remove the `hidden` attribute.

### Distribution

- [x] **Minify the bundle at build time (external tool permitted here).**
      Done — `tools/minify.mjs` (terser), opt-in, never called by `build.bat`.
      Top-level names are never mangled, output is ES5-only, property access and
      quoted keys are left alone (ES3 rejects reserved words as bare property
      names). Before writing, it verifies every public top-level name survived
      and that the output parses as ES5. `dist/launcher.min.js` is gitignored —
      `build.js` wipes `dist/` on every run, so it is transient by construction.
      CI (`minified` job) runs `test-core.js` through the minified bundle on
      `windows-latest` and diffs an example's output against the plain bundle.
      Obfuscation was deliberately not attempted: whitespace/comment stripping
      plus local mangling only, which is what the eval-scope model can take.

- [x] **Attach the built artifacts to every GitHub release.** Done —
      `.github/workflows/release.yml`, triggered on a `v*` tag push. Convention
      is `vMAJOR.MINOR.PATCH` (recorded in CLAUDE.md "Releasing" and in the
      CHANGELOG header): promote `Unreleased` to `## [X.Y.Z] - YYYY-MM-DD`,
      then tag. The job runs the suite, builds, minifies, smoke-tests both
      bundles under `cscript.exe`, and uploads `launcher.js`,
      `launcher.min.js`, `launcher.bat` via the runner's `gh` CLI. Release
      notes come from `tools/changelog-notes.mjs`, which fails the release when
      the version has no CHANGELOG section.
      **Untested end-to-end:** nothing short of pushing a real tag exercises
      the workflow, and that is the repo owner's call, not a task an agent
      should take. The notes extractor has its own unit checks and was run
      against the real CHANGELOG.

- [x] **`--compile`: bundle libs + a user script into one standalone `.js`.**
      Done — `cscript.exe build.js --compile myscript.js [--out path] [--all-libs]`.
      Shares `dist/`'s emitters, scans `load("...")` calls and inlines those libs
      in `libNames` order, falls back to every lib when a `load()` argument is
      not a literal, errors on a lib that does not exist, and emits
      `_jsw_hta_inline_libs` only when `ui` is included. Arguments go through
      `libs/minimist.js`, loaded with `new Function` (build.js has no `load()`),
      with a long-flags-only fallback parser if that fails.
      `bin/tests/test-build.js` gained 16 tests: a compiled fixture is inspected
      *and executed* as a subprocess.

### Examples

- [x] **`examples/share-folder.js`: folder → zip → anonymous upload → QR.**
      Done. Zips with `tar.exe` through `exec_command()` (and stops with a clear
      message on a pre-1803 machine rather than guessing — the
      `Shell.Application` `CopyHere` fallback is documented in a comment, not
      implemented). Uploads to 0x0.st as `multipart/form-data`, with the body
      assembled in `ADODB.Stream` and sent through `MSXML2.ServerXMLHTTP`. The
      returned URL is drawn as a QR code by `libs/qrcode.js` — offline, no image
      fetched — on the console and in a window. The header and a pre-upload
      prompt both spell out that the file becomes public and expires; nothing is
      sent without an explicit `yes`.
      **Not executed:** it needs Windows, the network and a desktop, and its
      acceptance test is a human scanning the code. Example-only, no suite.

- [x] **Offline QR code generation.** Done — `libs/qrcode.js`, written from
      scratch rather than vendored: versions 1-40, L/M/Q/H, numeric /
      alphanumeric / byte (UTF-8), Reed-Solomon over GF(256), all eight masks
      with the spec's penalty scoring, BCH format/version information, plus
      ASCII / HTML / SVG renderers. `examples/qr-code-generator.js` now encodes
      locally and needs no network.
      **Known scope limit:** one segment per symbol — the encoder picks a single
      mode for the whole string rather than splitting mixed text into
      alphanumeric and numeric runs. Symbols stay valid and scannable; a mixed
      string just uses a slightly larger version than an optimising encoder
      would. Worth revisiting only if symbol size becomes a real constraint.

### Research

- [x] **Study: distribute the bundle as a base64 payload run through cscript.**
      Done — findings in `studies/base64-payload.md`, working proof of concept in
      `studies/make-base64-bundle.js` (it runs `test-core.js` and the examples
      from a base64 payload).
      **Recommendation: do not productise.** The outer file has to stay JScript,
      as predicted; base64 is a flat ~41% size tax (236 KB bundle → 333 KB
      payload, 117 KB minified → 166 KB); `eval` of one giant string collapses
      every error in the bundle onto one line of the wrapper; and it hides
      nothing. `--compile` already ships one file, smaller and debuggable, and
      the minifier already halves it. No `--base64` mode.
      Worth keeping in mind for a different problem: embedding *binary* assets
      in a distributable `.js`, where encoding bytes as text has no alternative.
