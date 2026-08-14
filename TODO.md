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

- [ ] **`examples/share-folder.js`: folder → zip → anonymous upload → QR.**
      Decision: upload to an **anonymous, no-signup file host** (0x0.st or
      file.io) — no API key, simplest to demo. Files are **public and expire**;
      state this plainly in the script header and in a prompt before uploading.
      - Select a folder (reuse the prompt helpers in `libs/helpers.js`, or a
        simple `read_line`).
      - Zip it with **no external download**: prefer `tar.exe` (built into
        Windows 10 1803+) via `exec_command` from `libs/win.js`
        (`tar -a -c -f out.zip -C parent folder`) — reliable and synchronous.
        Note the Win10+ requirement. The older `Shell.Application` "compressed
        folder" `CopyHere` trick is the fallback but needs an empty-zip header
        stub and a poll for its async copy; document that if used.
      - Upload the zip bytes as `multipart/form-data` via
        `MSXML2.ServerXMLHTTP`, assembling the body with `ADODB.Stream`
        (raw bytes can't live in a JScript string safely). Read the returned
        URL from the response.
      - QR the URL by reusing `examples/qr-code-generator.js`'s `open_hta`
        approach (api.qrserver.com today; switch to the offline generator below
        once it exists).
      - Acceptance: run it, pick a folder, scan the QR, and the URL downloads a
        zip identical to the source folder. Network + desktop required, so mark
        it `skip()` in any suite and keep it example-only.

- [ ] **Offline QR code generation.** `examples/qr-code-generator.js` (and the
      share-folder example above) currently render the QR via api.qrserver.com,
      which needs the network. Implement a real ES3 QR encoder — data-encoding
      modes, Reed–Solomon error correction, mask selection — or vendor an
      existing ES3-compatible generator, so the examples work with no network
      access. This is the largest single item here; a vendored, license-clean
      encoder is the pragmatic path.

### Research

- [ ] **Study: distribute the bundle as a base64 payload run through cscript.**
      A spike, not a committed feature. Goal: ship libs + script as one base64
      blob and execute it.
      - Reality check up front: `cscript.exe` cannot run a raw base64/text file
        — it needs a JScript (`.js`/`.wsf`) entry point. The feasible shape is a
        small JScript bootstrap that embeds the base64 string, decodes it with
        `libs/base64.js` (or an inline decoder), and `eval`s the result. So the
        outer file is still JScript; only the payload is base64.
      - Trade-offs to measure: base64 inflates size ~33% (partly offset by
        minifying first); `eval` of one large string loses line numbers in stack
        traces; net benefit over a plain minified bundle is unclear.
      - Deliverable of the *study*: a short findings note (feasible shape, real
        size numbers vs. the plain and minified bundles, error-handling caveats,
        recommendation) plus a working proof-of-concept that runs an example
        from a base64 payload. If it proves worthwhile, productize it later as a
        `--base64` mode on the compile tool above.