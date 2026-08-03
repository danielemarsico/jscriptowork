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

- [ ] **GitHub Pages: a hand-written landing page.**
      One self-contained `docs/index.html`, served via Settings → Pages →
      "Deploy from a branch", `main` / `/docs`. No static-site generator, no
      Jekyll — a single page, in keeping with the zero-dependency ethos.
      - Content: hero (name + the tagline "Bring the power of modern JavaScript
        to Windows CScript"), a short "what it is", the lib/feature table from
        the README, a quick-start (`cscript.exe bin\launcher.js yourscript.js`),
        and links to the README, CHANGELOG, latest release, and the repo.
      - Self-contained: inline all CSS, no external fonts/CDN/scripts, so the
        page renders offline and can't be broken by a third-party outage.
      - Responsive; light/dark via `prefers-color-scheme`.
      - Acceptance: enabling Pages serves the page; it renders with the network
        blocked (no external requests in the page source).

- [ ] **Ko-fi donate button on the landing page.** Depends on the page above.
      **Blocked: needs the Ko-fi handle from the repo owner** — wire it as a
      plain styled link to `https://ko-fi.com/<HANDLE>` (no external widget
      script, to keep the page self-contained), placed in the header or footer.
      Until the handle is supplied, leave a clearly-marked `<!-- TODO: ko-fi
      handle -->` placeholder rather than a guessed URL.

### Distribution

- [ ] **Minify the bundle at build time (external tool permitted here).**
      Decision: a Node/npm minifier (e.g. terser) is allowed at **build time
      only** — the shipped artifact stays pure JScript, so end users still need
      nothing but Windows. `build.js` (run under cscript) remains the canonical
      bundler; minification is a **separate, opt-in** maintainer step, never
      required by `build.bat`.
      - Add a `tools/minify.mjs` (or an npm script) that reads
        `dist/launcher.js` and writes `dist/launcher.min.js`.
      - Terser config MUST target ES5/ES3 output (no ES6 emitted) and MUST NOT
        break JScript's eval-scoping model: the top-level bare-assignment
        globals (`foo = function(){}`) that survive `load()`'s eval scope must
        keep their exact names — do **not** mangle or scope top-level names.
        Keep `'use strict'` handling in mind (JScript parses but doesn't
        enforce it).
      - Obfuscation is a secondary, optional goal. Given the eval-scope quirks,
        aggressive global name-mangling is risky; if attempted, it must preserve
        every public global the launcher and user scripts reference. Recommend
        whitespace/comment stripping + local mangling only.
      - Acceptance: `dist/launcher.min.js` runs an example (e.g.
        `examples/hello-world.js` or a headless one) with identical output to
        `dist/launcher.js`; CI runs at least one suite through the minified
        bundle on the `windows-latest` runner to prove JScript still accepts it.

- [ ] **Attach the built artifacts to every GitHub release.** Depends on the
      minify step.
      - Establish a versioning convention first: the repo has no releases yet
        and the CHANGELOG uses `Unreleased` + dated sections. Adopt
        `vMAJOR.MINOR.PATCH` tags and, on release, promote CHANGELOG
        `Unreleased` to a version heading.
      - Add a release workflow (`.github/workflows/release.yml`) triggered on
        `v*` tag push: build `dist/`, run the minifier, and upload
        `launcher.js`, `launcher.min.js`, and `launcher.bat` as release assets.
      - Acceptance: pushing a `vX.Y.Z` tag produces a GitHub Release carrying
        those three assets.

- [ ] **`--compile`: bundle libs + a user script into one standalone `.js`.**
      Produce a single file that runs via `cscript.exe myscript.bundled.js`
      with no `libs/` folder and no launcher — the natural extension of what
      `build.js` already does for the generic launcher.
      - Reuse `build.js`'s inlining machinery. Prepend the launcher bootstrap
        (`log`, `CURRENT_FOLDER`/`ROOT_FOLDER`, `read_all_text_file`, and
        `load()` as a no-op), inline the needed libs, then append the user
        script body in place of the argument-driven executor.
      - Which libs to inline: scan the user script for `load("x")` calls and
        include those (in `libNames` order, so `core` precedes `polyfills`,
        etc.); fall back to "all libs" if scanning is ambiguous. Preserve load
        order — it matters (e.g. `log.js` must come after `console.js`).
      - CLI shape: `cscript.exe build.js --compile myscript.js [--out path]`,
        or a dedicated `tools/compile.js`. Reuse `libs/minimist.js` for args.
      - Acceptance: the compiled file runs standalone and produces the same
        output as running the source through `bin/launcher.js`; extend
        `bin/tests/test-build.js` to compile a fixture script and assert the
        expected libs are inlined and `load()` is a no-op.

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