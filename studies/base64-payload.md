# Study: distributing the bundle as a base64 payload

**Status: complete. Recommendation: do not productise.**

A spike, asked for in `TODO.md`: could jscriptowork ship as a single base64 blob
run through `cscript.exe`? This note records what the shape actually has to be,
what it costs, what breaks, and why the answer is no.

The proof of concept is `studies/make-base64-bundle.js` and it works — a
base64-payload bundle really does run the test suite and the examples.

## The feasible shape

The idea "ship one base64 file and run it" cannot be taken literally.
`cscript.exe` dispatches on the file extension and needs a script engine to
hand the file to; a `.b64` of arbitrary bytes has no engine. So the outer file
is a normal `.js` either way. Only the *payload* is base64:

```js
var _payload = [
'Ly8ganNjcmlwdG93b3JrIGJ1bmRsZWQgbGF1bmNoZXI...',
'IGZyb20gbGlicy8gLSBkbyBub3QgZWRpdCBieSBoYW5...',
// ... a few thousand more lines
].join('');

function _decode_payload(b64) { /* ~20 lines, or a call into MSXML */ }

eval(_decode_payload(_payload));
```

Two details that matter and both turn out fine:

- **`eval` here is global scope.** At the top level of a `.js` that
  `cscript.exe` runs directly, a direct `eval` evaluates in global scope, so the
  payload's `var`s and function declarations become real globals — exactly as if
  the source had been pasted in. This is what makes the trick viable at all: the
  bundle's whole contract is its top-level names.
- **Chunking, not one giant literal.** The payload goes into an array of 76-character
  strings joined at runtime. A single 300 KB string literal on one line is legal
  but unreadable, awful in diffs, and leans on limits nobody documents. The join
  costs nothing worth measuring.

## Size, measured

Numbers from this repository, with `libs/qrcode.js` in the bundle:

| Artifact | Bytes | vs. plain bundle |
|---|---:|---:|
| `dist/launcher.js` (plain bundle) | 236,344 | 1.00x |
| `dist/launcher.min.js` (terser) | 117,398 | 0.50x |
| plain bundle as a base64 payload | 333,294 | 1.41x |
| minified bundle as a base64 payload | 166,354 | 0.70x |
| one example via `build.js --compile` | 65,246 | 0.28x |
| that compiled file as a base64 payload | 93,159 | 0.39x |

The 1.41x breaks down as base64's unavoidable 4/3 inflation, plus about 5% for
the quotes, commas and line breaks of the chunked array, plus a fixed ~1.5 KB of
bootstrap.

The one interesting row is base64-of-minified at 166 KB — smaller than the plain
bundle it replaces. But that saving is the minifier's, not base64's: the same
minified bundle is 117 KB on its own. **Base64 never wins on size. It is a 41%
tax on whatever you wrap.**

## What it costs

1. **Error reporting collapses.** Everything inside the payload is one `eval`,
   so a failure anywhere in 236 KB of source reports as a single line of the
   wrapper. This was not theoretical during the spike: an early bug in the
   generator (emitting the decoder but never calling it) surfaced as a syntax
   error at "line 1" whose text was a wall of base64. Debugging a real bundle in
   that state would be miserable, and `dist/launcher.js` already `eval`s the
   *user* script, so this would stack a second layer of the same problem.
2. **Startup does interpreted work before any of your code runs.** The
   hand-rolled decoder loops over 315,128 characters in JScript on every single
   run. The MSXML route (`MSXML2.DOMDocument` with `dataType="bin.base64"`, then
   `ADODB.Stream` to turn bytes back into text) pushes that into native code and
   is what a productised version should use — but it is COM, on a project whose
   pitch is that a target machine needs nothing but Windows. Both decoders are
   implemented; only the JScript one has actually been executed.
3. **Encoding is one more thing to get wrong.** The payload round-trips as bytes
   and is turned back into text as `iso-8859-1`, one byte to one character. That
   is exact for ASCII source, which is what `libs/` is. Anything else inherits
   whatever the ANSI code page decides — the same caveat that already applies to
   reading scripts through `FileSystemObject`.
4. **It hides nothing.** If the appeal is obfuscation: base64 is an encoding, not
   a cipher, and every reader of this note can decode it in one line. Anyone who
   can run the file can read it.

## What already solves the problems this was reaching for

- *"I want to ship one file."* `build.js --compile myscript.js` already produces
  one standalone `.js` — 65 KB for the example above, versus 93 KB base64'd,
  and it is readable and debuggable.
- *"I want it smaller."* `node tools/minify.mjs` halves it. Base64 then adds
  41% back.
- *"I want the source hidden."* Base64 does not do this. Nothing available in
  pure JScript does.

## Recommendation

Do not add a `--base64` mode. The wrapper works, and it is kept here as a
runnable spike, but it makes every artifact bigger, makes every error harder to
read, and solves nothing that `--compile` and the minifier do not already solve
better.

The same machinery *is* worth remembering for a different problem: embedding
**binary assets** — an icon, a small zip, a certificate — inside a distributable
`.js`, where there is no alternative to encoding the bytes as text. That is a
real use for `base64_encode_bytes` and the MSXML decoder. It is just not a use
for wrapping source code.

## Reproducing the numbers

```bat
cscript.exe build.js
node tools\minify.mjs

cscript.exe bin\launcher.js studies\make-base64-bundle.js --in dist\launcher.js
cscript.exe bin\launcher.js studies\make-base64-bundle.js --in dist\launcher.min.js

:: and then, to prove it runs:
cscript.exe dist\launcher.b64.js bin\tests\test-core.js
```
