// tools/minify.mjs - optional, maintainer-only minification of the bundle.
//
//   cd tools && npm install
//   node tools/minify.mjs                     (from the project root)
//   node tools/minify.mjs --in dist/launcher.js --out dist/launcher.min.js
//
// This is the ONE place in the project where Node.js and npm are allowed, and
// only at *build* time. The artifact it produces is still plain JScript, so a
// target machine continues to need nothing but Windows. `build.js` (run under
// cscript.exe) stays the canonical bundler; `build.bat` never calls this.
//
// Why the terser settings below are what they are - every one of them exists
// because JScript 5.8 is an ES3 engine and because dist/launcher.js is loaded
// through eval():
//
//   ecma: 5                  never emit ES6 (arrows, shorthand, `let`) - all of
//                            it is a *parse* error in JScript, which takes down
//                            the whole file, not just the offending line.
//   mangle.toplevel: false   the bundle's public API is its top-level names.
//                            `foo = function(){}` globals and `function foo(){}`
//                            declarations are what user scripts call after the
//                            launcher eval()s them; renaming any of them breaks
//                            every caller. Locals inside functions are fair game.
//   compress.properties:false  stops `obj["default"]` collapsing to `obj.default`.
//                            ES5 allows reserved words as property names; ES3 -
//                            and so JScript - does not.
//   format.quote_keys: true  same hazard from the other direction: keeps
//                            `{ "default": 1 }` quoted instead of unquoting a key
//                            that ES3 would reject.
//   format.ascii_only: true  the bundle is read by cscript.exe with the system
//                            code page; escaping non-ASCII avoids an encoding
//                            round-trip deciding what a byte means.
//
// After minifying it verifies two things and fails loudly on either:
//   1. every public top-level name found in the input still appears in the output
//   2. the output parses as ES5 (acorn, ecmaVersion 5) - a smoke test for
//      "did terser emit modern syntax anyway"
//
// Neither check is a substitute for running the minified bundle under
// cscript.exe; CI does that on windows-latest.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { minify } from "terser";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

function parseArgs(argv) {
    const out = { in: null, out: null, quiet: false };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === "--in" || a === "-i") { out.in = argv[++i]; }
        else if (a === "--out" || a === "-o") { out.out = argv[++i]; }
        else if (a === "--quiet" || a === "-q") { out.quiet = true; }
        else if (a === "--help" || a === "-h") { out.help = true; }
        else { throw new Error("unknown argument: " + a); }
    }
    return out;
}

const USAGE = [
    "Usage: node tools/minify.mjs [--in <file>] [--out <file>] [--quiet]",
    "",
    "  --in   source bundle   (default: dist/launcher.js)",
    "  --out  minified output (default: <in> with .min.js)",
].join("\n");

// Public names the bundle exposes at top level. Two shapes matter:
//   function foo(...)   - a declaration; global because dist/ inlines at top level
//   foo = function(...) - a bare assignment; the shape libs/ must use, because
//                         bin/launcher.js eval()s them inside load()'s scope
// `var foo = ...` counts too: in dist/ those are genuine globals (CURRENT_PATH,
// ROOT_FOLDER, _jsw_hta_inline_libs, ...) that scripts and ui.js read.
function publicNames(source) {
    const names = new Set();
    const lines = source.split(/\r?\n/);
    const patterns = [
        /^function\s+([A-Za-z_$][\w$]*)\s*\(/,
        /^([A-Za-z_$][\w$]*)\s*=\s*[^=]/,
        /^var\s+([A-Za-z_$][\w$]*)\s*=/,
    ];
    for (const line of lines) {
        for (const re of patterns) {
            const m = re.exec(line);
            if (m) { names.add(m[1]); break; }
        }
    }
    return names;
}

function missingNames(names, minified) {
    const missing = [];
    for (const name of names) {
        const re = new RegExp("(^|[^\\w$.])" + name.replace(/\$/g, "\\$") + "($|[^\\w$])");
        if (!re.test(minified)) { missing.push(name); }
    }
    return missing.sort();
}

// Optional: acorn is a dev convenience, not a hard requirement. If it is not
// installed the ES5 parse check is reported as skipped rather than failing.
async function checkEs5(code) {
    let acorn;
    try {
        acorn = await import("acorn");
    } catch (e) {
        return { skipped: true };
    }
    try {
        acorn.parse(code, { ecmaVersion: 5, allowReturnOutsideFunction: true });
        return { ok: true };
    } catch (e) {
        return { ok: false, error: e.message };
    }
}

async function main() {
    const args = parseArgs(process.argv.slice(2));
    if (args.help) { console.log(USAGE); return 0; }

    const inPath = resolve(ROOT, args.in || "dist/launcher.js");
    const outPath = resolve(ROOT, args.out || inPath.replace(/\.js$/, ".min.js"));

    if (!existsSync(inPath)) {
        console.error("minify: no such file: " + inPath);
        console.error("Build the bundle first:  cscript.exe build.js   (on Windows)");
        return 1;
    }
    if (outPath === inPath) {
        console.error("minify: refusing to overwrite the source bundle in place");
        return 1;
    }

    const source = readFileSync(inPath, "utf8");

    const result = await minify(source, {
        ecma: 5,
        toplevel: false,
        compress: {
            ecma: 5,
            properties: false,
            // `dist/launcher.js` ends in an IIFE that eval()s the user script.
            // terser already refuses to rename anything reachable from a direct
            // eval; keeping unused code is the belt to that braces, since a user
            // script can reference a helper that looks unreferenced from here.
            unused: false,
            toplevel: false,
        },
        mangle: {
            toplevel: false,
            eval: false,
            reserved: ["WScript", "ActiveXObject", "Enumerator", "GetObject"],
        },
        format: {
            ecma: 5,
            comments: false,
            quote_keys: true,
            ascii_only: true,
            preamble: "// jscriptowork bundled launcher (minified by tools/minify.mjs).\n" +
                      "// Generated - do not edit. Source: dist/launcher.js, itself built from libs/.",
        },
    });

    if (result.error) { throw result.error; }
    const code = result.code;

    const missing = missingNames(publicNames(source), code);
    if (missing.length) {
        console.error("minify: FAILED - these public top-level names vanished from the output:");
        console.error("  " + missing.join(", "));
        console.error("The bundle's globals are its API; a minified bundle that drops one is broken.");
        return 1;
    }

    const es5 = await checkEs5(code);
    if (es5.ok === false) {
        console.error("minify: FAILED - output does not parse as ES5: " + es5.error);
        console.error("JScript would reject the whole file at parse time.");
        return 1;
    }

    writeFileSync(outPath, code, "utf8");

    if (!args.quiet) {
        const before = Buffer.byteLength(source, "utf8");
        const after = Buffer.byteLength(code, "utf8");
        const pct = (100 - (after / before) * 100).toFixed(1);
        console.log("minify: " + inPath);
        console.log("     -> " + outPath);
        console.log("   size: " + before.toLocaleString() + " -> " + after.toLocaleString() +
                    " bytes (-" + pct + "%)");
        console.log("  names: " + publicNames(source).size + " public top-level names preserved");
        console.log("    es5: " + (es5.skipped ? "check skipped (acorn not installed)" : "output parses as ES5"));
    }
    return 0;
}

main().then(function (code) { process.exit(code); }, function (err) {
    console.error("minify: " + (err && err.stack ? err.stack : err));
    process.exit(1);
});
