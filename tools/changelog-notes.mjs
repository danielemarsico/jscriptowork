// tools/changelog-notes.mjs - pull one version's section out of CHANGELOG.md.
//
//   node tools/changelog-notes.mjs v1.2.3                  (prints to stdout)
//   node tools/changelog-notes.mjs 1.2.3 --out notes.md
//
// Used by .github/workflows/release.yml to turn the CHANGELOG into GitHub
// release notes. It exits non-zero when the version has no section, which is
// what keeps the release convention honest: promote `Unreleased` to
// `## [X.Y.Z] - YYYY-MM-DD` before tagging, or the release fails loudly instead
// of shipping with empty notes.
//
// Maintainer-only tooling, like everything else in tools/ - nothing here is
// needed to run jscriptowork.

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const USAGE = [
    "Usage: node tools/changelog-notes.mjs <version> [--out <file>] [--file <changelog>]",
    "",
    "  <version>  v1.2.3 or 1.2.3 - the leading 'v' is optional",
    "  --out      write to a file instead of stdout",
    "  --file     changelog to read (default: CHANGELOG.md)",
].join("\n");

// Heading shapes accepted, all equivalent:
//   ## [1.2.3] - 2026-01-01
//   ## [1.2.3]
//   ## 1.2.3 - 2026-01-01
//   ## 1.2.3
function headingRe(version) {
    const v = version.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp("^##\\s+\\[?" + v + "\\]?(\\s|$)");
}

export function extractNotes(changelog, version) {
    const clean = String(version).replace(/^v/, "");
    const lines = changelog.split(/\r?\n/);
    const wanted = headingRe(clean);

    let start = -1;
    for (let i = 0; i < lines.length; i++) {
        if (wanted.test(lines[i])) { start = i + 1; break; }
    }
    if (start < 0) { return null; }

    let end = lines.length;
    for (let i = start; i < lines.length; i++) {
        if (/^##\s/.test(lines[i])) { end = i; break; }
    }

    // Drop a trailing horizontal rule: the CHANGELOG separates sections with
    // `---`, which reads as a stray divider at the bottom of release notes.
    const body = lines.slice(start, end);
    while (body.length && (body[body.length - 1].trim() === "" || body[body.length - 1].trim() === "---")) {
        body.pop();
    }
    while (body.length && body[0].trim() === "") { body.shift(); }

    return body.join("\n");
}

function main(argv) {
    const args = { version: null, out: null, file: "CHANGELOG.md" };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === "--out" || a === "-o") { args.out = argv[++i]; }
        else if (a === "--file" || a === "-f") { args.file = argv[++i]; }
        else if (a === "--help" || a === "-h") { console.log(USAGE); return 0; }
        else if (args.version === null) { args.version = a; }
        else { console.error("changelog-notes: unexpected argument: " + a); return 1; }
    }

    if (!args.version) { console.error(USAGE); return 1; }

    const path = resolve(ROOT, args.file);
    if (!existsSync(path)) {
        console.error("changelog-notes: no such file: " + path);
        return 1;
    }

    const version = args.version.replace(/^v/, "");
    const notes = extractNotes(readFileSync(path, "utf8"), version);

    if (notes === null) {
        const today = new Date().toISOString().slice(0, 10);
        console.error("changelog-notes: " + args.file + " has no section for " + version + ".");
        console.error("Promote the Unreleased section to '## [" + version + "] - " + today + "' before tagging.");
        return 1;
    }
    if (notes.trim() === "") {
        console.error("changelog-notes: the section for " + version + " is empty.");
        return 1;
    }

    if (args.out) {
        writeFileSync(resolve(ROOT, args.out), notes + "\n", "utf8");
        console.log("changelog-notes: wrote " + notes.split("\n").length + " lines to " + args.out);
    } else {
        process.stdout.write(notes + "\n");
    }
    return 0;
}

// Only run when invoked directly, so the extractor can be imported by a test.
if (process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url))) {
    process.exit(main(process.argv.slice(2)));
}
