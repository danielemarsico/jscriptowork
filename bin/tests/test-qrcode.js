// test-qrcode.js - Tests for libs/qrcode.js
//
// Needs nothing: no network, no desktop, no Office. That is the point of the
// lib - the examples used to fetch QR images from api.qrserver.com.
//
// The golden matrices below were produced by this implementation and then
// checked module-for-module against an independent QR encoder, and the
// resulting symbols were decoded back to their input text with an independent
// QR *decoder*. They are the regression anchor: if a refactor changes a single
// module, these fail.
//
// Non-ASCII characters are written as \u escapes so this file stays pure
// ASCII: the launcher reads scripts through FileSystemObject in ASCII mode,
// which would mangle literal multi-byte characters.
//
// Run via:  cscript.exe launcher.js test-qrcode.js

load("core");
load("polyfills");
load("qrcode");
load("minitest");

// Renders a QR object as an array of "0101..." strings, one per row.
function rows_of(qr) {
    var out = [];
    for (var r = 0; r < qr.size; r++) {
        var s = "";
        for (var c = 0; c < qr.size; c++) { s += qr.modules[r][c] ? "1" : "0"; }
        out.push(s);
    }
    return out;
}

// ---------------------------------------------------------------------------
// GF(256) and Reed-Solomon
// ---------------------------------------------------------------------------

describe("qrcode - GF(256) arithmetic", function() {

    it("log and exp are inverses over the whole field", function() {
        for (var x = 1; x < 256; x++) {
            assert.equal(_qr_exp[_qr_log[x]], x, "round-trip failed for " + x);
        }
    });

    it("a^0 is 1, and log(1) is 0", function() {
        assert.equal(_qr_exp[0], 1);
        assert.equal(_qr_log[1], 0);
    });

    it("a^255 wraps back to a^0 (the group has order 255, not 256)", function() {
        assert.equal(_qr_exp[255], _qr_exp[0]);
        assert.equal(_qr_exp[300], _qr_exp[45]);
    });

    it("multiplication matches the primitive polynomial 0x11D", function() {
        assert.equal(_qr_gf_mul(2, 2), 4);
        assert.equal(_qr_gf_mul(0x80, 2), 0x1D);   // overflow folds in 0x11D
        assert.equal(_qr_gf_mul(0, 123), 0);
        assert.equal(_qr_gf_mul(123, 0), 0);
        assert.equal(_qr_gf_mul(1, 123), 123);
    });

    it("multiplication is commutative and associative", function() {
        assert.equal(_qr_gf_mul(37, 91), _qr_gf_mul(91, 37));
        assert.equal(_qr_gf_mul(_qr_gf_mul(3, 5), 7), _qr_gf_mul(3, _qr_gf_mul(5, 7)));
    });

});

describe("qrcode - Reed-Solomon generator polynomials", function() {

    // The spec's generator polynomials, as exponents of a. These are the
    // published values for 7, 10 and 13 error-correction codewords.
    it("degree 7 matches the spec", function() {
        var got = [];
        var gen = _qr_rs_generator(7);
        for (var i = 0; i < gen.length; i++) { got.push(_qr_log[gen[i]]); }
        assert.deepEqual(got, [0, 87, 229, 146, 149, 238, 102, 21]);
    });

    it("degree 10 matches the spec", function() {
        var got = [];
        var gen = _qr_rs_generator(10);
        for (var i = 0; i < gen.length; i++) { got.push(_qr_log[gen[i]]); }
        assert.deepEqual(got, [0, 251, 67, 46, 61, 118, 70, 64, 94, 32, 45]);
    });

    it("degree 13 matches the spec", function() {
        var got = [];
        var gen = _qr_rs_generator(13);
        for (var i = 0; i < gen.length; i++) { got.push(_qr_log[gen[i]]); }
        assert.deepEqual(got, [0, 74, 152, 176, 100, 86, 100, 106, 104, 130, 218, 206, 140, 78]);
    });

    it("produces the right number of error-correction codewords", function() {
        assert.equal(_qr_rs_remainder([32, 33, 204, 57], 10).length, 10);
        assert.equal(_qr_rs_remainder([1, 2, 3], 17).length, 17);
    });

    it("computes the known EC codewords for a version 1-M block", function() {
        // Data codewords for "AAAA" at 1-M, and the EC block that follows them.
        var data = [32, 33, 204, 57, 128, 236, 17, 236, 17, 236,
                    17, 236, 17, 236, 17, 236];
        assert.deepEqual(_qr_rs_remainder(data, 10),
                         [137, 44, 92, 140, 203, 121, 201, 74, 147, 178]);
    });

});

// ---------------------------------------------------------------------------
// Mode detection and payload sizing
// ---------------------------------------------------------------------------

describe("qrcode - mode detection", function() {

    it("picks numeric for digits", function() {
        assert.equal(_qr_pick_mode("8675309"), "numeric");
    });

    it("picks alphanumeric for the 45-character set", function() {
        assert.equal(_qr_pick_mode("HELLO WORLD"), "alphanumeric");
        assert.equal(_qr_pick_mode("$%*+-./:"), "alphanumeric");
    });

    it("picks byte for anything else, lower case included", function() {
        assert.equal(_qr_pick_mode("hello"), "byte");
        assert.equal(_qr_pick_mode("https://example.com/"), "byte");
        assert.equal(_qr_pick_mode("caff\u00E8"), "byte");
    });

    it("picks byte for the empty string", function() {
        assert.equal(_qr_pick_mode(""), "byte");
    });

});

describe("qrcode - UTF-8 encoding", function() {

    it("encodes ASCII unchanged", function() {
        assert.deepEqual(_qr_utf8_bytes("Ab1"), [0x41, 0x62, 0x31]);
    });

    it("encodes a two-byte character", function() {
        assert.deepEqual(_qr_utf8_bytes("\u00E8"), [0xC3, 0xA8]);   // e-grave
    });

    it("encodes a three-byte character", function() {
        assert.deepEqual(_qr_utf8_bytes("\u65E5"), [0xE6, 0x97, 0xA5]);  // CJK "day"
    });

    it("encodes a surrogate pair as one four-byte character", function() {
        // U+1F600, the grinning face
        assert.deepEqual(_qr_utf8_bytes("\uD83D\uDE00"), [0xF0, 0x9F, 0x98, 0x80]);
    });

    it("returns nothing for the empty string", function() {
        assert.deepEqual(_qr_utf8_bytes(""), []);
    });

});

describe("qrcode - capacity", function() {

    it("derives the documented total codeword counts", function() {
        // The spec's total codewords per version, derived here from the
        // function-pattern layout rather than tabulated.
        assert.equal(_qr_total_codewords(1), 26);
        assert.equal(_qr_total_codewords(2), 44);
        assert.equal(_qr_total_codewords(7), 196);
        assert.equal(_qr_total_codewords(10), 346);
        assert.equal(_qr_total_codewords(40), 3706);
    });

    it("derives the documented data codeword counts", function() {
        assert.equal(_qr_data_codewords(1, "L"), 19);
        assert.equal(_qr_data_codewords(1, "M"), 16);
        assert.equal(_qr_data_codewords(1, "Q"), 13);
        assert.equal(_qr_data_codewords(1, "H"), 9);
        assert.equal(_qr_data_codewords(40, "L"), 2956);
        assert.equal(_qr_data_codewords(40, "H"), 1276);
    });

    it("sizes are 4 * version + 17", function() {
        assert.equal(_qr_size_for(1), 21);
        assert.equal(_qr_size_for(7), 45);
        assert.equal(_qr_size_for(40), 177);
    });

    it("picks the smallest version the data fits in", function() {
        assert.equal(qr_encode("A").version, 1);
        assert.equal(qr_encode("https://example.com/", { ec_level: "Q" }).version, 2);
    });

    it("a higher error correction level needs a bigger symbol", function() {
        var l = qr_encode("The quick brown fox jumps over the lazy dog", { ec_level: "L" });
        var h = qr_encode("The quick brown fox jumps over the lazy dog", { ec_level: "H" });
        assert.ok(h.version > l.version,
            "H should need a bigger version than L for the same text");
    });

    it("honours a version floor without shrinking below the data", function() {
        assert.equal(qr_encode("A", { version: 10 }).version, 10);
        // Too much data for version 1: the floor is a minimum, not a cap.
        assert.ok(qr_encode("x", { version: 1 }).version === 1);
        assert.ok(qr_encode(new Array(200).join("x"), { version: 1 }).version > 1);
    });

});

// ---------------------------------------------------------------------------
// Golden matrices
// ---------------------------------------------------------------------------

describe("qrcode - golden symbols", function() {

    it("encodes HELLO WORLD at level M exactly", function() {
        var qr = qr_encode("HELLO WORLD", { ec_level: "M" });
        assert.equal(qr.version, 1);
        assert.equal(qr.mode, "alphanumeric");
        assert.equal(qr.mask, 0);
        assert.deepEqual(rows_of(qr), [
            "111111100010101111111",
            "100000101110001000001",
            "101110100010101011101",
            "101110100010101011101",
            "101110101011101011101",
            "100000100111001000001",
            "111111101010101111111",
            "000000000000000000000",
            "101010100100100010010",
            "011110001001000010001",
            "000111111101001011000",
            "111101011001110101110",
            "010011110101001110101",
            "000000001010001000101",
            "111111100000100101100",
            "100000100110001101000",
            "101110101100101111111",
            "101110100011010100010",
            "101110101111011101001",
            "100000100001110001011",
            "111111101101011100001"
        ]);
    });

    it("encodes 12345678 at level L with mask 0 exactly", function() {
        var qr = qr_encode("12345678", { ec_level: "L", mask: 0 });
        assert.equal(qr.version, 1);
        assert.equal(qr.mode, "numeric");
        assert.deepEqual(rows_of(qr), [
            "111111100010101111111",
            "100000100000101000001",
            "101110101010001011101",
            "101110100000101011101",
            "101110100101101011101",
            "100000100111001000001",
            "111111101010101111111",
            "000000001010000000000",
            "111011111010111000100",
            "100001010111010100001",
            "111101100101011101000",
            "010011001011110111001",
            "101100110011011100011",
            "000000001010001001010",
            "111111101100100010001",
            "100000101000001000011",
            "101110101100101011001",
            "101110100101010101010",
            "101110101101011100101",
            "100000101011110111000",
            "111111101001011100101"
        ]);
    });

});

// ---------------------------------------------------------------------------
// Structure every symbol must have
// ---------------------------------------------------------------------------

describe("qrcode - structural invariants", function() {

    var cases = [
        { text: "A", opts: { ec_level: "L" } },
        { text: "https://example.com/", opts: { ec_level: "Q" } },
        { text: "1234567890123456789012345", opts: { ec_level: "H" } },
        { text: "caff\u00E8 latte", opts: { ec_level: "M" } },
        { text: "x", opts: { ec_level: "M", version: 7 } },   // version info blocks
        { text: "y", opts: { ec_level: "M", version: 14 } }   // more alignment patterns
    ];

    for (var i = 0; i < cases.length; i++) {
        (function(c, index) {

            var qr = qr_encode(c.text, c.opts);

            it("case " + index + ": matrix is square and the declared size", function() {
                assert.equal(qr.modules.length, qr.size);
                for (var r = 0; r < qr.size; r++) {
                    assert.equal(qr.modules[r].length, qr.size);
                }
                assert.equal(qr.size, qr.version * 4 + 17);
            });

            it("case " + index + ": has all three finder patterns", function() {
                var corners = [[0, 0], [0, qr.size - 7], [qr.size - 7, 0]];
                for (var k = 0; k < corners.length; k++) {
                    var top = corners[k][0], left = corners[k][1];
                    assert.ok(qr.modules[top][left], "finder corner is light");
                    assert.ok(qr.modules[top + 3][left + 3], "finder centre is light");
                    assert.notOk(qr.modules[top + 1][left + 1], "finder ring is dark");
                    assert.ok(qr.modules[top + 6][left + 6], "finder outer edge is light");
                }
            });

            it("case " + index + ": timing patterns alternate", function() {
                for (var p = 8; p < qr.size - 8; p++) {
                    assert.equal(qr.modules[6][p], p % 2 === 0, "row 6 at " + p);
                    assert.equal(qr.modules[p][6], p % 2 === 0, "column 6 at " + p);
                }
            });

            it("case " + index + ": the always-dark module is dark", function() {
                assert.ok(qr.modules[qr.size - 8][8]);
            });

            it("case " + index + ": reports a mask in range", function() {
                assert.ok(qr.mask >= 0 && qr.mask <= 7);
            });

        }(cases[i], i));
    }

    it("version 7 and up carries version information blocks", function() {
        // The version bits are BCH-encoded; check the encoding itself, since
        // the placement is covered by the golden matrices at lower versions.
        assert.equal(_qr_version_bits(7),  0x07C94);
        assert.equal(_qr_version_bits(21), 0x15683);
        assert.equal(_qr_version_bits(40), 0x28C69);
    });

    it("format information is BCH-encoded and masked with 0x5412", function() {
        // Published format strings: level M mask 0, and level L mask 0.
        assert.equal(_qr_format_bits("M", 0), 0x5412);
        assert.equal(_qr_format_bits("L", 0), 0x77C4);
        assert.equal(_qr_format_bits("H", 7), 0x083B);
    });

});

// ---------------------------------------------------------------------------
// Masking
// ---------------------------------------------------------------------------

describe("qrcode - masking", function() {

    it("each mask formula matches the spec", function() {
        assert.equal(_qr_mask_at(0, 0, 0), true);    // (i+j) % 2
        assert.equal(_qr_mask_at(0, 0, 1), false);
        assert.equal(_qr_mask_at(1, 2, 5), true);    // i % 2
        assert.equal(_qr_mask_at(1, 3, 5), false);
        assert.equal(_qr_mask_at(2, 5, 3), true);    // j % 3
        assert.equal(_qr_mask_at(2, 5, 4), false);
        assert.equal(_qr_mask_at(3, 1, 2), true);    // (i+j) % 3
        assert.equal(_qr_mask_at(4, 0, 0), true);
        assert.equal(_qr_mask_at(5, 0, 0), true);
        assert.equal(_qr_mask_at(6, 0, 0), true);
        assert.equal(_qr_mask_at(7, 0, 0), true);
        assert.equal(_qr_mask_at(7, 0, 1), false);
    });

    it("forcing a mask is honoured", function() {
        for (var m = 0; m < 8; m++) {
            assert.equal(qr_encode("mask test", { mask: m }).mask, m);
        }
    });

    it("different masks produce different symbols", function() {
        var a = rows_of(qr_encode("mask test", { mask: 0 })).join("");
        var b = rows_of(qr_encode("mask test", { mask: 5 })).join("");
        assert.notEqual(a, b);
    });

    it("the automatic choice is the lowest-penalty mask", function() {
        var qr = qr_encode("penalty check");
        var chosen = _qr_penalty(qr.modules);
        for (var m = 0; m < 8; m++) {
            var other = _qr_penalty(qr_encode("penalty check", { mask: m }).modules);
            assert.ok(chosen <= other,
                "mask " + qr.mask + " scored " + chosen + " but mask " + m + " scored " + other);
        }
    });

    it("penalty scoring rises for an obviously bad symbol", function() {
        // An all-dark grid trips rules 1, 2 and 4 hard.
        var grid = _qr_new_grid(21, true);
        assert.ok(_qr_penalty(grid) > 1000);
    });

});

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

describe("qrcode - error handling", function() {

    it("rejects an unknown error correction level", function() {
        assert.throws(function() { qr_encode("x", { ec_level: "Z" }); });
    });

    it("rejects an unknown mode", function() {
        assert.throws(function() { qr_encode("x", { mode: "kanji" }); });
    });

    it("rejects a mode the text does not fit", function() {
        assert.throws(function() { qr_encode("abc", { mode: "numeric" }); });
        assert.throws(function() { qr_encode("abc", { mode: "alphanumeric" }); });
    });

    it("rejects an out-of-range version or mask", function() {
        assert.throws(function() { qr_encode("x", { version: 0 }); });
        assert.throws(function() { qr_encode("x", { version: 41 }); });
        assert.throws(function() { qr_encode("x", { mask: 8 }); });
        assert.throws(function() { qr_encode("x", { mask: -1 }); });
    });

    it("rejects text too long for any symbol", function() {
        var huge = new Array(3000).join("abcdefghij");
        assert.throws(function() { qr_encode(huge, { ec_level: "H" }); });
    });

    it("accepts the empty string and both levels of nothing", function() {
        assert.doesNotThrow(function() { qr_encode(""); });
        assert.equal(qr_encode("").version, 1);
    });

    it("takes the widest payload each level allows at version 40", function() {
        // Byte mode capacity at 40-L is 2953 bytes.
        var text = new Array(2954).join("a");   // 2953 characters
        assert.doesNotThrow(function() { qr_encode(text, { ec_level: "L" }); });
        assert.equal(qr_encode(text, { ec_level: "L" }).version, 40);
        assert.throws(function() { qr_encode(text + "a", { ec_level: "L" }); });
    });

});

// ---------------------------------------------------------------------------
// Renderers
// ---------------------------------------------------------------------------

describe("qrcode - renderers", function() {

    var qr = qr_encode("HELLO WORLD", { ec_level: "M" });

    it("qr_to_matrix gives 0/1 rather than booleans", function() {
        var m = qr_to_matrix(qr);
        assert.equal(m.length, qr.size);
        assert.equal(m[0].length, qr.size);
        assert.equal(m[0][0], 1);      // finder corner
        assert.equal(m[7][7], 0);      // separator
    });

    it("qr_to_ascii lays out one line per row plus the quiet zone", function() {
        var lines = qr_to_ascii(qr, { quiet_zone: 2 }).split("\n");
        assert.equal(lines.length, qr.size + 4);
        // Two characters per module, quiet zone included.
        assert.equal(lines[0].length, (qr.size + 4) * 2);
    });

    it("qr_to_ascii honours custom characters", function() {
        var art = qr_to_ascii(qr, { dark: "X", light: ".", quiet_zone: 0 });
        var lines = art.split("\n");
        assert.equal(lines.length, qr.size);
        assert.equal(lines[0].length, qr.size);
        assert.equal(lines[0].charAt(0), "X");
        assert.equal(lines[7].charAt(7), ".");
    });

    it("qr_to_html emits one cell per module", function() {
        var html = qr_to_html(qr, { scale: 3 });
        assert.ok(html.indexOf("<table") === 0);
        assert.ok(html.indexOf("</table>") !== -1);

        var cells = html.split("<td").length - 1;
        assert.equal(cells, qr.size * qr.size);

        var trs = html.split("<tr").length - 1;
        assert.equal(trs, qr.size);
    });

    it("qr_to_html uses the colours it is given", function() {
        var html = qr_to_html(qr, { dark: "#123456", light: "#abcdef" });
        assert.ok(html.indexOf("#123456") !== -1);
        assert.ok(html.indexOf("#abcdef") !== -1);
    });

    it("qr_to_svg emits one rect per dark module plus the background", function() {
        var svg = qr_to_svg(qr, { scale: 2, quiet_zone: 1 });
        assert.ok(svg.indexOf("<svg") !== -1);
        assert.ok(svg.indexOf("</svg>") !== -1);

        var dark = 0;
        for (var r = 0; r < qr.size; r++) {
            for (var c = 0; c < qr.size; c++) { if (qr.modules[r][c]) { dark++; } }
        }
        assert.equal(svg.split("<rect").length - 1, dark + 1);

        var side = (qr.size + 2) * 2;
        assert.ok(svg.indexOf('width="' + side + '"') !== -1);
    });

    it("renderers do not mutate the symbol", function() {
        var before = rows_of(qr).join("");
        qr_to_ascii(qr);
        qr_to_html(qr);
        qr_to_svg(qr);
        qr_to_matrix(qr);
        assert.equal(rows_of(qr).join(""), before);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
