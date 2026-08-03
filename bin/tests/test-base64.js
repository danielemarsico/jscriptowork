// test-base64.js - Tests for libs/base64.js
//
// Known-answer vectors: RFC 4648, Section 10 (https://www.rfc-editor.org/rfc/rfc4648)
//
// Run via:  cscript.exe launcher.js test-base64.js

load("core");
load("polyfills");
load("base64");
load("minitest");

// ---------------------------------------------------------------------------
// base64_encode / base64_decode — RFC 4648 known-answer vectors
// ---------------------------------------------------------------------------

describe("base64_encode/decode - RFC 4648 test vectors", function() {

    var vectors = [
        ["",       ""],
        ["f",      "Zg=="],
        ["fo",     "Zm8="],
        ["foo",    "Zm9v"],
        ["foob",   "Zm9vYg=="],
        ["fooba",  "Zm9vYmE="],
        ["foobar", "Zm9vYmFy"]
    ];

    vectors.forEach(function(v) {
        it("encodes " + JSON.stringify(v[0]) + " as " + JSON.stringify(v[1]), function() {
            assert.equal(base64_encode(v[0]), v[1]);
        });

        it("decodes " + JSON.stringify(v[1]) + " back to " + JSON.stringify(v[0]), function() {
            assert.equal(base64_decode(v[1]), v[0]);
        });
    });

});

// ---------------------------------------------------------------------------
// base64_encode / base64_decode — general behaviour
// ---------------------------------------------------------------------------

describe("base64_encode/decode", function() {

    it("returns a string", function() {
        assert.equal(typeof base64_encode("hello"), "string");
    });

    it("round-trips arbitrary ASCII text", function() {
        var src = "The quick brown fox jumps over the lazy dog. 0123456789!@#$%^&*()";
        assert.equal(base64_decode(base64_encode(src)), src);
    });

    it("output only contains base64 alphabet characters and padding", function() {
        assert.ok(/^[A-Za-z0-9+\/]*={0,2}$/.test(base64_encode("foobar")));
    });

    it("output length is always a multiple of 4", function() {
        for (var i = 0; i < 8; i++) {
            var s = new Array(i + 1).join("x");
            assert.equal(base64_encode(s).length % 4, 0, "length for input of " + i + " chars");
        }
    });

    it("round-trips a 2-byte UTF-8 character (café)", function() {
        var src = "caf" + String.fromCharCode(0xE9);
        assert.equal(base64_decode(base64_encode(src)), src);
    });

    it("round-trips 3-byte UTF-8 characters (CJK)", function() {
        var src = String.fromCharCode(0x4E2D) + String.fromCharCode(0x6587);
        assert.equal(base64_decode(base64_encode(src)), src);
    });

    it("round-trips a surrogate pair (emoji, 4-byte UTF-8)", function() {
        var src = String.fromCharCode(0xD83D, 0xDE00); // U+1F600
        assert.equal(base64_decode(base64_encode(src)), src);
    });

    it("ignores embedded whitespace/newlines when decoding", function() {
        assert.equal(base64_decode("Zm9v\r\nYmFy"), "foobar");
    });

});

// ---------------------------------------------------------------------------
// base64_encode_bytes / base64_decode_bytes
// ---------------------------------------------------------------------------

describe("base64_encode_bytes/decode_bytes", function() {

    it("agrees with base64_encode on ASCII bytes", function() {
        var bytes = [102, 111, 111, 98, 97, 114]; // "foobar"
        assert.equal(base64_encode_bytes(bytes), base64_encode("foobar"));
    });

    it("round-trips every byte value 0x00-0xFF", function() {
        var bytes = [];
        for (var i = 0; i < 256; i++) { bytes.push(i); }
        assert.deepEqual(base64_decode_bytes(base64_encode_bytes(bytes)), bytes);
    });

    it("round-trips an empty byte array", function() {
        assert.deepEqual(base64_decode_bytes(base64_encode_bytes([])), []);
    });

    it("does not mutate the input array", function() {
        var bytes = [1, 2, 3];
        base64_encode_bytes(bytes);
        assert.deepEqual(bytes, [1, 2, 3]);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
