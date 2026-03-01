// test-crypto.js - Tests for sha256, sha256_bytes, hmac_sha256 in libs/crypto.js
//
// Known-answer test vectors:
//   SHA-256:    NIST FIPS 180-4, Appendix B (https://doi.org/10.6028/NIST.FIPS.180-4)
//   HMAC-SHA-256: RFC 4231, Section 4 (https://www.rfc-editor.org/rfc/rfc4231)
//
// Run via:  cscript.exe launcher.js test-crypto.js

load("core");
load("polyfills");
load("system");
load("crypto");
load("minitest");

// ---------------------------------------------------------------------------
// sha256 — structural
// ---------------------------------------------------------------------------

describe("sha256 - structural", function() {

    it("returns a string", function() {
        assert.equal(typeof sha256(""), "string");
    });

    it("returns exactly 64 characters", function() {
        assert.equal(sha256("").length,   64);
        assert.equal(sha256("abc").length, 64);
    });

    it("output contains only lowercase hex characters", function() {
        assert.ok(/^[0-9a-f]{64}$/.test(sha256("hello world")));
    });

    it("is deterministic — same input gives same output", function() {
        assert.equal(sha256("test"), sha256("test"));
    });

    it("different inputs give different outputs", function() {
        assert.notEqual(sha256("abc"), sha256("ABC"));
        assert.notEqual(sha256("a"),   sha256("b"));
    });

    it("single-character change produces a completely different hash", function() {
        var h1 = sha256("hello");
        var h2 = sha256("hellp");
        assert.notEqual(h1, h2);
    });

});

// ---------------------------------------------------------------------------
// sha256 — NIST FIPS 180-4 known-answer vectors
// ---------------------------------------------------------------------------

describe("sha256 - NIST FIPS 180-4 test vectors", function() {

    // Empty string (widely published, consistent with NIST zero-length test)
    it("empty string", function() {
        assert.equal(sha256(""),
            "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
    });

    // NIST FIPS 180-4 Appendix B.1 — single 512-bit block
    it("\"abc\" (NIST B.1 — one block)", function() {
        assert.equal(sha256("abc"),
            "ba7816bf8f01cfea414140de5dae2ec73b00361bbef0469f11fd94e5df54f6b8");
    });

    // NIST FIPS 180-4 Appendix B.2 — two 512-bit blocks (56-byte input)
    it("\"abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq\" (NIST B.2 — two blocks)", function() {
        assert.equal(
            sha256("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"),
            "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
        );
    });

});

// ---------------------------------------------------------------------------
// sha256_bytes
// ---------------------------------------------------------------------------

describe("sha256_bytes", function() {

    it("returns exactly 64 hex characters", function() {
        assert.equal(sha256_bytes([]).length,          64);
        assert.equal(sha256_bytes([1, 2, 3]).length,   64);
    });

    it("sha256_bytes([97,98,99]) equals sha256(\"abc\")", function() {
        // 97=a  98=b  99=c  in ASCII — which is also the UTF-8 encoding
        assert.equal(sha256_bytes([97, 98, 99]), sha256("abc"));
    });

    it("sha256_bytes([]) equals sha256(\"\")", function() {
        assert.equal(sha256_bytes([]), sha256(""));
    });

    it("does not mutate the input array", function() {
        var input = [1, 2, 3];
        sha256_bytes(input);
        assert.equal(input.length, 3);
        assert.equal(input[0], 1);
        assert.equal(input[2], 3);
    });

    it("all-zero bytes of varying lengths produce different hashes", function() {
        assert.notEqual(sha256_bytes([0]),       sha256_bytes([0, 0]));
        assert.notEqual(sha256_bytes([0, 0, 0]), sha256_bytes([0, 0]));
    });

});

// ---------------------------------------------------------------------------
// hmac_sha256 — structural
// ---------------------------------------------------------------------------

describe("hmac_sha256 - structural", function() {

    it("returns exactly 64 hex characters", function() {
        assert.equal(hmac_sha256("key", "message").length, 64);
    });

    it("output is lowercase hex", function() {
        assert.ok(/^[0-9a-f]{64}$/.test(hmac_sha256("k", "m")));
    });

    it("is deterministic", function() {
        assert.equal(hmac_sha256("key", "msg"), hmac_sha256("key", "msg"));
    });

    it("different messages produce different MACs", function() {
        assert.notEqual(hmac_sha256("key", "msg1"), hmac_sha256("key", "msg2"));
    });

    it("different keys produce different MACs", function() {
        assert.notEqual(hmac_sha256("key1", "msg"), hmac_sha256("key2", "msg"));
    });

    it("result differs from plain sha256", function() {
        assert.notEqual(hmac_sha256("key", "msg"), sha256("msg"));
    });

});

// ---------------------------------------------------------------------------
// hmac_sha256 — RFC 4231 known-answer vectors
// ---------------------------------------------------------------------------

describe("hmac_sha256 - RFC 4231 known-answer vectors", function() {

    // RFC 4231 Test Case 2 — printable ASCII key and data
    //   Key  = "Jefe"
    //   Data = "what do ya want for nothing?"
    //   HMAC = 5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964a37827
    it("RFC 4231 test case 2 (key=\"Jefe\")", function() {
        assert.equal(
            hmac_sha256("Jefe", "what do ya want for nothing?"),
            "5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964a37827"
        );
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
