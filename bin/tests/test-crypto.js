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
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
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
// sha256 — message-padding boundaries
//
// SHA-256 pads to a multiple of 64 bytes and reserves the last 8 for the bit
// length, so 55/56/57 and 63/64/65 bytes are where an off-by-one in the padding
// code shows up.
// ---------------------------------------------------------------------------

describe("sha256 - padding boundaries", function() {

    // Repeated-character helper: a(3) === "aaa".
    function a(n) { return new Array(n + 1).join("a"); }

    it("55 bytes (largest input that still fits one block)", function() {
        assert.equal(sha256(a(55)),
            "9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318");
    });

    it("56 bytes (padding spills into a second block)", function() {
        assert.equal(sha256(a(56)),
            "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a");
    });

    it("57 bytes", function() {
        assert.equal(sha256(a(57)),
            "f13b2d724659eb3bf47f2dd6af1accc87b81f09f59f2b75e5c0bed6589dfe8c6");
    });

    it("63 bytes", function() {
        assert.equal(sha256(a(63)),
            "7d3e74a05d7db15bce4ad9ec0658ea98e3f06eeecf16b4c6fff2da457ddc2f34");
    });

    it("64 bytes (exactly one block)", function() {
        assert.equal(sha256(a(64)),
            "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb");
    });

    it("65 bytes", function() {
        assert.equal(sha256(a(65)),
            "635361c48bb9eab14198e76ea8ab7f1a41685d6ad62aa9146d301d4f17eb0ae0");
    });

    it("1000 bytes (many blocks)", function() {
        assert.equal(sha256(a(1000)),
            "41edece42d63e8d9bf515a9ba6932e1c20cbc9f5a5d134645adb5db1b9737ea3");
    });

});

// ---------------------------------------------------------------------------
// sha256 — UTF-8 encoding
//
// Non-ASCII characters are written as \u escapes so this file stays pure ASCII:
// the launcher reads scripts through FileSystemObject in ASCII mode, which
// would mangle literal multi-byte characters.
// ---------------------------------------------------------------------------

describe("sha256 - UTF-8 encoding", function() {

    it("encodes a 2-byte character (U+00E9 in \"café\")", function() {
        assert.equal(sha256("caf\u00E9"),
            "850f7dc43910ff890f8879c0ed26fe697c93a067ad93a7d50f466a7028a9bf4e");
    });

    it("encodes 3-byte characters (CJK)", function() {
        assert.equal(sha256("\u65E5\u672C\u8A9E"),
            "77710aedc74ecfa33685e33a6c7df5cc83004da1bdcef7fb280f5c2b2e97e0a5");
    });

    it("encodes a surrogate pair as 4 bytes (U+1F600)", function() {
        assert.equal(sha256("\uD83D\uDE00"),
            "f0443a342c5ef54783a111b51ba56c938e474c32324d90c3a60c9c8e3a37e2d9");
    });

    it("a non-ASCII string hashes differently from its ASCII prefix", function() {
        assert.notEqual(sha256("caf\u00E9"), sha256("caf"));
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

    it("hashes the high byte 0xFF", function() {
        assert.equal(sha256_bytes([255]),
            "a8100ae6aa1940d0b663bb31cd466142ebbdbd5187131b92d93818987832eb89");
    });

    it("hashes every byte value 0x00-0xFF (four blocks)", function() {
        var bytes = [];
        for (var i = 0; i < 256; i++) { bytes.push(i); }
        assert.equal(sha256_bytes(bytes),
            "40aff2e9d2d8922e47afd4648e6967497158785fbd1da870e7110266bf944880");
    });

    it("agrees with sha256() on hand-encoded UTF-8 bytes", function() {
        // "caf\u00E9" encodes to 63 61 66 C3 A9
        assert.equal(sha256_bytes([0x63, 0x61, 0x66, 0xC3, 0xA9]), sha256("caf\u00E9"));
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

    // RFC 4231 Test Case 1
    //   Key  = 0x0b repeated 20 times
    //   Data = "Hi There"
    // U+000B encodes to the single byte 0x0b in UTF-8, so the binary key is
    // expressible through the string API.
    it("RFC 4231 test case 1 (20-byte 0x0b key)", function() {
        var key = "";
        for (var i = 0; i < 20; i++) { key += String.fromCharCode(0x0b); }
        assert.equal(
            hmac_sha256(key, "Hi There"),
            "b0344c61d8db38535ca8afceaf0bf12b881dc200c9833da726e9376c2e32cff7"
        );
    });

    // RFC 4231 Test Case 2 — printable ASCII key and data
    //   Key  = "Jefe"
    //   Data = "what do ya want for nothing?"
    it("RFC 4231 test case 2 (key=\"Jefe\")", function() {
        assert.equal(
            hmac_sha256("Jefe", "what do ya want for nothing?"),
            "5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843"
        );
    });

    // RFC 4231 test cases 3, 4, 6 and 7 use keys made of bytes >= 0x80
    // (0xaa repeated). Those are not reachable through a UTF-8 string API —
    // U+00AA encodes to two bytes (0xC2 0xAA), not one — so they are skipped.
    skip("RFC 4231 test case 3 (20-byte 0xaa key)",
         "a raw 0xaa key byte cannot be expressed as a UTF-8 string; needs an hmac_sha256_bytes() overload");
    skip("RFC 4231 test case 6 (131-byte 0xaa key)",
         "same reason; the oversized-key path is covered by the ASCII cases below");

});

// ---------------------------------------------------------------------------
// hmac_sha256 — key-length handling
//
// HMAC's block size is 64 bytes: shorter keys are zero-padded, longer keys are
// replaced by their own SHA-256 digest. These vectors pin down both paths and
// the exact boundary.
// ---------------------------------------------------------------------------

describe("hmac_sha256 - key lengths", function() {

    function A(n) { return new Array(n + 1).join("A"); }

    it("empty key and empty message", function() {
        assert.equal(hmac_sha256("", ""),
            "b613679a0814d9ec772f95d778c35fc5ff1697c493715653c6c712144292c5ad");
    });

    it("64-byte key (exactly the block size, no hashing)", function() {
        assert.equal(hmac_sha256(A(64), "message"),
            "9d49949e1b0538128913db2e5ddec810197d4910b260f23b420fe2f1612e8d7f");
    });

    it("65-byte key (one over the block size, key gets hashed)", function() {
        assert.equal(hmac_sha256(A(65), "message"),
            "6cf573c19fdb91491350a76e7dde7c0490f2baab0959a543e87ccdc6d61291e4");
    });

    it("100-byte key", function() {
        assert.equal(hmac_sha256(A(100), "message"),
            "2df7a809aa6407a40f0aec209b77c2d81d3f7a55f8e200dc68f84a6c373dcb54");
    });

    it("keys either side of the block size give different MACs", function() {
        assert.notEqual(hmac_sha256(A(64), "message"), hmac_sha256(A(65), "message"));
    });

    it("handles a UTF-8 key and message", function() {
        assert.equal(hmac_sha256("cl\u00E9", "messag\u00E9"),
            "dca42cb126a735fd68436f3c76a4006501cde2569648b32a97e964979fb7718b");
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
