// crypto.js - Cryptographic hash functions for JScript / CScript
// ---------------------------------------------------------------------------
//   sha256(message)            SHA-256 of a UTF-8 string  →  64-char hex
//   sha256_bytes(byteArray)    SHA-256 of a byte array (integers 0-255)  →  64-char hex
//   hmac_sha256(key, message)  HMAC-SHA-256 of two UTF-8 strings  →  64-char hex
//
// Pure ES3 syntax — no arrow functions, let/const, template literals, etc.
// All public functions are assigned without 'var' so they survive load()'s eval scope.
//
// NOTE: No IIFE wrapper — JScript's eval() does not preserve closure scope for
// function declarations inside IIFEs, causing inner helpers to silently return
// undefined.  Instead, private helpers use the _sha256_ prefix as a namespace.
// sha256_bytes is fully self-contained (K constants and helpers defined locally).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Private helpers (_sha256_ prefix)
// ---------------------------------------------------------------------------

// Encode a JS string to a UTF-8 byte array (handles BMP + surrogate pairs).
_sha256_to_utf8 = function(str) {
    var bytes = [], i = 0;
    while (i < str.length) {
        var c = str.charCodeAt(i++);
        if (c >= 0xD800 && c <= 0xDBFF && i < str.length) {
            var lo = str.charCodeAt(i);
            if (lo >= 0xDC00 && lo <= 0xDFFF) { c = ((c - 0xD800) << 10) + (lo - 0xDC00) + 0x10000; i++; }
        }
        if (c < 0x80) {
            bytes.push(c);
        } else if (c < 0x800) {
            bytes.push((c >> 6) | 0xC0);
            bytes.push((c & 0x3F) | 0x80);
        } else if (c < 0x10000) {
            bytes.push((c >> 12) | 0xE0);
            bytes.push(((c >> 6) & 0x3F) | 0x80);
            bytes.push((c & 0x3F) | 0x80);
        } else {
            bytes.push((c >> 18) | 0xF0);
            bytes.push(((c >> 12) & 0x3F) | 0x80);
            bytes.push(((c >> 6) & 0x3F) | 0x80);
            bytes.push((c & 0x3F) | 0x80);
        }
    }
    return bytes;
};

// Convert a 64-char hex string to a 32-element byte array.
_sha256_from_hex = function(hex) {
    var bytes = [];
    for (var i = 0; i < hex.length; i += 2) {
        bytes.push(parseInt(hex.substr(i, 2), 16));
    }
    return bytes;
};

// ---------------------------------------------------------------------------
// sha256_bytes — fully self-contained (K, add32, rotr32 defined locally)
// ---------------------------------------------------------------------------

sha256_bytes = function(input) {
    // Round constants: first 32 bits of cbrt(prime[0..63])
    var K = [
        0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
        0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
        0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
        0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
        0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
        0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
        0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
        0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
    ];

    // 32-bit helpers defined as var expressions (safe in JScript eval scope)
    var add32 = function(a, b) { return (a + b) | 0; };
    var rotr32 = function(x, n) { return (x >>> n) | (x << (32 - n)); };

    // Initial hash values: first 32 bits of sqrt(prime[0..7])
    var H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
             0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];

    // Work on a copy so the caller's array is not mutated
    var M = input.slice();

    // Pre-processing: pad to a multiple of 512 bits (64 bytes)
    var msgLen = M.length;
    M.push(0x80);                               // append 1-bit (0x80 byte)
    while (M.length % 64 !== 56) M.push(0x00); // zero-pad to 56 mod 64

    // Append original length in bits as a 64-bit big-endian integer
    var hi = (msgLen / 0x20000000) | 0;  // upper 32 bits of (msgLen * 8)
    var lo = msgLen << 3;                // lower 32 bits of (msgLen * 8)
    M.push((hi >>> 24) & 0xFF); M.push((hi >>> 16) & 0xFF);
    M.push((hi >>>  8) & 0xFF); M.push( hi         & 0xFF);
    M.push((lo >>> 24) & 0xFF); M.push((lo >>> 16) & 0xFF);
    M.push((lo >>>  8) & 0xFF); M.push( lo         & 0xFF);

    // Process each 64-byte block
    var blk, i, s0, s1;
    for (blk = 0; blk < M.length; blk += 64) {
        var W = [];

        // Prepare the 64-word message schedule
        for (i = 0; i < 16; i++) {
            W[i] = (M[blk + i*4]     << 24) | (M[blk + i*4 + 1] << 16) |
                   (M[blk + i*4 + 2] <<  8) |  M[blk + i*4 + 3];
        }
        for (i = 16; i < 64; i++) {
            s0   = rotr32(W[i-15],  7) ^ rotr32(W[i-15], 18) ^ (W[i-15] >>> 3);
            s1   = rotr32(W[i- 2], 17) ^ rotr32(W[i- 2], 19) ^ (W[i- 2] >>> 10);
            W[i] = add32(add32(add32(W[i-16], s0), W[i-7]), s1);
        }

        // Initialise working variables from current hash state
        var a = H[0], b = H[1], c = H[2], d = H[3];
        var e = H[4], f = H[5], g = H[6], h = H[7];

        // 64 compression rounds
        for (i = 0; i < 64; i++) {
            var S1    = rotr32(e,  6) ^ rotr32(e, 11) ^ rotr32(e, 25);
            var ch    = (e & f) ^ (~e & g);
            var temp1 = add32(add32(add32(add32(h, S1), ch), K[i]), W[i]);
            var S0    = rotr32(a,  2) ^ rotr32(a, 13) ^ rotr32(a, 22);
            var maj   = (a & b) ^ (a & c) ^ (b & c);
            var temp2 = add32(S0, maj);
            h = g; g = f; f = e; e = add32(d, temp1);
            d = c; c = b; b = a; a = add32(temp1, temp2);
        }

        // Add the compressed chunk to the current hash value
        H[0] = add32(H[0], a); H[1] = add32(H[1], b);
        H[2] = add32(H[2], c); H[3] = add32(H[3], d);
        H[4] = add32(H[4], e); H[5] = add32(H[5], f);
        H[6] = add32(H[6], g); H[7] = add32(H[7], h);
    }

    // Encode the eight 32-bit words as a 64-character lowercase hex string
    // NOTE: JScript does not support str[n] — use charAt() instead.
    var hex = '';
    var HEX = '0123456789abcdef';
    for (i = 0; i < 8; i++) {
        for (var k = 28; k >= 0; k -= 4) {
            hex += HEX.charAt((H[i] >>> k) & 0xF);
        }
    }
    return hex;
};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

// SHA-256 of a string (UTF-8 encoded).  Returns a 64-char lowercase hex string.
sha256 = function(message) {
    return sha256_bytes(_sha256_to_utf8(message));
};

// HMAC-SHA-256.  Both key and message are UTF-8 strings.
// Returns a 64-char lowercase hex string.
hmac_sha256 = function(key, message) {
    var keyBytes = _sha256_to_utf8(key);
    var msgBytes = _sha256_to_utf8(message);

    // Keys longer than the block size (64 bytes) are replaced by their hash
    if (keyBytes.length > 64) {
        keyBytes = _sha256_from_hex(sha256_bytes(keyBytes));
    }
    // Zero-pad the key to exactly one block
    while (keyBytes.length < 64) keyBytes.push(0x00);

    // Build inner and outer padded keys
    var ipad = [], opad = [], i;
    for (i = 0; i < 64; i++) {
        ipad.push(keyBytes[i] ^ 0x36);
        opad.push(keyBytes[i] ^ 0x5C);
    }

    // HMAC = SHA256(opad || SHA256(ipad || message))
    var inner = _sha256_from_hex(sha256_bytes(ipad.concat(msgBytes)));
    return sha256_bytes(opad.concat(inner));
};
