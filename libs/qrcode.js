// qrcode.js - QR code generation, from scratch, with no network and no
// dependencies. ISO/IEC 18004, versions 1-40, error correction L/M/Q/H,
// numeric / alphanumeric / byte (UTF-8) modes.
//
// The examples used to render QR codes by pointing an <img> at
// api.qrserver.com, which meant no network, no QR code. This encodes them
// locally instead.
//
//   var qr = qr_encode("https://example.com/");
//   log(qr_to_ascii(qr));                  // console
//   var html = qr_to_html(qr, { scale: 6 });  // an HTA window
//   var svg  = qr_to_svg(qr, { scale: 6 });   // a file
//
// Public helpers:
//   qr_encode(text, options)   -> { version, ec_level, mask, size, modules }
//   qr_to_ascii(qr, options)   -> string, two characters per module
//   qr_to_html(qr, options)    -> <table> markup, IE-safe (no SVG, no canvas)
//   qr_to_svg(qr, options)     -> SVG document as a string
//   qr_to_matrix(qr)           -> array of arrays of 0/1, for your own renderer
//
// options for qr_encode:
//   ec_level  "L" | "M" | "Q" | "H"   (default "M")
//   version   1..40, a *minimum* - a bigger one is used if the data needs it
//   mask      0..7, forces a mask instead of picking the best-scoring one
//   mode      "numeric" | "alphanumeric" | "byte", overrides auto-detection
//
// Every private helper is prefixed `_qr_` rather than hidden in an IIFE:
// JScript's eval (which is how load() pulls this file in) does not reliably
// preserve closure scope for function declarations inside an IIFE - see the
// note at the top of libs/crypto.js. Same reason every public name here is a
// bare assignment, not a function declaration.
//
// ES3 only: no let/const, no arrow functions, no template literals, no
// str[i] indexing (charAt), no Array.prototype.forEach/map (they come from
// core.js, which this file deliberately does not require).

// ---------------------------------------------------------------------------
// GF(256) - the field Reed-Solomon works in, x^8 + x^4 + x^3 + x^2 + 1 (0x11D)
// ---------------------------------------------------------------------------

_qr_exp = [];   // _qr_exp[i]  = a^i
_qr_log = [];   // _qr_log[x]  = i such that a^i = x

_qr_build_tables = function() {
    // The multiplicative group has order 255: a^255 == a^0 == 1. Running the
    // loop to 255 would overwrite log[1] with 255 instead of 0.
    var x = 1;
    for (var i = 0; i < 255; i++) {
        _qr_exp[i] = x;
        _qr_log[x] = i;
        x = x << 1;
        if (x & 0x100) { x = (x ^ 0x11D) & 0xFF; }
    }
    // Second lap, so the sum of two logs (up to 254 + 254) can be looked up
    // without a modulo.
    for (var j = 255; j < 512; j++) { _qr_exp[j] = _qr_exp[j - 255]; }
};

_qr_build_tables();

_qr_gf_mul = function(a, b) {
    if (a === 0 || b === 0) { return 0; }
    return _qr_exp[_qr_log[a] + _qr_log[b]];
};

// Generator polynomial for `degree` error-correction codewords:
// (x - a^0)(x - a^1)...(x - a^(degree-1)), coefficients high-order first.
_qr_rs_generator = function(degree) {
    var poly = [1];
    for (var d = 0; d < degree; d++) {
        var next = [];
        var i;
        for (i = 0; i <= poly.length; i++) { next[i] = 0; }
        // Multiplying (sum p_i x^(k-i)) by (x + a): the x term shifts every
        // coefficient up one power, the constant term scales it by a.
        for (i = 0; i < poly.length; i++) {
            next[i] = next[i] ^ poly[i];
            next[i + 1] = next[i + 1] ^ _qr_gf_mul(poly[i], _qr_exp[d]);
        }
        poly = next;
    }
    return poly;
};

// The remainder of data * x^degree divided by the generator polynomial: the
// error-correction codewords for one block.
_qr_rs_remainder = function(data, degree) {
    var gen = _qr_rs_generator(degree);
    var rem = [];
    var i, j;
    for (i = 0; i < degree; i++) { rem[i] = 0; }

    for (i = 0; i < data.length; i++) {
        var factor = data[i] ^ rem[0];
        for (j = 0; j < degree - 1; j++) { rem[j] = rem[j + 1]; }
        rem[degree - 1] = 0;
        for (j = 0; j < degree; j++) {
            rem[j] = rem[j] ^ _qr_gf_mul(gen[j + 1], factor);
        }
    }
    return rem;
};

// ---------------------------------------------------------------------------
// Spec tables
// ---------------------------------------------------------------------------

// Error correction per version, indexed [version - 1] then by level:
// [error-correction codewords per block, number of blocks].
// Everything else about the block layout follows from these two numbers plus
// the version's total codeword count, so there is no second table to keep in
// step with this one.
_qr_ec_table = [
    /*  1 */ { L: [7, 1],   M: [10, 1],  Q: [13, 1],  H: [17, 1]  },
    /*  2 */ { L: [10, 1],  M: [16, 1],  Q: [22, 1],  H: [28, 1]  },
    /*  3 */ { L: [15, 1],  M: [26, 1],  Q: [18, 2],  H: [22, 2]  },
    /*  4 */ { L: [20, 1],  M: [18, 2],  Q: [26, 2],  H: [16, 4]  },
    /*  5 */ { L: [26, 1],  M: [24, 2],  Q: [18, 4],  H: [22, 4]  },
    /*  6 */ { L: [18, 2],  M: [16, 4],  Q: [24, 4],  H: [28, 4]  },
    /*  7 */ { L: [20, 2],  M: [18, 4],  Q: [18, 6],  H: [26, 5]  },
    /*  8 */ { L: [24, 2],  M: [22, 4],  Q: [22, 6],  H: [26, 6]  },
    /*  9 */ { L: [30, 2],  M: [22, 5],  Q: [20, 8],  H: [24, 8]  },
    /* 10 */ { L: [18, 4],  M: [26, 5],  Q: [24, 8],  H: [28, 8]  },
    /* 11 */ { L: [20, 4],  M: [30, 5],  Q: [28, 8],  H: [24, 11] },
    /* 12 */ { L: [24, 4],  M: [22, 8],  Q: [26, 10], H: [28, 11] },
    /* 13 */ { L: [26, 4],  M: [22, 9],  Q: [24, 12], H: [22, 16] },
    /* 14 */ { L: [30, 4],  M: [24, 9],  Q: [20, 16], H: [24, 16] },
    /* 15 */ { L: [22, 6],  M: [24, 10], Q: [30, 12], H: [24, 18] },
    /* 16 */ { L: [24, 6],  M: [28, 10], Q: [24, 17], H: [30, 16] },
    /* 17 */ { L: [28, 6],  M: [28, 11], Q: [28, 16], H: [28, 19] },
    /* 18 */ { L: [30, 6],  M: [26, 13], Q: [28, 18], H: [28, 21] },
    /* 19 */ { L: [28, 7],  M: [26, 14], Q: [26, 21], H: [26, 25] },
    /* 20 */ { L: [28, 8],  M: [26, 16], Q: [30, 20], H: [28, 25] },
    /* 21 */ { L: [28, 8],  M: [26, 17], Q: [28, 23], H: [30, 25] },
    /* 22 */ { L: [28, 9],  M: [28, 17], Q: [30, 23], H: [24, 34] },
    /* 23 */ { L: [30, 9],  M: [28, 18], Q: [30, 25], H: [30, 30] },
    /* 24 */ { L: [30, 10], M: [28, 20], Q: [30, 27], H: [30, 32] },
    /* 25 */ { L: [26, 12], M: [28, 21], Q: [30, 29], H: [30, 35] },
    /* 26 */ { L: [28, 12], M: [28, 23], Q: [28, 34], H: [30, 37] },
    /* 27 */ { L: [30, 12], M: [28, 25], Q: [30, 34], H: [30, 40] },
    /* 28 */ { L: [30, 13], M: [28, 26], Q: [30, 35], H: [30, 42] },
    /* 29 */ { L: [30, 14], M: [28, 28], Q: [30, 38], H: [30, 45] },
    /* 30 */ { L: [30, 15], M: [28, 29], Q: [30, 40], H: [30, 48] },
    /* 31 */ { L: [30, 16], M: [28, 31], Q: [30, 43], H: [30, 51] },
    /* 32 */ { L: [30, 17], M: [28, 33], Q: [30, 45], H: [30, 54] },
    /* 33 */ { L: [30, 18], M: [28, 35], Q: [30, 48], H: [30, 57] },
    /* 34 */ { L: [30, 19], M: [28, 37], Q: [30, 51], H: [30, 60] },
    /* 35 */ { L: [30, 19], M: [28, 38], Q: [30, 53], H: [30, 63] },
    /* 36 */ { L: [30, 20], M: [28, 40], Q: [30, 56], H: [30, 66] },
    /* 37 */ { L: [30, 21], M: [28, 43], Q: [30, 59], H: [30, 70] },
    /* 38 */ { L: [30, 22], M: [28, 45], Q: [30, 62], H: [30, 74] },
    /* 39 */ { L: [30, 24], M: [28, 47], Q: [30, 65], H: [30, 77] },
    /* 40 */ { L: [30, 25], M: [28, 49], Q: [30, 68], H: [30, 81] }
];

// Row/column centres of the alignment patterns, indexed [version - 1].
// Version 1 has none.
_qr_align_table = [
    [], [6, 18], [6, 22], [6, 26], [6, 30], [6, 34],
    [6, 22, 38], [6, 24, 42], [6, 26, 46], [6, 28, 50], [6, 30, 54],
    [6, 32, 58], [6, 34, 62],
    [6, 26, 46, 66], [6, 26, 48, 70], [6, 26, 50, 74], [6, 30, 54, 78],
    [6, 30, 56, 82], [6, 30, 58, 86], [6, 34, 62, 90],
    [6, 28, 50, 72, 94], [6, 26, 50, 74, 98], [6, 30, 54, 78, 102],
    [6, 28, 54, 80, 106], [6, 32, 58, 84, 110], [6, 30, 58, 86, 114],
    [6, 34, 62, 90, 118],
    [6, 26, 50, 74, 98, 122], [6, 30, 54, 78, 102, 126],
    [6, 26, 52, 78, 104, 130], [6, 30, 56, 82, 108, 134],
    [6, 34, 60, 86, 112, 138], [6, 30, 58, 86, 114, 142],
    [6, 34, 62, 90, 118, 146],
    [6, 30, 54, 78, 102, 126, 150], [6, 24, 50, 76, 102, 128, 154],
    [6, 28, 54, 80, 106, 132, 158], [6, 32, 58, 84, 110, 136, 162],
    [6, 26, 54, 82, 110, 138, 166], [6, 30, 58, 86, 114, 142, 170]
];

_qr_ec_levels = { L: 1, M: 0, Q: 3, H: 2 };  // 2-bit format-info indicators

_qr_alnum_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:";

// ---------------------------------------------------------------------------
// Bit buffer
// ---------------------------------------------------------------------------

_qr_bits_new = function() {
    return { bits: [], length: 0 };
};

_qr_bits_push = function(buf, value, count) {
    for (var i = count - 1; i >= 0; i--) {
        buf.bits[buf.length++] = (value >>> i) & 1;
    }
};

// ---------------------------------------------------------------------------
// Modes
// ---------------------------------------------------------------------------

_qr_is_numeric = function(text) {
    return text.length > 0 && /^[0-9]+$/.test(text);
};

_qr_is_alphanumeric = function(text) {
    if (text.length === 0) { return false; }
    for (var i = 0; i < text.length; i++) {
        if (_qr_alnum_chars.indexOf(text.charAt(i)) === -1) { return false; }
    }
    return true;
};

_qr_pick_mode = function(text) {
    if (_qr_is_numeric(text))      { return "numeric"; }
    if (_qr_is_alphanumeric(text)) { return "alphanumeric"; }
    return "byte";
};

// UTF-8 bytes of a JScript (UTF-16) string, surrogate pairs included.
_qr_utf8_bytes = function(text) {
    var bytes = [];
    for (var i = 0; i < text.length; i++) {
        var c = text.charCodeAt(i);

        if (c >= 0xD800 && c <= 0xDBFF && i + 1 < text.length) {
            var low = text.charCodeAt(i + 1);
            if (low >= 0xDC00 && low <= 0xDFFF) {
                c = 0x10000 + ((c - 0xD800) << 10) + (low - 0xDC00);
                i++;
            }
        }

        if (c < 0x80) {
            bytes.push(c);
        } else if (c < 0x800) {
            bytes.push(0xC0 | (c >> 6), 0x80 | (c & 0x3F));
        } else if (c < 0x10000) {
            bytes.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F));
        } else {
            bytes.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 0x3F),
                       0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F));
        }
    }
    return bytes;
};

_qr_mode_indicator = function(mode) {
    if (mode === "numeric")      { return 1; }
    if (mode === "alphanumeric") { return 2; }
    return 4;
};

// Bit width of the character-count indicator: it grows with the version.
_qr_count_bits = function(mode, version) {
    var group = version <= 9 ? 0 : (version <= 26 ? 1 : 2);
    if (mode === "numeric")      { return [10, 12, 14][group]; }
    if (mode === "alphanumeric") { return [9, 11, 13][group]; }
    return [8, 16, 16][group];
};

// How many bits the payload itself takes, excluding mode and count indicators.
_qr_data_bits = function(mode, text) {
    if (mode === "numeric") {
        var groups = Math.floor(text.length / 3);
        var rest   = text.length % 3;
        return groups * 10 + (rest === 0 ? 0 : (rest === 1 ? 4 : 7));
    }
    if (mode === "alphanumeric") {
        return Math.floor(text.length / 2) * 11 + (text.length % 2) * 6;
    }
    return _qr_utf8_bytes(text).length * 8;
};

_qr_write_payload = function(buf, mode, text) {
    var i;
    if (mode === "numeric") {
        for (i = 0; i + 3 <= text.length; i += 3) {
            _qr_bits_push(buf, parseInt(text.substr(i, 3), 10), 10);
        }
        var rest = text.length - i;
        if (rest === 1) { _qr_bits_push(buf, parseInt(text.substr(i, 1), 10), 4); }
        else if (rest === 2) { _qr_bits_push(buf, parseInt(text.substr(i, 2), 10), 7); }
        return;
    }
    if (mode === "alphanumeric") {
        for (i = 0; i + 2 <= text.length; i += 2) {
            var pair = _qr_alnum_chars.indexOf(text.charAt(i)) * 45 +
                       _qr_alnum_chars.indexOf(text.charAt(i + 1));
            _qr_bits_push(buf, pair, 11);
        }
        if (i < text.length) {
            _qr_bits_push(buf, _qr_alnum_chars.indexOf(text.charAt(i)), 6);
        }
        return;
    }
    var bytes = _qr_utf8_bytes(text);
    for (i = 0; i < bytes.length; i++) { _qr_bits_push(buf, bytes[i], 8); }
};

// ---------------------------------------------------------------------------
// Geometry
//
// Total codeword capacity is derived from the function-pattern layout rather
// than tabulated: build the reserved-module map for a version, count what is
// left, and that is the data area. One less table to get wrong.
// ---------------------------------------------------------------------------

_qr_size_for = function(version) { return version * 4 + 17; };

_qr_new_grid = function(size, value) {
    var grid = [];
    for (var r = 0; r < size; r++) {
        var row = [];
        for (var c = 0; c < size; c++) { row[c] = value; }
        grid[r] = row;
    }
    return grid;
};

// true wherever a function pattern (finder, separator, timing, alignment,
// format/version reservation, dark module) lives - i.e. not data.
_qr_reserved_map = function(version) {
    var size = _qr_size_for(version);
    var map  = _qr_new_grid(size, false);
    var r, c, i;

    function block(top, left, height, width) {
        for (var y = 0; y < height; y++) {
            for (var x = 0; x < width; x++) {
                var ry = top + y, rx = left + x;
                if (ry >= 0 && ry < size && rx >= 0 && rx < size) { map[ry][rx] = true; }
            }
        }
    }

    // Finder patterns plus their separators, and the format-info strips that
    // sit right against them.
    block(0, 0, 9, 9);
    block(0, size - 8, 9, 8);
    block(size - 8, 0, 8, 9);

    // Timing patterns.
    for (i = 0; i < size; i++) { map[6][i] = true; map[i][6] = true; }

    // Alignment patterns, except where they would collide with a finder.
    var centres = _qr_align_table[version - 1];
    for (var a = 0; a < centres.length; a++) {
        for (var b = 0; b < centres.length; b++) {
            r = centres[a];
            c = centres[b];
            if ((r <= 8 && c <= 8) || (r <= 8 && c >= size - 9) || (r >= size - 9 && c <= 8)) {
                continue;
            }
            block(r - 2, c - 2, 5, 5);
        }
    }

    // Version information blocks (version 7 and up).
    if (version >= 7) {
        block(0, size - 11, 6, 3);
        block(size - 11, 0, 3, 6);
    }

    return map;
};

_qr_data_capacity_bits = function(version) {
    var map   = _qr_reserved_map(version);
    var size  = map.length;
    var free  = 0;
    for (var r = 0; r < size; r++) {
        for (var c = 0; c < size; c++) { if (!map[r][c]) { free++; } }
    }
    return free;
};

_qr_total_codewords = function(version) {
    return Math.floor(_qr_data_capacity_bits(version) / 8);
};

// Data codewords available once the error-correction codewords are subtracted.
_qr_data_codewords = function(version, ecLevel) {
    var ec = _qr_ec_table[version - 1][ecLevel];
    return _qr_total_codewords(version) - ec[0] * ec[1];
};

// ---------------------------------------------------------------------------
// Codeword assembly
// ---------------------------------------------------------------------------

_qr_choose_version = function(mode, text, ecLevel, minVersion) {
    for (var v = minVersion; v <= 40; v++) {
        var needed = 4 + _qr_count_bits(mode, v) + _qr_data_bits(mode, text);
        if (needed <= _qr_data_codewords(v, ecLevel) * 8) { return v; }
    }
    return -1;
};

_qr_make_codewords = function(mode, text, version, ecLevel) {
    var capacity = _qr_data_codewords(version, ecLevel) * 8;
    var buf = _qr_bits_new();
    var i;

    _qr_bits_push(buf, _qr_mode_indicator(mode), 4);
    _qr_bits_push(buf,
        mode === "byte" ? _qr_utf8_bytes(text).length : text.length,
        _qr_count_bits(mode, version));
    _qr_write_payload(buf, mode, text);

    // Terminator: up to four zero bits.
    var terminator = Math.min(4, capacity - buf.length);
    for (i = 0; i < terminator; i++) { _qr_bits_push(buf, 0, 1); }

    // Pad to a whole codeword, then alternate the two specified pad bytes.
    while (buf.length % 8 !== 0) { _qr_bits_push(buf, 0, 1); }

    var pad = [0xEC, 0x11];
    var p = 0;
    while (buf.length < capacity) {
        _qr_bits_push(buf, pad[p], 8);
        p = 1 - p;
    }

    var words = [];
    for (i = 0; i < buf.length; i += 8) {
        var byteValue = 0;
        for (var b = 0; b < 8; b++) { byteValue = (byteValue << 1) | buf.bits[i + b]; }
        words.push(byteValue);
    }
    return words;
};

// Split the data into blocks, compute each block's error-correction codewords,
// then interleave: first codeword of every block, second of every block, ...
_qr_interleave = function(dataWords, version, ecLevel) {
    var ec         = _qr_ec_table[version - 1][ecLevel];
    var ecPerBlock = ec[0];
    var numBlocks  = ec[1];

    var shortLen  = Math.floor(dataWords.length / numBlocks);
    var numLong   = dataWords.length % numBlocks;   // blocks with one extra codeword
    var numShort  = numBlocks - numLong;

    var dataBlocks = [];
    var ecBlocks   = [];
    var offset = 0;
    var i, j;

    for (i = 0; i < numBlocks; i++) {
        var len   = shortLen + (i < numShort ? 0 : 1);
        var block = dataWords.slice(offset, offset + len);
        offset += len;
        dataBlocks.push(block);
        ecBlocks.push(_qr_rs_remainder(block, ecPerBlock));
    }

    var out = [];
    var longest = shortLen + (numLong > 0 ? 1 : 0);
    for (j = 0; j < longest; j++) {
        for (i = 0; i < numBlocks; i++) {
            if (j < dataBlocks[i].length) { out.push(dataBlocks[i][j]); }
        }
    }
    for (j = 0; j < ecPerBlock; j++) {
        for (i = 0; i < numBlocks; i++) { out.push(ecBlocks[i][j]); }
    }
    return out;
};

// ---------------------------------------------------------------------------
// Matrix
// ---------------------------------------------------------------------------

_qr_draw_function_patterns = function(modules, version) {
    var size = modules.length;
    var i, j;

    function finder(top, left) {
        for (var y = -1; y <= 7; y++) {
            for (var x = -1; x <= 7; x++) {
                var r = top + y, c = left + x;
                if (r < 0 || r >= size || c < 0 || c >= size) { continue; }
                var edge = (y === 0 || y === 6) && x >= 0 && x <= 6;
                var side = (x === 0 || x === 6) && y >= 0 && y <= 6;
                var core = y >= 2 && y <= 4 && x >= 2 && x <= 4;
                modules[r][c] = edge || side || core;
            }
        }
    }

    finder(0, 0);
    finder(0, size - 7);
    finder(size - 7, 0);

    // Timing patterns: alternating, starting dark at index 6.
    for (i = 8; i < size - 8; i++) {
        modules[6][i] = (i % 2 === 0);
        modules[i][6] = (i % 2 === 0);
    }

    // Alignment patterns.
    var centres = _qr_align_table[version - 1];
    for (var a = 0; a < centres.length; a++) {
        for (var b = 0; b < centres.length; b++) {
            var cr = centres[a], cc = centres[b];
            if ((cr <= 8 && cc <= 8) || (cr <= 8 && cc >= size - 9) ||
                (cr >= size - 9 && cc <= 8)) { continue; }
            for (i = -2; i <= 2; i++) {
                for (j = -2; j <= 2; j++) {
                    var ring = Math.max(Math.abs(i), Math.abs(j));
                    modules[cr + i][cc + j] = (ring !== 1);
                }
            }
        }
    }

    // The one module that is always dark.
    modules[size - 8][8] = true;

    if (version >= 7) {
        var bits = _qr_version_bits(version);
        for (i = 0; i < 18; i++) {
            var bit = ((bits >> i) & 1) === 1;
            var row = Math.floor(i / 3);
            var col = size - 11 + (i % 3);
            modules[row][col] = bit;
            modules[col][row] = bit;
        }
    }
};

// BCH(18,6) over the version number, generator 0x1F25.
_qr_version_bits = function(version) {
    var rem = version;
    for (var i = 0; i < 12; i++) {
        rem = (rem << 1) ^ ((rem >>> 11) * 0x1F25);
    }
    return ((version << 12) | rem) >>> 0;
};

// BCH(15,5) over (ec level, mask), generator 0x537, then XOR 0x5412 so an
// all-zero format never produces an all-light strip.
_qr_format_bits = function(ecLevel, mask) {
    var data = (_qr_ec_levels[ecLevel] << 3) | mask;
    var rem  = data;
    for (var i = 0; i < 10; i++) {
        rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
    }
    return (((data << 10) | rem) ^ 0x5412) >>> 0;
};

_qr_draw_format = function(modules, ecLevel, mask) {
    var size = modules.length;
    var bits = _qr_format_bits(ecLevel, mask);
    var i;

    // Copy 1: down the left of the top-left finder, then along the top of it.
    for (i = 0; i <= 5; i++) { modules[i][8] = ((bits >> i) & 1) === 1; }
    modules[7][8] = ((bits >> 6) & 1) === 1;
    modules[8][8] = ((bits >> 7) & 1) === 1;
    modules[8][7] = ((bits >> 8) & 1) === 1;
    for (i = 9; i < 15; i++) { modules[8][14 - i] = ((bits >> i) & 1) === 1; }

    // Copy 2: the low bits run right-to-left along row 8 under the top-right
    // finder, the high bits run down column 8 beside the bottom-left one.
    for (i = 0; i < 8; i++) { modules[8][size - 1 - i] = ((bits >> i) & 1) === 1; }
    for (i = 8; i < 15; i++) { modules[size - 15 + i][8] = ((bits >> i) & 1) === 1; }
};

// Zigzag from the bottom-right: two columns at a time, right to left, skipping
// the vertical timing column, alternating upward and downward.
_qr_place_data = function(modules, reserved, words) {
    var size = modules.length;
    var bitIndex = 0;
    var upward = true;

    for (var right = size - 1; right >= 1; right -= 2) {
        if (right === 6) { right = 5; }  // column 6 is the timing pattern
        for (var vert = 0; vert < size; vert++) {
            var row = upward ? (size - 1 - vert) : vert;
            for (var k = 0; k < 2; k++) {
                var col = right - k;
                if (reserved[row][col]) { continue; }
                var bit = false;
                if (bitIndex < words.length * 8) {
                    bit = ((words[bitIndex >> 3] >>> (7 - (bitIndex & 7))) & 1) === 1;
                }
                modules[row][col] = bit;
                bitIndex++;
            }
        }
        upward = !upward;
    }
};

_qr_mask_at = function(mask, row, col) {
    switch (mask) {
        case 0: return (row + col) % 2 === 0;
        case 1: return row % 2 === 0;
        case 2: return col % 3 === 0;
        case 3: return (row + col) % 3 === 0;
        case 4: return (Math.floor(row / 2) + Math.floor(col / 3)) % 2 === 0;
        case 5: return ((row * col) % 2) + ((row * col) % 3) === 0;
        case 6: return (((row * col) % 2) + ((row * col) % 3)) % 2 === 0;
        default: return ((((row + col) % 2) + ((row * col) % 3)) % 2) === 0;
    }
};

_qr_apply_mask = function(modules, reserved, mask) {
    var size = modules.length;
    for (var r = 0; r < size; r++) {
        for (var c = 0; c < size; c++) {
            if (reserved[r][c]) { continue; }
            if (_qr_mask_at(mask, r, c)) { modules[r][c] = !modules[r][c]; }
        }
    }
};

// The four penalty rules from the spec. Lower is better.
_qr_penalty = function(modules) {
    var size = modules.length;
    var score = 0;
    var r, c, i;

    // Rule 1: runs of five or more same-coloured modules.
    for (r = 0; r < size; r++) {
        var runColour = modules[r][0], runLength = 1;
        for (c = 1; c < size; c++) {
            if (modules[r][c] === runColour) { runLength++; }
            else {
                if (runLength >= 5) { score += 3 + (runLength - 5); }
                runColour = modules[r][c];
                runLength = 1;
            }
        }
        if (runLength >= 5) { score += 3 + (runLength - 5); }
    }
    for (c = 0; c < size; c++) {
        var vColour = modules[0][c], vLength = 1;
        for (r = 1; r < size; r++) {
            if (modules[r][c] === vColour) { vLength++; }
            else {
                if (vLength >= 5) { score += 3 + (vLength - 5); }
                vColour = modules[r][c];
                vLength = 1;
            }
        }
        if (vLength >= 5) { score += 3 + (vLength - 5); }
    }

    // Rule 2: 2x2 blocks of one colour.
    for (r = 0; r < size - 1; r++) {
        for (c = 0; c < size - 1; c++) {
            var v = modules[r][c];
            if (modules[r][c + 1] === v && modules[r + 1][c] === v &&
                modules[r + 1][c + 1] === v) { score += 3; }
        }
    }

    // Rule 3: the finder-lookalike 1:1:3:1:1 pattern with four light modules
    // on either side, in rows or columns.
    var pattern  = [true, false, true, true, true, false, true, false, false, false, false];
    var reversed = [false, false, false, false, true, false, true, true, true, false, true];

    function matches(get, start, want) {
        for (var k = 0; k < 11; k++) { if (get(start + k) !== want[k]) { return false; } }
        return true;
    }

    for (r = 0; r < size; r++) {
        (function(row) {
            var get = function(idx) { return modules[row][idx]; };
            for (var start = 0; start + 11 <= size; start++) {
                if (matches(get, start, pattern))  { score += 40; }
                if (matches(get, start, reversed)) { score += 40; }
            }
        }(r));
    }
    for (c = 0; c < size; c++) {
        (function(col) {
            var get = function(idx) { return modules[idx][col]; };
            for (var start = 0; start + 11 <= size; start++) {
                if (matches(get, start, pattern))  { score += 40; }
                if (matches(get, start, reversed)) { score += 40; }
            }
        }(c));
    }

    // Rule 4: how far the dark/light balance is from 50%.
    var dark = 0;
    for (r = 0; r < size; r++) {
        for (c = 0; c < size; c++) { if (modules[r][c]) { dark++; } }
    }
    var percent   = dark * 100 / (size * size);
    var deviation = Math.floor(Math.abs(percent - 50) / 5);
    score += deviation * 10;

    return score;
};

_qr_copy_grid = function(grid) {
    var out = [];
    for (var r = 0; r < grid.length; r++) { out[r] = grid[r].slice(0); }
    return out;
};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

qr_encode = function(text, options) {
    options = options || {};

    if (text === null || typeof text === "undefined") { text = ""; }
    text = String(text);

    var ecLevel = String(options.ec_level || "M").toUpperCase();
    if (!_qr_ec_levels.hasOwnProperty(ecLevel)) {
        throw new Error("qr_encode: unknown error correction level '" + ecLevel +
                        "' (expected L, M, Q or H)");
    }

    var mode = options.mode || _qr_pick_mode(text);
    if (mode !== "numeric" && mode !== "alphanumeric" && mode !== "byte") {
        throw new Error("qr_encode: unknown mode '" + mode + "'");
    }
    if (mode === "numeric" && !_qr_is_numeric(text)) {
        throw new Error("qr_encode: numeric mode needs digits only");
    }
    if (mode === "alphanumeric" && !_qr_is_alphanumeric(text)) {
        throw new Error("qr_encode: alphanumeric mode allows only 0-9 A-Z and $%*+-./: and space");
    }

    var minVersion = 1;
    if (typeof options.version !== "undefined" && options.version !== null) {
        minVersion = Number(options.version);
        // Not `options.version ? ... : 1`: version 0 is invalid, and truthiness
        // would quietly turn it into 1.
        if (!(minVersion >= 1 && minVersion <= 40)) {
            throw new Error("qr_encode: version must be between 1 and 40");
        }
    }

    var version = _qr_choose_version(mode, text, ecLevel, minVersion);
    if (version === -1) {
        throw new Error("qr_encode: " + text.length + " characters do not fit in any " +
                        "QR version at error correction level " + ecLevel);
    }

    var words    = _qr_interleave(_qr_make_codewords(mode, text, version, ecLevel),
                                  version, ecLevel);
    var size     = _qr_size_for(version);
    var reserved = _qr_reserved_map(version);

    var base = _qr_new_grid(size, false);
    _qr_draw_function_patterns(base, version);
    _qr_place_data(base, reserved, words);

    var chosenMask = -1;
    var best = null;

    if (typeof options.mask === "number") {
        chosenMask = options.mask;
        if (chosenMask < 0 || chosenMask > 7) {
            throw new Error("qr_encode: mask must be between 0 and 7");
        }
        best = _qr_copy_grid(base);
        _qr_apply_mask(best, reserved, chosenMask);
        _qr_draw_format(best, ecLevel, chosenMask);
    } else {
        var bestScore = -1;
        for (var m = 0; m < 8; m++) {
            var candidate = _qr_copy_grid(base);
            _qr_apply_mask(candidate, reserved, m);
            _qr_draw_format(candidate, ecLevel, m);
            var score = _qr_penalty(candidate);
            if (bestScore === -1 || score < bestScore) {
                bestScore  = score;
                best       = candidate;
                chosenMask = m;
            }
        }
    }

    return {
        version:  version,
        ec_level: ecLevel,
        mode:     mode,
        mask:     chosenMask,
        size:     size,
        modules:  best
    };
};

qr_to_matrix = function(qr) {
    var out = [];
    for (var r = 0; r < qr.size; r++) {
        var row = [];
        for (var c = 0; c < qr.size; c++) { row[c] = qr.modules[r][c] ? 1 : 0; }
        out[r] = row;
    }
    return out;
};

// Two characters per module, because console cells are twice as tall as wide.
qr_to_ascii = function(qr, options) {
    options = options || {};
    var dark   = typeof options.dark   === "string" ? options.dark   : "##";
    var light  = typeof options.light  === "string" ? options.light  : "  ";
    var quiet  = typeof options.quiet_zone === "number" ? options.quiet_zone : 2;
    var lines  = [];
    var blank  = "";
    var r, c, i;

    for (i = 0; i < qr.size + quiet * 2; i++) { blank += light; }
    for (i = 0; i < quiet; i++) { lines.push(blank); }

    for (r = 0; r < qr.size; r++) {
        var line = "";
        for (i = 0; i < quiet; i++) { line += light; }
        for (c = 0; c < qr.size; c++) { line += qr.modules[r][c] ? dark : light; }
        for (i = 0; i < quiet; i++) { line += light; }
        lines.push(line);
    }

    for (i = 0; i < quiet; i++) { lines.push(blank); }
    return lines.join("\n");
};

// A <table> rather than SVG or canvas: HTAs render in an old IE engine, and a
// table of coloured cells is the one thing every version of it draws correctly.
qr_to_html = function(qr, options) {
    options = options || {};
    var scale = options.scale || 4;
    var dark  = options.dark  || "#000000";
    var light = options.light || "#ffffff";
    var quiet = typeof options.quiet_zone === "number" ? options.quiet_zone : 4;
    var side  = (qr.size + quiet * 2) * scale;

    var html = ['<table cellpadding="0" cellspacing="0" border="0" ',
                'style="border-collapse:collapse;background:', light,
                ';width:', side, 'px;height:', side, 'px;padding:', quiet * scale, 'px;">'];

    for (var r = 0; r < qr.size; r++) {
        html.push('<tr style="height:', scale, 'px;">');
        for (var c = 0; c < qr.size; c++) {
            html.push('<td style="width:', scale, 'px;height:', scale,
                      'px;background:', (qr.modules[r][c] ? dark : light), ';"></td>');
        }
        html.push('</tr>');
    }
    html.push('</table>');
    return html.join("");
};

qr_to_svg = function(qr, options) {
    options = options || {};
    var scale = options.scale || 4;
    var dark  = options.dark  || "#000000";
    var light = options.light || "#ffffff";
    var quiet = typeof options.quiet_zone === "number" ? options.quiet_zone : 4;
    var side  = (qr.size + quiet * 2) * scale;

    var svg = ['<?xml version="1.0" encoding="UTF-8"?>\n',
               '<svg xmlns="http://www.w3.org/2000/svg" width="', side, '" height="', side,
               '" viewBox="0 0 ', side, ' ', side, '" shape-rendering="crispEdges">',
               '<rect width="', side, '" height="', side, '" fill="', light, '"/>'];

    for (var r = 0; r < qr.size; r++) {
        for (var c = 0; c < qr.size; c++) {
            if (!qr.modules[r][c]) { continue; }
            svg.push('<rect x="', (c + quiet) * scale, '" y="', (r + quiet) * scale,
                     '" width="', scale, '" height="', scale, '" fill="', dark, '"/>');
        }
    }
    svg.push('</svg>');
    return svg.join("");
};
