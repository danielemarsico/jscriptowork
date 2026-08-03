// csv.js - CSV reading and writing for JScript / CScript
//
//   load("core"); load("system"); load("csv");
//
//   csv_parse(text [, options])          -> array of rows (or of objects)
//   csv_format(rows [, options])         -> CSV string
//   read_csv_file(path [, options])      -> array of rows (or of objects)
//   write_csv_file(path, rows [, opts])  -> writes a CSV file
//
// Until now every caller that wanted tabular data went through Excel COM
// (do_in_excel + read_sheet_data), which means Excel has to be installed and a
// COM server has to start just to read a few thousand rows. These helpers need
// nothing but the file system.
//
// Follows RFC 4180: fields containing the delimiter, a double quote, CR or LF
// are wrapped in double quotes, and a literal double quote inside a quoted
// field is written as two double quotes.
//
// Options (all optional):
//   delimiter    single character field separator.        default ","
//   headers      parse: true  -> return objects keyed by the first row.
//                format: true -> derive a header row from the first object's
//                        keys; an array -> use exactly those keys, in order.
//                                                          default false
//   trim         trim whitespace around every parsed field. default false
//   eol          format only: line terminator.             default "\r\n"
//   quote_all    format only: quote every field.           default false
//   unicode      write_csv_file only: write UTF-16.         default false
//
// Parsing notes:
//   - Accepts CRLF, LF or CR line endings, mixed within one file.
//   - A UTF-8 BOM at the start of the text is stripped.
//   - A trailing newline does not produce a final empty row, but a blank line
//     in the middle of the file does produce a one-empty-field row - that is a
//     real record in CSV, not a formatting artefact.
//   - Every value is a string; no number or date coercion is attempted.

// ---------------------------------------------------------------------------
// Private helpers (_csv_ prefix, per the load()/eval scoping note in CLAUDE.md)
// ---------------------------------------------------------------------------

_csv_option = function(options, name, fallback) {
    if (!options) { return fallback; }
    return (typeof options[name] === 'undefined') ? fallback : options[name];
};

// Quotes one field if it needs quoting (or if quote_all was requested).
_csv_quote_field = function(value, delimiter, quote_all) {
    var s = (value === null || typeof value === 'undefined') ? '' : String(value);
    var needs = quote_all ||
                s.indexOf(delimiter) !== -1 ||
                s.indexOf('"')       !== -1 ||
                s.indexOf('\n')      !== -1 ||
                s.indexOf('\r')      !== -1;
    if (!needs) { return s; }
    return '"' + s.replace(/"/g, '""') + '"';
};

// Collects the key order for an array of objects: the first object's keys,
// then any key later objects introduce, so no column is silently dropped.
_csv_derive_headers = function(rows) {
    var keys = [];
    var seen = {};
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        for (var k in row) {
            if (Object.prototype.hasOwnProperty.call(row, k) && !seen[k]) {
                seen[k] = true;
                keys.push(k);
            }
        }
    }
    return keys;
};

// ---------------------------------------------------------------------------
// csv_parse(text [, options])
// ---------------------------------------------------------------------------

// Returns an array of arrays of strings, or - with options.headers - an array
// of objects keyed by the first row. Returns [] for empty input.
csv_parse = function(text, options) {

    var delimiter = _csv_option(options, 'delimiter', ',');
    var headers   = _csv_option(options, 'headers',   false);
    var trim      = _csv_option(options, 'trim',      false);

    if (text === null || typeof text === 'undefined') { return []; }
    text = String(text);

    // Strip a leading BOM. Two shapes reach us: U+FEFF when the file was
    // decoded as Unicode, and the raw EF BB BF bytes when a UTF-8 file was
    // read as ASCII. Written as escapes so this file stays pure ASCII.
    if (text.charAt(0) === '\uFEFF') {
        text = text.substring(1);
    } else if (text.substring(0, 3) === '\u00EF\u00BB\u00BF') {
        text = text.substring(3);
    }

    if (text.length === 0) { return []; }

    var rows     = [];
    var row      = [];
    var field    = '';
    var inQuotes = false;
    var quoted   = false;   // this field had quotes: never trim it
    var i        = 0;

    // Ends the current field. trim applies only to unquoted fields - quoting is
    // how a CSV author says "this whitespace is data".
    // Declared as a var expression, not a function declaration: see the note at
    // the top of libs/crypto.js about JScript's eval() and inner functions.
    var push_field = function() {
        row.push((trim && !quoted) ? field.trim() : field);
        field  = '';
        quoted = false;
    };

    // Character-at-a-time scan: a regex split cannot see whether a delimiter or
    // newline is inside a quoted field. Uses charAt(), never str[i], which
    // JScript does not support.
    while (i < text.length) {
        var c = text.charAt(i);

        if (inQuotes) {
            if (c === '"') {
                if (text.charAt(i + 1) === '"') {   // escaped quote
                    field += '"';
                    i += 2;
                    continue;
                }
                inQuotes = false;
                i++;
                continue;
            }
            field += c;
            i++;
            continue;
        }

        if (c === '"') {
            inQuotes = true;
            quoted   = true;
            i++;
            continue;
        }

        if (c === delimiter) {
            push_field();
            i++;
            continue;
        }

        if (c === '\r' || c === '\n') {
            push_field();
            rows.push(row);
            row = [];
            // CRLF counts as one terminator.
            if (c === '\r' && text.charAt(i + 1) === '\n') { i += 2; } else { i++; }
            continue;
        }

        field += c;
        i++;
    }

    // Whatever is left after the last terminator is the final row - unless the
    // text ended exactly on a terminator, in which case there is nothing left.
    if (field.length > 0 || row.length > 0 || quoted) {
        push_field();
        rows.push(row);
    }

    if (!headers) { return rows; }

    if (rows.length === 0) { return []; }

    var keys    = rows[0];
    var records = [];
    for (var r = 1; r < rows.length; r++) {
        var record = {};
        for (var c2 = 0; c2 < keys.length; c2++) {
            // Short rows yield "" rather than undefined, so every record has
            // the same shape whatever the file looks like.
            record[keys[c2]] = (c2 < rows[r].length) ? rows[r][c2] : '';
        }
        records.push(record);
    }
    return records;
};

// ---------------------------------------------------------------------------
// csv_format(rows [, options])
// ---------------------------------------------------------------------------

// rows may be an array of arrays or an array of objects (use options.headers
// to control the column order for objects). Returns a string with a trailing
// end-of-line, or "" for an empty input array.
csv_format = function(rows, options) {

    var delimiter = _csv_option(options, 'delimiter', ',');
    // null means "not specified": an array of objects then gets a header row
    // derived from its keys. Pass headers:false to suppress it outright.
    var headers   = _csv_option(options, 'headers',   null);
    var eol       = _csv_option(options, 'eol',       '\r\n');
    var quote_all = _csv_option(options, 'quote_all', false);

    if (!rows || rows.length === 0) { return ''; }

    var lines = [];
    var keys  = null;

    // Array of objects: work out the columns, and emit a header row unless the
    // caller explicitly passed headers:false.
    if (!Array.isArray(rows[0])) {
        keys = Array.isArray(headers) ? headers : _csv_derive_headers(rows);
        if (headers !== false) {
            var head = [];
            for (var h = 0; h < keys.length; h++) {
                head.push(_csv_quote_field(keys[h], delimiter, quote_all));
            }
            lines.push(head.join(delimiter));
        }
    }

    for (var i = 0; i < rows.length; i++) {
        var row  = rows[i];
        var cells = [];
        if (keys) {
            for (var k = 0; k < keys.length; k++) {
                cells.push(_csv_quote_field(row[keys[k]], delimiter, quote_all));
            }
        } else {
            for (var c = 0; c < row.length; c++) {
                cells.push(_csv_quote_field(row[c], delimiter, quote_all));
            }
        }
        lines.push(cells.join(delimiter));
    }

    return lines.join(eol) + eol;
};

// ---------------------------------------------------------------------------
// File helpers
// ---------------------------------------------------------------------------

// Reads a CSV file and parses it. Same options as csv_parse.
// Encoding is handled by read_text_file, which auto-detects a Unicode BOM.
read_csv_file = function(path, options) {
    return csv_parse(read_text_file(path), options);
};

// Formats rows and writes them to a file. Same options as csv_format, plus
// `unicode` to write UTF-16 instead of ASCII.
write_csv_file = function(path, rows, options) {
    write_text_to_file(csv_format(rows, options), path, _csv_option(options, 'unicode', false));
};
