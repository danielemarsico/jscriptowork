// json-encode-decode.js - JSON.stringify / JSON.parse, local and over HTTP
//
// Part 1 encodes a plain object to a JSON string and decodes it back,
// entirely offline.
//
// Part 2 fetches a real JSON document from jsonplaceholder.typicode.com (a
// free public API made for exactly this kind of example/testing use) and
// parses it with JSON.parse. It needs network access; failures there are
// caught and reported without stopping the script.
//
// Run via:
//   examples\run.bat json-encode-decode.js
//   cscript.exe bin\launcher.js examples\json-encode-decode.js
//   cscript.exe dist\launcher.js examples\json-encode-decode.js

load("core");
load("polyfills");
load("system");

// ---------------------------------------------------------------------------
// Part 1 - local encode / decode round-trip
// ---------------------------------------------------------------------------

log("--- Part 1: local JSON encode/decode ---");

var ticket = {
    id:       42,
    title:    "Replace printer toner",
    done:     false,
    tags:     ["office", "urgent"],
    reported: new Date(2026, 0, 15)
};

var encoded = JSON.stringify(ticket, null, 2);
log("Encoded:");
log(encoded);

var decoded = JSON.parse(encoded);
log("Decoded id:    " + decoded.id);
log("Decoded title: " + decoded.title);
log("Decoded tags:  " + decoded.tags.join(", "));

// ---------------------------------------------------------------------------
// Part 2 - decode JSON fetched from a public sample API
// ---------------------------------------------------------------------------

log("");
log("--- Part 2: JSON from jsonplaceholder.typicode.com ---");

try {

    http_request(
        "https://jsonplaceholder.typicode.com/todos/1",
        "GET",
        function(responseText, statusCode) {

            if (statusCode !== 200) {
                log("Request failed with HTTP " + statusCode);
                return;
            }

            var todo = JSON.parse(responseText);
            log("Raw response: " + responseText);
            log("Parsed userId:    " + todo.userId);
            log("Parsed id:        " + todo.id);
            log("Parsed title:     " + todo.title);
            log("Parsed completed: " + todo.completed);
        },
        null,
        null,
        5000 // 5s timeout
    );

} catch (ex) {
    log("Could not reach the network: " + ex.message);
}
