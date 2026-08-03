// test-http.js - Integration tests for http_request (GET / POST over HTTPS)
// Uses httpbin.org as the echo server by default (requires network access).
//
// Offline mode: set JSW_TEST_HTTP_OFFLINE=1 (or run under CI, which sets
// CI=true automatically) to replace http_request with a local stub that
// mimics httpbin's responses for exactly the requests this suite makes, so
// the whole file runs with no network access and is immune to httpbin.org
// rate-limiting/downtime.
//
// Run via:  cscript.exe launcher.js test-http.js

load("core");
load("polyfills");
load("system");
load("minitest");

var BASE = "https://httpbin.org";

// ---------------------------------------------------------------------------
// Offline stub — swaps out http_request (bare assignment, per the load()/eval
// scoping rule in CLAUDE.md) with a local fake that answers the exact request
// shapes used below, without touching the network.
// ---------------------------------------------------------------------------

// Bare assignment so the async/binary block further down can see it too.
OFFLINE = (function() {
    var shellEnv = new ActiveXObject("WScript.Shell").Environment("Process");
    return shellEnv("JSW_TEST_HTTP_OFFLINE") !== "" || shellEnv("CI") === "true";
}());

(function() {
    var offline = OFFLINE;
    if (!offline) { return; }

    log("http_request tests running in OFFLINE stub mode - httpbin.org is not contacted.");

    function parseQueryString(qs) {
        var out = {};
        if (!qs) { return out; }
        var pairs = qs.split("&");
        for (var i = 0; i < pairs.length; i++) {
            var kv = pairs[i].split("=");
            out[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || "");
        }
        return out;
    }

    function headerValue(headers, name) {
        if (!headers) { return ""; }
        for (var key in headers) {
            if (Object.prototype.hasOwnProperty.call(headers, key) && key.toLowerCase() === name) {
                return headers[key];
            }
        }
        return "";
    }

    // Overrides the global http_request defined by system.js above.
    http_request = function(url, method, reqListener, body, headers, timeout) {
        if (["GET", "POST", "PUT", "DELETE"].indexOf(method) === -1) {
            throw "method not recognized:" + method;
        }
        // 192.0.2.1 (RFC 5737 TEST-NET-1) is reserved and never routes, so a
        // real request always exceeds any timeout - the stub reproduces that
        // outcome directly instead of actually waiting it out.
        if (url.indexOf("192.0.2.1") !== -1) {
            throw "offline stub: simulated timeout for " + url;
        }

        var qIndex = url.indexOf("?");
        var path   = qIndex === -1 ? url : url.substring(0, qIndex);
        var query  = qIndex === -1 ? ""  : url.substring(qIndex + 1);

        var status   = 200;
        var response = { url: url, args: parseQueryString(query), headers: headers || {} };

        if (path.indexOf("/status/") !== -1) {
            status = parseInt(path.substring(path.lastIndexOf("/") + 1), 10);
        } else if (method === "POST" && path.indexOf("/post") !== -1) {
            if (headerValue(headers, "content-type").indexOf("application/json") !== -1) {
                response.json = JSON.parse(body);
                response.form = {};
            } else {
                response.json = null;
                response.form = parseQueryString(body);
            }
        }

        reqListener(JSON.stringify(response), status, "Content-Type: application/json\r\n");
    };
}());

// ---------------------------------------------------------------------------
// Helper: run a request and capture response text + status synchronously.
// Returns { body: string, status: number }
// ---------------------------------------------------------------------------
function request(url, method, body, headers, timeout) {
    var result = { body: null, status: null, headers: null };
    http_request(url, method, function(responseText, statusCode, responseHeaders) {
        result.body    = responseText;
        result.status  = statusCode;
        result.headers = responseHeaders;
    }, body, headers, timeout);
    return result;
}

function parseJSON(str) {
    return JSON.parse(str);
}

// ---------------------------------------------------------------------------
// Error handling (no network required)
// ---------------------------------------------------------------------------

describe("http_request - error handling", function() {

    it("throws on unrecognised HTTP method", function() {
        assert.throws(function() {
            http_request(BASE + "/get", "PATCH", function() {});
        });
    });

    it("throws on empty method string", function() {
        assert.throws(function() {
            http_request(BASE + "/get", "", function() {});
        });
    });

    it("accepts GET without throwing", function() {
        assert.doesNotThrow(function() {
            http_request(BASE + "/get", "GET", function() {});
        });
    });

});

// ---------------------------------------------------------------------------
// GET
// ---------------------------------------------------------------------------

describe("http_request - GET", function() {

    it("returns HTTP 200 for a valid URL", function() {
        var res = request(BASE + "/get", "GET");
        assert.equal(res.status, 200);
    });

    it("returns a non-empty response body", function() {
        var res = request(BASE + "/get", "GET");
        assert.ok(res.body && res.body.length > 0);
    });

    it("response body is valid JSON", function() {
        var res = request(BASE + "/get", "GET");
        var data = parseJSON(res.body);
        assert.ok(typeof data === "object" && data !== null);
    });

    it("response contains the requested url", function() {
        var res = request(BASE + "/get", "GET");
        var data = parseJSON(res.body);
        assert.ok(data.url.indexOf("/get") !== -1);
    });

    it("custom request header is echoed back by httpbin", function() {
        // Append timestamp to bust WinInet's local HTTP cache so the header is actually sent fresh
        var res = request(BASE + "/get?_=" + Date.now(), "GET", null, { "X-Test-Header": "jscriptowork" });
        var data = parseJSON(res.body);
        assert.equal(data.headers["X-Test-Header"], "jscriptowork");
    });

    it("returns HTTP 404 for a missing resource", function() {
        var res = request(BASE + "/status/404", "GET");
        assert.equal(res.status, 404);
    });

    it("returns HTTP 500 for a server-error endpoint", function() {
        var res = request(BASE + "/status/500", "GET");
        assert.equal(res.status, 500);
    });

    it("query string parameters are included in response url", function() {
        var res = request(BASE + "/get?foo=bar&n=42", "GET");
        var data = parseJSON(res.body);
        assert.equal(data.args.foo, "bar");
        assert.equal(data.args.n,   "42");
    });

    it("exposes response headers to the callback", function() {
        var res = request(BASE + "/get", "GET");
        assert.equal(typeof res.headers, "string");
        assert.ok(res.headers.toLowerCase().indexOf("content-type") !== -1, res.headers);
    });

    it("throws when the request exceeds the given timeout", function() {
        // 192.0.2.1 is TEST-NET-1 (RFC 5737): reserved so it never routes on
        // the real internet. This makes the timeout deterministic and
        // independent of httpbin's own responsiveness.
        assert.throws(function() {
            request("http://192.0.2.1/", "GET", null, null, 200);
        });
    });

});

// ---------------------------------------------------------------------------
// POST - JSON body
// ---------------------------------------------------------------------------

describe("http_request - POST (JSON body)", function() {

    var payload     = JSON.stringify({ message: "hello", value: 42 });
    var jsonHeaders = { "Content-Type": "application/json" };

    it("returns HTTP 200 for POST", function() {
        var res = request(BASE + "/post", "POST", payload, jsonHeaders);
        assert.equal(res.status, 200);
    });

    it("response is valid JSON", function() {
        var res = request(BASE + "/post", "POST", payload, jsonHeaders);
        var data = parseJSON(res.body);
        assert.ok(typeof data === "object" && data !== null);
    });

    it("httpbin echoes the parsed JSON body", function() {
        var res  = request(BASE + "/post", "POST", payload, jsonHeaders);
        var data = parseJSON(res.body);
        assert.equal(data.json.message, "hello");
        assert.equal(data.json.value,   42);
    });

    it("response url points to /post endpoint", function() {
        var res  = request(BASE + "/post", "POST", payload, jsonHeaders);
        var data = parseJSON(res.body);
        assert.ok(data.url.indexOf("/post") !== -1);
    });

});

// ---------------------------------------------------------------------------
// POST - form-encoded body
// ---------------------------------------------------------------------------

describe("http_request - POST (form body)", function() {

    var formHeaders = { "Content-Type": "application/x-www-form-urlencoded" };

    it("sends form field and httpbin echoes it back", function() {
        var res  = request(BASE + "/post", "POST", "field=jscriptowork", formHeaders);
        var data = parseJSON(res.body);
        assert.equal(data.form.field, "jscriptowork");
    });

    it("sends multiple form fields", function() {
        var res  = request(BASE + "/post", "POST", "a=1&b=2", formHeaders);
        var data = parseJSON(res.body);
        assert.equal(data.form.a, "1");
        assert.equal(data.form.b, "2");
    });

});

// ---------------------------------------------------------------------------
// Async and binary HTTP
//
// http_request_async / http_request_bytes / http_download_file talk to
// ServerXMLHTTP directly and are not covered by the offline stub (which only
// replaces http_request), so the network-dependent assertions are skipped when
// running offline. The argument checking in front of them needs no network and
// always runs: those calls throw before any COM object is created.
// ---------------------------------------------------------------------------

describe("http_request_async - argument checking", function() {

    it("is defined", function() {
        assert.equal(typeof http_request_async, "function");
        assert.equal(typeof http_wait_all,      "function");
    });

    it("rejects an unsupported method before opening a connection", function() {
        assert.throws(function() {
            http_request_async("http://example.invalid/", "PATCH", function() {});
        });
        assert.throws(function() {
            http_request_async("http://example.invalid/", "", function() {});
        });
    });

});

describe("http_request_bytes / http_download_file - argument checking", function() {

    it("are defined", function() {
        assert.equal(typeof http_request_bytes,  "function");
        assert.equal(typeof http_download_file,  "function");
    });

    it("http_request_bytes rejects an unsupported method", function() {
        assert.throws(function() {
            http_request_bytes("http://example.invalid/", "PATCH", function() {});
        });
    });

    it("http_download_file rejects an unsupported method", function() {
        assert.throws(function() {
            http_download_file("http://example.invalid/", "C:\\nope.bin", { method: "PATCH" });
        });
    });

});

describe("http_request_async - over the network", function() {

    if (OFFLINE) {
        var REASON = "needs real network; the offline stub only replaces http_request";
        skip("returns a handle that is not done yet",        REASON);
        skip("wait() fires the callback and reports done",   REASON);
        skip("exposes status, text and headers on the handle", REASON);
        skip("wait() is idempotent - the callback fires once", REASON);
        skip("http_wait_all resolves several requests",      REASON);
        skip("wait(seconds) returns false when it times out", REASON);
        skip("abort() leaves the handle not done",           REASON);
        return;
    }

    it("returns a handle that is not done yet", function() {
        var h = http_request_async(BASE + "/get", "GET", function() {});
        assert.equal(h.done, false);
        assert.equal(h.url,  BASE + "/get");
        assert.equal(typeof h.wait,  "function");
        assert.equal(typeof h.abort, "function");
        h.wait(30);
    });

    it("wait() fires the callback and reports done", function() {
        var fired = 0;
        var h = http_request_async(BASE + "/get", "GET", function() { fired++; });
        assert.equal(h.wait(30), true);
        assert.equal(fired,   1);
        assert.equal(h.done,  true);
    });

    it("exposes status, text and headers on the handle", function() {
        var h = http_request_async(BASE + "/get", "GET", function() {});
        h.wait(30);
        assert.equal(h.status, 200);
        assert.ok(h.text.length > 0);
        assert.ok(h.headers.length > 0);
    });

    it("wait() is idempotent - the callback fires once", function() {
        var fired = 0;
        var h = http_request_async(BASE + "/get", "GET", function() { fired++; });
        h.wait(30);
        h.wait(30);
        assert.equal(fired, 1);
    });

    it("http_wait_all resolves several requests", function() {
        var seen = [];
        var handles = [
            http_request_async(BASE + "/get?n=1", "GET", function() { seen.push(1); }),
            http_request_async(BASE + "/get?n=2", "GET", function() { seen.push(2); })
        ];
        assert.equal(http_wait_all(handles, 30), true);
        assert.equal(seen.length, 2);
        assert.equal(handles[0].status, 200);
        assert.equal(handles[1].status, 200);
    });

    it("wait(seconds) returns false when it times out", function() {
        // 192.0.2.1 is RFC 5737 TEST-NET-1: reserved, never routes.
        var h = http_request_async("http://192.0.2.1/", "GET", function() {});
        assert.equal(h.wait(1), false);
        assert.equal(h.done,    false);
        h.abort();
    });

    it("abort() leaves the handle not done", function() {
        var h = http_request_async("http://192.0.2.1/", "GET", function() {});
        h.abort();
        assert.equal(h.done, false);
        assert.equal(h.wait(1), false);
    });

});

describe("http_request_bytes / http_download_file - over the network", function() {

    if (OFFLINE) {
        var REASON2 = "needs real network; the offline stub only replaces http_request";
        skip("hands the body to the callback as a byte array", REASON2);
        skip("byte values are all in 0-255",                   REASON2);
        skip("reports the status code",                        REASON2);
        skip("downloads a file to disk",                       REASON2);
        skip("returns the status and writes nothing on 404",   REASON2);
        return;
    }

    var fso  = new ActiveXObject("Scripting.FileSystemObject");
    var TMPD = fso.GetSpecialFolder(2).Path + "\\jscriptowork_http_test";
    if (!fso.FolderExists(TMPD)) { fso.CreateFolder(TMPD); }

    it("hands the body to the callback as a byte array", function() {
        var got = null;
        http_request_bytes(BASE + "/bytes/64", "GET", function(bytes) { got = bytes; });
        assert.ok(Array.isArray(got));
        assert.equal(got.length, 64);
    });

    it("byte values are all in 0-255", function() {
        var got = null;
        http_request_bytes(BASE + "/bytes/32", "GET", function(bytes) { got = bytes; });
        var ok = true;
        for (var i = 0; i < got.length; i++) {
            if (typeof got[i] !== "number" || got[i] < 0 || got[i] > 255) { ok = false; }
        }
        assert.ok(ok, "byte array contains a value outside 0-255");
    });

    it("reports the status code", function() {
        var status = null;
        http_request_bytes(BASE + "/bytes/8", "GET", function(b, s) { status = s; });
        assert.equal(status, 200);
    });

    it("downloads a file to disk", function() {
        var path = TMPD + "\\downloaded.bin";
        if (file_exists(path)) { delete_file(path); }
        var status = http_download_file(BASE + "/bytes/128", path);
        assert.equal(status, 200);
        assert.ok(file_exists(path));
        assert.equal(read_binary_file(path).length, 128);
        delete_file(path);
    });

    it("returns the status and writes nothing on 404", function() {
        var path = TMPD + "\\missing.bin";
        if (file_exists(path)) { delete_file(path); }
        var status = http_download_file(BASE + "/status/404", path);
        assert.equal(status, 404);
        assert.notOk(file_exists(path));
    });

    fso.DeleteFolder(TMPD, true);

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
