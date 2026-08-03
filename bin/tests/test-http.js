// test-http.js - Integration tests for http_request (GET / POST over HTTPS)
// Requires network access. Uses httpbin.org as the echo server.
// Run via:  cscript.exe launcher.js test-http.js

load("core");
load("polyfills");
load("system");
load("minitest");

var BASE = "https://httpbin.org";

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
        assert.throws(function() {
            request(BASE + "/delay/5", "GET", null, null, 1);
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
// Summary
// ---------------------------------------------------------------------------

_test.summary();
