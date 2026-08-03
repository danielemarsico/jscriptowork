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

(function() {
    var shellEnv = new ActiveXObject("WScript.Shell").Environment("Process");
    var offline  = shellEnv("JSW_TEST_HTTP_OFFLINE") !== "" || shellEnv("CI") === "true";
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
// Summary
// ---------------------------------------------------------------------------

_test.summary({ exit: true });
