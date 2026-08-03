// test-ext.js - Tests for libs/ext.js (Ext.util.JSON / Ext.encode / Ext.decode)
//
// ext.js is a compatibility shim nothing else in the project uses; it depends
// on JSON.stringify/JSON.parse from core.js, so core is loaded first.
//
// Run via:  cscript.exe launcher.js test-ext.js

load("core");
load("ext");
load("minitest");

describe("Ext.encode / Ext.decode", function() {

    it("exposes the Ext namespace", function() {
        assert.equal(typeof Ext,           "object");
        assert.equal(typeof Ext.util.JSON, "object");
    });

    it("Ext.encode matches JSON.stringify", function() {
        var src = { a: 1, b: [2, 3] };
        assert.equal(Ext.encode(src), JSON.stringify(src));
    });

    it("Ext.decode matches JSON.parse", function() {
        assert.deepEqual(Ext.decode('{"a":1}'), JSON.parse('{"a":1}'));
    });

    it("round-trips", function() {
        var src = { name: "test", values: [1, 2, 3] };
        assert.deepEqual(Ext.decode(Ext.encode(src)), src);
    });

});

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

_test.summary();
