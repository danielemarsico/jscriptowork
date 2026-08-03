// ext.js - Ext.util.JSON / Ext.encode / Ext.decode compatibility shim
//
// Nothing in this project uses these; they exist only for scripts written
// against the old ExtJS-style JSON API. Load explicitly with load("ext") if
// you need them - they are not pulled in by core.js.

if (!this.Ext) {

    Ext={};
    Ext.util = {};

    /**
     * @class Ext.util.JSON
     * Modified version of Douglas Crockford"s json.js that doesn"t
     * mess with the Object prototype
     * http://www.json.org/js.html
     * @singleton
     */
    Ext.util.JSON = {
        encode : function(o) {
            return JSON.stringify(o);
        },

        decode : function(s) {
            return JSON.parse(s);
        }
    };

    /**
     * Shorthand for {@link Ext.util.JSON#encode}
     * @param {Mixed} o The variable to encode
     * @return {String} The JSON string
     * @member Ext
     * @method encode
     */
    Ext.encode = Ext.util.JSON.encode;
    /**
     * Shorthand for {@link Ext.util.JSON#decode}
     * @param {String} json The JSON string
     * @param {Boolean} safe (optional) Whether to return null or throw an exception if the JSON is invalid.
     * @return {Object} The resulting object
     * @member Ext
     * @method decode
     */
    Ext.decode = Ext.util.JSON.decode;

}
