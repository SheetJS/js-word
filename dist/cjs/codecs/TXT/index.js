"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
function write_para_elt_str(elt, opts) {
    var RS = opts && opts.RS || "\n";
    switch (elt.t) {
        case "s": return elt.v;
        case "t":
            return elt.r.map(function (tr) {
                return tr.c.map(function (tc) {
                    return tc.p.map(function (p) {
                        return write_para_str(p);
                    }).join(RS);
                }).join(RS);
            }).join(RS);
        default: throw "Cannot generate plaintext for " + elt.t + " elements";
    }
}
function write_para_str(para, opts) {
    return para.elts.map(function (elt) { return write_para_elt_str(elt, opts); }).join("");
}
function write_str(doc, opts) {
    var RS = opts && opts.RS || "\n";
    var o = doc.p.map(function (para) { return write_para_str(para, opts); }).join(RS) + RS;
    return o;
}
exports.write_str = write_str;
function write_buf(doc, opts) {
    return Buffer.from(write_str(doc, opts));
}
exports.write_buf = write_buf;
/* TODO: something more reasonable */
function parse_str(data) {
    var doc = { p: [] };
    var texts = data.split(/\r\n?|\n/);
    if (!texts[texts.length - 1] && data.slice(-2).match(/[\r\n]$/))
        texts = texts.slice(0, -1);
    texts.forEach(function (line) {
        var t = { t: "s", v: line };
        var para = { elts: [t] };
        doc.p.push(para);
    });
    return doc;
}
exports.parse_str = parse_str;
//# sourceMappingURL=index.js.map