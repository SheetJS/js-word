"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var jsdom_1 = require("jsdom");
function parse_str(data) {
    var dom = new jsdom_1.JSDOM(data);
    var doc = { p: [] };
    var para = { elts: [] };
    para.elts.push({ t: "s", v: dom.window.document.querySelector('body').textContent });
    doc.p.push(para);
    return doc;
}
exports.parse_str = parse_str;
function read(data) {
    return parse_str(data.toString());
}
exports.read = read;
//# sourceMappingURL=index.js.map