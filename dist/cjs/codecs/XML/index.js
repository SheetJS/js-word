"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var jsdom_1 = require("jsdom");
function parse_str(data) {
    // dom.window.document.children[0].tagName)
    var listOfParagraphs = [];
    var dom = new jsdom_1.JSDOM(data, { contentType: "text/xml" });
    var dfs = function (node) {
        if (node.hasChildNodes) {
            node.childNodes.forEach(function (node) {
                if (node.nodeName === "w:p")
                    listOfParagraphs.push(node);
                dfs(node);
            });
        }
    };
    dfs(dom.window.document);
    var out = { p: [] };
    listOfParagraphs.forEach(function (node) {
        node.childNodes.forEach(function (child) {
            var para = { elts: [] };
            para.elts.push({ t: "s", v: child.textContent });
            out.p.push(para);
        });
    });
    return out;
}
exports.parse_str = parse_str;
;
function read(data) {
    return parse_str(data.toString());
}
exports.read = read;
//# sourceMappingURL=index.js.map