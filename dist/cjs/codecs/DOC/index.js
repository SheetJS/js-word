"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var cfb_1 = require("cfb");
var fib_1 = require("./fib");
var clx_1 = require("./clx");
/** [MS-DOC] 2.4.1 Retrieving Text */
function getDocTxt(fib, docStream, tableStream) {
    var fibRgLw = fib.fibRgLw, fibRgFcLcbBlob = fib.fibRgFcLcbBlob;
    var fcClx = fibRgFcLcbBlob.fcClx, lcbClx = fibRgFcLcbBlob.lcbClx;
    var clx = tableStream.slice(fcClx, fcClx + lcbClx);
    var plcPcd = clx_1.parseClx(clx);
    var txt = clx_1.getTxt(fibRgLw, plcPcd, docStream);
    /* grab the body text */
    return txt.length == fibRgLw.ccpText ? txt.slice(0, -1) : txt.slice(0, fibRgLw.ccpText);
}
function parse_cfb(file) {
    /* [MS-DOC] 2.4.1 Retrieving Text */
    var wordDocument = cfb_1.find(file, "/WordDocument");
    var wordStream = wordDocument.content;
    var fib = fib_1.readFib(wordStream);
    var tableName = fib.base.fWhichTblStm === 1 ? "/1Table" : "/0Table";
    var table = cfb_1.find(file, tableName);
    var tableStream = table.content;
    var text = getDocTxt(fib, wordStream, tableStream);
    /* TODO: 2.8.25 strip fields */
    text = text.replace(/\x13.*?\x14(.*?)\x15/g, "$1");
    text = text.replace(/\x13.*?\x15/g, "");
    /* TODO: 1.3.5 Inline Picture 0x01, Floating 0x08 */
    text = text.replace(/[\x01\x08]/g, "");
    /* TODO: 2.4.3 Table cell mark is 0x07 */
    text = text.replace(/\x07/g, "\r");
    // TODO: correctly split into paragraphs
    // getParagraphs(fib, wordStream, tableStream);
    var doc = { p: [] };
    var para = { elts: [] };
    para.elts.push({ t: "s", v: text });
    doc.p.push(para);
    return doc;
}
exports.parse_cfb = parse_cfb;
function readFile(filePath) {
    var file = cfb_1.read(filePath, { type: "file" });
    return parse_cfb(file);
}
exports.readFile = readFile;
function read(data) {
    var file = cfb_1.read(data, { type: "buffer" });
    return parse_cfb(file);
}
exports.read = read;
//# sourceMappingURL=index.js.map