"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var cfb_1 = require("cfb");
var jsdom_1 = require("jsdom");
/* ECMA 17.3.1.22 p CT_P */
function process_para(child, root) {
    switch (child.nodeType) {
        case 1 /* ELEMENT_NODE */:
            var element = child;
            switch (element.tagName) {
                case "w:r":
                case "w:sdt":
                case "w:sdtContent":
                case "w:customXml":
                    element.childNodes.forEach(function (child) { return process_para(child, root); });
                    break;
                case "w:t":
                    root.elts.push({ t: "s", v: child.textContent });
                    break;
                case "w:hyperlink": // TODO: store actual hyperlink?
                    element.childNodes.forEach(function (child) { return process_para(child, root); });
                    break;
                case "w:br":
                    break;
                case "w:annotationRef":
                case "w:bookmarkEnd":
                case "w:bookmarkStart":
                case "w:commentRangeStart":
                case "w:commentRangeEnd":
                case "w:commentReference":
                case "w:del":
                case "w:drawing":
                case "w:endnoteReference":
                case "w:fldChar":
                case "w:fldSimple":
                case "w:footnoteReference":
                case "w:ins":
                case "w:instrText":
                case "w:lastRenderedPageBreak":
                case "w:moveFrom":
                case "w:moveFromRangeStart":
                case "w:moveFromRangeEnd":
                case "w:moveTo":
                case "w:moveToRangeStart":
                case "w:moveToRangeEnd":
                case "w:noBreakHyphen":
                case "w:object":
                case "w:pict":
                case "w:pPr":
                case "w:proofErr":
                case "w:rPr":
                case "w:ruby":
                case "w:sdtEndPr":
                case "w:sdtPr":
                case "w:sectPr":
                case "w:smartTag":
                case "w:softHyphen":
                case "w:sym": //TODO: Read documentation about this
                case "w:tab":
                case "mc:AlternateContent":
                case "m:oMath":
                case "m:oMathPara":
                case "w16se:sym":
                    break;
                default: throw "DOCX para unsupported " + element.tagName + " element";
            }
            break;
    }
}
;
function process_table(tablelt) {
    var table = { t: "t", r: [] };
    return table;
}
function process_body_elt(child, root) {
    if (root === void 0) { root = false; }
    var para = { elts: [] };
    switch (child.nodeType) {
        case 1: /* ELEMENT_NODE */
            var element = child;
            switch (element.tagName) {
                case "w:p":
                    element.childNodes.forEach(function (child) { return process_para(child, para); });
                    return para;
                case "w:tbl":
                    // para.elts.push(process_table(element));
                    console.log(element);
                case "w:customXML":
                    if (root)
                        break;
                case "w:sectPr":
                case "w:bookmarkStart":
                case "w:bookmarkEnd":
                case "w:commentRangeEnd":
                case "w:moveFromRangeEnd":
                case "w:sdt":
                case "w:altChunk": //TODO: implicit/explicit link handeling
                case "mc:AlternateContent":
                    break;
                default: throw "DOCX body unsupported " + element.tagName + " element";
            }
            break;
    }
}
function parse_cfb(file) {
    // Get content of document.xml
    var buf = cfb_1.find(file, "/word/document.xml").content;
    // Parse with JSDOM
    var dom = new jsdom_1.JSDOM(buf.toString(), { contentType: "text/xml" });
    var docx = { p: [] };
    var rootelt = dom.window.document.children[0];
    var bodyelt = rootelt.querySelector("w\\:document > w\\:body");
    bodyelt.childNodes.forEach(function (child) {
        var res = process_body_elt(child, true);
        if (res)
            docx.p.push(res);
    });
    return docx;
    // const paragraphs = dom.window.document.querySelectorAll("w\\:p");
    // const para = parse_para(paragraphs);
}
exports.parse_cfb = parse_cfb;
//# sourceMappingURL=index.js.map