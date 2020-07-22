"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var cfb_1 = require("cfb");
var jsdom_1 = require("jsdom");
/* 5.1.3 <text:p> children */
function process_para(child, root) {
    switch (child.nodeType) {
        case 1 /*ELEMENT_NODE*/:
            var element = child;
            if (element.tagName.match(/draw:/))
                return;
            if (element.tagName == "text:s")
                root.elts.push({ t: "s", v: " ".repeat(+element.getAttribute("c") || 1) });
            if (element.tagName == "text:tab")
                root.elts.push({ t: "s", v: "\t".repeat(+element.getAttribute("c") || 1) });
            element.childNodes.forEach(function (child) { return process_para(child, root); });
            break;
        case 3 /*TEXT_NODE*/:
            root.elts.push({ t: "s", v: child.textContent });
            break;
        default: throw "unsupported node type " + child.nodeType;
    }
}
/* 9.1.4 <table:table-cell> children */
function process_td(tdelt) {
    var tablecell = { t: "c", p: [] };
    tdelt.childNodes.forEach(function (child) {
        var data = process_body_elt(child);
        if (data)
            tablecell.p.push(data);
    });
    return tablecell;
}
/* 9.1.3 <table:table-row> children */
function process_tr(trelt) {
    var tablerow = { t: "r", c: [] };
    trelt.childNodes.forEach(function (child) {
        if (child.nodeType != 1)
            return;
        var element = child;
        switch (element.tagName) {
            case "table:table-cell":
                tablerow.c.push(process_td(element));
                break;
            // table:covered-table-cell
            default: throw "ODT tablerow unsupported " + element.tagName + " element";
        }
    });
    return tablerow;
}
/* 9.1.2 <table:table> children */
function process_table(tablelt) {
    var table = { t: "t", r: [] };
    tablelt.childNodes.forEach(function (child) {
        if (child.nodeType != 1)
            return;
        var element = child;
        switch (element.tagName) {
            case "table:table-column": break;
            case "table:table-row":
                table.r.push(process_tr(element));
                break;
            default: throw "ODT table unsupported " + element.tagName + " element";
        }
    });
    return table;
}
/* 8.4 <text:illustration-index> children */
function process_illustration_index(child, root) {
    switch (child.nodeType) {
        case 1 /*ELEMENT_NODE*/:
            var element = child;
            switch (element.tagName) {
                case "text:illustration-index-source": break;
                case "text:index-body":
                    element.childNodes.forEach(function (child) { return process_index_body(child, root); });
                    break;
                default: throw "ODT illustration-index unsupported " + element.tagName + " element";
            }
    }
}
/* 8.2.2 <text:index-body> children */
function process_index_body(child, root) {
    switch (child.nodeType) {
        case 1:
            var element = child;
            switch (element.tagName) {
                case "text:index-title":
                    element.childNodes.forEach(function (child) { return process_index_body(child, root); });
                    break;
                case "text:p":
                case "text:h":
                    var para_1 = { elts: [] };
                    element.childNodes.forEach(function (child) {
                        process_para(child, para_1);
                    });
                    root.push(para_1);
                    break;
            }
    }
}
/* 3.4 <office:text> children */
function process_body_elt(child, root) {
    if (root === void 0) { root = false; }
    var para = { elts: [] };
    switch (child.nodeType) {
        case 1 /*ELEMENT_NODE*/:
            var element = child;
            switch (element.tagName) {
                // dr3d:scene
                // draw:*
                // office:forms
                // table:*
                // text:*
                case "text:p":
                case "text:h":
                    element.childNodes.forEach(function (child) { return process_para(child, para); });
                    return para;
                case "table:table":
                    para.elts.push(process_table(element));
                    return para;
                // text:illustration-index
                case "text:illustration-index":
                    var paraArray_1 = [];
                    console.log();
                    element.childNodes.forEach(function (child) { return process_illustration_index(child, paraArray_1); });
                    return paraArray_1;
                // case "text:illustration-index-body":
                // recursively parse children
                // para.elts.push(process_illustration_index(element, para));
                // return para;
                // text:list
                case "text:illustration-index-source":
                case "text:sequence-decls":
                case "text:variable-decls":
                case "text:tracked-changes":
                    if (root)
                        break;
                default: throw "ODT body unsupported " + element.tagName + " element";
            }
            break;
    }
}
/**
 * Grabs the text content of an odt file
 *
 * @param {string} file path to .odt file
 * @return {string} text content of file
 */
function parse_cfb(file) {
    // Read the content.xml of the file
    var buf = cfb_1.find(file, '/content.xml').content;
    // Parse with JSDOM
    var dom = new jsdom_1.JSDOM(buf.toString(), { contentType: "text/xml" });
    var doc = { p: [] };
    // Use querySelector to grab elements from content.xml
    // We can grab any element, not just text:p
    /* 3.1.3.2 <office:document-content> */
    var rootelt = dom.window.document.children[0];
    /* 3.3 <office:body> */
    var bodyelt = rootelt.querySelector("office\\:document-content > office\\:body");
    /* 3.4 <office:text> */
    var textelt = bodyelt.querySelector("office\\:body > office\\:text");
    textelt.childNodes.forEach(function (child) {
        var res = process_body_elt(child, true);
        if (res) {
            if (res.push == undefined) {
                doc.p.push(res);
            }
            else {
                res.forEach(function (para) {
                    doc.p.push(para);
                });
            }
        }
    });
    return doc;
}
exports.parse_cfb = parse_cfb;
//# sourceMappingURL=index.js.map