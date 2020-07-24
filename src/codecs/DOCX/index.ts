import { read as readCFB, find, CFB$Container } from "cfb";
import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara, WJSTable, WJSTableRow, WJSTableCell, WJSRel } from "../../types";

/* ECMA 17.3.1.22 p CT_P */
function process_para(child: Node, root: WJSPara, file:CFB$Container, path: string) {
  switch (child.nodeType) {
    case 1 /* ELEMENT_NODE */:
      const element = (child as Element);
      switch (element.tagName) {
        case "w:r":
        case "w:sdt":
        case "w:sdtContent":
        case "w:customXml":
          element.childNodes.forEach((child) => process_para(child, root, file, path));
          break;
        case "w:t":
          root.elts.push({ t: "s", v: child.textContent }); break;
        case "w:hyperlink": // TODO: store actual hyperlink?
          element.childNodes.forEach((child) => process_para(child, root, file, path));
          break;
        case "w:br":
          break;

        case "w:annotationRef":
        case "w:bookmarkEnd":
        case "w:bookmarkStart":
        case "w:commentRangeStart":
        case "w:commentRangeEnd":
        case "w:commentReference": //TODO: add reference support
        case "w:del":
        case "w:drawing":
        case "w:endnoteReference":
          parse_relationship(file, path);
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
        case "w:sym": 
        case "w:tab": //TODO: Add tab support
        case "mc:AlternateContent":
        case "m:oMath":
        case "m:oMathPara":
        case "w16se:sym":
          break;
        default: throw `DOCX para unsupported ${element.tagName} element`
      }
      break;
  }
};

function process_tc(tcelt: Element, file:CFB$Container, path: string): WJSTableCell {
  const tableCell: WJSTableCell = { t: "c", p: [] };
  const para: WJSPara = {elts: []};
  tcelt.childNodes.forEach(child => {
    const data = process_body_elt(child, false, file, path);
    if (data) tableCell.p.push(data);
    // console.log(tableCell.p[0]);
  })
  return tableCell;
}

function process_tr(trelt: Element, file:CFB$Container, path: string): WJSTableRow {
  const tableRow: WJSTableRow = { t: "r", c: [] };
  trelt.childNodes.forEach(child => {
    if(child.nodeType != 1) return;
    const element = (child as Element);
    switch(element.tagName) {
      case "w:trPr": 
      case "w:sdt":
      case "w:tblPrEx": 
      case "w:commentRangeEnd": 
      break;
      case "w:tc" :
        tableRow.c.push(process_tc(element, file, path));
        // console.log("cells: ", tableRow.c);
        break;
      default: throw `DOCX tablerow unsupported ${element.tagName} element`
    }
  });
  return tableRow

}

function process_table(tablelt: Element, file:CFB$Container, path: string): WJSTable {
  const table: WJSTable = { t: "t", r: [] };
  tablelt.childNodes.forEach(child => {
    if (child.nodeType != 1) return;
    const element = (child as Element);
    switch (element.tagName) {
      case "w:tblPr":
      case "w:tblGrid":
      case "w:bookmarkEnd":
        break;
      case "w:tr":
        table.r.push(process_tr(element, file, path));
        // console.log("rows: ", table.r);
        break;
      default: throw `DOCX table unsuported ${element.tagName} element`
    }
  });
  return table;
}

function process_body_elt(child: ChildNode, root: boolean = false, file:CFB$Container, path: string): WJSPara | void {
  const para: WJSPara = { elts: [] };
  switch (child.nodeType) {
    case 1: /* ELEMENT_NODE */
      const element = (child as Element);
      switch (element.tagName) {
        case "w:p":
          element.childNodes.forEach((child) => process_para(child, para, file, path));
          return para;
        case "w:tbl":
          para.elts.push(process_table(element, file, path));
          return para;
        // console.log("tables: ", para.elts);
        case "w:customXML":
          if (root) break;
        case "w:sectPr":
        case "w:bookmarkStart":
        case "w:bookmarkEnd":
        case "w:commentRangeEnd":
        case "w:moveFromRangeEnd":
        case "w:tcPr":
        case "w:sdt":
        case "w:altChunk": //TODO: implicit/explicit link handeling
        case "mc:AlternateContent":
          break;
        default: throw `DOCX body unsupported ${element.tagName} element`
      }
      break;
  }
}

function parse_relationship(file: CFB$Container, path: string): any {
  const n = path.lastIndexOf("/");
  path = path.slice(0, n+1) + "_rels/" + path.slice(n+1) + ".rels";
  const dom = get_file_content(file, path);
  const rootelt = dom.window.document.firstChild;
  rootelt.childNodes.forEach(element => {
    const id = element.attributes
  });
}

function get_file_content(file: CFB$Container, path: string): JSDOM {
  const buf = find(file, path).content;

  // Parse with JSDOM
  return new JSDOM((buf as Buffer).toString(), { contentType: "text/xml" });
} 

export function parse_cfb(file: CFB$Container): WJSDoc {
  const docx: WJSDoc = { p: [], rels: [] }

  const path = '/word/document.xml'

  const dom = get_file_content(file, path);

  const rootelt = dom.window.document.children[0];

  const bodyelt = rootelt.querySelector("w\\:document > w\\:body");

  bodyelt.childNodes.forEach(child => {
    const res = process_body_elt(child, true, file, path);
    if (res) docx.p.push(res);
  });

  return docx;

}