import { CFB$Container, find } from "cfb";
import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara, WJSTable, WJSTableRow, WJSTableCell } from "../../types";

/* 5.1.3 <text:p> children */
function process_para(child: Node, root: WJSPara) {
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      if(element.tagName.match(/draw:/)) return;
      if(element.tagName == "text:s") root.elts.push({t:"s", v: " ".repeat(+element.getAttribute("c") || 1)});
      element.childNodes.forEach((child) => process_para(child, root));
      break;
    case 3 /*TEXT_NODE*/: root.elts.push({t:"s", v: child.textContent}); break;
    default: throw "unsupported node type " + child.nodeType;
  }
}

/* 9.1.4 <table:table-cell> children */
function process_td(tdelt: Element): WJSTableCell {
  const tablecell: WJSTableCell = { t: "c", p: [] };
  tdelt.childNodes.forEach(child => {
    const data = process_body_elt(child);
    if(data) tablecell.p.push(data);
  });
  return tablecell;
}

/* 9.1.3 <table:table-row> children */
function process_tr(trelt: Element): WJSTableRow {
  const tablerow: WJSTableRow = { t: "r", c: [] };
  trelt.childNodes.forEach(child => {
    if(child.nodeType != 1) return;
    const element = (child as Element);
    switch(element.tagName) {
      case "table:table-cell":
        tablerow.c.push(process_td(element));
        break;
      // table:covered-table-cell
      default: throw `ODT tablerow unsupported ${element.tagName} element`
    }
  });
  return tablerow;
}

/* 9.1.2 <table:table> children */
function process_table(tablelt: Element): WJSTable {
  const table: WJSTable = { t: "t", r: [] };
  tablelt.childNodes.forEach(child => {
    if(child.nodeType != 1) return;
    const element = (child as Element);
    switch(element.tagName) {
      case "table:table-column": break;
      case "table:table-row":
        table.r.push(process_tr(element));
        break;
      default: throw `ODT table unsupported ${element.tagName} element`
    }
  });
  return table;
}

/* 3.4 <office:text> children */
function process_body_elt(child: ChildNode, root: boolean = false): WJSPara|void {
  const para: WJSPara = { elts: []};
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      switch(element.tagName) {
        // dr3d:scene
        // draw:*
        // office:forms
        // table:*
        // text:*
        case "text:p":
        case "text:h":
          element.childNodes.forEach((child) => process_para(child, para));
          return para;
        case "table:table":
          para.elts.push(process_table(element));
          return para;
        // text:list
        case "text:sequence-decls":
        case "text:variable-decls":
        case "text:illustration-index":
        case "text:tracked-changes":
          if(root) break;
        default: throw `ODT body unsupported ${element.tagName} element`
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
export function parse_cfb(file: CFB$Container): WJSDoc {
  // Read the content.xml of the file
  const buf = find(file, '/content.xml').content;

  // Parse with JSDOM
  const dom = new JSDOM((buf as Buffer).toString(), {contentType: "text/xml"});

  const doc: WJSDoc = {p: []};

  // Use querySelector to grab elements from content.xml
  // We can grab any element, not just text:p

  /* 3.1.3.2 <office:document-content> */
  const rootelt = dom.window.document.children[0];

  /* 3.3 <office:body> */
  const bodyelt = rootelt.querySelector("office\\:document-content > office\\:body");

  /* 3.4 <office:text> */
  const textelt = bodyelt.querySelector("office\\:body > office\\:text");
  textelt.childNodes.forEach(child => {
    const res = process_body_elt(child, true);
    if(res) doc.p.push(res);
  });

  return doc;
}
