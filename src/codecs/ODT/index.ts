import { CFB$Container, find } from "cfb";
import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara, WJSTable, WJSTableRow, WJSTableCell, WJSParaElement } from "../../types";

/* 5.1.3 <text:p> children */
function process_para(child: Node, root: WJSPara) {
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      if(element.tagName.match(/draw:/)) return;
      switch(element.tagName) {
        case "text:s":
          root.elts.push({ t: "s", v: " ".repeat(+element.getAttribute("c") || 1) });
          break;
        case "text:tab":
          root.elts.push({ t: "s", v: "\t".repeat(+element.getAttribute("c") || 1) });
          break;
        case "text:note":
          element.childNodes.forEach((child) => {
            process_note(child, root);
          });
          return;
        default:
          element.childNodes.forEach((child) => {
            process_para(child, root);
          });
          break;
      }
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
    if(data) tablecell.p.push(data as WJSPara);
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

/* 8.4 <text:illustration-index> children */
function process_illustration_index(child: Node, root: WJSPara[]) {
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      switch(element.tagName) {
        case "text:illustration-index-source": break;
        case "text:index-body":
          element.childNodes.forEach((child) => process_index_body(child, root));
          break;
        default: throw `ODT illustration-index unsupported ${element.tagName} element`
      }
  }
}

/* 8.2.2 <text:index-body> children */
function process_index_body(child: Node, root: WJSPara[]) {
  switch(child.nodeType) {
    case 1:
      const element = (child as Element);
      switch(element.tagName) {
        case "text:index-title":
          element.childNodes.forEach((child) => process_index_body(child, root));
          break;
        case "text:p":
        case "text:h":
          const para: WJSPara = { elts: [] };
          element.childNodes.forEach((child) => {
            process_para(child, para);
          });
          root.push(para);
          break;
      }
  }
}

/* <text:list> children */
function process_list(child: Node, root: WJSPara[], index: number) {
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      switch(element.tagName) {
        // Recursive case
        case "text:list-item":
          element.childNodes.forEach((child) => {
            process_list_item(child, root, index);
            index++;
          });
          break;
      }
  }
}

/* <text:list-item> children */
function process_list_item(child: Node, root: WJSPara[], iter: number) {
  switch(child.nodeType) {
    case 1:
      const element = (child as Element);
      switch(element.tagName) {
        case "text:list":
          let index : number = 1;
          element.childNodes.forEach((child) => {
            process_list(child, root, index);
            index++;
          });
          break;
        case "text:h":
        case "text:p":
          const para : WJSPara = { elts: [] };
          para.elts.push({t: "s", v: "" + iter + ". "});
          element.childNodes.forEach((child) => {
            process_para(child, para);
          });
          root.push(para);
          break;
      }
  }
}

function process_note(child: Node, root: WJSPara) {
  switch(child.nodeType) {
    case 1 /*ELEMENT_NODE*/:
      const element = (child as Element);
      switch(element.tagName) {
        case "text:note-citation":
          element.childNodes.forEach((child) => process_para(child, root));
          break;
        case "text:note-body":
          element.childNodes.forEach((child) => process_note(child, root));
          break;
        case "text:p":
        case "text:h":
          element.childNodes.forEach((child) => process_para(child, root));
          break;
      }
      break;
  }
}

/* 3.4 <office:text> children */
function process_body_elt(child: ChildNode, root: boolean = false): WJSPara[]|WJSPara|void {
  const para: WJSPara = { elts: []};
  const paraArray: WJSPara[] = [];
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
        // text:illustration-index
        case "text:illustration-index":
          element.childNodes.forEach((child) => process_illustration_index(child, paraArray));
          return paraArray;
        // text:list
        case "text:list":
          let index : number = 1;
          element.childNodes.forEach((child) => {
            process_list(child, paraArray, index);
            index++;
          });
          return paraArray;
        case "text:note":
        case "text:sequence-decls":
        case "text:variable-decls":
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
    if(res) {
      // Check if res is an array
      if ((res as WJSPara[]).push == undefined) {
        // If not, push to doc
        doc.p.push(res as WJSPara);
      } else {
        // If yes, push each paragraph to doc
        (res as WJSPara[]).forEach(para => {
          doc.p.push(para);
        });
      }
    }
  });

  return doc;
}
