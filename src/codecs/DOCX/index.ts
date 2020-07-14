import { read as readCFB, find, CFB$Container } from "cfb";
import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara } from "../../types";

/* ECMA 17.3.1.22 p CT_P */
function process_para(child: Node, root: WJSPara) {
  switch(child.nodeType) {
    case 1 /* ELEMENT_NODE */:
      const element = (child as Element);
      switch(element.tagName) {
        case "w:r":
        case "w:sdt":
        case "w:sdtContent":
          element.childNodes.forEach((child) => process_para(child, root));
          break;
        case "w:t":
          root.elts.push({t: "s", v: child.textContent}); break;
        case "w:hyperlink": // TODO: store actual hyperlink?
          element.childNodes.forEach((child) => process_para(child, root));
          break;
        case "w:br":
          break;
        case "w:pPr":
        case "w:rPr":
        case "w:bookmarkStart":
        case "w:bookmarkEnd":
        case "w:sectPr":
        case "w:pict":
        case "w:lastRenderedPageBreak":
        case "w:drawing":
          break;
        default: throw `DOCX para unsupported ${element.tagName} element`
      }
    break;
  }
};

function process_body_elt(child: ChildNode, root: boolean = false): WJSPara|void {
  const para: WJSPara = {elts : []};
  switch(child.nodeType) {
    case 1: /* ELEMENT_NODE */
    const element = (child as Element);
    switch(element.tagName) {
      case "w:p":
        element.childNodes.forEach((child) => process_para(child, para));
        return para;
      case "w:tbl":
      case "w:customXML":
        if(root) break;
      case "w:sectPr": break;
      default: throw `DOCX body unsupported ${element.tagName} element`
    }
    break;
  }
}

export function parse_cfb(file: CFB$Container): WJSDoc {
  // Get content of document.xml
  const buf = find(file, "/word/document.xml").content;

  // Parse with JSDOM
  const dom = new JSDOM((buf as Buffer).toString(), {contentType: "text/xml"});

  const docx: WJSDoc = {p: []}

  const rootelt = dom.window.document.children[0];

  const bodyelt = rootelt.querySelector("w\\:document > w\\:body");

  bodyelt.childNodes.forEach(child => {
    const res = process_body_elt(child, true);
    if(res) docx.p.push(res);
  })

  return docx;

  // const paragraphs = dom.window.document.querySelectorAll("w\\:p");

  // const para = parse_para(paragraphs);

}