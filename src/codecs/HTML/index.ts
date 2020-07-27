import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara } from "../../types";

export function parse_str(data: string): WJSDoc {
  const dom: JSDOM = new JSDOM(data);
  const doc: WJSDoc = {p:[]};
  const para: WJSPara = {elts:[]};

  // Assuming first child is always <html...>
  const root = dom.window.document.childNodes[0];
  dfs(root, para);
  doc.p.push(para);
  return doc;
}

const dfs = (element: ChildNode , para: WJSPara): WJSPara => {
  element.childNodes.forEach(child => {
    switch (child.nodeName) {
      case "P":
        para.elts.push({t: "s", v: child.textContent});
      default: /*throw `DOCX table unsuported ${child.nodeName} element`*/
    }
    dfs(child, para);
  })
  return para;
}

export function read(data: Buffer): WJSDoc {
  return parse_str(data.toString());
}