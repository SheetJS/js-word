import { JSDOM } from "jsdom";
import { WJSDoc, WJSPara } from "../../types";

export function parse_str(data: string): WJSDoc {
  const dom: JSDOM = new JSDOM(data);
  const doc: WJSDoc = {p:[]};
  const para: WJSPara = {elts:[]};
  para.elts.push({t: "s", v: dom.window.document.querySelector('body').textContent});
  doc.p.push(para);
  return doc;
}

export function read(data: Buffer): WJSDoc {
  return parse_str(data.toString());
}