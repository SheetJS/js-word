import { JSDOM } from 'jsdom';
import { WJSDoc, WJSPara } from '../../types';

export function parse_str(data: string): WJSDoc {
  // dom.window.document.children[0].tagName)
  const listOfParagraphs: Node[] = [];
  const dom: JSDOM = new JSDOM(data, {contentType: "text/xml"});
  const dfs = (node: Node): void => {
    if (node.hasChildNodes) {
      node.childNodes.forEach((node: ChildNode) => {
        if (node.nodeName === "w:p") listOfParagraphs.push(node);
        dfs(node);
      });
    }
  };
  dfs(dom.window.document)

  const out: WJSDoc = {p: []};
  listOfParagraphs.forEach((node: Node) => {
    node.childNodes.forEach((child) => {
      const para: WJSPara = { elts: [] };
      para.elts.push({ t: "s", v: child.textContent });
      out.p.push(para);
    });
  });

  return out;
};

export function read(data: Buffer): WJSDoc {
  return parse_str(data.toString());
}