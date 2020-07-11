import { WJSDoc, WJSParaElement, WJSPara, WJSTextRun } from "../../types";

function write_para_elt_str(elt: WJSParaElement, opts?: any): string {
  const RS = opts && opts.RS || "\n";
  switch(elt.t) {
    case "s": return elt.v;
    case "t":
      return elt.r.map(tr =>
        tr.c.map(tc =>
          tc.p.map(p =>
            write_para_str(p)
          ).join(RS)
        ).join(RS)
      ).join(RS);
    default: throw `Cannot generate plaintext for ${(elt as any).t} elements`;
  }
}

function write_para_str(para: WJSPara, opts?: any): string {
  return para.elts.map(elt => write_para_elt_str(elt, opts)).join("")
}

export function write_str(doc: WJSDoc, opts?: any): string {
  const RS = opts && opts.RS || "\n";
  var o = doc.p.map(para => write_para_str(para, opts)).join(RS) + RS;
  return o;
}

export function write_buf(doc: WJSDoc, opts?: any): Buffer {
  return Buffer.from(write_str(doc, opts));
}

/* TODO: something more reasonable */
export function parse_str(data: string): WJSDoc {
  const doc: WJSDoc = { p: [] };
  let texts = data.split(/\r\n?|\n/);
  if(!texts[texts.length-1] && data.slice(-2).match(/[\r\n]$/)) texts = texts.slice(0, -1);
  texts.forEach(line => {
    const t: WJSTextRun = { t: "s", v: line };
    const para: WJSPara = { elts: [t] };
    doc.p.push(para);
  });
  return doc;
}