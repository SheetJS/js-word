import { read as readCFB, find, CFB$Container } from "cfb";
import { readFib, Fib } from "./fib";
import { parseClx, getTxt } from "./clx";
import { WJSDoc, WJSPara } from "../../types";

/** [MS-DOC] 2.4.1 Retrieving Text */
function getDocTxt(fib: Fib, docStream: Buffer, tableStream: Buffer): string {
  const { fibRgLw, fibRgFcLcbBlob } = fib;
  const { fcClx, lcbClx } = fibRgFcLcbBlob;
  const clx = tableStream.slice(fcClx, fcClx + lcbClx);
  const plcPcd = parseClx(clx);
  const txt = getTxt(fibRgLw, plcPcd, docStream);

  /* grab the body text */
  return txt.length == fibRgLw.ccpText ? txt.slice(0, -1) : txt.slice(0, fibRgLw.ccpText);
}


export function parse_cfb(file: CFB$Container): WJSDoc {
  /* [MS-DOC] 2.4.1 Retrieving Text */
  const wordDocument = find(file, "/WordDocument");
  const wordStream = wordDocument.content as Buffer;
  const fib = readFib(wordStream);

  const tableName = fib.base.fWhichTblStm === 1 ? "/1Table" : "/0Table";
  const table = find(file, tableName);
  const tableStream = table.content as Buffer;

  let text = getDocTxt(fib, wordStream, tableStream);

  /* TODO: 2.8.25 strip fields */
  text = text.replace(/\x13[^\x13]*\x14(.*?)\x15/g, "$1")
  text = text.replace(/\x13.*?\x15/g, "")

  /* TODO: 1.3.5 Inline Picture 0x01, Floating 0x08 */
  text = text.replace(/[\x01\x08]/g, "");

  /* TODO: 2.4.3 Table cell mark is 0x07 */
  text = text.replace(/\x07/g, "\r");

  // TODO: correctly split into paragraphs
  // getParagraphs(fib, wordStream, tableStream);

  const doc: WJSDoc = { p: [] };
  const para: WJSPara = { elts: [] };
  para.elts.push({ t: "s", v: text });
  doc.p.push(para);
  return doc;
}

export function readFile(filePath: string): WJSDoc {
  const file = readCFB(filePath, { type: "file" });
  return parse_cfb(file);
}

export function read(data: Buffer): WJSDoc {
  const file = readCFB(data, { type: "buffer" });
  return parse_cfb(file);
}
