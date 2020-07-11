import { read as readCFB, find, CFB$Container } from "cfb";
import { parse_cfb as parse_DOCX } from "./DOCX";
import { parse_cfb as parse_DOC } from "./DOC";
import { parse_cfb as parse_MHT } from "./MHT";
import { parse_cfb as parse_ODT } from "./ODT";
import { parse_str as parse_TXT } from "./TXT";
import { read as read_HTML, parse_str as parse_HTML } from "./HTML";
import { read as read_RTF } from "./RTF";
import { read as read_XML, parse_str as parse_XML } from "./XML";
import { readFileSync } from "fs";
import { WJSDoc } from "../types";

export function parse_cfb(file: CFB$Container): WJSDoc {
  if(find(file, "/WordDocument")) return parse_DOC(file);
  if(find(file, "/CONTENTS")) throw "Unsupported Works WPS file";
  if(find(file, "/MM") || find(file, "/MN0")) throw "Unsupported Works WPS file";
  throw "Unsupported CFB file";
}

export function parse_zip(file: CFB$Container): WJSDoc {
  if(find(file, "/[Content_Types].xml")) return parse_DOCX(file);
  if(find(file, "/META-INF/manifest.xml")) return parse_ODT(file);
  throw "Unsupported ZIP file";
}

/** read JS string */
export function read_str(data: string): WJSDoc {
  const header = data.slice(0,17);

  /* MIME text is technically 7-bit so type: "binary" is acceptable */
  if(header == "MIME-Version: 1.0") return parse_MHT(readCFB(data, {type: "binary"}));
  if(header.slice(0,5) == "<?xml") return parse_XML(data);
  if(header.slice(0,5) == "<html") return parse_HTML(data);

  /* TODO: more formats here */
  if(header.split("").map(c => c.charCodeAt(0)).every(cc => cc == 9 || cc == 10 || cc == 13 || cc >= 0x20)) return parse_TXT(data.toString());

  if(!header.length) return { p: [] };
  throw "Unsupported string";
}

// TODO: replace this with a proper structure
export function read(data: Buffer): WJSDoc {
  const header = data.slice(0,17).toString("binary");

  if(header.slice(0,3) == "\xef\xbb\xbf") return read_str(data.slice(3).toString());
  if(header.slice(0,2) == "\xff\xfe") return read_str(data.slice(2).toString("utf16le"));
  /* One convenient use of buf.swap16() is to perform a fast in-place conversion between UTF-16 little-endian and UTF-16 big-endian */
  if(header.slice(0,2) == "\xfe\xff") return read_str(data.slice(2).swap16().toString("utf16le"));

  if(header == "MIME-Version: 1.0") return parse_MHT(readCFB(data, {type: "buffer"}));
  if(header.slice(0,6) == "{\\rtf1") return read_RTF(data);
  if(header.slice(0,5) == "<?xml") return read_XML(data);
  if(header.slice(0,5) == "<html") return read_HTML(data);
  if(header.slice(0,4) == "\xD0\xCF\x11\xE0") return parse_cfb(readCFB(data, {type: "buffer"}));
  if(header.slice(0,4) == "PK\x03\x04") return parse_zip(readCFB(data, {type: "buffer"}));

  // TODO: better plaintext check
  if(header.split("").map(c => c.charCodeAt(0)).every(cc => cc == 9 || cc == 10 || cc == 13 || cc >= 0x20 && cc <= 0x7F)) return parse_TXT(data.toString());

  throw "Unsupported file";
}

export function readFile(path: string): WJSDoc {
  return read(readFileSync(path));
}