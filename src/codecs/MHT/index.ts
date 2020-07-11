import { CFB$Container, CFB$Entry } from "cfb";
import { WJSDoc } from "../../types";
import { read as read_HTML } from "../HTML";

export function parse_cfb(file: CFB$Container): WJSDoc {
  const firstHtmlIdx: number = file.FullPaths.findIndex((path: string) => /\.html?$/.test(path));
  const entry: CFB$Entry = file.FileIndex[firstHtmlIdx];
  if(!entry || !entry.content) throw "MHT missing an HTML entry";
  return read_HTML(Buffer.from(entry.content as Buffer));
}