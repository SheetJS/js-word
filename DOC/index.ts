import { read, find } from "cfb";
import { readFib } from "./fib";
import { getDocTxt } from "./clx";

function parseDoc(filePath: string) {
  const file = read(filePath, {type: "file"});
  const wordDocument = find(file, "/WordDocument");
  const wordStream = wordDocument.content as Buffer;
  const fib = readFib(wordStream);
  const tableName = fib.base.fWhichTblStm === 1 ? "/1Table" : "/0Table";
  const table = find(file, tableName);
  const tableStream = table.content as Buffer;
  const text = getDocTxt(fib, wordStream, tableStream);
  return text;
}

console.log(parseDoc("test1.doc"));
