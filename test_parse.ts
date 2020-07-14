import { readFile } from "./src";

const filename = process.argv[2];
const doc = readFile(filename);
console.log(doc);
