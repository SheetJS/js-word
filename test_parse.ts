import { readFile } from "./src";

const filename = process.argv[2];
console.log(readFile(filename));
