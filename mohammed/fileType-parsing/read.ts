 import { fileTypeHandler } from './fileTypeHandler';
 import * as fs from 'fs';
 import {Trie, trieNode} from './trie';

const signatureToType: { [byte: number]: string} = {
    0x50: 'zip',
    0xd0: 'doc',
    0x7b: 'rtf',
    0x3c: 'xml',
}

const keys = Object.keys(fileTypeHandler);

const read = (buffer: Buffer): string  => {
    for (const key of keys) {
        if (fileTypeHandler[key].validate(buffer)) {
            return (`${signatureToType[fileTypeHandler[key].signature]}`);
        }
    }   
}

function readFile(path: string): void {
    console.log("It is a", read(fs.readFileSync(path)), "file");
}

const paths: string[] = [
    './testFiles/test.rtf',
    './testFiles/test.xml',
    './testFiles/test.doc',
    './testFiles/test.docm',
]

paths.forEach(path => readFile(path))