// File meant to be run on its own
import * as fs from 'fs'

enum FileType {
  docx = "504b0304",
  docm = "504b0304",
  doc = "d0cf11e0",
}

function readFile(path: string): void {
  read(fs.readFileSync(path));
}

function read(buffer: Buffer): void {
  const magicHexAsString: string = buffer.toString('hex');
  if (magicHexAsString.slice(0, FileType.docm.length) == FileType.docm) {
    console.log("It is a .docm file");
  } else if (magicHexAsString.slice(0, FileType.doc.length) == FileType.doc) {
    console.log("It is a .doc file");
  } else {
    console.log("Unknown file type");
  }
}

const path: string = './testFiles/test.doc'
readFile(path);