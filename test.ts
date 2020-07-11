import 'mocha';
import * as assert from 'assert';
import * as fs from 'fs';
import * as WORD from "./src";
import { sync } from 'glob';

const formats = process.env.FMTS ? process.env.FMTS.split(",") : ["doc", "docx", "htm", "html", "mht", "odt", "rtf", "xml", "txt"];
describe("test_files", () => {
  formats.filter(fmt => fmt != "txt").forEach(fmt => {
    describe(fmt, () => {
      const files = sync(`test_files/**/*.${fmt}`);
      files.forEach(fn => {
        if(fs.existsSync(fn + ".skip")) return;
        (fs.existsSync(fn + ".txt") ? it : it.skip)(fn, () => {
          const doc: WORD.WJSDoc = WORD.readFile(fn);
          const result: string = "\ufeff" + WORD.to_text(doc, { RS: "\r" });
          const baseline = fs.readFileSync(fn + ".txt", "utf8");
          assert.equal(result, baseline);
        });
      });
    });
  });
  if(formats.indexOf("txt") > -1) describe("txt", () => {
    const files = sync(`test_files/**/*.txt`);
    files.forEach(fn => {
      it(fn, () => {
        const doc: WORD.WJSDoc = WORD.readFile(fn);
        const result: string = "\ufeff" + WORD.to_text(doc, { RS: "\r" });
        const baseline = fs.readFileSync(fn, "utf8");
        assert.equal(result, baseline);
      });
    });
  });
});
