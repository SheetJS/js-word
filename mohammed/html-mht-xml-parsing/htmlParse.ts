import * as jsdom from 'jsdom'
import * as fs from 'fs';
import * as CFB from 'cfb';

const { JSDOM } = jsdom;

const cfb: CFB.CFB$Container = CFB.read("testFiles/demo.mht", { type: "file" });
console.log(cfb.FullPaths);

const firstHtmlIdx:number = cfb.FullPaths.findIndex((path: string) => /\.html?$/.test(path));
const entry: CFB.CFB$Entry = cfb.FileIndex[firstHtmlIdx];
const htmlAsString = entry.content.toString();
const dom: jsdom.JSDOM = new JSDOM(htmlAsString);
fs.writeFile("testFiles/output.txt", dom.window.document.querySelector('body').textContent, (err: NodeJS.ErrnoException) => console.error(err))
// console.log(dom.window.document.querySelector('body').textContent.replace("\n", ''));

/*fs.readFile('./Hello_World.demo', 'utf8', (err: NodeJS.ErrnoException, data: string) => {
    dom.window.document.children.item(0).childNodes

})*/