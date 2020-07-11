import * as jsdom from 'jsdom'
import * as fs from 'fs';

const { JSDOM } = jsdom;

fs.readFile('testFiles/Hello_World.xml', 'utf8', (err: NodeJS.ErrnoException, data: string) => {
    const listOfParagraphs: Node[] = [];
    const dom: jsdom.JSDOM = new JSDOM(data, {contentType: "text/xml"});
    const dfs = (node: Node): void => {
        if (node.hasChildNodes) {
            node.childNodes.forEach((node: ChildNode) => {
                if (node.nodeName === "w:p") listOfParagraphs.push(node);
                dfs(node);
            })
        }
    }
    dfs(dom.window.document)
    
    listOfParagraphs.forEach((node: Node) => {
        node.childNodes.forEach((child) => console.log(child.textContent))
        console.log('')
    })
    
})