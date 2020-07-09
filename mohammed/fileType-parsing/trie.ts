import { fileTypeHandler } from './fileTypeHandler';
const keys = Object.keys(fileTypeHandler);

export class trieNode {
    value: string;
    fileType?: string;
}

export class Trie {
    private tree: trieNode;
    constructor() {
        this.tree = new trieNode();
    }
    private _traversePrint(keys: string[], node: trieNode, level: number): void {
        keys = keys.filter(key => key !== "value");
        for (let [index, key] of keys.entries()) {
            //Prints values with indentation
            console.log(`${"| ".repeat(level - 1)}|-> ${level}.${index} ${node[key].value}`);
            
            //Recursion step
            this._traversePrint(Object.keys(node[key]), node[key], level + 1);
        }
        return;
    }

    private _traverseInsert(node: trieNode, string: string, index: number, fileType: string): trieNode {
        if (index === string.length - 1) {
            return (node[string[index]] = { value: string, fileType: fileType })
        } else {
            if (!node[string[index]]) {
                node[string[index]] = { value: string.slice(0, index + 1) }
            }

            node[string[index]][string[index + 1]] = {
                ...node[string[index]][string[index + 1]],
                ...this._traverseInsert(node[string[index]], string, index + 1, fileType)
            }
        }

        return node[string[index]];
    }

    private _traverseSearch(string: string, node: trieNode, index: number): string {
        if (!node) {
            console.error("Not there lul!");
            return "INVALID";
        }

        if (index === string.length) {
            console.log("Got em'");
            return node.value;
        } else {
            return this._traverseSearch(string, node[string[index]], index + 1);
        }
    }

    public insert(string: string, fileType: string): void {
        this._traverseInsert(this.tree, string, 0, fileType)
    }

    public search(string: string): void {
        console.log(this._traverseSearch(string, this.tree, 0));
    }

    public print(): void {
        console.log("Tree");
        this._traversePrint(Object.keys(this.tree), this.tree, 1);
    }
}

const testTrie = new Trie();
// const testWordsOne: string[] = ["test", "teseract","tennis","testing","insertion","insert","inseparable","international"]
// const testWordsTwo: string[] = ["to", "teal", "ted", "ten", "an", "ion", "io", "inn"]
// testWordsOne.forEach(word => testTrie.insert(word))
keys.forEach(key => {
    testTrie.insert(fileTypeHandler[key].signature.toString(), fileTypeHandler[key].fileExtension)
})
testTrie.print();
console.log(testTrie);
