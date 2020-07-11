class TrieNode {
    value: number;
    filename: string;
    children: { [key: string]: TrieNode};
    constructor(value?: number, filename?: string) { this.filename = filename; this.value = value }
}

function testfunction() {
    const a = new TrieNode();
    a.value  = 0x50dcef;
    a.filename = 'doc'
    
    const b = new TrieNode();
    b.value = 0xdc;
    a.filename = 'docx'
    
    const c = new TrieNode(0xef, 'doas');
    
    a.children = {'b': b}
    // console.log(a);
    
    a.children = {...a.children,'c': c}
    // console.log(a)
    
    
    a.children = {...a.children,'c': c}
    console.log(a)
    b.children = {'d': c}
    console.log(a.children['b'].children['d'])
    console.log(a);
    console.log(a.children['t']);
}



const numberFromHex = (number: number, index: number):number => parseInt(number.toString(16).slice(index, index + 2), 16);

const test = 0x50dcef
const index = 2
// console.log(numberFromHex(test, index));


class Trie {
    root: TrieNode;
    constructor() {
        this.root = new TrieNode();
    }

    insert(fileName: string, signature: number, index: number, node: TrieNode): TrieNode {
        if (index === signature.toString(16).length/2 - 1) {
            node.filename = fileName;
            node.value = signature;
            return (node)
        } else {
            if (!node.children[numberFromHex(signature, index)]) {
                node.children = {
                    ...node.children,

                }
            }
        }

        // if (index === string.length - 1) {
        //     return (node[string[index]] = { value: string })
        // } else {
        //     if (!node[string[index]]) {
        //         node[string[index]] = { value: string.slice(0, index + 1) }
        //     }

        //     node[string[index]][string[index + 1]] = {
        //         ...node[string[index]][string[index + 1]],
        //         ...this._traverseInsert(node[string[index]], string, index + 1)
        //     }
        // }

        // return node[string[index]];
    }

}

const testTrie = new Trie();

testTrie.insert('xls', 0x50, 0, testTrie.root);
console.log(testTrie);



// testfunction();