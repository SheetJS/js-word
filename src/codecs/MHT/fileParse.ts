import * as fs from 'fs';
import * as rd from 'readline'
import { Fields } from './fileEnums'

const charsetExample: string = 'Content-Type: text/html; charset="utf-8"'

let location: string = "";
let contentType: string = "";
let contentTransferEncoding: string = "";

const getContentType = (line: string): string => {
    return ""
} 
const getContentTransferEncoding = (line: string): string => {
    return ""
} 
const getContentLocation = (line: string): string => {
    return ""
} 

const reader = rd.createInterface(fs.createReadStream("./sample.mht"))
reader.on("line", (line: string) => {
    console.log(line);
    
})

