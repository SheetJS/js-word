import { CFB$Container } from "cfb";
import { WJSDoc } from "../types";
export declare function parse_cfb(file: CFB$Container): WJSDoc;
export declare function parse_zip(file: CFB$Container): WJSDoc;
/** read JS string */
export declare function read_str(data: string): WJSDoc;
export declare function read(data: Buffer): WJSDoc;
export declare function readFile(path: string): WJSDoc;
