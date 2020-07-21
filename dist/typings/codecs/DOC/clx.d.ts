import { FibRgLw97 } from "./fib";
/**
 * [MS-DOC] 2.9.38 Clx
 */
declare function parseClx(clx: Buffer): Buffer;
export declare function getTxt(fibRgLw: FibRgLw97, plcPcd: Buffer, doc: Buffer): string;
export { parseClx };
