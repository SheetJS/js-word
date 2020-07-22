/**
 * [MS-DOC] 2.5.1 Fib
 */
interface Fib {
    base: FibBase;
    fibRgLw: FibRgLw97;
    fibRgFcLcbBlob: FibRgFcLcb;
    fibRgCswNew?: FibRgCswNew;
}
/**
 * [MS-DOC] 2.5.2 FibBase
 */
interface FibBase {
    /**
     * nFib (2 bytes): An unsigned integer that specifies the version number of
     * the file format used. Superseded by FibRgCswNew.nFibNew if it is present.
     * This value SHOULD be 0x00C1. Could possibly be 0x00C0 or 0x00C2 but should
     * be treated as if it were 0x00C1.
     */
    nFib: number;
    /**
     * G - fWhichTblStm (1 bit): Specifies the Table stream to which the FIB
     * refers. When this value is set to 1, use 1Table; when this value is set to
     * 0, use 0Table.
     */
    fWhichTblStm: number;
}
/**
 * [MS-DOC] 2.5.4 FibRgLw97
 */
interface FibRgLw97 {
    /**
     * ccpText (4 bytes): A signed integer that specifies the count of CPs in the
     * main document. This value MUST be zero, 1, or greater.
     */
    ccpText: number;
    /**
     * ccpFtn (4 bytes): A signed integer that specifies the count of CPs in the
     * footnote subdocument. This value MUST be zero, 1, or greater.
     */
    ccpFtn: number;
    /**
     * ccpHdd (4 bytes): A signed integer that specifies the count of CPs in the
     * header subdocument. This value MUST be zero, 1, or greater.
     */
    ccpHdd: number;
    /**
     * ccpAtn (4 bytes): A signed integer that specifies the count of CPs in the
     * comment subdocument. This value MUST be zero, 1, or greater.
     */
    ccpAtn: number;
    /**
     * ccpEdn (4 bytes): A signed integer that specifies the count of CPs in the
     * endnote subdocument. This value MUST be zero, 1, or greater.
     */
    ccpEdn: number;
    /**
     * ccpTxbx (4 bytes): A signed integer that specifies the count of CPs in the
     * textbox subdocument of the main document. This value MUST be zero, 1, or
     * greater.
     */
    ccpTxbx: number;
    /**
     * ccpHdrTxbx (4 bytes): A signed integer that specifies the count of CPs in
     * the textbox subdocument of the header. This value MUST be zero, 1, or
     * greater.
     */
    ccpHdrTxbx: number;
}
/**
 * [MS-DOC] 2.5.6 FibRgFcLcb97
 */
interface FibRgFcLcb {
    /**
     * fcClx (4 bytes):  An unsigned integer that specifies an offset in the
     * Table Stream. A Clx begins at this offset.
     */
    fcClx: number;
    /**
     * lcbClx (4 bytes):  An unsigned integer that specifies the size, in bytes,
     * of the Clx at offset fcClx in the Table Stream. This value MUST be greater
     * than zero.
     */
    lcbClx: number;
    /**
     * fcPlcfBtePapx (4 bytes): An unsigned integer that specifies an offset in
     * the Table Stream. A PlcBtePapx begins at the offset. fcPlcfBtePapx MUST be
     * greater than zero, and MUST be a valid offset in the Table Stream.
     */
    fcPlcfBtePapx: number;
    /**
     * lcbPlcfBtePapx (4 bytes): An unsigned integer that specifies the size, in
     * bytes, of the PlcBtePapx at offset fcPlcfBtePapx in the Table Stream.
     * lcbPlcfBteChpx MUST be greater than zero.
     */
    lcbPlcfBtePapx: number;
}
/**
 * [MS-DOC] 2.5.11 FibRgCswNew
 */
interface FibRgCswNew {
    /**
     * nFibNew (2 bytes): An unsigned integer that specifies the version number
     * of the file format that is used. This value MUST be one of the following.
     * 0x00D9, 0x0101, 0x010C, 0x0112.
     */
    nFibNew: number;
}
/**
 * [MS-DOC] 2.5.1 Fib
 */
declare function readFib(buffer: Buffer): Fib;
export { Fib, FibRgLw97, FibRgFcLcb, FibRgCswNew, readFib };
