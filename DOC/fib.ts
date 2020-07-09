/* [MS-DOC] 2.5.1 Fib */
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
function readFib(buffer: Buffer): Fib {
  let offset = 0;
  const base = readBase(buffer.slice(offset, offset += 32));
  offset += 32;
  const fibRgLw = readFibRgLw(buffer.slice(offset, offset += 88));
  const cbRgFcLcb = buffer.readUInt16LE(offset);
  offset += 2;
  const fibRgFcLcbBlob = readFibRgFcLcbBlob(buffer.slice(offset, offset += cbRgFcLcb * 8));
  const fib: Fib = {
    base,
    fibRgLw,
    fibRgFcLcbBlob
  };
  const cswNew = buffer.readUInt16LE(offset);
  offset += 2;
  if (cswNew > 0) {
    fib.fibRgCswNew = readFibRgCswNew(buffer.slice(offset, offset += cswNew * 2))
  }

  return fib;
}

/**
 * [MS-DOC] 2.5.2 FibBase
 */
function readBase(buffer: Buffer): FibBase {
  let offset = 2;
  const nFib = buffer.readUInt16LE(offset);
  offset += 9;

  const bits = buffer.readUInt8(offset);
  const fWhichTblStm = (bits >> 1) & 0x1;

  return {
    nFib,
    fWhichTblStm
  };
}

/**
 * [MS-DOC] 2.5.4 FibRgLw97
 */
function readFibRgLw(buffer: Buffer): FibRgLw97 {
  let offset = 12;
  const ccpText = buffer.readInt32LE(offset);
  offset += 4;

  const ccpFtn = buffer.readInt32LE(offset);
  offset += 4;

  const ccpHdd = buffer.readInt32LE(offset);
  offset += 8;

  const ccpAtn = buffer.readInt32LE(offset);
  offset += 4;

  const ccpEdn = buffer.readInt32LE(offset);
  offset += 4;

  const ccpTxbx = buffer.readInt32LE(offset);
  offset += 4;

  const ccpHdrTxbx = buffer.readInt32LE(offset);

  return {
    ccpText,
    ccpFtn,
    ccpHdd,
    ccpAtn,
    ccpEdn,
    ccpTxbx,
    ccpHdrTxbx
  };
}

/**
 * [MS-DOC] 2.5.6 FibRgFcLcb97
 */
function readFibRgFcLcbBlob(buffer: Buffer): FibRgFcLcb {
  let offset = 264;
  const fcClx = buffer.readUInt32LE(offset);
  offset += 4;

  const lcbClx = buffer.readUInt32LE(offset);

  return {
    fcClx,
    lcbClx
  };
}

/**
 * [MS-DOC] 2.5.11 FibRgCswNew
 */
function readFibRgCswNew(buffer: Buffer): FibRgCswNew {
  const nFibNew = buffer.readUInt16LE(0);

  return {
    nFibNew
  };
}

export {
  Fib,
  FibRgLw97,
  FibRgFcLcb,
  FibRgCswNew,
  readFib
};
