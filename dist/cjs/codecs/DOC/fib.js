"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * [MS-DOC] 2.5.1 Fib
 */
function readFib(buffer) {
    var offset = 0;
    var base = readBase(buffer.slice(offset, offset += 32));
    offset += 32;
    var fibRgLw = readFibRgLw(buffer.slice(offset, offset += 88));
    var cbRgFcLcb = buffer.readUInt16LE(offset);
    offset += 2;
    var fibRgFcLcbBlob = readFibRgFcLcbBlob(buffer.slice(offset, offset += cbRgFcLcb * 8));
    var fib = {
        base: base,
        fibRgLw: fibRgLw,
        fibRgFcLcbBlob: fibRgFcLcbBlob
    };
    var cswNew = buffer.readUInt16LE(offset);
    offset += 2;
    if (cswNew > 0) {
        fib.fibRgCswNew = readFibRgCswNew(buffer.slice(offset, offset += cswNew * 2));
    }
    return fib;
}
exports.readFib = readFib;
/**
 * [MS-DOC] 2.5.2 FibBase
 */
function readBase(buffer) {
    var offset = 2;
    var nFib = buffer.readUInt16LE(offset);
    offset += 9;
    var bits = buffer.readUInt8(offset);
    var fWhichTblStm = (bits >> 1) & 0x1;
    return {
        nFib: nFib,
        fWhichTblStm: fWhichTblStm
    };
}
/**
 * [MS-DOC] 2.5.4 FibRgLw97
 */
function readFibRgLw(buffer) {
    var offset = 12;
    var ccpText = buffer.readInt32LE(offset);
    offset += 4;
    var ccpFtn = buffer.readInt32LE(offset);
    offset += 4;
    var ccpHdd = buffer.readInt32LE(offset);
    offset += 8;
    var ccpAtn = buffer.readInt32LE(offset);
    offset += 4;
    var ccpEdn = buffer.readInt32LE(offset);
    offset += 4;
    var ccpTxbx = buffer.readInt32LE(offset);
    offset += 4;
    var ccpHdrTxbx = buffer.readInt32LE(offset);
    return {
        ccpText: ccpText,
        ccpFtn: ccpFtn,
        ccpHdd: ccpHdd,
        ccpAtn: ccpAtn,
        ccpEdn: ccpEdn,
        ccpTxbx: ccpTxbx,
        ccpHdrTxbx: ccpHdrTxbx
    };
}
/**
 * [MS-DOC] 2.5.6 FibRgFcLcb97
 */
function readFibRgFcLcbBlob(buffer) {
    var offset = 0;
    offset += 8; // StshfOrig
    offset += 8; // Stshf
    offset += 8; // PlcffndRef
    offset += 8; // PlcffndTxt
    offset += 8; // PlcfandRef
    offset += 8; // PlcfandTxt
    offset += 8; // PlcfSed
    offset += 8; // PlcPad
    offset += 8; // PlcfPhe
    offset += 8; // SttbfGlsy
    offset += 8; // PlcfGlsy
    offset += 8; // PlcfHdd
    offset += 8; // PlcfBteChpx
    var fcPlcfBtePapx = buffer.readUInt32LE(offset);
    offset += 4;
    var lcbPlcfBtePapx = buffer.readUInt32LE(offset);
    offset += 4;
    offset += 8; // PlcfSea
    offset += 8; // SttbfFfn
    offset += 8; // PlcfFldMom
    offset += 8; // PlcfFldHdr
    offset += 8; // PlcfFldFtn
    offset += 8; // PlcfFldAtn
    offset += 8; // PlcfFldMcr
    offset += 8; // SttbfBkmk
    offset += 8; // PlcfBkf
    offset += 8; // PlcfBkl
    offset += 8; // Cmds
    offset += 8; // Unused1
    offset += 8; // SttbfMcr
    offset += 8; // PrDrvr
    offset += 8; // PrEnvPort
    offset += 8; // PrEnvLand
    offset += 8; // Wss
    offset += 8; // Dop
    offset += 8; // SttbfAssoc
    if (offset != 33 * 8)
        throw new Error("Could not read FibRgFcLcb");
    var fcClx = buffer.readUInt32LE(offset);
    offset += 4;
    var lcbClx = buffer.readUInt32LE(offset);
    offset += 4;
    return {
        fcPlcfBtePapx: fcPlcfBtePapx,
        lcbPlcfBtePapx: lcbPlcfBtePapx,
        fcClx: fcClx,
        lcbClx: lcbClx
    };
}
/**
 * [MS-DOC] 2.5.11 FibRgCswNew
 */
function readFibRgCswNew(buffer) {
    var nFibNew = buffer.readUInt16LE(0);
    return {
        nFibNew: nFibNew
    };
}
//# sourceMappingURL=fib.js.map