"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * [MS-DOC] 2.9.38 Clx
 */
function parseClx(clx) {
    /* Skip RgPrc to get Pcdt */
    var offset = 0;
    /* [MS-DOC] 2.9.209 Prc */
    var firstByte = clx.readUInt8(offset);
    if (firstByte !== 0x1 && firstByte !== 0x2) {
        throw Error("Invalid first byte of Clx.");
    }
    /* Not empty RgPrc */
    while (clx.readUInt8(offset) === 0x1) {
        /* [MS-DOC] 2.9.210 PrcData */
        offset++;
        var cbGrpGpl = clx.readInt16LE(offset);
        offset += 2;
        /* cbGrpGpl must be less than or equal to 0x3fa2 */
        console.assert(cbGrpGpl <= 0x3fa2);
        offset += cbGrpGpl;
    }
    /* [MS-DOC] 2.9.178 Pcdt */
    var pcdt = clx.slice(offset);
    /* clxt (first byte of Pcdt) must be 0x2 */
    console.assert(pcdt.readUInt8(0) === 0x2);
    var lcb = pcdt.readUInt32LE(1);
    /* [MS-DOC] 2.8.35 PlcPcd */
    var plcPcd = pcdt.slice(5, 5 + lcb);
    return plcPcd;
}
exports.parseClx = parseClx;
/**
 * [MS-DOC] 2.8.35 PlcPcd
 */
function getLastCp(fibRgLw) {
    var fibMeta = Object.values(fibRgLw);
    var ccpText = fibMeta[0], ccpOther = fibMeta.slice(1);
    var ccpSum = ccpOther.reduce(function (a, b) { return a + b; }, 0);
    return ccpSum !== 0 ? ccpSum + ccpText + 1 : ccpText;
}
function getTxt(fibRgLw, plcPcd, doc) {
    var cpSizeBytes = 4;
    var lastCp = getLastCp(fibRgLw);
    var offset = 0;
    var pcdCount = -1;
    while (plcPcd.readUInt32LE(offset) <= lastCp) {
        offset += cpSizeBytes;
        pcdCount++;
    }
    /* [MS-DOC] 2.8.35 PlcPcd */
    var acp = plcPcd.slice(0, offset);
    var pcdSizeBytes = 8;
    var upperBound = offset + pcdCount * pcdSizeBytes;
    var acpIndex = 0;
    var finalTxt = "";
    while (offset < upperBound) {
        var pcd = plcPcd.slice(offset, (offset += pcdSizeBytes));
        var fcCompressed = pcd.readUInt32LE(2);
        var fc = fcCompressed & ~(0x1 << 30);
        var strlen = acp.readUInt32LE((acpIndex + 1) * 4) - acp.readUInt32LE(acpIndex * 4);
        if ((fcCompressed >> 30) & 0x1) {
            finalTxt += getTxtCompressed(doc, fc, strlen);
        }
        else {
            finalTxt += getTxtNotCompressed(doc, fc, strlen);
        }
        acpIndex++;
    }
    return finalTxt;
}
exports.getTxt = getTxt;
/* [MS-DOC] 2.9.73 FcCompressed */
function getTxtCompressed(doc, fc, strlen) {
    return fixFcString(doc.slice(fc / 2, fc / 2 + strlen).toString("binary"));
}
function getTxtNotCompressed(doc, fc, strlen) {
    return doc.slice(fc, fc + 2 * strlen).toString("utf16le");
}
function fixFcString(str) {
    var replacements = {
        "\x82": "\u201A",
        "\x83": "\u0192",
        "\x84": "\u201E",
        "\x85": "\u2026",
        "\x86": "\u2020",
        "\x87": "\u2021",
        "\x88": "\u02C6",
        "\x89": "\u2030",
        "\x8A": "\u0160",
        "\x8B": "\u2039",
        "\x8C": "\u0152",
        "\x91": "\u2018",
        "\x92": "\u2019",
        "\x93": "\u201C",
        "\x94": "\u201D",
        "\x95": "\u2022",
        "\x96": "\u2013",
        "\x97": "\u2014",
        "\x98": "\u02DC",
        "\x99": "\u2122",
        "\x9A": "\u0161",
        "\x9B": "\u203A",
        "\x9C": "\u0153",
        "\x9F": "\u0178",
    };
    return str.replace(/[\x82-\x8C\x91-\x9C\x9F]/g, function ($$) { return replacements[$$]; });
}
//# sourceMappingURL=clx.js.map