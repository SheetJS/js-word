/**
 * [MS-DOC] 2.4.1 Retrieving Text
 *
 * @param {Buffer} fib - FIB Structure
 * @param {Buffer} doc - Document Content
 * @param {Buffer} tableStream - Either 1Table or 0Table
 * @return {string} Document Text
 */
function getDocTxt(fib, doc, tableStream) {
  const { fibRgLw, fibRgFcLcbBlob } = fib;
  const { fcClx, lcbClx } = fibRgFcLcbBlob;
  const clx = tableStream.slice(fcClx, fcClx + lcbClx);
  const plcPcd = parseClx(clx);
  const txt = getTxt(fibRgLw, plcPcd, doc);
  return txt;
}

/**
 * [MS-DOC] 2.9.38 Clx
 *
 * @param {Buffer} clx - Clx Structure
 * @return {string} Document Text
 */
function parseClx(clx) {
  /* Skip RgPrc to get Pcdt */
  let offset = 0;

  /* [MS-DOC] 2.9.209 Prc */
  const firstByte = clx.readUInt8(offset);
  if (firstByte !== 0x1 && firstByte !== 0x2) {
    throw Error("Invalid first byte of Clx.");
  }

  /* Not empty RgPrc */
  while (clx.readUInt8(offset) === 0x1) {
    /* [MS-DOC] 2.9.210 PrcData */
    offset++;
    const cbGrpGpl = clx.readInt16LE(offset);
    offset += 2;

    /* cbGrpGpl must be less than or equal to 0x3fa2 */
    console.assert(cbGrpGpl <= 0x3fa2);
    offset += cbGrpGpl;
  }

  /* [MS-DOC] 2.9.178 Pcdt */
  const pcdt = clx.slice(offset);

  /* clxt (first byte of Pcdt) must be 0x2 */
  console.assert(pcdt.readUInt8(0) === 0x2);
  const lcb = pcdt.readUInt32LE(1);

  /* [MS-DOC] 2.8.35 PlcPcd */
  const plcPcd = pcdt.slice(5, 5 + lcb);
  return plcPcd;
}

/**
 * [MS-DOC] 2.8.35 PlcPcd
 *
 * @param {Buffer} fibRgLw - FibRgLw Structure
 * @return {number} Value of Last CP
 */
function getLastCp(fibRgLw) {
  const fibMeta = Object.values(fibRgLw);
  const [ccpText, ...ccpOther] = fibMeta;
  const ccpSum = ccpOther.reduce((a, b) => a + b, 0);
  return ccpSum !== 0 ? ccpSum + ccpText + 1 : ccpText;
}

/**
 * @param {Buffer} fibRgLw - FibRgLw Structure
 * @param {Buffer} plcPcd - PlcPcd Structure
 * @param {Buffer} doc - Document Content
 * @return {string} Document Text
 */
function getTxt(fibRgLw, plcPcd, doc) {
  const cpSizeBytes = 4;
  const lastCp = getLastCp(fibRgLw);
  let offset = 0;
  let pcdCount = -1;

  while (plcPcd.readUInt32LE(offset) <= lastCp) {
    offset += cpSizeBytes;
    pcdCount++;
  }

  /* [MS-DOC] 2.8.35 PlcPcd */
  const acp = plcPcd.slice(0, offset);

  const pcdSizeBytes = 8;
  const upperBound = offset + pcdCount * pcdSizeBytes;
  let acpIndex = 0;
  let finalTxt = "";
  while (offset < upperBound) {
    const pcd = plcPcd.slice(offset, (offset += pcdSizeBytes));
    const fcCompressed = pcd.readUInt32LE(2);
    const fc = fcCompressed & ~(0x1 << 30);
    const strlen =
      acp.readUInt32LE((acpIndex + 1) * 4) - acp.readUInt32LE(acpIndex * 4);
    if ((fcCompressed >> 30) & 0x1) {
      finalTxt += getTxtCompressed(doc, fc, strlen);
    } else {
      finalTxt += getTxtNotCompressed(doc, fc, strlen);
    }
    acpIndex++;
  }

  return finalTxt;
}

/* [MS - DOC] 2.9.73 FcCompressed */

/**
 * @param {Buffer} doc - Compressed String
 * @param {number} fc - Document Offset
 * @param {number} strlen - Text Length
 * @return {string} Document Text
 */

function getTxtCompressed(doc, fc, strlen) {
  return fixFcString(doc.slice(fc / 2, fc / 2 + strlen).toString("binary"));
}

function getTxtNotCompressed(doc, fc, strlen) {
  return doc.slice(fc, fc + 2 * strlen).toString("utf16le");
}

/**
 * @param {string} str - Compressed String
 * @return {string} Fixed Value
 */
function fixFcString(str) {
  const replacements = {
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

  return str.replace(/[\x82-\x8C\x91-\x9C\x9F]/g, ($$) => replacements[$$]);
}

module.exports = {
  getDocTxt,
};
