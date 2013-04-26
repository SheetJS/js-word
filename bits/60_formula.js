
/* Small helpers */
function parseread(l) { return function(blob, length) { blob.l+=l; return; }; }
function parseread1(blob, length) { blob.l+=1; return; }

/* Rgce Helpers */

/* 2.5.51 */
function parse_ColRelU(blob, length) {
	var c = blob.read_shift(2);
	return [c & 0x3FFF, (c >> 14) & 1, (c >> 15) & 1];
}

/* 2.5.198.105 */
function parse_RgceArea(blob, length) {
	var read = blob.read_shift.bind(blob);
	var r=read(2), R=read(2);
	var c=parse_ColRelU(blob, 2);
	var C=parse_ColRelU(blob, 2);
	return { s:{r:r, c:c[0], cRel:c[1], rRel:c[2]}, e:{r:R, c:C[0], cRel:C[1], rRel:C[2]} };
}

/* 2.5.198.109 */
function parse_RgceLoc(blob, length) {
	var r = blob.read_shift(2);
	var c = parse_ColRelU(blob, 2);
	return {r:r, c:c[0], cRel:c[1], rRel:c[2]};
}

/* 2.5.198.111 */
function parse_RgceLocRel(blob, length) {
	var r = blob.read_shift(2);
	var cl = blob.read_shift(2);
	var cRel = (cl & 0x8000) >> 15, rRel = (cl & 0x4000) >> 14;
	cl &= 0x3FFF;
	if(cRel !== 0) while(cl >= 0x100) cl -= 0x100;
	return {r:r,c:cl,cRel:cRel,rRel:rRel};
}

/* Ptg Tokens */

/* 2.5.198.27 */
function parse_PtgArea(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceArea(blob, 8);
	return [type, area];
}

/* 2.5.198.27 */
function parse_PtgArea3d(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2);
	var area = parse_RgceArea(blob, 8);
	return [type, ixti, area];
}

/* 2.5.198.84 TODO */
function parse_PtgRef(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLoc(blob,4);
	return [type, loc];
}

/* 2.5.198.88 TODO */
function parse_PtgRefN(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLocRel(blob,4);
	return [type, loc];
}

/* 2.5.198.85 TODO */
function parse_PtgRef3d(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var ixti = blob.read_shift(2); // XtiIndex
	var loc = parse_RgceLoc(blob,4);
	return [type, ixti, loc];
}


/* 2.5.198.35 TODO */
function parse_PtgAttrGoto(blob, length) {
	blob.l += 2;
	return blob.read_shift(2);
}

/* 2.5.198.62 TODO */
function parse_PtgFunc(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var iftab = blob.read_shift(2);
	return [FtabArgc[iftab], Ftab[iftab]];
}
/* 2.5.198.63 TODO */
function parse_PtgFuncVar(blob, length) {
	blob.l++;
	var cparams = blob.read_shift(1), tab = parsetab(blob);
	return [cparams, (tab[0] === 0 ? Ftab : Cetab)[tab[1]]];
}

function parsetab(blob, length) {
	return [blob[blob.l+1]>>7, blob.read_shift(2) & 0x7FFF];
}

/* 2.5.198.36 */
var parse_PtgAttrIf = parseread(4);
/* 2.5.198.37 */
var parse_PtgAttrSemi = parseread(4);
/* 2.5.198.41 */
var parse_PtgAttrSum = parseread(4);
/* 2.5.198.43 */
var parse_PtgConcat = parseread1;

/* 2.5.198.58 */
function parse_PtgExp(blob, length) {
	blob.l++;
	var row = blob.read_shift(2);
	var col = blob.read_shift(2);
	return [row, col];
}

/* 2.5.198.66 TODO */
function parse_PtgInt(blob, length) { blob.l++; return blob.read_shift(2); }

/* 2.5.198.42 */
function parse_PtgBool(blob, length) { blob.l++; return blob.read_shift(1)!==0;}

/* 2.5.198.79 */
function parse_PtgNum(blob, length) { blob.l++; return parse_Xnum(blob, 8); }

/* 2.5.198.89 */
function parse_PtgStr(blob, length) { blob.l++; return parse_ShortXLUnicodeString(blob); }

/* 2.5.198.57 */
function parse_PtgErr(blob, length) { blob.l++; return BERR[blob.read_shift(1)]; }

/* 2.5.198.76 */
function parse_PtgName(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var nameindex = blob.read_shift(4);
	return [type, 0, nameindex];
}

/* 2.5.198.77 */
function parse_PtgNameX(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var ixti = blob.read_shift(2); // XtiIndex
	var nameindex = blob.read_shift(4);
	return [type, ixti, nameindex];
}

/* 2.5.198.26 */
var parse_PtgAdd = parseread1;
/* 2.5.198.45 */
var parse_PtgDiv = parseread1;
/* 2.5.198.56 */
var parse_PtgEq = parseread1;
/* 2.5.198.64 */
var parse_PtgGe = parseread1;
/* 2.5.198.65 */
var parse_PtgGt = parseread1;
/* 2.5.198.67 */
var parse_PtgIsect = parseread1;
/* 2.5.198.68 */
var parse_PtgLe = parseread1;
/* 2.5.198.69 */
var parse_PtgLt = parseread1;
/* 2.5.198.74 */
var parse_PtgMissArg = parseread1;
/* 2.5.198.75 */
var parse_PtgMul = parseread1;
/* 2.5.198.78 */
var parse_PtgNe = parseread1;
/* 2.5.198.80 */
var parse_PtgParen = parseread1;
/* 2.5.198.81 */
var parse_PtgPercent = parseread1;
/* 2.5.198.82 */
var parse_PtgPower = parseread1;
/* 2.5.198.83 */
var parse_PtgRange = parseread1;
/* 2.5.198.90 */
var parse_PtgSub = parseread1;
/* 2.5.198.93 */
var parse_PtgUminus = parseread1;
/* 2.5.198.94 */
var parse_PtgUnion = parseread1;
/* 2.5.198.95 */
var parse_PtgUplus = parseread1;

/* 2.5.198.29 */
var parse_PtgAreaErr = parsenoop;
/* 2.5.198.30 */
var parse_PtgAreaErr3d = parsenoop;
/* 2.5.198.31 */
var parse_PtgAreaN = parsenoop;
/* 2.5.198.32 */
var parse_PtgArray = parsenoop;
/* 2.5.198.33 */
var parse_PtgAttrBaxcel = parsenoop;
/* 2.5.198.34 */
var parse_PtgAttrChoose = parsenoop;
/* 2.5.198.38 */
var parse_PtgAttrSpace = parsenoop;
/* 2.5.198.39 */
var parse_PtgAttrSpaceSemi = parsenoop;
/* 2.5.198.70 */
var parse_PtgMemArea = parsenoop;
/* 2.5.198.71 */
var parse_PtgMemErr = parsenoop;
/* 2.5.198.72 */
var parse_PtgMemFunc = parsenoop;
/* 2.5.198.73 */
var parse_PtgMemNoMem = parsenoop;
/* 2.5.198.86 */
var parse_PtgRefErr = parsenoop;
/* 2.5.198.87 */
var parse_PtgRefErr3d = parsenoop;
/* 2.5.198.92 */
var parse_PtgTbl = parsenoop;

/* 2.5.198.25 */
var PtgTypes = {
	0x01: { n:'PtgExp', f:parse_PtgExp },
	0x02: { n:'PtgTbl', f:parse_PtgTbl },
	0x03: { n:'PtgAdd', f:parse_PtgAdd },
	0x04: { n:'PtgSub', f:parse_PtgSub },
	0x05: { n:'PtgMul', f:parse_PtgMul },
	0x06: { n:'PtgDiv', f:parse_PtgDiv },
	0x07: { n:'PtgPower', f:parse_PtgPower },
	0x08: { n:'PtgConcat', f:parse_PtgConcat },
	0x09: { n:'PtgLt', f:parse_PtgLt },
	0x0A: { n:'PtgLe', f:parse_PtgLe },
	0x0B: { n:'PtgEq', f:parse_PtgEq },
	0x0C: { n:'PtgGe', f:parse_PtgGe },
	0x0D: { n:'PtgGt', f:parse_PtgGt },
	0x0E: { n:'PtgNe', f:parse_PtgNe },
	0x0F: { n:'PtgIsect', f:parse_PtgIsect },
	0x10: { n:'PtgUnion', f:parse_PtgUnion },
	0x11: { n:'PtgRange', f:parse_PtgRange },
	0x12: { n:'PtgUplus', f:parse_PtgUplus },
	0x13: { n:'PtgUminus', f:parse_PtgUminus },
	0x14: { n:'PtgPercent', f:parse_PtgPercent },
	0x15: { n:'PtgParen', f:parse_PtgParen },
	0x16: { n:'PtgMissArg', f:parse_PtgMissArg },
	0x17: { n:'PtgStr', f:parse_PtgStr },
	0x1C: { n:'PtgErr', f:parse_PtgErr },
	0x1D: { n:'PtgBool', f:parse_PtgBool },
	0x1E: { n:'PtgInt', f:parse_PtgInt },
	0x1F: { n:'PtgNum', f:parse_PtgNum },
	0x20: { n:'PtgArray', f:parse_PtgArray },
	0x21: { n:'PtgFunc', f:parse_PtgFunc },
	0x22: { n:'PtgFuncVar', f:parse_PtgFuncVar },
	0x23: { n:'PtgName', f:parse_PtgName },
	0x24: { n:'PtgRef', f:parse_PtgRef },
	0x25: { n:'PtgArea', f:parse_PtgArea },
	0x26: { n:'PtgMemArea', f:parse_PtgMemArea },
	0x27: { n:'PtgMemErr', f:parse_PtgMemErr },
	0x28: { n:'PtgMemNoMem', f:parse_PtgMemNoMem },
	0x29: { n:'PtgMemFunc', f:parse_PtgMemFunc },
	0x2A: { n:'PtgRefErr', f:parse_PtgRefErr },
	0x2B: { n:'PtgAreaErr', f:parse_PtgAreaErr },
	0x2C: { n:'PtgRefN', f:parse_PtgRefN },
	0x2D: { n:'PtgAreaN', f:parse_PtgAreaN },
	0x39: { n:'PtgNameX', f:parse_PtgNameX },
	0x3A: { n:'PtgRef3d', f:parse_PtgRef3d },
	0x3B: { n:'PtgArea3d', f:parse_PtgArea3d },
	0x3C: { n:'PtgRefErr3d', f:parse_PtgRefErr3d },
	0x3D: { n:'PtgAreaErr3d', f:parse_PtgAreaErr3d },
	0xFF: {}
};
/* These are duplicated in the PtgTypes table */
var PtgDupes = {
	0x40: 0x20, 0x60: 0x20,
	0x41: 0x21, 0x61: 0x21,
	0x42: 0x22, 0x62: 0x22,
	0x43: 0x23, 0x63: 0x23,
	0x44: 0x24, 0x64: 0x24,
	0x45: 0x25, 0x65: 0x25,
	0x46: 0x26, 0x66: 0x26,
	0x47: 0x27, 0x67: 0x27,
	0x48: 0x28, 0x68: 0x28,
	0x49: 0x29, 0x69: 0x29,
	0x4A: 0x2A, 0x6A: 0x2A,
	0x4B: 0x2B, 0x6B: 0x2B,
	0x4C: 0x2C, 0x6C: 0x2C,
	0x4D: 0x2D, 0x6D: 0x2D,
	0x59: 0x39, 0x79: 0x39,
	0x5A: 0x3A, 0x7A: 0x3A,
	0x5B: 0x3B, 0x7B: 0x3B,
	0x5C: 0x3C, 0x7C: 0x3C,
	0x5D: 0x3D, 0x7D: 0x3D
};
for(var y in PtgDupes) PtgTypes[y] = PtgTypes[PtgDupes[y]];

var Ptg18 = {};
var Ptg19 = {
	0x01: { n:'PtgAttrSemi', f:parse_PtgAttrSemi },
	0x02: { n:'PtgAttrIf', f:parse_PtgAttrIf },
	0x04: { n:'PtgAttrChoose', f:parse_PtgAttrChoose },
	0x08: { n:'PtgAttrGoto', f:parse_PtgAttrGoto },
	0x10: { n:'PtgAttrSum', f:parse_PtgAttrSum },
	0x20: { n:'PtgAttrBaxcel', f:parse_PtgAttrBaxcel },
	0x40: { n:'PtgAttrSpace', f:parse_PtgAttrSpace },
	0x41: { n:'PtgAttrSpaceSemi', f:parse_PtgAttrSpaceSemi },
	0xFF: {}
};

