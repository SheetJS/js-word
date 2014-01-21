
/* sections refer to MS-XLS unless otherwise stated */

/* --- Simple Utilities --- */
function parsenoop(blob, length) { blob.read_shift(length); return; }
function parsenoop2(blob, length) { blob.read_shift(length); return null; }

function parslurp(blob, length, cb) {
	var arr = [], target = blob.l + length;
	while(blob.l < target) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw "Slurp error";
	return arr;
}

function parslurp2(blob, length, cb) {
	var arr = [], target = blob.l + length, len = blob.read_shift(2);
	while(len-- !== 0) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw "Slurp error";
	return arr;
}

function parsebool(blob, length) { return blob.read_shift(length) === 0x1; }

function parseuint16(blob, length) { return blob.read_shift(2, 'u'); }
function parseuint16a(blob, length) { return parslurp(blob,length,parseuint16);}

/* --- 2.5 Structures --- */

/* [MS-XLS] 2.5.14 Boolean */
var parse_Boolean = parsebool;

/* [MS-XLS] 2.5.10 Bes (boolean or error) */
function parse_Bes(blob) {
	var v = blob.read_shift(1), t = blob.read_shift(1);
	return t === 0x01 ? BERR[v] : v === 0x01;
}

/* [MS-XLS] 2.5.240 ShortXLUnicodeString */
function parse_ShortXLUnicodeString(blob) {
	var read = blob.read_shift.bind(blob);
	var cch = read(1);
	var fHighByte = read(1);
	var retval;
	var width = 1 + (fHighByte === 0 ? 0 : 1), encoding = fHighByte ? 'dbcs' : 'sbcs';
	retval = cch ? read(encoding, cch) : "";
	return retval;
}

/* 2.5.293 XLUnicodeRichExtendedString */
function parse_XLUnicodeRichExtendedString(blob) {
	var read_shift = blob.read_shift.bind(blob);
	var cch = read_shift(2), flags = read_shift(1);
	var fHighByte = flags & 0x1, fExtSt = flags & 0x4, fRichSt = flags & 0x8;
	var width = 1 + (flags & 0x1); // 0x0 -> utf8, 0x1 -> dbcs
	var cRun, cbExtRst;
	if(fRichSt) cRun = read_shift(2);
	if(fExtSt) cbExtRst = read_shift(4);
	var encoding = (flags & 0x1) ? 'dbcs' : 'sbcs';
	var msg = cch === 0 ? "" : read_shift(encoding, cch);
	if(fRichSt) blob.l += 4 * cRun; //TODO: parse this
	if(fExtSt) blob.l += cbExtRst; //TODO: parse this
	return msg;
}

/* 2.5.296 XLUnicodeStringNoCch */
function parse_XLUnicodeStringNoCch(blob, cch) {
	var read = blob.read_shift.bind(blob);
	var fHighByte = read(1);
	var retval;
	if(fHighByte===0) { retval = __utf8(blob,blob.l, blob.l+cch); blob.l += cch; }
	else { retval = blob.read_shift('dbcs', cch); }
	return retval;
}

/* 2.5.294 XLUnicodeString */
function parse_XLUnicodeString(blob) {
	var cch = blob.read_shift(2);
	if(cch === 0) { blob.l++; return ""; }
	return parse_XLUnicodeStringNoCch(blob, cch);
}

/* 2.5.342 Xnum */
function parse_Xnum(blob, length) { return blob.read_shift('ieee754'); }

