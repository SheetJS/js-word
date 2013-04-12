
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

/* 2.5.14 */
var parse_Boolean = parsebool;

/* 2.5.240 */
function parse_ShortXLUnicodeString(blob) {
	var read = blob.read_shift.bind(blob);
	var cch = read(1);
	var fHighByte = read(1);
	var retval;
	if(fHighByte===0) { retval = blob.utf8(blob.l, blob.l+cch); blob.l += cch; }
	else { retval = blob.utf16le(blob.l, blob.l + 2*cch); blob.l += 2*cch; }
	return retval;
}

/* 2.5.293 */
function parse_XLUnicodeRichExtendedString(blob) {
	var read_shift = blob.read_shift.bind(blob);
	var cch = read_shift(2), flags = read_shift(1);
	var width = 1 + (flags & 0x1);
	// fRichSt
	if(flags & 0x08) {
		// TODO: cRun
		// TODO: cbExtRst
		read_shift(6);
	}
	var encoding = (flags & 0x1) ? 'dbcs' : 'utf8';
	var msg = read_shift(encoding, cch);
	return msg;
}

/* 2.5.294 */
function parse_XLUnicodeString(blob) {
	var read = blob.read_shift.bind(blob);
	var cch = read(2);
	var fHighByte = read(1);
	var retval = undefined;
	if(fHighByte===0) { retval = blob.utf8(blob.l, blob.l+cch); blob.l += cch; }
	else { retval = blob.read_shift('dbcs', cch); }
	return retval;
}

/* 2.5.342 */
function parse_Xnum(blob, length) { return blob.read_shift('ieee754'); }

