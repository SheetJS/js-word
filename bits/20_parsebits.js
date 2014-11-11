
/* sections refer to MS-XLS unless otherwise stated */

/* --- Simple Utilities --- */
function parsenoop(blob, length) { blob.read_shift(length); return; }
function parsenoop2(blob, length) { blob.read_shift(length); return null; }

function parslurp(blob, length, cb) {
	var arr = [], target = blob.l + length;
	while(blob.l < target) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw new Error("Slurp error");
	return arr;
}

function parslurp2(blob, length, cb) {
	var arr = [], target = blob.l + length, len = blob.read_shift(2);
	while(len-- !== 0) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw new Error("Slurp error");
	return arr;
}

function parsebool(blob, length) { return blob.read_shift(length) === 0x1; }

function parseuint16(blob) { return blob.read_shift(2, 'u'); }
function parseuint16a(blob, length) { return parslurp(blob,length,parseuint16);}

/* --- 2.5 Structures --- */

/* [MS-XLS] 2.5.14 Boolean */
var parse_Boolean = parsebool;

/* [MS-XLS] 2.5.10 Bes (boolean or error) */
function parse_Bes(blob) {
	var v = blob.read_shift(1), t = blob.read_shift(1);
	return t === 0x01 ? v : v === 0x01;
}

/* [MS-XLS] 2.5.240 ShortXLUnicodeString */
function parse_ShortXLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(1);
	var width = 1, encoding = 'sbcs';
	if(opts === undefined || opts.biff !== 5) {
		var fHighByte = blob.read_shift(1);
		if(fHighByte) { width = 2; encoding = 'dbcs'; }
	}
	return cch ? blob.read_shift(cch, encoding) : "";
}

/* 2.5.293 XLUnicodeRichExtendedString */
function parse_XLUnicodeRichExtendedString(blob) {
	var cch = blob.read_shift(2), flags = blob.read_shift(1);
	var fHighByte = flags & 0x1, fExtSt = flags & 0x4, fRichSt = flags & 0x8;
	var width = 1 + (flags & 0x1); // 0x0 -> utf8, 0x1 -> dbcs
	var cRun, cbExtRst;
	var z = {};
	if(fRichSt) cRun = blob.read_shift(2);
	if(fExtSt) cbExtRst = blob.read_shift(4);
	var encoding = (flags & 0x1) ? 'dbcs' : 'sbcs';
	var msg = cch === 0 ? "" : blob.read_shift(cch, encoding);
	if(fRichSt) blob.l += 4 * cRun; //TODO: parse this
	if(fExtSt) blob.l += cbExtRst; //TODO: parse this
	z.t = msg;
	if(!fRichSt) { z.raw = "<t>" + z.t + "</t>"; z.r = z.t; }
	return z;
}

/* 2.5.296 XLUnicodeStringNoCch */
function parse_XLUnicodeStringNoCch(blob, cch, opts) {
	var retval;
	var fHighByte = blob.read_shift(1);
	if(fHighByte===0) { retval = blob.read_shift(cch, 'sbcs'); }
	else { retval = blob.read_shift(cch, 'dbcs'); }
	return retval;
}

/* 2.5.294 XLUnicodeString */
function parse_XLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(opts !== undefined && opts.biff > 0 && opts.biff < 8 ? 1 : 2);
	if(cch === 0) { blob.l++; return ""; }
	return parse_XLUnicodeStringNoCch(blob, cch, opts);
}
/* BIFF5 override */
function parse_XLUnicodeString2(blob, length, opts) {
	if(opts.biff !== 5 && opts.biff !== 2) return parse_XLUnicodeString(blob, length, opts);
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, 'sbcs');
}

/* 2.5.342 Xnum */
function parse_Xnum(blob) { return blob.read_shift(8, 'f'); }

/* 2.5.61 ControlInfo */
var parse_ControlInfo = parsenoop;

/* [MS-OSHARED] 2.3.7.6 URLMoniker TODO: flags */
var parse_URLMoniker = function(blob, length) {
	var len = blob.read_shift(4), start = blob.l;
	var extra = false;
	if(len > 24) {
		/* look ahead */
		blob.l += len - 24;
		if(blob.read_shift(16) === "795881f43b1d7f48af2c825dc4852763") extra = true;
		blob.l = start;
	}
	var url = blob.read_shift((extra?len-24:len)>>1, 'utf16le').replace(chr0,"");
	if(extra) blob.l += 24;
	return url;
};

/* [MS-OSHARED] 2.3.7.8 FileMoniker TODO: all fields */
var parse_FileMoniker = function(blob, length) {
	var cAnti = blob.read_shift(2);
	var ansiLength = blob.read_shift(4);
	var ansiPath = blob.read_shift(ansiLength, 'cstr');
	var endServer = blob.read_shift(2);
	var versionNumber = blob.read_shift(2);
	var cbUnicodePathSize = blob.read_shift(4);
	if(cbUnicodePathSize === 0) return ansiPath.replace(/\\/g,"/");
	var cbUnicodePathBytes = blob.read_shift(4);
	var usKeyValue = blob.read_shift(2);
	var unicodePath = blob.read_shift(cbUnicodePathBytes>>1, 'utf16le').replace(chr0,"");
	return unicodePath;
};

/* [MS-OSHARED] 2.3.7.2 HyperlinkMoniker TODO: all the monikers */
var parse_HyperlinkMoniker = function(blob, length) {
	var clsid = blob.read_shift(16); length -= 16;
	switch(clsid) {
		case "e0c9ea79f9bace118c8200aa004ba90b": return parse_URLMoniker(blob, length);
		case "0303000000000000c000000000000046": return parse_FileMoniker(blob, length);
		default: throw "unsupported moniker " + clsid;
	}
};

/* [MS-OSHARED] 2.3.7.9 HyperlinkString */
var parse_HyperlinkString = function(blob, length) {
	var len = blob.read_shift(4);
	var o = blob.read_shift(len, 'utf16le').replace(chr0, "");
	return o;
};

/* [MS-OSHARED] 2.3.7.1 Hyperlink Object TODO: unify params with XLSX */
var parse_Hyperlink = function(blob, length) {
	var end = blob.l + length;
	var sVer = blob.read_shift(4);
	if(sVer !== 2) throw new Error("Unrecognized streamVersion: " + sVer);
	var flags = blob.read_shift(2);
	blob.l += 2;
	var displayName, targetFrameName, moniker, oleMoniker, location, guid, fileTime;
	if(flags & 0x0010) displayName = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0080) targetFrameName = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0101) moniker = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0001) oleMoniker = parse_HyperlinkMoniker(blob, end - blob.l);
	if(flags & 0x0008) location = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0020) guid = blob.read_shift(16);
	if(flags & 0x0040) fileTime = parse_FILETIME(blob, 8);
	blob.l = end;
	var target = (targetFrameName||moniker||oleMoniker);
	if(location) target+="#"+location;
	return {Target: target};
};

/* 2.5.178 LongRGBA */
function parse_LongRGBA(blob, length) { var r = blob.read_shift(1), g = blob.read_shift(1), b = blob.read_shift(1), a = blob.read_shift(1); return [r,g,b,a]; }

/* 2.5.177 LongRGB */
function parse_LongRGB(blob, length) { var x = parse_LongRGBA(blob, length); x[3] = 0; return x; }
