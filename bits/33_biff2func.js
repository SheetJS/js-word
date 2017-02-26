/* BIFF2_??? where ??? is the name from [XLS] */
function parse_BIFF2STR(blob, length, opts) {
	var cell = parse_Cell(blob, 6);
	++blob.l;
	var str = parse_XLUnicodeString2(blob, length-7, opts);
	cell.t = 'str';
	cell.val = str;
	return cell;
}

function parse_BIFF2NUM(blob, length, opts) {
	var cell = parse_Cell(blob, 6);
	++blob.l;
	var num = parse_Xnum(blob, 8);
	cell.t = 'n';
	cell.val = num;
	return cell;
}

function parse_BIFF2INT(blob, length) {
	var cell = parse_Cell(blob, 6);
	++blob.l;
	var num = blob.read_shift(2);
	cell.t = 'n';
	cell.val = num;
	return cell;
}

function parse_BIFF2STRING(blob, length) {
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, 'sbcs-cont');
}

/* TODO: convert to BIFF8 font struct */
function parse_BIFF2FONTXTRA(blob, length) {
	blob.l += 6; // unknown
	blob.l += 2; // font weight "bls"
	blob.l += 1; // charset
	blob.l += 3; // unknown
	blob.l += 1; // font family
	blob.l += length - 9;
}

/* TODO: parse rich text runs */
function parse_RString(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_Cell(blob, 6);
	var cch = blob.read_shift(2);
	var str = parse_XLUnicodeStringNoCch(blob, cch, opts);
	blob.l = end;
	cell.t = 'str';
	cell.val = str;
	return cell;
}
