/* BIFF2_??? where ??? is the name from [XLS] */
function parse_BIFF2STR(blob, length, opts) {
	var cell = parse_Cell(blob, 6);
	++blob.l;
	var str = parse_XLUnicodeString2(blob, length-7, opts);
	cell.val = str;
	return cell;
}

function parse_BIFF2NUM(blob, length, opts) {
	var cell = parse_Cell(blob, 6);
	++blob.l;
	var num = parse_Xnum(blob, 8);
	cell.val = num;
	return cell;
}
