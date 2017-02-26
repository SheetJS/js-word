/* 2.5.217 */
function parse_RkNumber(data)/*:number*/ {
	var b = data.slice(data.l, data.l+4);
	var fX100 = b[0] & 1, fInt = b[0] & 2;
	data.l+=4;
	b[0] &= 0xFC;
	var RK = fInt === 0 ? __double([0,0,0,0,b[0],b[1],b[2],b[3]],0) : __readInt32LE(b,0)>>2;
	return fX100 ? RK/100 : RK;
}

/* 2.5.198.2 */
var BErr = {
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0x2B: "#GETTING_DATA", /* Noted in 2.5.10 but not in 2.5.198.2 */
	0xFF: "#WTF?"
};
var RBErr = evert(BErr);

