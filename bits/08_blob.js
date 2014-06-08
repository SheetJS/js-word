function new_buf(len) {
	/* jshint -W056 */
	return new (typeof Buffer !== 'undefined' ? Buffer : Array)(len);
	/* jshint +W056 */
}
function readIEEE754(buf, idx, isLE, nl, ml) {
	if(isLE === undefined) isLE = true;
	if(!nl) nl = 8;
	if(!ml && nl === 8) ml = 52;
	var e, m, el = nl * 8 - ml - 1, eMax = (1 << el) - 1, eBias = eMax >> 1;
	var bits = -7, d = isLE ? -1 : 1, i = isLE ? (nl - 1) : 0, s = buf[idx + i];

	i += d;
	e = s & ((1 << (-bits)) - 1); s >>>= (-bits); bits += el;
	for (; bits > 0; e = e * 256 + buf[idx + i], i += d, bits -= 8);
	m = e & ((1 << (-bits)) - 1); e >>>= (-bits); bits += ml;
	for (; bits > 0; m = m * 256 + buf[idx + i], i += d, bits -= 8);
	if (e === eMax) return m ? NaN : ((s ? -1 : 1) * Infinity);
	else if (e === 0) e = 1 - eBias;
	else { m = m + Math.pow(2, ml); e = e - eBias; }
	return (s ? -1 : 1) * m * Math.pow(2, e - ml);
}

var Base64 = (function(){
	var map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	return {
		/* (will need this for writing) encode: function(input, utf8) {
			var o = "";
			var c1, c2, c3, e1, e2, e3, e4;
			for(var i = 0; i < input.length; ) {
				c1 = input.charCodeAt(i++);
				c2 = input.charCodeAt(i++);
				c3 = input.charCodeAt(i++);
				e1 = c1 >> 2;
				e2 = (c1 & 3) << 4 | c2 >> 4;
				e3 = (c2 & 15) << 2 | c3 >> 6;
				e4 = c3 & 63;
				if (isNaN(c2)) { e3 = e4 = 64; }
				else if (isNaN(c3)) { e4 = 64; }
				o += map.charAt(e1) + map.charAt(e2) + map.charAt(e3) + map.charAt(e4);
			}
			return o;
		},*/
		decode: function(input, utf8) {
			var o = "";
			var c1, c2, c3;
			var e1, e2, e3, e4;
			input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");
			for(var i = 0; i < input.length;) {
				e1 = map.indexOf(input.charAt(i++));
				e2 = map.indexOf(input.charAt(i++));
				e3 = map.indexOf(input.charAt(i++));
				e4 = map.indexOf(input.charAt(i++));
				c1 = e1 << 2 | e2 >> 4;
				c2 = (e2 & 15) << 4 | e3 >> 2;
				c3 = (e3 & 3) << 6 | e4;
				o += String.fromCharCode(c1);
				if (e3 != 64) { o += String.fromCharCode(c2); }
				if (e4 != 64) { o += String.fromCharCode(c3); }
			}
			return o;
		}
	};
})();

function s2a(s) {
	if(typeof Buffer !== 'undefined') return new Buffer(s, "binary");
	var w = s.split("").map(function(x){ return x.charCodeAt(0) & 0xff; });
	return w;
}

var __toBuffer, ___toBuffer;
__toBuffer = ___toBuffer = function(bufs) {
	var x = [];
	for(var i = 0; i != bufs[0].length; ++i) { x = x.concat(bufs[0][i]); }
	return x;
};
if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function(s,e) {return this.toString('hex',s,e);};
	Buffer.prototype.utf16le= function(s,e){return this.toString('utf16le',s,e).replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!');};
	Buffer.prototype.utf8 = function(s,e) { return this.toString('utf8',s,e); };
	Buffer.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return len > 0 ? this.utf8(i+4,i+4+len-1) : "";};
	Buffer.prototype.lpwstr = function(i) { var len = 2*this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};
	if(typeof cptable !== "undefined") Buffer.prototype.lpstr = function(i) {
		var len = this.readUInt32LE(i);
		if(len === 0) return "";
		return cptable.utils.decode(current_codepage,this.slice(i+4,i+4+len-1));
	};
	__toBuffer = function(bufs) { try { return Buffer.concat(bufs[0]); } catch(e) { return ___toBuffer(bufs);} };
}

var __readUInt8 = function(b, idx) { return b.readUInt8 ? b.readUInt8(idx) : b[idx]; };
var __readUInt16LE = function(b, idx) { return b.readUInt16LE ? b.readUInt16LE(idx) : b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = __readUInt16LE(b,idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
var __readUInt32LE = function(b, idx) { return b.readUInt32LE ? b.readUInt32LE(idx) : b[idx+3]*(1<<24)+b[idx+2]*(1<<16)+b[idx+1]*(1<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { if(b.readInt32LE) return b.readInt32LE(idx); var u = __readUInt32LE(b,idx); if(!(u & 0x80000000)) return u; return (0xffffffff - u + 1) * -1; };
var __readDoubleLE = function(b, idx) { return b.readDoubleLE ? b.readDoubleLE(idx) : readIEEE754(b, idx||0);};

var __hexlify = function(b,l) { if(b.hexlify) return b.hexlify((b.l||0), (b.l||0)+l); return b.slice(b.l||0,(b.l||0)+16).map(function(x){return (x<16?"0":"") + x.toString(16);}).join(""); };
var __unhexlify = function(s) { if(typeof Buffer !== 'undefined') return new Buffer(s, 'hex'); return s.match(/../g).map(function(x) { return parseInt(x,16);}); };

var __utf16le = function(b,s,e) { if(b.utf16le) return b.utf16le(s,e); var ss=[]; for(var i=s; i<e; i+=2) ss.push(String.fromCharCode(__readUInt16LE(b,i))); return ss.join("").replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!'); };

var __utf8 = function(b,s,e) { if(b.utf8) return b.utf8(s,e); var ss=[]; for(var i=s; i<e; i++) ss.push(String.fromCharCode(__readUInt8(b,i))); return ss.join(""); };

var __lpstr = function(b,i) { if(b.lpstr) return b.lpstr(i); var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var __lpwstr = function(b,i) { if(b.lpwstr) return b.lpwstr(i); var len = 2*__readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};

if(typeof cptable !== 'undefined') {
	__utf16le = function(b,s,e) { if(b.utf16le) return b.utf16le(s,e); return cptable.utils.decode(1200, b.slice(s,e)).replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!'); };
	__utf8 = function(b,s,e) { if(b.utf8) return b.utf8(s,e); return cptable.utils.decode(65001, b.slice(s,e)).replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!'); };
	__lpstr = function(b,i) { if(b.lpstr) return b.lpstr(i); var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i+4, i+4+len-1)) : "";};
	__lpwstr = function(b,i) { if(b.lpwstr) return b.lpwstr(i); var len = 2*__readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i+4,i+4+len-1)) : "";};
}

function bconcat(bufs) { return (typeof Buffer !== 'undefined') ? Buffer.concat(bufs) : [].concat.apply([], bufs); }

function ReadShift(size, t) {
	var o, oo=[], w, vv, i, loc; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o=(t==='u' ? __readUInt16LE : __readInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __readDoubleLE(this, this.l); break; }
		/* falls through */
		case 16: o = __hexlify(this, 16); break;

		case 'utf8': size = t; o = __utf8(this, this.l, this.l + size); break;
		case 'utf16le': size=2*t; o = __utf16le(this, this.l, this.l + size); break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr': o = __lpstr(this, this.l); size = 5 + o.length; break;

		case 'lpwstr': o = __lpwstr(this, this.l); size = 5 + o.length; if(o[o.length-1] == '\u0000') size += 2; break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		/* TODO: DBCS http://msdn.microsoft.com/en-us/library/cc194788.aspx */
		case 'dbcs': size = 2*t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); break;

		case 'sbcs': size = t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt8(this, loc)));
				loc+=1;
			} o = oo.join(""); break;

		case 'cstr': size = 0; o = "";
			while((w=__readUInt8(this, this.l + size++))!==0) oo.push(_getchar(w));
			o = oo.join(""); break;
		case 'wstr': size = 0; o = "";
			while((w=__readUInt16LE(this,this.l +size))!==0){oo.push(_getchar(w));size+=2;}
			size+=2; o = oo.join(""); break;
	}
	this.l+=size; return o;
}

function CheckField(hexstr, fld) {
	var b = this.slice(this.l, this.l+hexstr.length/2);
	var m = __hexlify(b,hexstr.length/2);
	if(m !== hexstr) throw (fld||"") + 'Expected ' + hexstr + ' saw ' + m;
	this.l += hexstr.length/2;
}

function prep_blob(blob, pos) {
	blob.l = pos || 0;
	//var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	//blob.read_shift = read;
	//blob.chk = chk;
	blob.read_shift = ReadShift;
	blob.chk = CheckField;
	//return [read, chk];
}

