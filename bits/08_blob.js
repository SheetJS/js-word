
var Buffers = Array;

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
		encode: function(input, utf8) {
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
		},
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
	var w = s.split("").map(function(x){return x.charCodeAt(0);});
	return w;
}

if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function() { return this.toString('hex'); };
	Buffer.prototype.utf16le= function(s,e){return this.toString('utf16le',s,e).replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!');};
	Buffer.prototype.utf8 = function(s,e) { return this.toString('utf8',s,e); };
	Buffer.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return len > 0 ? this.utf8(i+4,i+4+len-1) : "";};
	Buffer.prototype.lpwstr = function(i) { var len = 2*this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};
	if(typeof cptable !== "undefined") Buffer.prototype.lpstr = function(i) {
		var len = this.readUInt32LE(i);
		if(len === 0) return "";
		if(typeof current_cptable === "undefined") return this.utf8(i+4,i+4+len-1);
		var t = Array(this.slice(i+4,i+4+len-1));
		//1console.log("start", this.l, len, t);
		var c, j = i+4, o = "", cc;
		for(;j!=i+4+len;++j) {
			c = this.readUInt8(j);
			cc = current_cptable.dec[c];
			if(typeof cc === 'undefined') {
				c = c*256 + this.readUInt8(++j);
				cc = current_cptable.dec[c];
			}
			if(typeof cc === 'undefined') throw "Unrecognized character " + c.toString(16);
			if(c === 0) break;
			o += cc;
		//1console.log(cc, cc.charCodeAt(0), o, this.l);
		}
		return o;
	};
}

Array.prototype.readUInt8 = function(idx) { return this[idx]; };
Array.prototype.readUInt16LE = function(idx) { return this[idx+1]*(1<<8)+this[idx]; };
Array.prototype.readInt16LE = function(idx) { var u = this.readUInt16LE(idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
Array.prototype.readUInt32LE = function(idx) { return this[idx+3]*(1<<24)+this[idx+2]*(1<<16)+this[idx+1]*(1<<8)+this[idx]; };
Array.prototype.readInt32LE = function(idx) { var u = this.readUInt32LE(idx); if(!(u & 0x80000000)) return u; return (0xffffffff - u + 1) * -1; };
Array.prototype.readDoubleLE = function(idx) { return readIEEE754(this, idx||0);};

Array.prototype.hexlify = function() { return this.map(function(x){return (x<16?"0":"") + x.toString(16);}).join(""); };

Array.prototype.utf16le = function(s,e) { var str = ""; for(var i=s; i<e; i+=2) str += String.fromCharCode(this.readUInt16LE(i)); return str.replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!'); };

Array.prototype.utf8 = function(s,e) { var str = ""; for(var i=s; i<e; i++) str += String.fromCharCode(this.readUInt8(i)); return str; };

Array.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return len > 0 ? this.utf8(i+4,i+4+len-1) : "";};
Array.prototype.lpwstr = function(i) { var len = 2*this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};

function bconcat(bufs) { return (typeof Buffer !== 'undefined') ? Buffer.concat(bufs) : [].concat.apply([], bufs); }

function ReadShift(size, t) {
	var o, w, vv, i, loc; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = this.readUInt8(this.l); break;
		case 2: o=t==='u'?this.readUInt16LE(this.l):this.readInt16LE(this.l);break;
		case 4: o = this.readUInt32LE(this.l); break;
		case 8: if(t === 'f') { o = this.readDoubleLE(this.l); break; }
		/* falls through */
		case 16: o = this.toString('hex', this.l,this.l+size); break;

		case 'utf8': size = t; o = this.utf8(this.l, this.l + size); break;
		case 'utf16le': size = 2*t; o = this.utf16le(this.l, this.l + size); break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr': o = this.lpstr(this.l); size = 5 + o.length; break;

		case 'lpwstr': o = this.lpwstr(this.l); size = 5 + o.length; if(o[o.length-1] == '\u0000') size += 2; break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		/* TODO: DBCS http://msdn.microsoft.com/en-us/library/cc194788.aspx */
		case 'dbcs': size = 2*t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = this.readUInt8(loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return o + vv;
				}
				o += String.fromCharCode(this.readUInt16LE(loc));
				loc+=2;
			} break;

		case 'sbcs': size = t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = this.readUInt8(loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return o + vv;
				}
				o += String.fromCharCode(this.readUInt8(loc));
				loc+=1;
			} break;

		case 'cstr': size = 0; o = "";
			while((w=this.readUInt8(this.l + size++))!==0) o+= String.fromCharCode(w);
			break;
		case 'wstr': size = 0; o = "";
			while((w=this.readUInt16LE(this.l +size))!==0){o+= String.fromCharCode(w);size+=2;}
			size+=2; break;
	}
	this.l+=size; return o;
}

function CheckField(hexstr, fld) {
	var m = this.slice(this.l, this.l+hexstr.length/2).hexlify('hex');
	if(m !== hexstr) throw (fld||"") + 'Expected ' + hexstr + ' saw ' + m;
	this.l += hexstr.length/2;
}

function WarnField(hexstr, fld) {
	var m = this.slice(this.l, this.l+hexstr.length/2).hexlify('hex');
	if(m !== hexstr) console.error((fld||"") + 'Expected ' + hexstr +' saw ' + m);
	this.l += hexstr.length/2;
}

function prep_blob(blob, pos) {
	blob.read_shift = ReadShift.bind(blob);
	blob.chk = CheckField;
	blob.l = pos || 0;
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	return [read, chk];
}

