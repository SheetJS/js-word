/* vim: set ts=2: */
/*jshint eqnull:true */

var Buffers;

function readIEEE754(buf, idx, isLE, nl, ml) {
	if(isLE === undefined) isLE = true;
	if(!nl) nl = 8;
	if(!ml && nl === 8) ml = 52;
  var e, m, el = nl * 8 - ml - 1, eMax = (1 << el) - 1, eBias = eMax >> 1,
      bits = -7, d = isLE ? -1 : 1, i = isLE ? (nl - 1) : 0, s = buf[idx + i];

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
	var w = s.split("").map(function(x){return x.charCodeAt(0);}); 
	return w;
}

if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function() { return this.toString('hex'); };
	Buffer.prototype.utf16le= function(s,e){return this.toString('utf16le',s,e).replace(/\u0000/,'').replace(/[\u0001-\u0005]/,'!');};
	Buffer.prototype.utf8 = function(s,e) { return this.toString('utf8',s,e); };
	Buffer.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};
}

Array.prototype.readUInt8 = function(idx) { return this[idx]; };
Array.prototype.readUInt16LE = function(idx) { return this[idx+1]*(1<<8)+this[idx]; };
Array.prototype.readInt16LE = function(idx) { var u = this.readUInt16LE(idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
Array.prototype.readUInt32LE = function(idx) { return this[idx+3]*(1<<24)+this[idx+2]*(1<<16)+this[idx+1]*(1<<8)+this[idx]; };
Array.prototype.readDoubleLE = function(idx) { return readIEEE754(this, idx||0);};

Array.prototype.hexlify = function() { return this.map(function(x){return (x<16?"0":"") + x.toString(16);}).join(""); };

Array.prototype.utf16le = function(s,e) { var str = ""; for(var i=s; i<e; i+=2) str += String.fromCharCode(this.readUInt16LE(i)); return str.replace(/\u0000/,'').replace(/[\u0001-\u0005]/,'!'); };

Array.prototype.utf8 = function(s,e) { var str = ""; for(var i=s; i<e; i++) str += String.fromCharCode(this.readUInt8(i)); return str; };
Array.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};


function ReadShift(size, t) {
	var o; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = this.readUInt8(this.l); break;
		case 2:o=t==='u'?this.readUInt16LE(this.l):this.readInt16LE(this.l);break;
		case 4: o = this.readUInt32LE(this.l); break;
		case 8: if(t === 'f') { o = this.readDoubleLE(this.l); break; }
		/* falls through */
		case 16: o = this.toString('hex', this.l,this.l+size); break;
		case 'lpstr': o = this.lpstr(this.l); size = 5 + o.length; break;
		case 'utf8': size = t; o = this.utf8(this.l, this.l + size); break;
		case 'utf16le': size = 2*t; o = this.utf16le(this.l, this.l + size); break;
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

// Implements [MS-CFB] v20120705
var CFB = (function(){

function parse(file) {


var mver = 3; // major version
var ssz = 512; // sector size
var mssz = 64; // mini sector size
var nds = 0; // number of directory sectors
var nfs = 0; // number of FAT sectors
var nmfs = 0; // number of mini FAT sectors
var ndfs = 0; // number of DIFAT sectors
var dir_start = 0; // first directory sector location
var minifat_start = 0; // first mini FAT sector location
var difat_start = 0; // first mini FAT sector location

var ms_cutoff_size = 4096; // mini stream cutoff size
var minifat_store = 0; // first sector with minifat data
var minifat_size = 0; // size of minifat data

var fat_addrs = []; // locations of FAT sectors

/** Parse Header [MS-CFB] 2.2 */
var blob = file.slice(0,512);
prep_blob(blob);
var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
var wrn = WarnField.bind(blob);
var j = 0;

// header signature 8
chk(HEADER_SIGNATURE, 'Header Signature: ');

// clsid 16
chk(HEADER_CLSID, 'CLSID: ');
	
// minor version 2
wrn(HEADER_MINOR_VERSION, 'Minor Version: ');

// major version 3
mver = read(2);
switch(mver) {
	case 3: ssz = 512; break;
	case 4: ssz = 4096; break;
	default: throw "Major Version: Expected 3 or 4 saw " + mver;
}

// reprocess header
var pos = blob.l;
blob = file.slice(0,ssz);
prep_blob(blob,pos);
read = ReadShift.bind(blob);
chk = CheckField.bind(blob);
var header = file.slice(0,ssz);

// Byte Order TODO
chk('feff', 'Byte Order: ');

// Sector Shift
switch((q = read(2))) {
	case 0x09: if(mver !== 3) throw 'MajorVersion/SectorShift Mismatch'; break;
	case 0x0c: if(mver !== 4) throw 'MajorVersion/SectorShift Mismatch'; break;
	default: throw 'Sector Shift: Expected 9 or 12 saw ' + q;
}

// Mini Sector Shift
chk('0600', 'Mini Sector Shift: ');

// Reserved
chk('000000000000', 'Mini Sector Shift: ');

// Number of Directory Sectors
nds = read(4);
if(mver === 3 && nds !== 0) throw '# Directory Sectors: Expected 0 saw ' + nds;

// Number of FAT Sectors
nfs = read(4);

// First Directory Sector Location
dir_start = read(4);

// Transaction Signature TODO
read(4);

// Mini Stream Cutoff Size TODO
chk('00100000', 'Mini Stream Cutoff Size: ');

// First Mini FAT Sector Location
minifat_start = read(4);

// Number of Mini FAT Sectors
nmfs = read(4);

// First DIFAT sector location
difat_start = read(4);

// Number of DIFAT Sectors
ndfs = read(4);

// Grab FAT Sector Locations
for(j = 0; blob.l != 512; ) {
	if((q = read(4))>=MAXREGSECT) break;
	fat_addrs[j++] = q;
}


/** Break the file up into sectors */
if(file.length%ssz!==0) throw "File Length: Expected multiple of "+ssz;

var nsectors = (file.length - ssz)/ssz;
var sectors = [];
for(var i=1; i != nsectors + 1; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);

/** Chase down the rest of the DIFAT chain to build a comprehensive list
    DIFAT chains by storing the next sector number as the last 32 bytes */
function sleuth_fat(idx, cnt) {
	if(idx === ENDOFCHAIN) {
		if(cnt !== 0) throw "DIFAT chain shorter than expected";
		return;
	}
	var sector = sectors[idx];
	for(var i = 0; i != ssz/4-1; ++i) {
		if((q = sector.readUInt32LE(i*4)) === ENDOFCHAIN) break;
		fat_addrs.push(q);
	}
	sleuth_fat(sector.readUInt32LE(ssz-4),cnt - 1);
}
sleuth_fat(difat_start, ndfs);

/** DONT CAT THE FAT!  Just calculate where we need to go */
function get_buffer(byte_addr, bytes) {
	var addr = fat_addrs[Math.floor(byte_addr*4/ssz)];
	if(ssz - (byte_addr*4 % ssz) < (bytes || 0))
		throw "FAT boundary crossed: " + byte_addr + " "+bytes+" "+ssz;
	return sectors[addr].slice((byte_addr*4 % ssz));
}

function get_buffer_u32(byte_addr) {
	return get_buffer(byte_addr,4).readUInt32LE(0);
}

function get_next_sector(idx) { return get_buffer_u32(idx); }

/** Chains */
var chkd = new Array(sectors.length), sector_list = [];
for(i=0; i != sectors.length; ++i) {
	var buf = [];
	if(chkd[i]) continue;
	for(j=i; j<=MAXREGSECT; buf.push(j),j=get_next_sector(j)) chkd[j] = true;
	sector_list[i] = {nodes: buf};
	sector_list[i].data = Buffers(buf.map(function(k){return sectors[k];})).toBuffer();
}
sector_list[dir_start].name = "!Directory";
if(nmfs > 0) sector_list[minifat_start].name = "!MiniFAT";
sector_list[fat_addrs[0]].name = "!FAT";

/** read directory structure */
var files = {}, Paths = [];
function read_directory(idx) {
	var blob, read;
	var sector = sector_list[idx].data;
	for(var i = 0; i != sector.length; i+= 128, l = 64) {
		blob = sector.slice(i, i+128);
		prep_blob(blob, 64);
		read = ReadShift.bind(blob);
		var namelen = read(2);
		if(namelen === 0) return;
		var name = blob.utf16le(0,namelen-(Paths.length?2:0)); // OLE
		Paths.push(name);
		var o = { name: name };
		o.type = EntryTypes[read(1)];
		o.color = read(1);
		o.left = read(4); if(o.left === NOSTREAM) delete o.left;
		o.right = read(4); if(o.right === NOSTREAM) delete o.right;
		o.child = read(4); if(o.child === NOSTREAM) delete o.child;
		o.clsid = read(16);
		o.state = read(4);
		o.ctime = read(8);
		o.mtime = read(8);
		o.start = read(4);
		o.size = read(4);
		if(o.type === 'root') { //root entry
			minifat_store = o.start;
			if(nmfs > 0) sector_list[minifat_store].name = "!StreamData";
			minifat_size = o.size;
		} else if(o.size >= ms_cutoff_size) {
			o.storage = 'fat';
			sector_list[o.start].name = o.name;
			o.content = sector_list[o.start].data.slice(0,o.size);
			prep_blob(o.content);
		} else {
			o.storage = 'minifat';
			w = o.start * mssz;
			o.content = sector_list[minifat_store].data.slice(w,w+o.size);
			prep_blob(o.content);
		}
		files[name] = o;
	}
}
read_directory(dir_start);

var root_name = Paths.shift();

var rval = {
	raw: {header: header, sectors: sectors},
	Paths: Paths,
	Directory: files 
};

for(var name in files) {
	switch(name) {
		case '!DocumentSummaryInformation': /* MS-OSHARED 2.3.3.2.2 */
			rval.DocSummary = parse_PSS(files[name],DocSummaryPIDDSI); break; 
		case '!SummaryInformation': /* MS-OSHARED 2.3.3.2.1 */ 
			rval.Summary = parse_PSS(files[name], SummaryPIDSI); break; 
	}
}

return rval;
} // parse


// 2.3.3.1.4
function parse_lpstr(blob) {
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var str = read('lpstr');
	return str;
}

function parse_string(blob, type) {
	switch(type) {
		case VT_LPSTR: return parse_lpstr(blob);
		case VT_LPWSTR: return parse_lpwstr(blob);
		default: throw "Unrecognized string type " + type;
	}
}

function parse_unaligned_vec_lpstr(blob) {
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var length = read(4);
	var ret = [];
	for(var i = 0; i != length; ++i) ret[i] = read('lpstr');
	return ret;
}

/* 2.3.3.1.14 VtVecHeadingPairValue */
function parse_heading_pair_vec(blob) {
	// TODO: 
}

/* [MS-OLEPS] 2.8 FILETIME */
function parse_datetime(blob) {
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var dwLowDateTime = read(4), dwHighDateTime = read(4);
	return [dwLowDateTime, dwHighDateTime];
}

/* [MS-OLEPS] 2.18.1 Dictionary (2.17, 2.16) */
function parse_dictionary(blob,CodePage) {
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var cnt = read(4);
	var dict = {};
	for(var j = 0; j != cnt; ++j) {
		var pid = read(4);
		var len = read(4);
		dict[pid] = read((CodePage === 0x4B0 ?'utf16le':'utf8'), len).replace(/\u0000/g,'').replace(/[\u0001-\u0005]/g,'!');
	}
	if(blob.l % 4) blob.l = (blob.l>>2+1)<<2;
	return dict;
}

/* [MS-OLEPS] 2.9 */
function parse_blob(blob) {
	var size = blob.read_shift(4);
	var bytes = blob.slice(blob.l,blob.l+size);
	if(blob.l % 4) blob.l = (blob.l>>2+1)<<2;
	return bytes;
}

function parse_property(blob, type) {
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var t = read(2), ret;
	read(2);
	if(t !== type && VT_CUSTOM.indexOf(type)===-1) throw 'Expected type ' + type + ' saw ' + t;
	switch(type) {
		case VT_I2: ret = read(2, 'i'); read(2); return ret;
		case VT_I4: ret = read(4, 'i'); return ret;
		case VT_LPSTR: return parse_lpstr(blob, t).replace(/\u0000/g,'');
		case VT_STRING: return parse_string(blob, t).replace(/\u0000/g,'');
		case VT_BOOL: return read(4) !== 0x0;
		case VT_FILETIME: return parse_datetime(blob);
		case VT_VECTOR | VT_LPSTR: return parse_unaligned_vec_lpstr(blob);
		case VT_VECTOR | VT_VARIANT: return parse_heading_pair_vec(blob);
	}
}

/* 2.20 PropertySet */
function parse_PSet(blob, PIDSI) {
	var start_addr = blob.l;
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
	var size = read(4);
	var NumProps = read(4);
	var Props = [], i = 0;
	var CodePage = 0;
	var Dictionary = -1, DictObj;
	for(i = 0; i != NumProps; ++i) {
		var PropID = read(4);
		var Offset = read(4);
		Props[i] = [PropID, Offset + start_addr];
	}
	var PropH = {};
	for(i = 0; i != NumProps; ++i) {
		if(blob.l !== Props[i][1]) throw "Read Error: Expected address " + Props[i][1] + ' at ' + blob.l + ' (' + i;
		if(PIDSI) {
			var piddsi = PIDSI[Props[i][0]];
			PropH[piddsi.n] = parse_property(blob, piddsi.t);
		} else {
			if(Props[i][0] === 0x1) {
				CodePage = PropH.CodePage = parse_property(blob, VT_I2);
				if(Dictionary !== -1) {
					var oldpos = blob.l;
					blob.l = Props[Dictionary][1];
					DictObj = parse_dictionary(blob,CodePage);
					blob.l = oldpos;
				}
			} else if(Props[i][0] === 0) {
				if(CodePage === 0) { Dictionary = i; blob.l = Props[i+1][1]; continue; }
				DictObj = parse_dictionary(blob,CodePage);
			} else {
				var name = DictObj[Props[i][0]];
				var val;
				switch(blob[blob.l]) {
					//TODO
					case 0x41: val = parse_blob(blob); break;
				}
				PropH[name] = val;
			}
		}
	}
	blob.l = start_addr + size; /* step ahead to skip padding */
	return PropH;
}

/* MS-OLEPS 2.21 PropertySetStream */
function parse_PSS(file, PIDSI) {
	var blob = file.content;
	prep_blob(blob);
	var read = ReadShift.bind(blob), chk = CheckField.bind(blob);

	var NumSets, FMTID0, FMTID1, Offset0, Offset1;
	chk('feff', 'Byte Order: ');
	
	var vers = read(2); // TODO: check version 
	var SystemIdentifier = read(4);	
	chk(HEADER_CLSID, 'CLSID: ');
	NumSets = read(4); 
	if(NumSets !== 1 && NumSets !== 2) throw "Unrecognized #Sets: " + NumSets;
	FMTID0 = read(16); Offset0 = read(4);

	if(NumSets === 1 && Offset0 !== blob.l) throw "Length mismatch";
	else if(NumSets === 2) { FMTID1 = read(16); Offset1 = read(4); }
	var PSet0 = parse_PSet(blob, PIDSI);

	var rval = { SystemIdentifier: SystemIdentifier };
	for(var y in PSet0) rval[y] = PSet0[y]; 
	//rval.blob = blob;
	rval.FMTID = FMTID0;
	//rval.PSet0 = PSet0;
	if(NumSets === 1) return rval;
	var PSet1 = parse_PSet(blob, null);
	for(var y in PSet1) rval[y] = PSet1[y]; 
	rval.FMTID = [FMTID0, FMTID1]; // TODO: verify FMTID0/1
	return rval;
}

function readFileSync(filename) {
	var fs = require('fs');
	var file = fs.readFileSync(filename);
	//var file = s2a(Base64.decode(fs.readFileSync(filename).toString('base64')));
	return parse(file);
}

function readSync(blob, options) {
	var o = options || {};
	switch((o.type || "base64")) {
		case "file": return readFileSync(blob);
		case "base64": blob = Base64.decode(blob); break; 
		case "binary": blob = s2a(blob); break;
	}
	return parse(blob);
}

this.read = readSync;
this.parse = parse;
return this;
})();

{ /* Constants */
// Property Types 2.2 Table
{
	var VT_EMPTY    = 0x0000;
	var VT_NULL     = 0x0001;
	var VT_I2       = 0x0002;
	var VT_I4       = 0x0003;
	var VT_R4       = 0x0004;
	var VT_R8       = 0x0005;
	var VT_CY       = 0x0006;
	var VT_DATE     = 0x0007;
	var VT_BSTR     = 0x0008;
	var VT_ERROR    = 0x000A;
	var VT_BOOL     = 0x000B;
	var VT_VARIANT  = 0x000C;
	var VT_DECIMAL  = 0x000E;
	var VT_I1       = 0x0010;
	var VT_UI1      = 0x0011;
	var VT_UI2      = 0x0012;
	var VT_UI4      = 0x0013;
	var VT_I8       = 0x0014;
	var VT_UI8      = 0x0015;
	var VT_INT      = 0x0016;
	var VT_UINT     = 0x0017;
	var VT_LPSTR    = 0x001E;
	var VT_LPWSTR   = 0x001F;
	var VT_FILETIME = 0x0040;
	var VT_BLOB     = 0x0041;
	var VT_STREAM   = 0x0042;
	var VT_STORAGE  = 0x0043;
	var VT_STREAMED_Object  = 0x0044;
	var VT_STORED_Object    = 0x0045;
	var VT_BLOB_Object      = 0x0046;
	var VT_CF       = 0x0047;
	var VT_CLSID    = 0x0048;
	var VT_VERSIONED_STREAM = 0x0049;
	var VT_VECTOR   = 0x1000;
	var VT_ARRAY    = 0x2000;

	var VT_STRING   = 0x0050; // 2.3.3.1.11 VtString
	var VT_CUSTOM   = [VT_STRING];
}

/* MS-OSHARED 2.3.3.2.2.1 Document Summary Information PIDDSI */
var DocSummaryPIDDSI = {
	0x01: { n: 'CodePage', t: VT_I2 },
	0x02: { n: 'Category', t: VT_STRING },
	0x03: { n: 'PresentationFormat', t: VT_STRING },
	0x04: { n: 'ByteCount', t: VT_I4 },
	0x05: { n: 'LineCount', t: VT_I4 },
	0x06: { n: 'ParagraphCount', t: VT_I4 },
	0x07: { n: 'SlideCount', t: VT_I4 },
	0x08: { n: 'NoteCount', t: VT_I4 },
	0x09: { n: 'HiddenCount', t: VT_I4 },
	0x0a: { n: 'MultimediaClipCount', t: VT_I4 },
	0x0b: { n: 'Scale', t: VT_BOOL },
	0x0c: { n: 'HeadingPair', t: VT_VECTOR | VT_VARIANT },
	0x0d: { n: 'DocParts', t: VT_VECTOR | VT_LPSTR },
	0x0e: { n: 'Manager', t: VT_STRING },
	0x0f: { n: 'Company', t: VT_STRING },
	0x10: { n: 'LinksDirty', t: VT_BOOL },
	0x11: { n: 'CharacterCount', t: VT_I4 },
	0x13: { n: 'SharedDoc', t: VT_BOOL },
	0x16: { n: 'HLinksChanged', t: VT_BOOL },
	0x17: { n: 'Version', t: VT_I4 },
	0xFF: {}
};

var SummaryPIDSI = {
	0x01: { n: 'CodePage', t: VT_I2 },
	0x02: { n: 'Title', t: VT_LPSTR },
	0x03: { n: 'Subject', t: VT_LPSTR },
	0x04: { n: 'Author', t: VT_LPSTR },
	0x05: { n: 'Keywords', t: VT_LPSTR },
	0x06: { n: 'Comments', t: VT_LPSTR },
	0x07: { n: 'Template', t: VT_LPSTR },
	0x08: { n: 'LastAuthor', t: VT_LPSTR },
	0x09: { n: 'RevNumber', t: VT_LPSTR },
	0x0A: { n: 'EditTime', t: VT_FILETIME },
	0x0B: { n: 'LastPrinted', t: VT_FILETIME },
	0x0C: { n: 'CreateTime', t: VT_FILETIME },
	0x0D: { n: 'SaveTime', t: VT_FILETIME },
	0x0E: { n: 'PageCount', t: VT_I4 },
	0x0F: { n: 'WordCount', t: VT_I4 },
	0x10: { n: 'CharCount', t: VT_I4 },
	0x11: { n: 'Thumbnail', t: VT_CF },
	0x12: { n: 'ApplicationName', t: VT_LPSTR },
	0x13: { n: 'DocumentSecurity', t: VT_I4 },
	0xFF: {}	
};

/* CFB Constants */
{
	var MAXREGSECT = 0xFFFFFFFA; 
	var DIFSECT = 0xFFFFFFFC; 
	var FATSECT = 0xFFFFFFFD;
	var ENDOFCHAIN = 0xFFFFFFFE;
	var FREESECT = 0xFFFFFFFF;
	var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
	var HEADER_MINOR_VERSION = '3e00';
	var MAXREGSID = 0xFFFFFFFA;
	var NOSTREAM = 0xFFFFFFFF;
	var HEADER_CLSID = '00000000000000000000000000000000';

	var EntryTypes = ['unknown','storage','stream',null,null,'root'];
}

}

if(typeof require !== 'undefined' && typeof exports !== 'undefined') {
	Buffers = require('buffers');
	var fs = require('fs');
	exports.read = CFB.read;
	exports.parse = CFB.parse;
	exports.main = function(args) {
		var cfb = CFB.read(args[0], {type:'file'});
		console.log(cfb);
	};
	if(typeof module !== 'undefined' && require.main === module)
		exports.main(process.argv.slice(2));
} else {
	Buffers = Array;
	Buffers.prototype.toBuffer = function() {
		var x = [];
		for(var i = 0; i != this[0].length; ++i) { x = x.concat(this[0][i]); }
		return x;
	};
}
