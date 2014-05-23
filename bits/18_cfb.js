/* [MS-CFB] v20130118 */
/*if(typeof module !== "undefined" && typeof require !== 'undefined') CFB = require('cfb');
else*/ var CFB = (function(){
this.version = '0.9.1';
var exports = {};
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

/* [MS-CFB] 2.2 Compound File Header */
var blob = file.slice(0,512);
prep_blob(blob);
var read = ReadShift.bind(blob), chk = CheckField.bind(blob);
var j = 0, q;

// header signature 8
chk(HEADER_SIGNATURE, 'Header Signature: ');

// clsid 16
chk(HEADER_CLSID, 'CLSID: ');

// minor version 2
read(2);

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
var nsectors = Math.ceil((file.length - ssz)/ssz);
var sectors = [];
for(var i=1; i != nsectors; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);
sectors[nsectors-1] = file.slice(nsectors*ssz);

/** Chase down the rest of the DIFAT chain to build a comprehensive list
    DIFAT chains by storing the next sector number as the last 32 bytes */
function sleuth_fat(idx, cnt) {
	if(idx === ENDOFCHAIN) {
		if(cnt !== 0) throw "DIFAT chain shorter than expected";
		return;
	}
	if(idx !== FREESECT) {
		var sector = sectors[idx];
		for(var i = 0; i != ssz/4-1; ++i) {
			if((q = __readUInt32LE(sector,i*4)) === ENDOFCHAIN) break;
			fat_addrs.push(q);
		}
		sleuth_fat(__readUInt32LE(sector,ssz-4),cnt - 1);
	}
}
sleuth_fat(difat_start, ndfs);

/** DONT CAT THE FAT!  Just calculate where we need to go */
function get_buffer(byte_addr, bytes) {
	var addr = fat_addrs[Math.floor(byte_addr*4/ssz)];
	if(ssz - (byte_addr*4 % ssz) < (bytes || 0)) throw "FAT boundary crossed: " + byte_addr + " "+bytes+" "+ssz;
	return sectors[addr].slice((byte_addr*4 % ssz));
}

function get_buffer_u32(byte_addr) {
	return __readUInt32LE(get_buffer(byte_addr,4), 0);
}

function get_next_sector(idx) { return get_buffer_u32(idx); }

/** Chains */
var chkd = new Array(sectors.length), sector_list = [];
var get_sector = function get_sector(k) { return sectors[k]; };
for(i=0; i != sectors.length; ++i) {
	var buf = [], k = (i + dir_start) % sectors.length;
	if(chkd[k]) continue;
	for(j=k; j<=MAXREGSECT; buf.push(j),j=get_next_sector(j)) chkd[j] = true;
	sector_list[k] = {nodes: buf};
	sector_list[k].data = __toBuffer(Array(buf.map(get_sector)));
}
sector_list[dir_start].name = "!Directory";
if(nmfs > 0 && minifat_start !== ENDOFCHAIN) sector_list[minifat_start].name = "!MiniFAT";
sector_list[fat_addrs[0]].name = "!FAT";

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
var files = {}, Paths = [], FileIndex = [], FullPaths = [], FullPathDir = {};
function read_directory(idx) {
	var blob, read, w;
	var sector = sector_list[idx].data;
	for(var i = 0; i != sector.length; i+= 128) {
		blob = sector.slice(i, i+128);
		prep_blob(blob, 64);
		read = ReadShift.bind(blob);
		var namelen = read(2);
		if(namelen === 0) return;
		var name = __utf16le(blob,0,namelen-(Paths.length?2:0)); // OLE
		Paths.push(name);
		var o = { name: name };
		o.type = EntryTypes[read(1)];
		o.color = read(1);
		o.left = read(4); if(o.left === NOSTREAM) delete o.left;
		o.right = read(4); if(o.right === NOSTREAM) delete o.right;
		o.child = read(4); if(o.child === NOSTREAM) delete o.child;
		o.clsid = read(16);
		o.state = read(4);
		var ctime = read(8); if(ctime != "0000000000000000") o.ctime = ctime;
		var mtime = read(8); if(mtime != "0000000000000000") o.mtime = mtime;
		o.start = read(4);
		o.size = read(4);
		if(o.type === 'root') { //root entry
			minifat_store = o.start;
			if(nmfs > 0 && minifat_store !== ENDOFCHAIN) sector_list[minifat_store].name = "!StreamData";
			minifat_size = o.size;
		} else if(o.size >= ms_cutoff_size) {
			o.storage = 'fat';
			if(!sector_list[o.start] && dir_start > 0) o.start = (o.start + dir_start) % sectors.length;
			sector_list[o.start].name = o.name;
			o.content = sector_list[o.start].data.slice(0,o.size);
			prep_blob(o.content);
		} else {
			o.storage = 'minifat';
			w = o.start * mssz;
			if(minifat_store !== ENDOFCHAIN && o.start !== ENDOFCHAIN) {
				o.content = sector_list[minifat_store].data.slice(w,w+o.size);
				prep_blob(o.content);
			}
		}
		if(o.ctime) {
			var ct = blob.slice(blob.l-24, blob.l-16);
			var c2 = (__readUInt32LE(ct,4)/1e7)*Math.pow(2,32)+__readUInt32LE(ct,0)/1e7;
			o.ct = new Date((c2 - 11644473600)*1000);
		}
		if(o.mtime) {
			var mt = blob.slice(blob.l-16, blob.l-8);
			var m2 = (__readUInt32LE(mt,4)/1e7)*Math.pow(2,32)+__readUInt32LE(mt,0)/1e7;
			o.mt = new Date((m2 - 11644473600)*1000);
		}
		files[name] = o;
		FileIndex.push(o);
	}
}
read_directory(dir_start);

/* [MS-CFB] 2.6.4 Red-Black Tree */
function build_full_paths(Dir, pathobj, paths, patharr) {
	var i;
	var dad = new Array(patharr.length);

	var q = new Array(patharr.length);

	for(i=0; i != dad.length; ++i) { dad[i]=q[i]=i; paths[i]=patharr[i]; }

	while(q.length > 0) {
		for(i = q[0]; typeof i !== "undefined"; i = q.shift()) {
			if(dad[i] === i) {
				if(Dir[i].left && dad[Dir[i].left] != Dir[i].left) dad[i] = dad[Dir[i].left];
				if(Dir[i].right && dad[Dir[i].right] != Dir[i].right) dad[i] = dad[Dir[i].right];
			}
			if(Dir[i].child) dad[Dir[i].child] = i;
			if(Dir[i].left) { dad[Dir[i].left] = dad[i]; q.push(Dir[i].left); }
			if(Dir[i].right) { dad[Dir[i].right] = dad[i]; q.push(Dir[i].right); }
		}
		for(i=1; i != dad.length; ++i) if(dad[i] === i) {
			if(Dir[i].right && dad[Dir[i].right] != Dir[i].right) dad[i] = dad[Dir[i].right];
			else if(Dir[i].left && dad[Dir[i].left] != Dir[i].left) dad[i] = dad[Dir[i].left];
		}
	}

	for(i=1; i !== paths.length; ++i) {
		if(Dir[i].type === "unknown") continue;
		var j = dad[i];
		if(j === 0) paths[i] = paths[0] + "/" + paths[i];
		else while(j !== 0) {
			paths[i] = paths[j] + "/" + paths[i];
			j = dad[j];
		}
		dad[i] = 0;
	}

	paths[0] += "/";
	for(i=1; i !== paths.length; ++i) if(Dir[i].type !== 'stream') paths[i] += "/";
	for(i=0; i !== paths.length; ++i) pathobj[paths[i]] = FileIndex[i];
}
build_full_paths(FileIndex, FullPathDir, FullPaths, Paths);

var root_name = Paths.shift();
Paths.root = root_name;

/* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
function find_path(path) {
	if(path[0] === "/") path = root_name + path;
	var UCNames = (path.indexOf("/") !== -1 ? FullPaths : Paths).map(function(x) { return x.toUpperCase(); });
	var UCPath = path.toUpperCase();
	var w = UCNames.indexOf(UCPath);
	if(w === -1) return null;
	return path.indexOf("/") !== -1 ? FileIndex[w] : files[Paths[w]];
}

var rval = {
	raw: {header: header, sectors: sectors},
	FileIndex: FileIndex,
	FullPaths: FullPaths,
	FullPathDir: FullPathDir,
	find: find_path
};

for(var name in files) {
	switch(name) {
		/* [MS-OSHARED] 2.3.3.2.2 Document Summary Information Property Set */
		case '!DocumentSummaryInformation':
			try { rval.DocSummary = parse_PropertySetStream(files[name], DocSummaryPIDDSI); } catch(e) { } break;
		/* [MS-OSHARED] 2.3.3.2.1 Summary Information Property Set*/
		case '!SummaryInformation':
			try { rval.Summary = parse_PropertySetStream(files[name], SummaryPIDSI); } catch(e) { } break;
	}
}

return rval;
} // parse


function readFileSync(filename) {
	var fs = require('fs');
	var file = fs.readFileSync(filename);
	return parse(file);
}

function readSync(blob, options) {
	var o = options || {};
	switch((o.type || "base64")) {
		case "file": return readFileSync(blob);
		case "base64": blob = Base64.decode(blob);
		/* falls through */
		case "binary": blob = s2a(blob); break;
	}
	return parse(blob);
}

exports.read = readSync;
exports.parse = parse;
return exports;
})();

/** CFB Constants */
{
	/* 2.1 Compund File Sector Numbers and Types */
	var MAXREGSECT = 0xFFFFFFFA;
	var DIFSECT = 0xFFFFFFFC;
	var FATSECT = 0xFFFFFFFD;
	var ENDOFCHAIN = 0xFFFFFFFE;
	var FREESECT = 0xFFFFFFFF;
	/* 2.2 Compound File Header */
	var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
	var HEADER_MINOR_VERSION = '3e00';
	var MAXREGSID = 0xFFFFFFFA;
	var NOSTREAM = 0xFFFFFFFF;
	var HEADER_CLSID = '00000000000000000000000000000000';
	/* 2.6.1 Compound File Directory Entry */
	var EntryTypes = ['unknown','storage','stream','lockbytes','property','root'];
}

if(typeof require !== 'undefined' && typeof exports !== 'undefined') {
	var fs = require('fs');
	//exports.read = CFB.read;
	//exports.parse = CFB.parse;
	//exports.ReadShift = ReadShift;
	//exports.prep_blob = prep_blob;
}
