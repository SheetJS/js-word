var fs;
if(typeof exports !== 'undefined') {
	if(typeof module !== 'undefined' && module.exports) {
		fs = require('fs');
	}
}
function firstbyte(f,o) {
	switch((o||{}).type || "base64") {
		case 'buffer': return f[0];
		case 'base64': return Base64.decode(f.substr(0,12)).charCodeAt(0);
		case 'binary': return f.charCodeAt(0);
		case 'array': return f[0];
		default: throw new Error("Unrecognized type " + o.type);
	}
}

function xlsread(f, o) {
	var n=0;
	if(!o) o = {};
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(f)) ? "buffer" : "base64";
	switch((n = firstbyte(f, o))) {
		case 0xD0: return parse_xlscfb(CFB.read(f, o), o);
		case 0x09: return parse_xlscfb(s2a(o.type === 'base64' ? Base64.decode(f) : f), o);
		case 0x3C: return parse_xlml(f, o);
		case 0xEF: return parse_xlml(f, o);
		default: throw new Error("Unsupported file " + n);
	}
}
var readFile = function(f,o) {
	var d = fs.readFileSync(f);
	if(!o) o = {};
	switch(firstbyte(d, {type:'buffer'})) {
		case 0xD0: return parse_xlscfb(CFB.read(d,{type:'buffer'}),o);
		case 0x09: return parse_xlscfb(d, o);
		case 0x3C: return parse_xlml(d, (o.type="buffer",o));
		case 0xEF: return parse_xlml(d, (o.type="buffer",o));
		default: throw "Unsupported file";
	}
};

