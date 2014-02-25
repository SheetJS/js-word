function firstbyte(f,o) {
	switch((o||{}).type || "base64") {
		case 'buffer': return f[0];
		case 'base64': return Base64.decode(f.substr(0,12)).charCodeAt(0);
		case 'binary': return f.charCodeAt(0);
		default: return 0xD0;
	}
}

function xlsread(f, o) {
	/* TODO: enable browser xlml support */
	switch(firstbyte(f, o)) {
		case 0xD0: return parse_xlscfb(CFB.read(f, o), o);
		case 0x3C: return parse_xlml(f, o);
	}
}
var readFile = function(f,o) {
	var d = fs.readFileSync(f);
	if(!o) o = {};
	switch(firstbyte(d, {type:'buffer'})) {
		case 0xD0: return parse_xlscfb(CFB.read(d,{type:'buffer'}),o);
		case 0x3C: o.type = "file"; return parse_xlml(d, o);
		default: throw "Unsupported file";
	}
};
