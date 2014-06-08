// TODO: CP remap (need to read file version to determine OS)
function unescapexml(s){
	if(s.indexOf("&") > -1) s = s.replace(/&[a-z]*;/g, function($$) { return encodings[$$]; });
	return s.indexOf("_") === -1 ? s : s.replace(/_x([0-9a-fA-F]*)_/g,function(m,c) {return _chr(parseInt(c,16));});
}

function parsexmlbool(value, tag) {
	switch(value) {
		case '0': case 0: case 'false': case 'FALSE': return false;
		case '1': case 1: case 'true': case 'TRUE': return true;
		default: throw "bad boolean value " + value + " in "+(tag||"?");
	}
}

// matches <foo>...</foo> extracts content
function matchtag(f,g) {return new RegExp('<'+f+'(?: xml:space="preserve")?>([^\u2603]*)</'+f+'>',(g||"")+"m");}

/* TODO: handle codepages */
function fixstr(str) {
	str = str.replace(/&#([0-9]+);/g,function($$,$1) { return String.fromCharCode($1); });
	return (typeof cptable === "undefined") ? str : str;
}
