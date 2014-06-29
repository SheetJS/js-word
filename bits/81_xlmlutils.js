// TODO: CP remap (need to read file version to determine OS)
var encregex = /&[a-z]*;/g, coderegex = /_x([0-9a-fA-F]+)_/g;
function coderepl(m,c) {return _chr(parseInt(c,16));}
function encrepl($$) { return encodings[$$]; }
function unescapexml(s){
	if(s.indexOf("&") > -1) s = s.replace(encregex, encrepl);
	return s.indexOf("_") === -1 ? s : s.replace(coderegex,coderepl);
}

function parsexmlbool(value, tag) {
	switch(value) {
		case '1': case 'true': case 'TRUE': return true;
		/* case '0': case 'false': case 'FALSE':*/
		default: return false;
	}
}

// matches <foo>...</foo> extracts content
function matchtag(f,g) {return new RegExp('<'+f+'(?: xml:space="preserve")?>([^\u2603]*)</'+f+'>',(g||"")+"m");}

/* TODO: handle codepages */
var entregex = /&#(\d+);/g;
function entrepl($$,$1) { return String.fromCharCode(parseInt($1,10)); }
function fixstr(str) {
	return str.replace(entregex,entrepl);
}
