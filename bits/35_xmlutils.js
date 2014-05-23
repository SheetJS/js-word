var _chr = function(c) { return String.fromCharCode(c); };
var attregexg=/(\w+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
var attregex=/(\w+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
function parsexmltag(tag) {
	var words = tag.split(/\s+/);
	var z = {'0': words[0]};
	if(words.length === 1) return z;
	(tag.match(attregexg) || []).map(
		function(x){var y=x.match(attregex); z[y[1]] = unescapexml(y[2].substr(1,y[2].length-2)); });
	return z;
}

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = evert(encodings);
var rencstr = "&<>'\"".split("");

var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
