function isval(x) { return x !== undefined && x !== null; }

function keys(o) { return Object.keys(o); }

function evert(obj, arr) {
	var o = {};
	var K = keys(obj);
	for(var i = 0; i < K.length; ++i) {
		var k = K[i];
		if(!arr) o[obj[k]] = k;
		else (o[obj[k]]=o[obj[k]]||[]).push(k);
	}
	return o;
}

function rgb2Hex(rgb) {
	for(var i=0,o=1; i!=3; ++i) o = o*256 + (rgb[i]>255?255:rgb[i]<0?0:rgb[i]);
	return o.toString(16).toUpperCase().substr(1);
}
