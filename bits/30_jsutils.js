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

