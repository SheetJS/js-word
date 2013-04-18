/* Buffer.concat was added in the 0.8 series, so this is for older versions */
if(typeof Buffer !== "undefined" && !Buffer.concat) 
Buffer.concat = function(list, length) {
	if (!Array.isArray(list)) {
		throw new TypeError('Usage: Buffer.concat(list, [length])');
	}

	if (list.length === 0) {
		return new Buffer(0);
	} else if (list.length === 1) {
		return list[0];
	}

	if (typeof length !== 'number') {
		length = 0;
		for (var i = 0; i < list.length; i++) {
			var buf = list[i];
			length += buf.length;
		}
	}

	var buffer = new Buffer(length);
	var pos = 0;
	for (var i = 0; i < list.length; i++) {
		var buf = list[i];
		buf.copy(buffer, pos);
		pos += buf.length;
	}
	return buffer;
};
