function writeSync(wb, opts) {
	var o = opts||{};
	switch(o.bookType) {
		case 'xml': return write_xlml(wb, o);
		default: throw 'unsupported output format ' + o.bookType;
	}
}

function writeFileSync(wb, filename, opts) {
	var o = opts|{}; o.type = 'file';
	o.file = filename;
	switch(o.file.substr(-4).toLowerCase()) {
		case '.xls': o.bookType = 'xls'; break;
		case '.xml': o.bookType = 'xml'; break;
	}
	return writeSync(wb, o);
}

