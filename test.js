var XLS;
var fs = require('fs');
describe('source', function() { it('should load', function() { XLS = require('./'); }); });
describe('parsing test files', function() {
	var files = fs.readdirSync('test_files').filter(function(x){return x.substr(-4)==".xls";});
	files.forEach(function(x) {
		it('should parse ' + x, function() {
			XLS.readFile('./test_files/' + x);
		});
	});
});
