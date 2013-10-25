var XLS;
var fs = require('fs');
describe('source', function() { it('should load', function() { XLS = require('./'); }); });
var files = fs.readdirSync('test_files').filter(function(x){return x.substr(-4)==".xls";});
files.forEach(function(x) {
	describe(x, function() {
		it('should parse ' + x, function() {
			XLS.readFile('./test_files/' + x);
		});
	});
});
