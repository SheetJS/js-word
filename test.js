/* vim: set ts=2: */
var XLS;
var fs = require('fs');
describe('source', function() { it('should load', function() { XLS = require('./'); }); });

var files = fs.readdirSync('test_files').filter(function(x){return x.substr(-4)==".xls";});

function parsetest(x, wb) {
	describe(x + ' should generate correct output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var name = ('./test_files/' + x + '.' + i + '.csv');
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = XLS.utils.make_csv(wb.Sheets[ws]);
				if(file.replace(/"/g,"") != csv.replace(/"/g,"")) throw "CSV badness";
			} : null);
		});
	});
}

describe('should parse test files', function() {
	files.forEach(function(x) {
		it('should parse ' + x, function() {
			var wb = XLS.readFile('./test_files/' + x);
			parsetest(x, wb);
		});
	});
});
