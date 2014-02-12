/* vim: set ts=2: */
var XLS;
var fs = require('fs'), assert = require('assert');
describe('source',function(){ it('should load', function(){ XLS = require('./'); });});

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(function(x){return x.substr(-4)==".xls" || x.substr(-8)==".xls.b64";});

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x) { return x.substr(0,31); }

function normalizecsv(x) { return x.replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }

var dir = "./test_files/";

function parsetest(x, wb) {
	describe(x + ' should have all bits', function() {
		var sname = dir + '2011/' + x + '.sheetnames';
		it('should have all sheets', function() {
			wb.SheetNames.forEach(function(y) { assert(wb.Sheets[y], 'bad sheet ' + y); });
		});
		it('should have the right sheet names', fs.existsSync(sname) ? function() {
			var file = fs.readFileSync(sname, 'utf-8');
			var names = wb.SheetNames.map(fixsheetname).join("\n") + "\n";
			assert.equal(names, file);
		} : null);
	});
	describe(x + ' should generate CSV', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				if(wb.SSF) XLS.SSF.load_table(wb.SSF);
				var csv = XLS.utils.make_csv(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate JSON', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				if(wb.SSF) XLS.SSF.load_table(wb.SSF);
				var json = XLS.utils.sheet_to_row_object_array(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate formulae', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				if(wb.SSF) XLS.SSF.load_table(wb.SSF);
				var json = XLS.utils.get_formulae(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate correct output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var name = (dir + x + '.' + i + '.csv');
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				if(wb.SSF) XLS.SSF.load_table(wb.SSF);
				var csv = XLS.utils.make_csv(wb.Sheets[ws]);
				assert.equal(normalizecsv(csv), normalizecsv(file), "CSV badness");
			} : null);
		});
	});
}

describe('should parse test files', function() {
	files.forEach(function(x) {
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = x.substr(-4) == ".b64" ? XLS.read(fs.readFileSync(dir + x, 'utf8'), {type: 'base64'}) : XLS.readFile(dir + x);
			if(x.substr(-4) === ".xls") parsetest(x, wb);
		});
	});
});

describe('other features', function() {
	it('should fail on passwords', function() {
		assert.throws(function() { XLS.readFile(dir + 'apachepoi_password.xls'); });
	});
	it('should read binary strings', function() {
		XLS.read(fs.readFileSync(dir + 'formula_stress_test.xls', 'binary'), {type:'binary'});
	});
});
