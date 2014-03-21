/* vim: set ts=2: */
var XLS;
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){XLS=require('./');});});

var opts = {};
if(process.env.WTF) opts.WTF = true;
var ex = [".xls", ".xml"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){return ex.indexOf(x.substr(-4))>=0||exp.indexOf(x.substr(-12))>=0;}

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);

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
				var csv = XLS.utils.make_csv(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate JSON', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				var json = XLS.utils.sheet_to_row_object_array(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate formulae', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				var json = XLS.utils.get_formulae(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate correct output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var name = (dir + x + '.' + i + '.csv');
			if(x.substr(-4) === ".xls") {
				root = x.slice(0,-4);
				if(!fs.existsSync(name)) name=(dir + root + '.xlsx.'+i+'.csv');
				if(!fs.existsSync(name)) name=(dir + root + '.xlsm.'+i+'.csv');
				if(!fs.existsSync(name)) name=(dir + root + '.xlsb.'+i+'.csv');
			}
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = XLS.utils.make_csv(wb.Sheets[ws]);
				assert.equal(normalizecsv(csv), normalizecsv(file), "CSV badness");
			} : null);
		});
	});
	if(!fs.existsSync(dir + x + '.xml')) return;
	describe(x + '.xml from 2011', function() {
		it('should parse', function() {
			var xlsb = XLS.readFile(dir + x + '.xml', opts);
		});
	});
}

describe('should parse test files', function() {
	files.forEach(function(x) {
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = XLS.readFile(dir + x, opts);
			parsetest(x, wb);
		});
	});
});

describe('options', function() {
	before(function() {
		XLS = require('./');
	});
	describe('cell', function() {
		it('should generate formulae by default', function() {
			var wb = XLS.readFile(dir + 'formula_stress_test.xls');
			var found = false;
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					if(typeof ws[addr].f !== 'undefined') return found = true;
				});
			});
			assert(found);
		});
		it('should not generate formulae when requested', function() {
			var wb =XLS.readFile(dir+'formula_stress_test.xls',{cellFormula:false});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].f === 'undefined');
				});
			});
		});
		it('should not generate number formats by default', function() {
			var wb = XLS.readFile(dir+'number_format.xls');
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].z === 'undefined');
				});
			});
		});
		it('should generate number formats when requested', function() {
			var wb = XLS.readFile(dir+'number_format.xls', {cellNF: true});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(ws[addr].t!== 'n' || typeof ws[addr].z !== 'undefined');
				});
			});
		});
	});
	describe('sheet', function() {
		it('should read all cells by default', function() {
			var wb = XLS.readFile(dir+'formula_stress_test.xls');
			assert(typeof wb.Sheets.Text.A46 !== 'undefined');
			assert(typeof wb.Sheets.Text.B26 !== 'undefined');
			assert(typeof wb.Sheets.Text.C16 !== 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
			wb = XLS.readFile(dir+'formula_stress_test.xls.xml');
			assert(typeof wb.Sheets.Text.A46 !== 'undefined');
			assert(typeof wb.Sheets.Text.B26 !== 'undefined');
			assert(typeof wb.Sheets.Text.C16 !== 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
		});
		it('sheetRows n=20', function() {
			var wb = XLS.readFile(dir+'formula_stress_test.xls', {sheetRows:20});
			assert(typeof wb.Sheets.Text.A46 === 'undefined');
			assert(typeof wb.Sheets.Text.B26 === 'undefined');
			assert(typeof wb.Sheets.Text.C16 !== 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
			wb = XLS.readFile(dir+'formula_stress_test.xls.xml', {sheetRows:20});
			assert(typeof wb.Sheets.Text.A46 === 'undefined');
			assert(typeof wb.Sheets.Text.B26 === 'undefined');
			assert(typeof wb.Sheets.Text.C16 !== 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
		});
		it('sheetRows n=10', function() {
			var wb = XLS.readFile(dir+'formula_stress_test.xls', {sheetRows:10});
			assert(typeof wb.Sheets.Text.A46 === 'undefined');
			assert(typeof wb.Sheets.Text.B26 === 'undefined');
			assert(typeof wb.Sheets.Text.C16 === 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
			wb = XLS.readFile(dir+'formula_stress_test.xls.xml', {sheetRows:10});
			assert(typeof wb.Sheets.Text.A46 === 'undefined');
			assert(typeof wb.Sheets.Text.B26 === 'undefined');
			assert(typeof wb.Sheets.Text.C16 === 'undefined');
			assert(typeof wb.Sheets.Text.D2 !== 'undefined');
		});
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			var wb = XLS.readFile(dir+'merge_cells.xls', {bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps should not generate sheets', function() {
			var wb = XLS.readFile(dir+'number_format.xls', {bookProps:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			var wb = XLS.readFile(dir+'LONumbers.xls', {bookProps:true, bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookFiles should generate cfb', function() {
			var wb = XLS.readFile(dir+'formula_stress_test.xls', {bookFiles:true});
			assert(typeof wb.cfb !== 'undefined');
		});
	});
});

describe('input formats', function() {
	it('should read binary strings', function() {
		XLS.read(fs.readFileSync(dir+'formula_stress_test.xls', 'binary'), {type:'binary'});
		XLS.read(fs.readFileSync(dir+'formula_stress_test.xls.xml', 'binary'), {type:'binary'});
	});
	it('should read base64 strings', function() {
		XLS.read(fs.readFileSync(dir+'comments_stress_test.xls', 'base64'), {type: 'base64'});
		XLS.read(fs.readFileSync(dir+'comments_stress_test.xls.xml', 'base64'), {type: 'base64'});
	});
});

describe('features', function() {
	describe('should parse core properties and custom properties', function() {
		var wb;
		before(function() {
			XLS = require('./');
			wb = XLS.readFile(dir+'custom_properties.xls');
		});
		it('Must have read the core properties', function() {
			assert.equal(wb.Props.Company, 'Vector Inc');
			assert.equal(wb.Props.Author, 'Pony Foo'); /* XLSX uses Creator */
		});
		it('Must have read the custom properties', function() {
			assert.equal(wb.Custprops['I am a boolean'], true);
			/* The date test requires parsing FILETIME (64 bit integer) */
			//assert.equal(wb.Custprops['Date completed'], '1967-03-09T16:30:00Z');
			assert.equal(wb.Custprops.Status, 2);
			assert.equal(wb.Custprops.Counter, -3.14);
		});
	});
});

describe('invalid files', function() {
	it('should fail on passwords', function() {
		assert.throws(function() { XLS.readFile(dir + 'apachepoi_password.xls'); });
		assert.throws(function() { XLS.readFile(dir + 'apachepoi_xor-encryption-abc.xls'); });
	});
	it('should fail on XLSX files', function() {
		assert.throws(function() { XLS.readFile(dir + 'roo_type_excelx.xls'); });
	});
	it('should fail on ODS files', function() {
		assert.throws(function() { XLS.readFile(dir + 'roo_type_openoffice.xls');});
	});
	it('should fail on DOC files', function() {
		assert.throws(function() { XLS.readFile(dir + 'word_doc.doc');});
	});
});
