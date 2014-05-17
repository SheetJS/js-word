/* vim: set ts=2: */
var X;
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require('./');});});

var opts = {};
if(process.env.WTF) opts.WTF = true;
var ex = [".xls", ".xml"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){return ex.indexOf(x.substr(-4))>=0||exp.indexOf(x.substr(-12))>=0;}

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);
var fileA = (fs.existsSync('testA.lst') ? fs.readFileSync('testA.lst', 'utf-8').split("\n") : []).filter(test_file);

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x) { return x.substr(0,31); }

function fixcsv(x) { return x.replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }
function fixjson(x) { return x.replace(/[\r\n]+$/,""); }

var dir = "./test_files/";

var paths = {
	cp1:  dir + 'custom_properties.xls',
	cp2:  dir + 'custom_properties.xls.xml',
	cst1: dir + 'comments_stress_test.xls',
	cst2: dir + 'comments_stress_test.xls.xml',
	fst1: dir + 'formula_stress_test.xls',
	fst2: dir + 'formula_stress_test.xls.xml',
	fstb: dir + 'formula_stress_test.xls',
	hl1:  dir + 'hyperlink_stress_test_2011.xls',
	hl2:  dir + 'hyperlink_stress_test_2011.xml',
	lon1: dir + 'LONumbers.xls',
	mc1:  dir + 'merge_cells.xls',
	mc2:  dir + 'merge_cells.xls.xml',
	nf1:  dir + 'number_format.xls',
	nf2:  dir + 'number_format.xls.xml',
	swc1: dir + 'apachepoi_SimpleWithComments.xls',
	swc2: dir + '2011/apachepoi_SimpleWithComments.xls.xml'
};

var N1 = 'XLS';
var N2 = 'XML';

function parsetest(x, wb, full, ext) {
	ext = (ext ? " [" + ext + "]": "");
	describe(x + ext + ' should have all bits', function() {
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
	describe(x + ext + ' should generate CSV', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.make_csv(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ext + ' should generate JSON', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.sheet_to_row_object_array(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ext + ' should generate formulae', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.get_formulae(wb.Sheets[ws]);
			});
		});
	});
	if(!full) return;
	var getfile = function(dir, x, i, type) {
			var name = (dir + x + '.' + i + type);
			if(x.substr(-4) === ".xls") {
				root = x.slice(0,-4);
				if(!fs.existsSync(name)) name=(dir + root + '.xlsx.' + i + type);
				if(!fs.existsSync(name)) name=(dir + root + '.xlsm.' + i + type);
				if(!fs.existsSync(name)) name=(dir + root + '.xlsb.' + i + type);
			}
			return name;
	};
	describe(x + ext + ' should generate correct CSV output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var name = getfile(dir, x, i, ".csv");
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = X.utils.make_csv(wb.Sheets[ws]);
				assert.equal(fixcsv(csv), fixcsv(file), "CSV badness");
			} : null);
		});
	});
	describe(x + ext + ' should generate correct JSON output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var rawjson = getfile(dir, x, i, ".rawjson");
			if(fs.existsSync(rawjson)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(rawjson, 'utf-8');
				var json = X.utils.make_json(wb.Sheets[ws],{raw:true});
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});

			var jsonf = getfile(dir, x, i, ".json");
			if(fs.existsSync(jsonf)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(jsonf, 'utf-8');
				var json = X.utils.make_json(wb.Sheets[ws]);
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});
		});
	});
	if(fs.existsSync(dir + '2011/' + x + '.xml'))
	describe(x + ext + '.xml from 2011', function() {
		it('should parse', function() {
			var wb = X.readFile(dir + '2011/' + x + '.xml', opts);
		});
	});
	if(fs.existsSync(dir + x + '.xml' + ext))
	describe(x + '.xml', function() {
		it('should parse', function() {
			var wb = X.readFile(dir + x + '.xml', opts);
		});
	});
}

describe('should parse test files', function() {
	files.forEach(function(x) {
		if(!fs.existsSync(dir + x)) return;
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = X.readFile(dir + x, opts);
			parsetest(x, wb, true);
		});
	});
	fileA.forEach(function(x) {
		if(!fs.existsSync(dir + x)) return;
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = X.readFile(dir + x, {WTF:opts.wtf, sheetRows:10});
			parsetest(x, wb, false);
		});
	});
});

describe('parse options', function() {
	var html_cell_types = ['s'];
	before(function() {
		X = require('./');
	});
	describe('cell', function() {
		it.skip('should generate HTML by default', function() {
			var wb = X.readFile(paths.cst1);
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(html_cell_types.indexOf(ws[addr].t) === -1 || ws[addr].h);
			});
		});
		it.skip('should not generate HTML when requested', function() {
			var wb = X.readFile(paths.cst1, {cellHTML:false});
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(typeof ws[addr].h === 'undefined');
			});
		});
		it('should generate formulae by default', function() {
			var wb = X.readFile(paths.fstb);
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
			var wb =X.readFile(paths.fstb,{cellFormula:false});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].f === 'undefined');
				});
			});
		});
		it('should not generate number formats by default', function() {
			var wb = X.readFile(paths.nf1);
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].z === 'undefined');
				});
			});
		});
		it('should generate number formats when requested', function() {
			var wb = X.readFile(paths.nf1, {cellNF: true});
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
		it('should not generate sheet stubs by default', function() {
			var wb = X.readFile(paths.mc1);
			assert.throws(function() { wb.Sheets.Merge.A2.v; });
			wb = X.readFile(paths.mc2);
			assert.throws(function() { wb.Sheets.Merge.A2.v; });
		});
		it.skip('should generate sheet stubs when requested', function() {
			var wb = X.readFile(paths.mc1, {sheetStubs:true});
			assert(typeof wb.Sheets.Merge.A2.t !== 'undefined');
			wb = X.readFile(paths.mc2, {sheetStubs:true});
			assert(typeof wb.Sheets.Merge.A2.t !== 'undefined');
		});
		function checkcells(wb, A46, B26, C16, D2) {
			assert((typeof wb.Sheets.Text.A46 !== 'undefined') == A46);
			assert((typeof wb.Sheets.Text.B26 !== 'undefined') == B26);
			assert((typeof wb.Sheets.Text.C16 !== 'undefined') == C16);
			assert((typeof wb.Sheets.Text.D2  !== 'undefined') == D2);
		}
		it('should read all cells by default', function() {
			var wb = X.readFile(paths.fst1);
			checkcells(wb, true, true, true, true);
			wb = X.readFile(paths.fst2);
			checkcells(wb, true, true, true, true);
		});
		it('sheetRows n=20', function() {
			var wb = X.readFile(paths.fst1, {sheetRows:20});
			checkcells(wb, false, false, true, true);
			wb = X.readFile(paths.fst2, {sheetRows:20});
			checkcells(wb, false, false, true, true);
		});
		it('sheetRows n=10', function() {
			var wb = X.readFile(paths.fst1, {sheetRows:10});
			checkcells(wb, false, false, false, true);
			wb = X.readFile(paths.fst2, {sheetRows:10});
			checkcells(wb, false, false, false, true);
		});
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			var wb = X.readFile(paths.mc1, {bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
			var wb = X.readFile(paths.mc2, {bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps should not generate sheets', function() {
			var wb = X.readFile(paths.nf1, {bookProps:true});
			assert(typeof wb.Sheets === 'undefined');
			wb = X.readFile(paths.nf2, {bookProps:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			var wb = X.readFile(paths.lon1, {bookProps:true, bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('should not generate deps by default', function() {
			var wb = X.readFile(paths.fst1);
			assert(typeof wb.Deps === 'undefined' || !(wb.Deps.length>0));
			wb = X.readFile(paths.fst2);
			assert(typeof wb.Deps === 'undefined' || !(wb.Deps.length>0));
		});
		it.skip('bookDeps should generate deps', function() {
			var wb = X.readFile(paths.fst1, {bookDeps:true});
			assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
			wb = X.readFile(paths.fst2, {bookDeps:true});
			assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
		});
		var ckf = function(wb, fields, exists) { fields.forEach(function(f) {
				assert((typeof wb[f] !== 'undefined') == exists);
		}); };
		it('should not generate book files by default', function() {
			var wb = X.readFile(paths.fst1);
			ckf(wb, ['cfb'], false);
		});
		it('bookFiles should generate book files', function() {
			var wb = X.readFile(paths.fst1, {bookFiles:true});
			ckf(wb, ['cfb'], true);
		});
	});
});

describe('input formats', function() {
	it('should read binary strings', function() {
		X.read(fs.readFileSync(paths.cst1, 'binary'), {type: 'binary'});
		X.read(fs.readFileSync(paths.cst2, 'binary'), {type: 'binary'});
	});
	it('should read base64 strings', function() {
		X.read(fs.readFileSync(paths.cst1, 'base64'), {type: 'base64'});
		X.read(fs.readFileSync(paths.cst2, 'base64'), {type: 'base64'});
	});
	it('should read array', function() {
		X.read(fs.readFileSync(dir+'merge_cells.xls.xml', 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
	});
});

function coreprop(wb) {
	assert.equal(wb.Props.Title, 'Example with properties');
	assert.equal(wb.Props.Subject, 'Test it before you code it');
	assert.equal(wb.Props.Author, 'Pony Foo');
	assert.equal(wb.Props.Manager, 'Despicable Drew');
	assert.equal(wb.Props.Company, 'Vector Inc');
	assert.equal(wb.Props.Category, 'Quirky');
	assert.equal(wb.Props.Keywords, 'example humor');
	assert.equal(wb.Props.Comments, 'some comments');
	assert.equal(wb.Props.LastAuthor, 'Hugues');
}
function custprop(wb) {
	assert.equal(wb.Custprops['I am a boolean'], true);
	assert.equal(wb.Custprops['Date completed'].toISOString(), '1967-03-09T16:30:00.000Z');
	assert.equal(wb.Custprops.Status, 2);
	assert.equal(wb.Custprops.Counter, -3.14);
}

describe('parse features', function() {
	it('should have comment as part of cell properties', function(){
		var X = require('./');
		var sheet = 'Sheet1';
		var wb1=X.readFile(paths.swc1);
		var wb2=X.readFile(paths.swc2);

		[wb1,wb2].map(function(wb) { return wb.Sheets[sheet]; }).forEach(function(ws, i) {
			assert.equal(ws.B1.c.length, 1,"must have 1 comment");
			assert.equal(ws.B1.c[0].a, "Yegor Kozlov","must have the same author");
			assert.equal(ws.B1.c[0].t.replace(/\r\n/g,"\n").replace(/\r/g,"\n"), "Yegor Kozlov:\nfirst cell", "must have the concatenated texts");
			return;
			assert.equal(ws.B1.c[0].r, '<r><rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Yegor Kozlov:</t></r><r><rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\r\nfirst cell</t></r>', "must have the rich text representation");
			assert.equal(ws.B1.c[0].h, '<span style="font-weight: bold;">Yegor Kozlov:</span><span style=""><br/>first cell</span>', "must have the html representation");
		});
	});

	describe('should parse core properties and custom properties', function() {
		var wb1, wb2;
		before(function() {
			X = require('./');
			wb1 = X.readFile(paths.cp1);
			wb2 = X.readFile(paths.cp2);
		});

		it(N1 + ' should parse core properties', function() { coreprop(wb1); });
		it(N2 + ' should parse core properties', function() { coreprop(wb2); });
		it(N1 + ' should parse custom properties', function() { custprop(wb1); });
		it(N2 + ' should parse custom properties', function() { custprop(wb2); });
	});

	describe('sheetRows', function() {
		it('should use original range if not set', function() {
			var opts = {};
			var wb1 = X.readFile(paths.fst1, opts);
			var wb2 = X.readFile(paths.fst2, opts);
			[wb1, wb2].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!ref"],"A1:F49");
			});
		});
		it.skip('should adjust range if set', function() {
			var opts = {sheetRows:10};
			var wb1 = X.readFile(paths.fst1, opts);
			var wb2 = X.readFile(paths.fst2, opts);
			[wb1, wb2].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!fullref"],"A1:F49");
				assert.equal(wb.Sheets.Text["!ref"],"A1:F10");
			});
		});
		it.skip('should not generate comment cells', function() {
			var opts = {sheetRows:10};
			var wb1 = X.readFile(paths.cst1, opts);
			var wb2 = X.readFile(paths.cst2, opts);
			[wb1, wb2].forEach(function(wb) {
				assert.equal(wb.Sheets.Sheet7["!fullref"],"A1:N34");
				assert.equal(wb.Sheets.Sheet7["!ref"],"A1:A1");
			});
		});
	});

	describe('merge cells',function() {
		var wb1, wb2;
		before(function() {
			X = require('./');
			wb1 = X.readFile(paths.mc1);
			wb2 = X.readFile(paths.mc2);
		});
		it('should have !merges', function() {
			assert(wb1.Sheets.Merge['!merges']);
			assert(wb2.Sheets.Merge['!merges']);
			var m = [wb1, wb2].map(function(x) { return x.Sheets.Merge['!merges'].map(function(y) { return X.utils.encode_range(y); });});
			assert.deepEqual(m[0].sort(),m[1].sort());
		});
	});

	describe('should find hyperlinks', function() {
		var wb1, wb2;
		before(function() {
			X = require('./');
			wb1 = X.readFile(paths.hl1);
			wb2 = X.readFile(paths.hl2);
		});

		function hlink(wb) {
			var ws = wb.Sheets.Sheet1;
			assert.equal(ws.A1.l.Target, "http://www.sheetjs.com");
			assert.equal(ws.A2.l.Target, "http://oss.sheetjs.com");
			assert.equal(ws.A3.l.Target, "http://oss.sheetjs.com#foo");
			assert.equal(ws.A4.l.Target, "mailto:dev@sheetjs.com");
			assert.equal(ws.A5.l.Target, "mailto:dev@sheetjs.com?subject=hyperlink");
			assert.equal(ws.A6.l.Target, "../../sheetjs/Documents/Test.xlsx");
			assert.equal(ws.A7.l.Target, "http://sheetjs.com");
		}

		it(N1, function() { hlink(wb1); });
		it(N2, function() { hlink(wb2); });
	});

});

function password_file(x){return x.match(/^password.*\.xls$/); }
var password_files = fs.readdirSync('test_files').filter(password_file);
describe('invalid files', function() {
	describe('parse', function() { [
			['password', 'apachepoi_password.xls'],
			['passwords', 'apachepoi_xor-encryption-abc.xls'],
			['XLSX files', 'roo_type_excelx.xls'],
			['ODS files', 'roo_type_openoffice.xls'],
			['DOC files', 'word_doc.doc']
		].forEach(function(w) { it('should fail on ' + w[0], function() { assert.throws(function() { X.readFile(dir + w[1]); }); }); });
	});
});
describe('encryption', function() {
	password_files.forEach(function(x) {
		describe(x, function() {
			it('should throw with no password', function() {assert.throws(function() { X.readFile(dir + x); }); });
			it('should throw with wrong password', function() {assert.throws(function() { X.readFile(dir + x, {password:'passwor',WTF:opts.WTF}); }); });
			it('should recognize correct password', function() {
				try { X.readFile(dir + x, {password:'password',WTF:opts.WTF}); }
				catch(e) { if(e.message == "Password is incorrect") throw e; }
			});
			it.skip('should decrypt file', function() {
				var wb = X.readFile(dir + x, {password:'password',WTF:opts.WTF});
			});
		});
	});
});
