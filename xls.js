if(typeof require !== 'undefined') {
	var CFB= require('./cfb');
	var vm = require('vm'), fs = require('fs');
	vm.runInThisContext(fs.readFileSync(__dirname+'/xlsconsts.js'));
}

/* MS-OLEDS 2.3.8 CompObjStream TODO */
function parse_compobj(obj) {
	var v = {};
	var o = obj.content;
	var l = 28, m; // skip the 28 bytes
	m = o.lpstr(l); l += 5 + m.length; v.UserType = m;

	/* MS-OLEDS 2.3.1 ClipboardFormatOrAnsiString */
	m = o.readUInt32LE(l); l+= 4;
	switch(m) {
		case 0x00000000: break;
		case 0xffffffff: case 0xfffffffe: l+=4; break;
		default:
			if(m > 0x190) throw "Unsupported Clipboard: " + m;
			l += m;
	}

	m = o.lpstr(l); l += 5 + m.length; v.Reserved1 = m;

	if((m = o.readUInt32LE(l)) !== 0x71b2e9f4) return v;
	throw "Unsupported Unicode Extension";
}


function parse_xlscfb(cfb) {
var CompObj = cfb.Directory['!CompObj']; // OLE
var Summary = cfb.Directory['!SummaryInformation'];
var Workbook = cfb.Directory.Workbook; // OLE
var CompObjP, SummaryP, WorkbookP;

/* 2.2.2 + Magic TODO */
function parse_formula(formula, range) {
	range = range || {s:{c:0, r:0}};
	var stack = [], e1, e2, type, c, sht;
	if(!formula[0] || !formula[0][0]) return "";
	formula[0].forEach(function(f) {
		//console.log("++",f, formula[0])
		switch(f[0]) {
			/* Control Tokens -- ignore */
			case 'PtgAttrIf': case 'PtgAttrChoose': case 'PtgAttrGoto': break;

			case 'PtgRef':
				type = f[1][0], c = shift_cell(f[1][1], range);
				stack.push(encode_cell(c));
				break;
			case 'PtgRef3d': // TODO: lots of stuff
				type = f[1][0], sht = f[1][1], c = shift_cell(f[1][2], range);
				stack.push("!"+encode_cell(c));
				break;

			/* Function Call */
			case 'PtgFuncVar':
				var argc = f[1][0], func = f[1][1];
				var args = stack.slice(-argc);
				stack.length -= argc;
				stack.push(func + "(" + args.join(",") + ")");
				break;

			/* Binary Operators -- pop 2 push 1*/
			case 'PtgAdd':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"+"+e1);
				break;
			case 'PtgSub':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"-"+e1);
				break;
			case 'PtgMul':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"*"+e1);
				break;
			case 'PtgDiv':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"/"+e1);
				break;
			case 'PtgPower':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"^"+e1);
				break;
			case 'PtgConcat':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"&"+e1);
				break;
			case 'PtgLt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<"+e1);
				break;
			case 'PtgLe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<="+e1);
				break;
			case 'PtgEq':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"="+e1);
				break;
			case 'PtgGe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">="+e1);
				break;
			case 'PtgGt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">"+e1);
				break;
			case 'PtgNe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<>"+e1);
				break;

			case 'PtgInt':
				stack.push(f[1]); break;
			case 'PtgArea':
				type = f[1][0], r = shift_range(f[1][1], range);
				stack.push(encode_range(r));
				break;
			case 'PtgAttrSum':
				stack.push("SUM(" + stack.pop() + ")");
				break;
		}
	});
	//console.log("--",stack);
	return stack[0];
}

// 2.3.2
function parse_workbook(blob) {
	var wb = {opts:{}};
	var Sheets = {};
	var out = [];
	var read = blob.read_shift.bind(blob);
	var lst = [], seen = {};
	var Directory = {};
	var found_sheet = false;
	var range = {};
	var last_formula = null;
	var sst = [];
	var cur_sheet = "";
	var Preamble = {};
	function addline(line) { out.push(line); }
	while(blob.l < blob.length) {
		var s = blob.l;
		var RecordType = read(2);
		var length = read(2), y;
		var R = RecordEnum[RecordType];
		if(R && R.f) {
			if(R.r === 2 || R.r == 12) {
				var rt = read(2); length -= 2;
				if(rt !== RecordType) throw "rt mismatch";
				if(R.r == 12){ blob.l += 10; length -= 10; } // skip FRT
			}
			//console.log(R,blob.l,length,blob.length);
			var val = R.f(blob, length);
			switch(R.n) {
				/* Workbook Options */
				case 'Date1904': wb.opts.Date1904 = val; break;

				case 'BoundSheet8': {
					Directory[val.pos] = val;
				} break;
				case 'EOF': {
					var nout = [];
					for(y in out) if(out.hasOwnProperty(y)) nout[y] = out[y];
					if(cur_sheet === "") Preamble = nout; else Sheets[cur_sheet] = nout;
				} break;
				case 'BOF': {
					out = [];
					cur_sheet = (Directory[s] || {name:""}).name;
					lst.push([R.n, s, val, Directory[s]]);
				} break;
				case 'Number': {
					addline({cell:{c:val.c + range.s.c, r:val.r + range.s.r}, val:val.val});
				} break;
				case 'RK': {
					addline({cell:{c:val.c/* + range.s.c*/, r:val.r/* + range.s.r*/}, ixfe: val.ixfe, val:val.rknum});
				} break;
				case 'MulRk': {
					for(var j = val.c; j <= val.C; ++j) {
						addline({cell:{c:j/*+range.s.c*/, r:val.r/* + range.s.r*/}, ixfe: val.rkrec[j-val.c][0], val:val.rkrec[j-val.c][1]});
					}
				} break;
				case 'Formula': {
					if(val.val === "String") {
						last_formula = val;
					}
					else addline({cell:val.cell, ixfe: val.cell.ixfe, val:val.val, formula:parse_formula(val.formula, range)});
				} break;
				case 'String': {
					if(last_formula) {
						last_formula.val = val;
						addline({cell:last_formula.cell, ixfe: last_formula.cell.ixfe, val:JSON.stringify(last_formula.val), formula:parse_formula(last_formula.formula, range)});
						last_formula = null;
					}
				} break;
				case 'LabelSst': {
					addline({cell:{c:val.c, r:val.r}, ixfe:val.ixfe, val:JSON.stringify(sst[val.isst])});
				} break;
				case 'Dimensions': {
					range = val;
					out.range = range;
				} break;
				case 'SST': {
					sst = val;
				} break;
				case 'Scl': {
					//console.log("Zoom Level:", val[0]/val[1],val);
				} break;
			}
			lst.push([R.n, s, val]);
			continue;
		}
		lst.push(['Unrecognized', Number(RecordType).toString(16), RecordType]);
		read(length);
	}
	var sheetnamesraw = Object.keys(Directory).map(Number).sort().map(function(x){return Directory[x].name;});
	var sheetnames = []; sheetnamesraw.forEach(function(x){sheetnames.push(x);});
	//console.log(lst);
	//lst.filter(function(x) { return x[0] === 'Formula';}).forEach(function(x){console.log(x[2].cell,x[2].formula);});
	wb.Directory=sheetnamesraw;
	wb.SheetNames=sheetnamesraw;
	wb.Sheets=Sheets;
	wb.Preamble=Preamble;
	return wb;
}
if(Workbook) WorkbookP = parse_workbook(Workbook.content);

if(CompObj) CompObjP = parse_compobj(CompObj);

return WorkbookP;
}

function make_csv(ws) {
	var blob = [];
	for(var i = 0; i != ws.range.e.r; ++i) blob[i] = new Array(ws.range.e.c);
	ws.forEach(function(x) {
		var o = {v:x.val};
		if(x.formula) o.f = x.formula;
		blob[x.cell.r][x.cell.c] = o;
	});
	return blob.map(function(r) {return r.map(function(x){return x.v;}).join(",");}).join("\n");
}

function get_formulae(ws) {
	var cmds = [];
	ws.forEach(function(x) {
		var val = "";
		if(x.formula) val = x.formula;
		else if(typeof x.val === 'number') val = x.val;
		else val = x.val;
		cmds.push(encode_cell(x.cell) + "=" + val);
	});
	return cmds;
}

var utils = {
	make_csv: make_csv,
	get_formulae: get_formulae
};

var readFile = function(f) { return parse_xlscfb(CFB.read(f, {type:'file'})); }
if(typeof exports !== 'undefined') {
	exports.readFile = readFile;
	exports.utils = utils;
	if(typeof module !== 'undefined' && require.main === module ) {
		var wb = readFile(process.argv[2] || 'Book1.xls');
		var target_sheet = process.argv[3] || '';
		if(target_sheet === '') target_sheet = wb.Directory[0];
		var ws = wb.Sheets[target_sheet];
		console.log(target_sheet);
		console.log(make_csv(ws));
		//console.log(get_formulae(ws));
	}
}
