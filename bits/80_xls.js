/* [MS-OLEDS] 2.3.8 CompObjStream */
function parse_compobj(obj) {
	var v = {};
	var o = obj.content;

	/* [MS-OLEDS] 2.3.7 CompObjHeader -- All fields MUST be ignored */
	var l = 28, m;
	m = o.lpstr(l); l += 5 + m.length; v.UserType = m;

	/* [MS-OLEDS] 2.3.1 ClipboardFormatOrAnsiString */
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


/* 2.4.58 Continue logic */
function slurp(R, blob, length, opts) {
	var read = blob.read_shift.bind(blob);
	var l = length;
	var bufs = [blob.slice(blob.l,blob.l+l)];
	blob.l += length;
	var next = (RecordEnum[blob.readUInt16LE(blob.l)]);
	while(next && next.n === 'Continue') {
		l = blob.readUInt16LE(blob.l+2);
		bufs.push(blob.slice(blob.l+4,blob.l+4+l));
		blob.l += 4+l;
		next = (RecordEnum[blob.readUInt16LE(blob.l)]);
	}
	var b = (typeof Buffer !== 'undefined') ? Buffer.concat(bufs) : [].concat.apply([], bufs);
	prep_blob(b);
	return R.f(b, b.length, opts);
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
	var lastcell, last_cell;
	var shared_formulae = {};
	function addline(cell, line) {
		lastcell = cell;
		last_cell = encode_cell(cell);
		out[last_cell] = line;
	}
	var opts = {
		enc: false, // encrypted
		sbcch: 0, // cch in the preceding SupBook
		snames: [], // sheetnames
		sharedf: shared_formulae, // shared formulae by address
		rrtabid: [], // RRTabId
		lastuser: "", // Last User from WriteAccess
		codepage: 0, // CP from CodePage record
		winlocked: 0, // fLockWn from WinProtect
		wtf: false
	};
	var supbooks = [[]]; // 1-indexed, will hold extern names
	var sbc = 0, sbci = 0, sbcli = 0;
	supbooks.SheetNames = opts.snames;
	supbooks.sharedf = opts.sharedf;
	while(blob.l < blob.length) {
		var s = blob.l;
		var RecordType = read(2);
		var length = read(2), y;
		var R = RecordEnum[RecordType];
		if(R && R.f) {
			if(R.r === 2 || R.r == 12) {
				var rt = read(2); length -= 2;
				if(!opts.enc && rt !== RecordType) throw "rt mismatch";
				if(R.r == 12){ blob.l += 10; length -= 10; } // skip FRT
			}
			//console.error(R,blob.l,length,blob.length);
			var val;
			if(blob.l+length+2 >= blob.length) val = R.f(blob, length);
			else {
				var next = (RecordEnum[blob.readUInt16LE(blob.l+length)]);
				if(next && next.n === 'Continue') {
					val = slurp(R, blob, length, opts);
				} else {
					if(opts.enc) { parsenoop(blob, length); continue; }
					val = R.f(blob, length, opts);
				}
			}
			switch(R.n) {
				/* Workbook Options */
				case 'Date1904': wb.opts.Date1904 = val; break;
				case 'WriteProtect': wb.opts.WriteProtect = true; break;
				case 'FilePass': opts.enc = val; console.error("File is password-protected -- Cannot extract files (yet)"); break;
				case 'WriteAccess': opts.lastuser = val; break;
				case 'CodePage': opts.codepage = val; break;
				case 'RRTabId': opts.rrtabid = val; break;
				case 'WinProtect': opts.winlocked = val; break;

				case 'SupBook': supbooks[++sbc] = [val]; sbci = 0; break;
				case 'ExternName': supbooks[sbc][++sbci] = val; break;
				case 'Lbl': supbooks[0][++sbcli] = val; break;
				case 'ExternSheet': supbooks[sbc] = supbooks[sbc].concat(val); sbci += val.length; break;

				case 'Protect': out["!protect"] = val; break; /* for sheet or book */
				case 'Password': if(val !== 0) throw "Password protection unsupported";
				case 'Prot4Rev': case 'Prot4RevPass': break; /*TODO: Revision Control*/

				case 'BoundSheet8': {
					Directory[val.pos] = val;
					opts.snames.push(val.name);
				} break;
				case 'EOF': {
					var nout = {};
					if(range.e) {
						out["!range"] = range;
						if(range.e.r > 0 && range.e.c > 0) {
							range.e.r--; range.e.c--;
							out["!ref"] = encode_range(range);
							range.e.r++; range.e.c++;
						}
					}
					for(y in out) if(out.hasOwnProperty(y)) nout[y] = out[y];
					if(cur_sheet === "") Preamble = nout; else Sheets[cur_sheet] = nout;
				} break;
				case 'BOF': {
					out = {};
					cur_sheet = (Directory[s] || {name:""}).name;
					lst.push([R.n, s, val, Directory[s]]);
				} break;
				case 'Number': {
					addline({c:val.c + range.s.c, r:val.r + range.s.r}, {v:val.val, t:'n'});
				} break;
				case 'BoolErr': {
					addline({c:val.c + range.s.c, r:val.r + range.s.r}, {v:val.val, t:val.t});
				} break;
				case 'RK': {
					addline({c:val.c/* + range.s.c*/, r:val.r/* + range.s.r*/}, {ixfe: val.ixfe, v:val.rknum, t:'n'});
				} break;
				case 'MulRk': {
					for(var j = val.c; j <= val.C; ++j) {
						addline({c:j/*+range.s.c*/, r:val.r/* + range.s.r*/}, {ixfe: val.rkrec[j-val.c][0], v:val.rkrec[j-val.c][1], t:'n'});
					}
				} break;
				case 'Formula': {
					switch(val.val) {
						case 'String': last_formula = val; break;
						case 'Array Formula': throw "Array Formula unsupported";
						default: addline(val.cell, {v:val.val, f:stringify_formula(val.formula, range, val.cell, supbooks), ixfe: val.cell.ixfe, t:'n'}); // TODO: infer type from formula
					}
				} break;
				case 'String': {
					if(last_formula) {
						last_formula.val = val;
						addline(last_formula.cell, {v:JSON.stringify(last_formula.val), f:stringify_formula(last_formula.formula, range, last_formula.cell, supbooks), ixfe: last_formula.cell.ixfe, t:'s'});
						last_formula = null;
					}
				} break;
				case 'Array': {
					/* console.error(val); */
				} break;
				case 'ShrFmla': {
					out[last_cell].f = stringify_formula(val[0], range, lastcell, supbooks);
					shared_formulae[last_cell] = val[0];
				} break;
				case 'LabelSst': {
					addline({c:val.c, r:val.r}, {v:JSON.stringify(sst[val.isst]), ixfe:val.ixfe, t:'s'});
				} break;
				case 'Dimensions': {
					range = val;
				} break;
				case 'SST': {
					sst = val;
				} break;
				case 'Format': { /* val = [id, fmt] */
					SSF.load(val[1], val[0]);
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
	var sheetnamesraw = Object.keys(Directory).sort(function(a,b) { return Number(a) - Number(b); }).map(function(x){return Directory[x].name;});
	var sheetnames = []; sheetnamesraw.forEach(function(x){sheetnames.push(x);});
	//console.log(lst);
	//lst.filter(function(x) { return x[0] === 'Formula';}).forEach(function(x){console.log(x[2].cell,x[2].formula);});
	wb.Directory=sheetnamesraw;
	wb.SheetNames=sheetnamesraw;
	wb.Sheets=Sheets;
	wb.Preamble=Preamble;
	wb.Strings = sst;
	if(opts.enc) wb.Encryption = opts.enc;
	return wb;
}
if(Workbook) WorkbookP = parse_workbook(Workbook.content);

if(CompObj) CompObjP = parse_compobj(CompObj);

return WorkbookP;
}

function sheet_to_row_object_array(sheet){
	var val, rowObject, range, columnHeaders, emptyRow, C;
	var outSheet = [];
	if (sheet["!ref"]) {
		range = decode_range(sheet["!ref"]);

		columnHeaders = {};
		for (C = range.s.c; C <= range.e.c; ++C) {
			val = sheet[encode_cell({
				c: C,
				r: range.s.r
			})];
			if(val){
				switch(val.t) {
					case 's': case 'str': columnHeaders[C] = JSON.parse(val.v); break;
					case 'n': columnHeaders[C] = val.v; break;
				}
			}
		}

		for (var R = range.s.r + 1; R <= range.e.r; ++R) {
			emptyRow = true;
			//Row number is recorded in the prototype
			//so that it doesn't appear when stringified.
			rowObject = Object.create({ __rowNum__ : R });
			for (C = range.s.c; C <= range.e.c; ++C) {
				val = sheet[encode_cell({
					c: C,
					r: R
				})];
				if(val !== undefined) switch(val.t){
					case 's': case 'str':
						if(val.v !== undefined) val.v = JSON.parse(val.v);
					/* falls through */
					case 'b': case 'n':
						if(val.v !== undefined) {
							rowObject[columnHeaders[C]] = val.v;
							emptyRow = false;
						}
						break;
					case 'e': break; /* throw */
					default: throw 'unrecognized type ' + val.t;
				}
			}
			if(!emptyRow) {
				outSheet.push(rowObject);
			}
		}
	}
	return outSheet;
}

function sheet_to_csv(sheet) {
	var out = "";
	if(sheet["!ref"]) {
		var r = utils.decode_range(sheet["!ref"]);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var row = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var val = sheet[utils.encode_cell({c:C,r:R})];
				if(!val) { row.push(""); continue; }
				if(typeof val.v === 'boolean') val.v = val.v ? "TRUE" : "FALSE";
				row.push(String(val.v).replace(/\\n/g,"\n").replace(/\\t/g,"\t").replace(/\\\\/g,"\\").replace(/\\\"/g,"\"\""));
			}
			out += row.join(",") + "\n";
		}
	}
	return out;
}
var make_csv = sheet_to_csv;

function get_formulae(ws) {
	var cmds = [];
	for(var y in ws) if(y[0] !=='!' && ws.hasOwnProperty(y)) {
		var x = ws[y];
		var val = "";
		if(x.f) val = x.f;
		else if(typeof x.v === 'number') val = x.v;
		else val = x.v;
		cmds.push(y + "=" + val);
	}
	return cmds;
}

var utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	sheet_to_csv: sheet_to_csv,
	make_csv: sheet_to_csv,
	get_formulae: get_formulae,
	sheet_to_row_object_array: sheet_to_row_object_array
};

function xlsread(f, options) {
	return parse_xlscfb(CFB.read(f, options));
}
var readFile = function(f) { return parse_xlscfb(CFB.read(f, {type:'file'})); };
