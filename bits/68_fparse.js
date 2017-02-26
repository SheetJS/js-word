/* 2.4.127 TODO */
function parse_Formula(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_Cell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	if(opts.biff != 2) {
		blob.read_shift(1);
		if(opts.biff >= 5) {
			var chn = blob.read_shift(4);
		}
	}
	var cbf = parse_CellParsedFormula(blob, end - blob.l, opts);
	return {cell:cell, val:val[0], formula:cbf, shared: (flags >> 3) & 1, tt:val[1]};
}

/* 2.5.133 TODO: how to emit empty strings? */
function parse_FormulaValue(blob) {
	var b;
	if(__readUInt16LE(blob,blob.l + 6) !== 0xFFFF) return [parse_Xnum(blob),'n'];
	switch(blob[blob.l]) {
		case 0x00: blob.l += 8; return ["String", 's'];
		case 0x01: b = blob[blob.l+2] === 0x1; blob.l += 8; return [b,'b'];
		case 0x02: b = blob[blob.l+2]; blob.l += 8; return [b,'e'];
		case 0x03: blob.l += 8; return ["",'s'];
	}
	return [];
}

/* 2.5.198.103 */
function parse_RgbExtra(blob, length, rgce, opts) {
	if(opts.biff < 8) return parsenoop(blob, length);
	var target = blob.l + length;
	var o = [];
	for(var i = 0; i !== rgce.length; ++i) {
		switch(rgce[i][0]) {
			case 'PtgArray': /* PtgArray -> PtgExtraArray */
				rgce[i][1] = parse_PtgExtraArray(blob, 0, opts);
				o.push(rgce[i][1]);
				break;
			case 'PtgMemArea': /* PtgMemArea -> PtgExtraMem */
				rgce[i][2] = parse_PtgExtraMem(blob, rgce[i][1]);
				o.push(rgce[i][2]);
				break;
			default: break;
		}
	}
	length = target - blob.l;
	if(length !== 0) o.push(parsenoop(blob, length));
	return o;
}

/* 2.5.198.21 */
function parse_NameParsedFormula(blob, length, opts, cce) {
	var target = blob.l + length;
	var rgce = parse_Rgce(blob, cce, opts);
	var rgcb;
	if(target !== blob.l) rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.3 TODO */
function parse_CellParsedFormula(blob, length, opts) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.118 TODO */
function parse_SharedParsedFormula(blob, length, opts) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce, opts);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.1 TODO */
function parse_ArrayParsedFormula(blob, length, opts, ref) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.104 */
function parse_Rgce(blob, length, opts) {
	var target = blob.l + length;
	var R, id, ptgs = [];
	while(target != blob.l) {
		length = target - blob.l;
		id = blob[blob.l];
		R = PtgTypes[id];
		if(id === 0x18 || id === 0x19) {
			id = blob[blob.l + 1];
			R = (id === 0x18 ? Ptg18 : Ptg19)[id];
		}
		if(!R || !R.f) { ptgs.push(parsenoop(blob, length)); }
		else { ptgs.push([R.n, R.f(blob, length, opts)]); }
	}
	return ptgs;
}

function stringify_array(f) {
	var o = [];
	for(var i = 0; i < f.length; ++i) {
		var x = f[i], r = [];
		for(var j = 0; j < x.length; ++j) {
			var y = x[j];
			if(y) switch(y[0]) {
				// TODO: handle embedded quotes
				case 0x02: r.push('"' + y[1].replace(/"/g,'""') + '"'); break;
				default: r.push(y[1]);
			} else r.push("");
		}
		o.push(r.join(","));
	}
	return o.join(";");
}

/* 2.2.2 TODO */
var PtgBinOp = {
	PtgAdd: "+",
	PtgConcat: "&",
	PtgDiv: "/",
	PtgEq: "=",
	PtgGe: ">=",
	PtgGt: ">",
	PtgLe: "<=",
	PtgLt: "<",
	PtgMul: "*",
	PtgNe: "<>",
	PtgPower: "^",
	PtgSub: "-"
};
function stringify_formula(formula, range, cell, supbooks, opts) {
	var _range = {s:{c:0, r:0},e:{c:0, r:0}};
	var stack = [], e1, e2, type, c, ixti, nameidx, r, sname="";
	if(!formula[0] || !formula[0][0]) return "";
	var last_sp = -1, sp = "";
	//console.log("--",cell,formula[0])
	for(var ff = 0, fflen = formula[0].length; ff < fflen; ++ff) {
		var f = formula[0][ff];
		//console.log("++",f, stack)
		switch(f[0]) {
		/* 2.2.2.1 Unary Operator Tokens */
			/* 2.5.198.93 */
			case 'PtgUminus': stack.push("-" + stack.pop()); break;
			/* 2.5.198.95 */
			case 'PtgUplus': stack.push("+" + stack.pop()); break;
			/* 2.5.198.81 */
			case 'PtgPercent': stack.push(stack.pop() + "%"); break;

		/* 2.2.2.1 Binary Value Operator Token */
			case 'PtgAdd':    /* 2.5.198.26 */
			case 'PtgConcat': /* 2.5.198.43 */
			case 'PtgDiv':    /* 2.5.198.45 */
			case 'PtgEq':     /* 2.5.198.56 */
			case 'PtgGe':     /* 2.5.198.64 */
			case 'PtgGt':     /* 2.5.198.65 */
			case 'PtgLe':     /* 2.5.198.68 */
			case 'PtgLt':     /* 2.5.198.69 */
			case 'PtgMul':    /* 2.5.198.75 */
			case 'PtgNe':     /* 2.5.198.78 */
			case 'PtgPower':  /* 2.5.198.82 */
			case 'PtgSub':    /* 2.5.198.90 */
				e1 = stack.pop(); e2 = stack.pop();
				if(last_sp >= 0) {
					switch(formula[0][last_sp][1][0]) {
						case 0: sp = fill(" ", formula[0][last_sp][1][1]); break;
						case 1: sp = fill("\r", formula[0][last_sp][1][1]); break;
						default:
							sp = "";
							if(opts.WTF) throw new Error("Unexpected PtgSpace type " + formula[0][last_sp][1][0]);
					}
					e2 = e2 + sp;
					last_sp = -1;
				}
				stack.push(e2+PtgBinOp[f[0]]+e1);
				break;

		/* 2.2.2.1 Binary Reference Operator Token */
			/* 2.5.198.67 */
			case 'PtgIsect':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+" "+e1);
				break;
			case 'PtgUnion':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+","+e1);
				break;
			case 'PtgRange':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+":"+e1);
				break;

		/* 2.2.2.3 Control Tokens "can be ignored" */
			/* 2.5.198.34 */
			case 'PtgAttrChoose': break;
			/* 2.5.198.35 */
			case 'PtgAttrGoto': break;
			/* 2.5.198.36 */
			case 'PtgAttrIf': break;


			/* 2.5.198.84 */
			case 'PtgRef':
				type = f[1][0]; c = shift_cell(f[1][1], _range, opts);
				stack.push(encode_cell_xls(c));
				break;
			/* 2.5.198.88 */
			case 'PtgRefN':
				type = f[1][0]; c = shift_cell(f[1][1], cell, opts);
				stack.push(encode_cell_xls(c));
				break;
			case 'PtgRef3d': // TODO: lots of stuff
				type = f[1][0]; ixti = f[1][1]; c = shift_cell(f[1][2], _range, opts);
				sname = (supbooks && supbooks[1] ? supbooks[1][ixti+1] : "**MISSING**");
				stack.push(sname + "!" + encode_cell(c));
				break;

		/* Function Call */
			/* 2.5.198.62 */
			case 'PtgFunc':
			/* 2.5.198.63 */
			case 'PtgFuncVar':
				/* f[1] = [argc, func, type] */
				var argc = f[1][0], func = f[1][1];
				if(!argc) argc = 0;
				var args = argc == 0 ? [] : stack.slice(-argc);
				stack.length -= argc;
				if(func === 'User') func = args.shift();
				stack.push(func + "(" + args.join(",") + ")");
				break;

			/* 2.5.198.42 */
			case 'PtgBool': stack.push(f[1] ? "TRUE" : "FALSE"); break;
			/* 2.5.198.66 */
			case 'PtgInt': stack.push(f[1]); break;
			/* 2.5.198.79 TODO: precision? */
			case 'PtgNum': stack.push(String(f[1])); break;
			/* 2.5.198.89 */
			case 'PtgStr': stack.push('"' + f[1] + '"'); break;
			/* 2.5.198.57 */
			case 'PtgErr': stack.push(f[1]); break;
			/* 2.5.198.27 TODO: fixed points */
			case 'PtgArea':
				type = f[1][0]; r = shift_range(f[1][1], _range);
				stack.push(encode_range_xls(r));
				break;
			/* 2.5.198.28 */
			case 'PtgArea3d': // TODO: lots of stuff
				type = f[1][0]; ixti = f[1][1]; r = f[1][2];
				sname = (supbooks && supbooks[1] ? supbooks[1][ixti+1] : "**MISSING**");
				stack.push(sname + "!" + encode_range(r));
				break;
			/* 2.5.198.41 */
			case 'PtgAttrSum':
				stack.push("SUM(" + stack.pop() + ")");
				break;

		/* Expression Prefixes */
			/* 2.5.198.37 */
			case 'PtgAttrSemi': break;

			/* 2.5.97.60 TODO: do something different for revisions */
			case 'PtgName':
				/* f[1] = type, 0, nameindex */
				nameidx = f[1][2];
				var lbl = supbooks[0][nameidx];
				var name = lbl ? lbl.Name : "**MISSING**" + nameidx;
				if(name in XLSXFutureFunctions) name = XLSXFutureFunctions[name];
				stack.push(name);
				break;

			/* 2.5.97.61 TODO: do something different for revisions */
			case 'PtgNameX':
				/* f[1] = type, ixti, nameindex */
				var bookidx = f[1][1]; nameidx = f[1][2]; var externbook;
				/* TODO: Properly handle missing values */
				if(opts.biff == 5) {
					if(bookidx < 0) bookidx = -bookidx;
					if(supbooks[bookidx]) externbook = supbooks[bookidx][nameidx];
				} else {
					if(supbooks[bookidx+1]) externbook = supbooks[bookidx+1][nameidx];
					else if(supbooks[bookidx-1]) externbook = supbooks[bookidx-1][nameidx];
				}
				if(!externbook) externbook = {body: "??NAMEX??"};
				stack.push(externbook.body);
				break;

		/* 2.2.2.4 Display Tokens */
			/* 2.5.198.80 */
			case 'PtgParen':
				var lp = '(', rp = ')';
				if(last_sp >= 0) {
					sp = "";
					switch(formula[0][last_sp][1][0]) {
						case 2: lp = fill(" ", formula[0][last_sp][1][1]) + lp; break;
						case 3: lp = fill("\r", formula[0][last_sp][1][1]) + lp; break;
						case 4: rp = fill(" ", formula[0][last_sp][1][1]) + lp; break;
						case 5: rp = fill("\r", formula[0][last_sp][1][1]) + lp; break;
						default:
							if(opts.WTF) throw new Error("Unexpected PtgSpace type " + formula[0][last_sp][1][0]);
					}
					last_sp = -1;
				}
				stack.push(lp + stack.pop() + rp); break;

			/* 2.5.198.86 */
			case 'PtgRefErr': stack.push('#REF!'); break;

		/* */
			/* 2.5.198.58 TODO */
			case 'PtgExp':
				c = {c:f[1][1],r:f[1][0]};
				var q = {c: cell.c, r:cell.r};
				if(supbooks.sharedf[encode_cell(c)]) {
					var parsedf = (supbooks.sharedf[encode_cell(c)]);
					stack.push(stringify_formula(parsedf, _range, q, supbooks, opts));
				}
				else {
					var fnd = false;
					for(e1=0;e1!=supbooks.arrayf.length; ++e1) {
						/* TODO: should be something like range_has */
						e2 = supbooks.arrayf[e1];
						if(c.c < e2[0].s.c || c.c > e2[0].e.c) continue;
						if(c.r < e2[0].s.r || c.r > e2[0].e.r) continue;
						stack.push(stringify_formula(e2[1], _range, q, supbooks, opts));
						fnd = true;
						break;
					}
					if(!fnd) stack.push(f[1]);
				}
				break;

			/* 2.5.198.32 TODO */
			case 'PtgArray':
				stack.push("{" + stringify_array(f[1]) + "}");
				break;

		/* 2.2.2.5 Mem Tokens */
			/* 2.5.198.70 TODO: confirm this is a non-display */
			case 'PtgMemArea':
				//stack.push("(" + f[2].map(encode_range).join(",") + ")");
				break;

			/* 2.5.198.38 */
			case 'PtgAttrSpace':
			/* 2.5.198.39 */
			case 'PtgAttrSpaceSemi':
				last_sp = ff;
				break;

			/* 2.5.198.92 TODO */
			case 'PtgTbl': break;

			/* 2.5.198.71 */
			case 'PtgMemErr': break;

			/* 2.5.198.74 */
			case 'PtgMissArg':
				stack.push("");
				break;

			/* 2.5.198.29 TODO */
			case 'PtgAreaErr': break;

			/* 2.5.198.31 TODO */
			case 'PtgAreaN': stack.push(""); break;

			/* 2.5.198.87 TODO */
			case 'PtgRefErr3d': break;

			/* 2.5.198.72 TODO */
			case 'PtgMemFunc': break;

			default: throw 'Unrecognized Formula Token: ' + f;
		}
		var PtgNonDisp = ['PtgAttrSpace', 'PtgAttrSpaceSemi', 'PtgAttrGoto'];
		if(last_sp >= 0 && PtgNonDisp.indexOf(formula[0][ff][0]) == -1) {
			f = formula[0][last_sp];
			switch(f[1][0]) {
				case 0: sp = fill(" ", f[1][1]); break;
				case 1: sp = fill("\r", f[1][1]); break;
				default:
					sp = "";
					if(opts.WTF) throw new Error("Unexpected PtgSpace type " + f[1][0]);
			}
			stack.push(sp + stack.pop());
			last_sp = -1;
		}
		//console.log("::",f, stack)
	}
	//console.log("--",stack);
	if(stack.length > 1 && opts.WTF) throw new Error("bad formula stack");
	return stack[0];
}
