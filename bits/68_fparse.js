/* 2.4.127 TODO */
function parse_Formula(blob, length) {
	var cell = parse_Cell(blob, 6);
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	blob.read_shift(1);
	var chn = blob.read_shift(4);
	var cbf = parse_CellParsedFormula(blob, length-20);
	return {cell:cell, val:val, formula:cbf, shared: (flags >> 3) & 1};
}

/* 2.5.133 */
function parse_FormulaValue(blob) {
	var b;
	if(blob.readUInt16LE(blob.l + 6) !== 0xFFFF) return parse_Xnum(blob);
	switch(blob[blob.l]) {
		case 0x00: blob.l += 8; return "String";
		case 0x01: b = blob[blob.l+2] === 0x1; blob.l += 8; return b;
		case 0x02: b = BERR[blob[blob.l+2]]; blob.l += 8; return b;
		case 0x03: blob.l += 8; return "";
	}
}

/* 2.5.198.103 */
function parse_RgbExtra(blob, length, rgce) {
	var target = blob.l + length;
	var o = [];
	for(var i = 0; i !== rgce.length; ++i) {
		switch(rgce[i][0]) {
			case 'PtgArray': /* PtgArray -> PtgExtraArray */
				rgce[i][1] = parse_PtgExtraArray(blob);
				o.push(rgce[i][1]);
				break;
			default: break;
		}
	}
	length = target - blob.l;
	if(length !== 0) o.push(parsenoop(blob, length));
	return o;
}

/* 2.5.198.21 */
function parse_NameParsedFormula(blob, length, cce) {
	var target = blob.l + length;
	var rgce = parse_Rgce(blob, cce);
	var rgcb;
	if(target !== blob.l) rgcb = parse_RgbExtra(blob, target - blob.l, rgce);
	return [rgce, rgcb];
}

/* 2.5.198.3 TODO */
function parse_CellParsedFormula(blob, length) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce);
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, length - cce - 2, rgce);
	return [rgce, rgcb];
}

/* 2.5.198.118 TODO */
function parse_SharedParsedFormula(blob, length) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce);
	return [rgce, rgcb];
}

/* 2.5.198.104 */
function parse_Rgce(blob, length) {
	var target = blob.l + length;
	var R, id, ptgs = [];
	while(target != blob.l) {
		length = target - blob.l;
		id = blob[blob.l];
		R = PtgTypes[id];
		//console.log("ptg", id, R)
		if(id === 0x18 || id === 0x19) {
			id = blob[blob.l + 1];
			R = (id === 0x18 ? Ptg18 : Ptg19)[id];
		}
		if(!R) { ptgs.push(parsenoop(blob, length)); }
		else { ptgs.push([R.n, R.f(blob, length)]); }
	}
	return ptgs;
};

/* 2.2.2 + Magic TODO */
function stringify_formula(formula, range, cell, supbooks) {
	range = range || {s:{c:0, r:0}};
	var stack = [], e1, e2, type, c, ixti;
	if(!formula[0] || !formula[0][0]) return "";
	//console.log("--",formula[0])
	formula[0].forEach(function(f) {
		//console.log("++",f)
		switch(f[0]) {
		/* 2.2.2.1 Unary Operator Tokens */
			/* 2.5.198.93 */
			case 'PtgUminus': stack.push("-" + stack.pop()); break;
			/* 2.5.198.95 */
			case 'PtgUplus': stack.push("+" + stack.pop()); break;
			/* 2.5.198.81 */
			case 'PtgPercent': stack.push(stack.pop() + "%"); break;

		/* 2.2.2.1 Binary Value Operator Token */
			/* 2.5.198.26 */
			case 'PtgAdd':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"+"+e1);
				break;
			/* 2.5.198.90 */
			case 'PtgSub':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"-"+e1);
				break;
			/* 2.5.198.75 */
			case 'PtgMul':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"*"+e1);
				break;
			/* 2.5.198.45 */
			case 'PtgDiv':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"/"+e1);
				break;
			/* 2.5.198.82 */
			case 'PtgPower':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"^"+e1);
				break;
			/* 2.5.198.43 */
			case 'PtgConcat':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"&"+e1);
				break;
			/* 2.5.198.69 */
			case 'PtgLt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<"+e1);
				break;
			/* 2.5.198.68 */
			case 'PtgLe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<="+e1);
				break;
			/* 2.5.198.56 */
			case 'PtgEq':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"="+e1);
				break;
			/* 2.5.198.64 */
			case 'PtgGe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">="+e1);
				break;
			/* 2.5.198.65 */
			case 'PtgGt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">"+e1);
				break;
			/* 2.5.198.78 */
			case 'PtgNe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<>"+e1);
				break;

		/* 2.2.2.1 Binary Reference Operator Token */
			/* 2.5.198.67 */
			case 'PtgIsect':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+" "+e1);
				break;
			case 'PtgUnion':
			case 'PtgRange': break;

		/* 2.2.2.3 Control Tokens "can be ignored" */
			/* 2.5.198.34 */
			case 'PtgAttrChoose': break;
			/* 2.5.198.35 */
			case 'PtgAttrGoto': break;
			/* 2.5.198.36 */
			case 'PtgAttrIf': break;


			/* 2.5.198.84 */
			case 'PtgRef':
				type = f[1][0], c = shift_cell(decode_cell(encode_cell(f[1][1])), range);
				stack.push(encode_cell(c));
				break;
			/* 2.5.198.88 */
			case 'PtgRefN':
				type = f[1][0], c = shift_cell(decode_cell(encode_cell(f[1][1])), cell);
				stack.push(encode_cell(c));
				break;
			case 'PtgRef3d': // TODO: lots of stuff
				type = f[1][0], ixti = f[1][1], c = shift_cell(f[1][2], range);
				stack.push(supbooks[1][ixti+1]+"!"+encode_cell(c));
				break;

		/* Function Call */
			/* 2.5.198.62 */
			case 'PtgFunc':
			/* 2.5.198.63 */
			case 'PtgFuncVar':
				/* f[1] = [argc, func] */
				var argc = f[1][0], func = f[1][1];
				if(!argc) argc = 0;
				var args = stack.slice(-argc);
				stack.length -= argc;
				if(func === 'User') func = args.shift();
				stack.push(func + "(" + args.join(",") + ")");
				break;

			/* 2.5.198.66 */
			case 'PtgInt': stack.push(f[1]); break;
			/* 2.5.198.79 TODO: precision? */
			case 'PtgNum': stack.push(String(f[1])); break;
			/* 2.5.198.89 */
			case 'PtgStr': stack.push('"' + f[1] + '"'); break;

			/* 2.5.198.27 */
			case 'PtgArea':
				type = f[1][0], r = shift_range(f[1][1], range);
				stack.push(encode_range(r));
				break;
			/* 2.5.198.28 */
			case 'PtgArea3d': // TODO: lots of stuff
				type = f[1][0], ixti = f[1][1], r = f[1][2];
				stack.push(supbooks[1][ixti+1]+"!"+encode_range(r));
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
				var nameidx = f[1][2];
				var lbl = supbooks[0][nameidx];
				var name = lbl.Name;
				if(name in XLSXFutureFunctions) name = XLSXFutureFunctions[name];
				stack.push(name);
				break;

			/* 2.5.97.61 TODO: do something different for revisions */
			case 'PtgNameX':
				/* f[1] = type, ixti, nameindex */
				var bookidx = f[1][1], nameidx = f[1][2];
				var externbook = supbooks[bookidx+1][nameidx];
				stack.push(externbook.body);
				break;

		/* 2.2.2.4 Display Tokens */
			/* 2.5.198.80 */
			case 'PtgParen': stack.push('(' + stack.pop() + ')'); break;

		/* */
			/* 2.5.198.58 TODO */
			case 'PtgExp':
				var c = {c:f[1][1],r:f[1][0]};
				if(supbooks.sharedf[encode_cell(c)]) {
					var parsedf = (supbooks.sharedf[encode_cell(c)]);
					var q = {c: cell.c, r:cell.r};
					stack.push(stringify_formula(parsedf, range, q, supbooks));
				}
				else stack.push(f[1]);
				break;

			/* 2.5.198.32 TODO */
			case 'PtgArray':
				stack.push("{" + f[1].map(function(x) { return x.map(function(y) { return y[1];}).join(",");}).join(";") + "}");
				break;

			default: throw 'Unrecognized Formula Token: ' + f;
		}
	});
	//console.log("--",stack);
	return stack[0];
}
