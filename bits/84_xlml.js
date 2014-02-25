var everted_BERR = evert(BERR);

var magic_formats = {
	"General Number": "General",
	"General Date": SSF._table[22],
	"Long Date": "dddd, mmmm dd, yyyy",
	"Medium Date": SSF._table[15],
	"Short Date": SSF._table[14],
	"Long Time": SSF._table[19],
	"Medium Time": SSF._table[18],
	"Short Time": SSF._table[20],
	"Currency": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"Fixed": SSF._table[2],
	"Standard": SSF._table[4],
	"Percent": SSF._table[10],
	"Scientific": SSF._table[11],
	"Yes/No": '"Yes";"Yes";"No";@',
	"True/False": '"True";"True";"False";@',
	"On/Off": '"Yes";"Yes";"No";@',
};

function xlml_format(format, value) {
	return SSF.format(magic_formats[format] || unescapexml(format), value);
}

/* TODO: there must exist some form of OSP-blessed spec */
function parse_xlml_data(xml, data, cell, base, styles, o) {
	var nf = "General", sid = cell.StyleID; o = o || {};
	while(styles[sid]) {
		if(styles[sid].nf) nf = styles[sid].nf;
		if(!styles[sid].Parent) break;
		sid = styles[sid].Parent;
	}

	switch(data.Type) {
		case 'Boolean':
			cell.t = 'b'; cell.v = parsexmlbool(xml);
			break;
		case 'String':
			cell.t = 'str'; cell.v = fixstr(unescapexml(xml)); break;
		case 'DateTime':
			cell.v = (Date.parse(xml) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
			if(cell.v !== cell.v) cell.v = unescapexml(xml);
			else if(cell.v >= 1 && cell.v<60) cell.v = cell.v -1;
			/* falls through */
		case 'Number':
			if(typeof cell.v === 'undefined') cell.v=Number(xml);
			if(!cell.t) cell.t = 'n';
			break;
		case 'Error': cell.t = 'e'; cell.v = xml; cell.w = xml; break;
	}
	if(cell.t !== 'e') cell.w = xlml_format(nf||"General", cell.v);
	if(o.cellNF) cell.z = magic_formats[nf]||nf||"General";
	if(o.cellFormula && cell.Formula) {
		cell.f = rc_to_a1(unescapexml(cell.Formula), base); delete cell.Formula; }
	cell.ixfe = typeof cell.StyleID !== 'undefined' ? cell.StyleID : 'Default';

}

/* TODO: Everything */
function parse_xlml_xml(d, opts) {
	var str;
	if(typeof Buffer!=='undefined'&&d instanceof Buffer) str = d.toString('utf8');
	else if(typeof d === 'string') str = d;
	else throw "badf";
	var re = /<(\/?)([a-z]*:|)([A-Za-z]+)[^>]*>/mg, Rn;
	var state = [], tmp;
	var out = {};
	var sheets = {}, sheetnames = [], cursheet = {}, sheetname = "";
	var table = {}, cell = {}, row = {}, ddata = "", dtag, didx;
	var c = 0, r = 0;
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var styles = {}, stag = {};
	while((Rn = re.exec(str))) switch(Rn[3]) {
		case 'Data': {
			if(state[state.length-1][1]) break;
			if(Rn[1]==='/') parse_xlml_data(str.slice(didx, Rn.index), dtag, cell, {c:c,r:r}, styles, opts);
			else { dtag = parsexmltag(Rn[0]); didx = Rn.index + Rn[0].length; }
		} break;
		case 'Cell': {
			if(Rn[0].match(/\/>$/)) ++c;
			else if(Rn[1]==='/'){
				delete cell[0];
				cursheet[encode_cell({c:c,r:r})] = cell;
				++c;
			} else {
				cell = parsexmltag(Rn[0]);
				if(cell.Index) c = +cell.Index - 1;
				if(c < refguess.s.c) refguess.s.c = c;
				if(c > refguess.e.c) refguess.e.c = c;
			}
		} break;
		case 'Row': {
			if(Rn[1]==='/') {
				if(r < refguess.s.r) refguess.s.r = r;
				if(r > refguess.e.r) refguess.e.r = r;
				c = 0; ++r;
			} else { row = parsexmltag(Rn[0]); if(row.Index) r = +row.Index - 1; }
		} break;
		case 'Worksheet': {
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				sheetnames.push(sheetname);
				cursheet["!ref"] = encode_range(refguess);
				sheets[sheetname] = cursheet;
			} else {
				refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
				r = c = 0;
				state.push([Rn[3], false]);
				tmp = parsexmltag(Rn[0]);
				sheetname = tmp.Name;
				cursheet = {};
			}
		} break;
		case 'Table': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else {
				table = parsexmltag(Rn[0]);
				state.push([Rn[3], false]);
			}
		} break;

		case 'Style': {
			if(Rn[1]==='/') {
				styles[stag.ID] = stag;
			} else stag = parsexmltag(Rn[0]);
		} break;

		case 'NumberFormat': {
			stag.nf = parsexmltag(Rn[0]).Format || "General";
		} break;

		case 'NamedRange': break;
		case 'NamedCell': break;
		case 'Column': break;
		case 'B': break;
		case 'I': break;
		case 'U': break;
		case 'S': break;
		case 'Sub': break;
		case 'Sup': break;
		case 'Span': break;
		case 'Alignment': break;
		case 'Borders': break;
		case 'Font': break;
		case 'Interior': break;
		case 'Protection': break;

		case 'Styles':
		case 'Workbook': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else state.push([Rn[3], false]);
		} break;
		case 'Comment':
		case 'DocumentProperties':
		case 'CustomDocumentProperties':
		case 'OfficeDocumentSettings':
		case 'Names':
		case 'ExcelWorkbook':
		case 'WorkbookOptions':
		case 'WorksheetOptions': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else state.push([Rn[3], true]);
		} break;
		default: if(!state[state.length-1][1]) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
	}
	out.Sheets = sheets;
	out.SheetNames = sheetnames;
	out.SSF = SSF.get_table();
	return out;
}

function parse_xlml(data, opts) {
	switch((opts||{}).type||"base64") {
		case "base64": return parse_xlml_xml(Base64.decode(data), opts);
		case "binary": case "file": return parse_xlml_xml(data, opts);
		default: throw "dafuq";
	}
}
