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
function parse_xlml_data(xml, ss, data, cell, base, styles, o) {
	var nf = "General", sid = cell.StyleID; o = o || {};
	while(styles[sid]) {
		if(styles[sid].nf) nf = styles[sid].nf;
		if(!styles[sid].Parent) break;
		sid = styles[sid].Parent;
	}
	switch(data.Type) {
		case 'Boolean':
			cell.t = 'b';
			cell.v = parsexmlbool(xml);
			break;
		case 'String':
			cell.t = 'str'; cell.r = fixstr(unescapexml(xml));
			cell.v = xml.indexOf("<") > -1 ? ss : cell.r;
			break;
		case 'DateTime':
			cell.v = (Date.parse(xml) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
			if(cell.v !== cell.v) cell.v = unescapexml(xml);
			else if(cell.v >= 1 && cell.v<60) cell.v = cell.v -1;
			if(!nf || nf == "General") nf = "yyyy-mm-dd";
			/* falls through */
		case 'Number':
			if(typeof cell.v === 'undefined') cell.v=Number(xml);
			if(!cell.t) cell.t = 'n';
			/* TODO: the next line is undocumented black magic */
			//if((!nf || (nf == "General" && !cell.Formula)) && (cell.v != (cell.v|0))) { nf="#,##0.00"; console.log(cell.v, cell.v|0)}
			break;
		case 'Error': cell.t = 'e'; cell.v = xml; cell.w = xml; break;
		default: cell.t = 's'; cell.v = fixstr(ss); break;
	}
	if(cell.t !== 'e') try {
		cell.w = xlml_format(nf||"General", cell.v);
		if(o.cellNF) cell.z = magic_formats[nf]||nf||"General";
	} catch(e) { if(o.WTF) throw e; }
	if(o.cellFormula && cell.Formula) {
		cell.f = rc_to_a1(unescapexml(cell.Formula), base);
		delete cell.Formula;
	}
	cell.ixfe = typeof cell.StyleID !== 'undefined' ? cell.StyleID : 'Default';
}

/* TODO: Everything */
function parse_xlml_xml(d, opts) {
	var str;
	if(typeof Buffer!=='undefined'&&d instanceof Buffer) str = d.toString('utf8');
	else if(typeof d === 'string') str = d;
	else throw "badf";
	var re = /<(\/?)([a-z0-9]*:|)([A-Za-z_0-9]+)[^>]*>/mg, Rn;
	var state = [], tmp;
	var out = {};
	var sheets = {}, sheetnames = [], cursheet = {}, sheetname = "";
	var table = {}, cell = {}, row = {}, ddata = "", dtag, didx;
	var c = 0, r = 0;
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var styles = {}, stag = {};
	var ss = "", fidx = 0;
	var mergecells = [];
	var Props = {}, Custprops = {}, pidx = 0;
	var comments = [], comment = {};
	while((Rn = re.exec(str))) switch(Rn[3]) {
		case 'Data': {
			if(state[state.length-1][1]) break;
			if(Rn[1]==='/') parse_xlml_data(str.slice(didx, Rn.index), ss, dtag, state[state.length-1][0]=="Comment"?comment:cell, {c:c,r:r}, styles, opts);
			else { ss = ""; dtag = parsexmltag(Rn[0]); didx = Rn.index + Rn[0].length; }
		} break;
		case 'Cell': {
			if(Rn[1]==='/'){
				delete cell[0];
				if(comments.length > 0) cell.c = comments;
				if((!opts.sheetRows || opts.sheetRows > r) && typeof cell.v !== 'undefined') cursheet[encode_cell({c:c,r:r})] = cell;
				if(cell.HRef) {
					cell.l = {Target:cell.HRef, tooltip:cell.HRefScreenTip};
					delete cell.HRef; delete cell.HRefScreenTip;
				}
				if(cell.MergeAcross || cell.MergeDown) {
					var cc = c + Number(cell.MergeAcross||0);
					var rr = r + Number(cell.MergeDown||0);
					mergecells.push({s:{c:c,r:r},e:{c:cc,r:rr}});
				}
				++c;
				if(cell.MergeAcross) c += +cell.MergeAcross;
			} else {
				cell = parsexmltag(Rn[0]);
				if(cell.Index) c = +cell.Index - 1;
				if(c < refguess.s.c) refguess.s.c = c;
				if(c > refguess.e.c) refguess.e.c = c;
				if(Rn[0].match(/\/>$/)) ++c;
				comments = [];
			}
		} break;
		case 'Row': {
			if(Rn[1]==='/' || Rn[0].match(/\/>$/)) {
				if(r < refguess.s.r) refguess.s.r = r;
				if(r > refguess.e.r) refguess.e.r = r;
				if(Rn[0].match(/\/>$/)) {
					row = parsexmltag(Rn[0]);
					if(row.Index) r = +row.Index - 1;
				}
				c = 0; ++r;
			} else {
				row = parsexmltag(Rn[0]);
				if(row.Index) r = +row.Index - 1;
			}
		} break;
		case 'Worksheet': { /* TODO: read range from FullRows/FullColumns */
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				sheetnames.push(sheetname);
				cursheet["!ref"] = encode_range(refguess);
				if(mergecells.length) cursheet["!merges"] = mergecells;
				sheets[sheetname] = cursheet;
			} else {
				refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
				r = c = 0;
				state.push([Rn[3], false]);
				tmp = parsexmltag(Rn[0]);
				sheetname = tmp.Name;
				cursheet = {};
				mergecells = [];
			}
		} break;
		case 'Table': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else if(Rn[0].slice(-2) == "/>") break;
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
		case 'Border': break;
		case 'Alignment': break;
		case 'Borders': break;
		case 'Font': {
			if(Rn[0].match(/\/>$/)) break;
			else if(Rn[1]==="/") ss += str.slice(fidx, Rn.index);
			else fidx = Rn.index + Rn[0].length;
		} break;
		case 'Interior': break;
		case 'Protection': break;

		/* TODO: Normalize the properties */
		case 'Author':
		case 'Title':
		case 'Description':
		case 'Created':
		case 'Keywords':
		case 'Subject':
		case 'Category':
		case 'Company':
		case 'LastAuthor':
		case 'LastSaved':
		case 'LastPrinted':
		case 'Version':
		case 'Revision':
		case 'TotalTime':
		case 'HyperlinkBase':
		case 'Manager': {
			if(Rn[0].match(/\/>$/)) break;
			else if(Rn[1]==="/") Props[Rn[3]] = str.slice(pidx, Rn.index);
			else pidx = Rn.index + Rn[0].length;
		} break;
		case 'Paragraphs': break;

		/* OfficeDocumentSettings */
		case 'AllowPNG': break;
		case 'RemovePersonalInformation': break;
		case 'DownloadComponents': break;
		case 'LocationOfComponents': break;
		case 'Colors': break;
		case 'Color': break;
		case 'Index': break;
		case 'RGB': break;
		case 'PixelsPerInch': break;
		case 'TargetScreenSize': break;
		case 'ReadOnlyRecommended': break;

		/* ComponentOptions */
		case 'Toolbar': break;
		case 'HideOfficeLogo': break;
		case 'SpreadsheetAutoFit': break;
		case 'Label': break;
		case 'Caption': break;
		case 'MaxHeight': break;
		case 'MaxWidth': break;
		case 'NextSheetNumber': break;

		/* ExcelWorkbook */
		case 'WindowHeight': break;
		case 'WindowWidth': break;
		case 'WindowTopX': break;
		case 'WindowTopY': break;
		case 'TabRatio': break;
		case 'ProtectStructure': break;
		case 'ProtectWindows': break;
		case 'ActiveSheet': break;
		case 'DisplayInkNotes': break;
		case 'FirstVisibleSheet': break;
		case 'SupBook': break;
		case 'SheetName': break;
		case 'SheetIndex': break;
		case 'SheetIndexFirst': break;
		case 'SheetIndexLast': break;
		case 'Dll': break;
		case 'AcceptLabelsInFormulas': break;
		case 'DoNotSaveLinkValues': break;
		case 'Date1904': break;
		case 'Iteration': break;
		case 'MaxIterations': break;
		case 'MaxChange': break;
		case 'Path': break;
		case 'Xct': break;
		case 'Count': break;
		case 'SelectedSheets': break;
		case 'Calculation': break;
		case 'Uncalced': break;
		case 'StartupPrompt': break;
		case 'Crn': break;
		case 'ExternName': break;
		case 'Formula': break;
		case 'ColFirst': break;
		case 'ColLast': break;
		case 'WantAdvise': break;
		case 'Boolean': break;
		case 'Error': break;
		case 'Text': break;
		case 'OLE': break;
		case 'NoAutoRecover': break;
		case 'PublishObjects': break;

		/* WorkbookOptions */
		case 'OWCVersion': break;
		case 'Height': break;
		case 'Width': break;

		/* WorksheetOptions */
		case 'Unsynced': break;
		case 'Visible': break;
		case 'Print': break;
		case 'Panes': break;
		case 'Scale': break;
		case 'Pane': break;
		case 'Number': break;
		case 'Layout': break;
		case 'Header': break;
		case 'Footer': break;
		case 'PageSetup': break;
		case 'PageMargins': break;
		case 'Selected': break;
		case 'ProtectObjects': break;
		case 'EnableSelection': break;
		case 'ProtectScenarios': break;
		case 'ValidPrinterInfo': break;
		case 'HorizontalResolution': break;
		case 'VerticalResolution': break;
		case 'NumberofCopies': break;
		case 'ActiveRow': break;
		case 'ActiveCol': break;
		case 'ActivePane': break;
		case 'TopRowVisible': break;
		case 'TopRowBottomPane': break;
		case 'LeftColumnVisible': break;
		case 'LeftColumnRightPane': break;
		case 'FitToPage': break;
		case 'RangeSelection': break;
		case 'PaperSizeIndex': break;
		case 'PageLayoutZoom': break;
		case 'PageBreakZoom': break;
		case 'FilterOn': break;
		case 'DoNotDisplayGridlines': break;
		case 'SplitHorizontal': break;
		case 'SplitVertical': break;
		case 'FreezePanes': break;
		case 'FrozenNoSplit': break;
		case 'FitWidth': break;
		case 'FitHeight': break;
		case 'CommentsLayout': break;
		case 'Zoom': break;
		case 'LeftToRight': break;
		case 'Gridlines': break;
		case 'AllowSort': break;
		case 'AllowFilter': break;
		case 'AllowInsertRows': break;
		case 'AllowDeleteRows': break;
		case 'AllowInsertCols': break;
		case 'AllowDeleteCols': break;
		case 'AllowInsertHyperlinks': break;
		case 'AllowFormatCells': break;
		case 'AllowSizeCols': break;
		case 'AllowSizeRows': break;
		case 'RefModeR1C1': break;
		case 'NoSummaryRowsBelowDetail': break;
		case 'TabColorIndex': break;
		case 'DoNotDisplayHeadings': break;
		case 'ShowPageLayoutZoom': break;
		case 'NoSummaryColumnsRightDetail': break;
		case 'BlackAndWhite': break;
		case 'DoNotDisplayZeros': break;
		case 'DisplayPageBreak': break;
		case 'RowColHeadings': break;
		case 'DoNotDisplayOutline': break;
		case 'NoOrientation': break;
		case 'AllowUsePivotTables': break;
		case 'ZeroHeight': break;
		case 'ViewableRange': break;
		case 'Selection': break;
		case 'ProtectContents': break;

		/* PivotTable */
		case 'ImmediateItemsOnDrop': break;
		case 'ShowPageMultipleItemLabel': break;
		case 'CompactRowIndent': break;
		case 'Location': break;
		case 'PivotField': break;
		case 'Orientation': break;
		case 'LayoutForm': break;
		case 'LayoutSubtotalLocation': break;
		case 'LayoutCompactRow': break;
		case 'Position': break;
		case 'PivotItem': break;
		case 'DataType': break;
		case 'DataField': break;
		case 'SourceName': break;
		case 'ParentField': break;
		case 'PTLineItems': break;
		case 'PTLineItem': break;
		case 'CountOfSameItems': break;
		case 'Item': break;
		case 'ItemType': break;
		case 'PTSource': break;
		case 'CacheIndex': break;
		case 'ConsolidationReference': break;
		case 'FileName': break;
		case 'Reference': break;
		case 'NoColumnGrand': break;
		case 'NoRowGrand': break;
		case 'BlankLineAfterItems': break;
		case 'DoNotCalculateBeforeSave': break;
		case 'Hidden': break;
		case 'Subtotal': break;
		case 'BaseField': break;
		case 'MapChildItems': break;
		case 'Function': break;
		case 'RefreshOnFileOpen': break;
		case 'PrintSetTitles': break;
		case 'MergeLabels': break;

		/* PageBreaks */
		case 'ColBreaks': break;
		case 'ColBreak': break;
		case 'RowBreaks': break;
		case 'RowBreak': break;
		case 'ColStart': break;
		case 'ColEnd': break;
		case 'RowEnd': break;

		/* Version */
		case 'DefaultVersion': break;
		case 'RefreshName': break;
		case 'RefreshDate': break;
		case 'RefreshDateCopy': break;
		case 'VersionLastEdit': break;
		case 'VersionLastRefresh': break;
		case 'VersionLastUpdate': break;
		case 'VersionUpdateableMin': break;
		case 'VersionRefreshableMin': break;

		/* ConditionalFormatting */
		case 'Range': break;
		case 'Condition': break;
		case 'Qualifier': break;
		case 'Value1': break;
		case 'Value2': break;
		case 'Format': break;

		/* AutoFilter */
		case 'AutoFilter': break;
		case 'AutoFilterColumn': break;
		case 'AutoFilterCondition': break;
		case 'AutoFilterAnd': break;
		case 'AutoFilterOr': break;

		/* QueryTable */
		case 'Name': break;
		case 'Id': break;
		case 'AutoFormatFont': break;
		case 'AutoFormatPattern': break;
		case 'QuerySource': break;
		case 'QueryType': break;
		case 'EnableRedirections': break;
		case 'RefreshedInXl9': break;
		case 'URLString': break;
		case 'HTMLTables': break;
		case 'Connection': break;
		case 'CommandText': break;
		case 'RefreshInfo': break;
		case 'NoTitles': break;
		case 'NextId': break;
		case 'ColumnInfo': break;
		case 'OverwriteCells': break;
		case 'UseBlank': break;
		case 'DoNotPromptForFile': break;
		case 'TextWizardSettings': break;
		case 'Source': break;
		case 'Decimal': break;
		case 'ThousandSeparator': break;
		case 'TrailingMinusNumbers': break;
		case 'FormatSettings': break;
		case 'FieldType': break;
		case 'Delimiters': break;
		case 'Tab': break;
		case 'Comma': break;
		case 'AutoFormatName': break;

		/* DataValidation */
		case 'Type': break;
		case 'Min': break;
		case 'Max': break;
		case 'Sorting': break;
		case 'Sort': break;
		case 'Descending': break;
		case 'Order': break;
		case 'CaseSensitive': break;
		case 'Value': break;
		case 'ErrorStyle': break;
		case 'ErrorMessage': break;
		case 'ErrorTitle': break;
		case 'CellRangeList': break;
		case 'InputMessage': break;
		case 'InputTitle': break;
		case 'ComboHide': break;
		case 'InputHide': break;

		/* MapInfo (schema) */
		case 'Schema': break;
		case 'Map': break;
		case 'Entry': break;
		case 'XPath': break;
		case 'Field': break;
		case 'XSDType': break;
		case 'Aggregate': break;
		case 'ElementType': break;
		case 'AttributeType': break;
		/* These are from xsd (XML Schema Definition) */
		case 'schema':
		case 'element':
		case 'complexType':
		case 'datatype':
		case 'all':
		case 'attribute':
		case 'extends': break;

		case 'data': case 'row': break;

		case 'Styles':
		case 'Workbook': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else state.push([Rn[3], false]);
		} break;

		case 'Comment': {
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				comment.t = comment.v;
				delete comment.v; delete comment.w; delete comment.ixfe;
				comments.push(comment);
			} else {
				state.push([Rn[3], false]);
				tmp = parsexmltag(Rn[0]);
				comment = {a:tmp.Author};
			}
		} break;
		case 'ComponentOptions':
		case 'DocumentProperties':
		case 'CustomDocumentProperties':
		case 'OfficeDocumentSettings':
		case 'PivotTable':
		case 'PivotCache':
		case 'Names':
		case 'MapInfo':
		case 'PageBreaks':
		case 'QueryTable':
		case 'DataValidation':
		case 'ConditionalFormatting':
		case 'ExcelWorkbook':
		case 'WorkbookOptions':
		case 'WorksheetOptions': {
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
			else state.push([Rn[3], true]);
		} break;

		/* CustomDocumentProperties */
		default:
			if(!state[state.length-1][1]) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
			if(state[state.length-1][0]==='CustomDocumentProperties') {
				if(Rn[0].match(/\/>$/)) break;
				else if(Rn[1]==="/") Custprops[Rn[3].replace(/_x0020_/g," ")] = str.slice(pidx, Rn.index);
				else pidx = Rn.index + Rn[0].length;
				break;
			}
			if(opts.WTF) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
	}
	if(!opts.bookSheets && !opts.bookProps) out.Sheets = sheets;
	out.SheetNames = sheetnames;
	out.SSF = SSF.get_table();
	out.Props = Props;
	out.Custprops = Custprops;
	return out;
}

function parse_xlml(data, opts) {
	fixopts(opts=opts||{});
	switch(opts.type||"base64") {
		case "base64": return parse_xlml_xml(Base64.decode(data), opts);
		case "binary": case "file": return parse_xlml_xml(data, opts);
		case "array": return parse_xlml_xml(data.map(function(x) { return String.fromCharCode(x);}).join(""), opts);
		default: throw "dafuq";
	}
}
