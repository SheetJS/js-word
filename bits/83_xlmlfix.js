/* map from xlml named formats to SSF */
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
	"On/Off": '"Yes";"Yes";"No";@'
};

/* [MS-EXSPXML3] 2.4.54 ST_enmPattern */
var PatternTypeMap = {
	"None": "none",
	"Solid": "solid",
	"Gray50": "mediumGray",
	"Gray75": "darkGray",
	"Gray25": "lightGray",
	"HorzStripe": "darkHorizontal",
	"VertStripe": "darkVertical",
	"ReverseDiagStripe": "darkDown",
	"DiagStripe": "darkUp",
	"DiagCross": "darkGrid",
	"ThickDiagCross": "darkTrellis",
	"ThinHorzStripe": "lightHorizontal",
	"ThinVertStripe": "lightVertical",
	"ThinReverseDiagStripe": "lightDown",
	"ThinHorzCross": "lightGrid"
};

/* TODO: Normalize the properties */
function xlml_set_prop(Props, tag, val) {
	switch(tag) {
		case 'Description': tag = 'Comments'; break;
	}
	Props[tag] = val;
}

