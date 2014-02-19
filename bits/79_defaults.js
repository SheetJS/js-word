function fixopts(opts) {
	var defaults = [
		['cellNF', false], /* emit cell number format string as .z */
		['cellFormula', true], /* emit formulae as .f */

		['sheetRows', 0, 'n'], /* read n rows (0 = read all rows) */

		['bookSheets', false], /* only try to get sheet names (no Sheets) */
		['bookProps', false], /* only try to get properties (no Sheets) */

		['WTF', false] /* WTF mode (throws errors) */
	];
	defaults.forEach(function(d) {
		if(typeof opts[d[0]] === 'undefined') opts[d[0]] = d[1];
		if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
	});
}
