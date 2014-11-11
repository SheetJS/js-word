function fix_opts_func(defaults) {
	return function fix_opts(opts) {
		for(var i = 0; i != defaults.length; ++i) {
			var d = defaults[i];
			if(opts[d[0]] === undefined) opts[d[0]] = d[1];
			if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
		}
	};
}

var fix_read_opts = fix_opts_func([
	['cellNF', false], /* emit cell number format string as .z */
	['cellFormula', true], /* emit formulae as .f */
	['cellStyles', false], /* emits style/theme as .s */

	['sheetRows', 0, 'n'], /* read n rows (0 = read all rows) */

	['bookSheets', false], /* only try to get sheet names (no Sheets) */
	['bookProps', false], /* only try to get properties (no Sheets) */
	['bookFiles', false], /* include raw file structure (cfb) */

	['password',''], /* password */
	['WTF', false] /* WTF mode (throws errors) */
]);
