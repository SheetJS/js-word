/* TODO: it will be useful to parse the function str */
function rc_to_a1(fstr, base) {
	return fstr.replace(/(^|[^A-Za-z])R\[?(-?\d+|)\]?C\[?(-?\d+|)\]?/g,function($$,$1,$2,$3,$4) {
		var R = $2.length?+$2:0, C = $3.length?+$3:0;
		return ($1||"")+encode_cell({c:base.c+C,r:base.r+R});
	});
}
