/* TODO: it will be useful to parse the function str */
function rc_to_a1(fstr, base) {
	return fstr.replace(/(^|[^A-Za-z])R(\[?)(-?\d+|)\]?C(\[?)(-?\d+|)\]?/g,function($$,$1,$2,$3,$4,$5) {
		var R = $3.length?+$2:0, C = $5.length?+$4:0;
		if(C<0 && !$4) C=0;
		return ($1||"")+encode_cell({c:$4?base.c+C:C,r:$2?base.r+R:R});
	});
}
