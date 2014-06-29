var current_codepage = 1252, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('./dist/cpexcel');
	current_cptable = cptable[current_codepage];
}
function reset_cp() { set_cp(1252); }
function set_cp(cp) { current_codepage = cp; if(typeof cptable !== 'undefined') current_cptable = cptable[cp]; }

var _getchar = function _gc1(x) { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') _getchar = function _gc2(x) {
	if(x !== x || typeof x !== "number") throw new Error("dafuq");
	if(current_codepage === 1200) return String.fromCharCode(x);
	return cptable.utils.decode(current_codepage, [x&255,x>>8])[0];
};
