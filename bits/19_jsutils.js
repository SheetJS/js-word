function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined') return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || !o) return o;
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}

function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }
