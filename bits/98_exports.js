XLS.parse_xlscfb = parse_xlscfb;
XLS.read = xlsread;
XLS.readFile = readFile;
XLS.utils = utils;
XLS.CFB = CFB;
if(typeof module !== 'undefined' && require.main === module ) {
	var wb = readFile(process.argv[2] || 'Book1.xls');
	var target_sheet = process.argv[3] || '';
	if(target_sheet === '') target_sheet = wb.Directory[0];
	var ws = wb.Sheets[target_sheet];
	console.log(target_sheet);
	console.log(make_csv(ws));
	//console.log(get_formulae(ws));
}
