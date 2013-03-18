#!/usr/bin/env node

var XLS = require('../xls'); 
var wb = XLS.readFile(process.argv[2] || 'Book1.xls');
var target_sheet = process.argv[3] || '';
if(target_sheet === '') target_sheet = wb.Directory[0];
var ws = wb.Sheets[target_sheet];
console.log(target_sheet);
console.log(XLS.utils.make_csv(ws));
