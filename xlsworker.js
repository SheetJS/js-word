/* xls.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* uncomment the next line for encoding support */
//importScripts('dist/cpexcel.js');
importScripts('xls.js');
postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v;
  try {
    v = XLS.read(oEvent.data.d, {type: oEvent.data.b ? 'binary' : 'base64'});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
  postMessage({t:"xls", d:JSON.stringify(v)});
};
