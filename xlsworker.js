importScripts('xls.js');
postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v, cfb;
  try {
  /*
    cfb = XLS.CFB.read(oEvent.data, {type:"binary"});
    v = XLS.parse_xlscfb(cfb);
  */
    v = XLS.read(oEvent.data, {type:"binary"});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
  postMessage({t:"xls", d:v});
};
