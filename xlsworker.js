importScripts('xls.js');
postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v, cfb;
  try {
    cfb = XLS.CFB.read(oEvent.data, {type:"binary"});
    v = XLS.parse_xlscfb(cfb);
  } catch(e) { postMessage({t:"e",d:e.stack}); }
  postMessage({t:"xls", d:v});
};
