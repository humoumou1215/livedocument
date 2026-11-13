// Runs inside WPS's JavaScript add-in webview, not in the Node runtime.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
export function wpsMetadata(Application) {
  if (!Application) throw Error('WPS Application object is unavailable: launch this page from a WPS add-in');
  const pairs=[['spreadsheet','ActiveWorkbook'],['presentation','ActivePresentation'],['document','ActiveDocument']];
  for(const [kind,field] of pairs){
    try{const doc=Application[field];if(doc && typeof doc.Name==='string')return {host:'wps',kind,name:doc.Name};}catch{}
  }
  throw Error('No active WPS document available in this add-in context');
}
export function wpsExecute(Application,code){
  if(!Application)throw Error('WPS Application object unavailable');
  // Trusted-session execution, NOT a sandbox. Host's native API is supplied as Application.
  // A return statement is required for result extraction; async/await is allowed.
  const fn=new AsyncFunction('Application',`"use strict";\n${code}`);
  return fn(Application);
}
