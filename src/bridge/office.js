// Office.js is a different host API: Excel/Word/PowerPoint run + context.sync.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
export function officeMetadata(Office){
  if(!Office?.context?.host)throw Error('Office.js is not ready');
  const h=String(Office.context.host).toLowerCase();
  const kind=h.includes('excel')?'spreadsheet':h.includes('powerpoint')?'presentation':h.includes('word')?'document':null;
  if(!kind)throw Error(`Unsupported Office host: ${h}`);
  return {host:'office',kind,name:Office.context.document?.url?.split(/[\\/]/).pop()||`Active ${h} document`};
}
export async function officeExecute(globals,code,kind){
  const executor=kind==='spreadsheet'?globals.Excel:kind==='presentation'?globals.PowerPoint:globals.Word;
  if(!executor?.run)throw Error(`Office.js ${kind} run API is unavailable on this host`);
  const fn=new AsyncFunction('context',`"use strict";\n${code}`);
  return executor.run(async context=>{
    const value=await fn(context);
    await context.sync();
    return value;
  });
}
