import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs/promises';

const PLUGINS=[['LiveDocumentWriter','wps'],['LiveDocumentSheet','et'],['LiveDocumentSlide','wpp']];
const RESERVED=/^LiveDocument(?:Writer|Sheet|Slide)$/;
export function wpsAddonDirectory(platform=process.platform,env=process.env,home=os.homedir()){
  if(env.LIVEDOC_WPS_ADDONS_DIR)return path.resolve(env.LIVEDOC_WPS_ADDONS_DIR);
  if(platform==='win32')return path.join(env.APPDATA||path.join(home,'AppData','Roaming'),'kingsoft','wps','jsaddons');
  if(platform==='darwin')return path.join(home,'Library','Containers','com.kingsoft.wpsoffice.mac','Data','.kingsoft','wps','jsaddons');
  return path.join(home,'.local','share','Kingsoft','wps','jsaddons');
}
function xmlEscape(s){return s.replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;').replaceAll('>','&gt;');}
export function updateWpsPublishXml(existing,baseUrl,action='install'){
  if(!['install','uninstall'].includes(action))throw Error('Invalid action');
  const original=existing?.trim()||'<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n</jsplugins>';
  if((original.match(/<jsplugins(?:\s[^>]*)?>/gi)||[]).length!==1 || (original.match(/<\/jsplugins>/gi)||[]).length!==1)throw Error('Unrecognized publish.xml: must have exactly one jsplugins root; no file changed');
  const pattern=/<jspluginonline\b[^>]*\/\s*>/gi;
  let filtered=original.replace(pattern,element=>{
    const m=element.match(/\bname\s*=\s*["']([^"']+)["']/i);
    return m&&RESERVED.test(m[1])?'':element;
  });
  let addition='';
  if(action==='install'){
    const u=new URL(baseUrl);
    if(!['http:','https:'].includes(u.protocol)||!['127.0.0.1','localhost'].includes(u.hostname))throw Error('Only local HTTP(S) WPS add-in URLs are supported');
    if(!u.pathname.endsWith('/'))throw Error('WPS add-in URL must end with /');
    addition='\n'+PLUGINS.map(([name,type])=>`  <jspluginonline name="${name}" type="${type}" url="${xmlEscape(u.href)}" enable="true"/>`).join('\n')+'\n';
  }
  filtered=filtered.replace(/\s*<\/jsplugins>/i,`${addition}</jsplugins>`);
  return filtered.endsWith('\n')?filtered:filtered+'\n';
}
export async function writeWpsPublish({action='install',baseUrl='http://127.0.0.1:18765/wps-addon/',directory=wpsAddonDirectory(),dryRun=false}={}){
  const file=path.join(directory,'publish.xml');let current=null;
  try{current=await fs.readFile(file,'utf8');}catch(e){if(e.code!=='ENOENT')throw e;}
  if(action==='uninstall' && current===null)return {file,changed:false,reason:'No publish.xml exists'};
  const next=updateWpsPublishXml(current,baseUrl,action);
  if(next===current)return {file,changed:false};
  if(dryRun)return {file,changed:true,preview:next};
  await fs.mkdir(directory,{recursive:true});
  let backup=null;
  if(current!==null){
    backup=path.join(directory,`publish.xml.livedoc-backup-${Date.now()}`);
    await fs.copyFile(file,backup,fs.constants.COPYFILE_EXCL);
  }
  const tmp=path.join(directory,`publish.xml.livedoc-${process.pid}-${Date.now()}.tmp`);
  try{await fs.writeFile(tmp,next,{encoding:'utf8',flag:'wx'});await fs.rename(tmp,file);}
  catch(e){await fs.rm(tmp,{force:true}).catch(()=>{});throw e;}
  return {file,backup,changed:true,action};
}
