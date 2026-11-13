import {startBridge} from './client.js';
import {wpsMetadata,wpsExecute} from './wps.js';
import {officeMetadata,officeExecute} from './office.js';

const $=id=>document.getElementById(id);
const isOffice=document.body.dataset.host==='office';
let controller;
function setStatus(s){$('status').textContent=s;}
async function attach(){
  if(controller)return;
  const code=$('code').value.trim();
  if(!/^\d{8}$/.test(code)){setStatus('Enter the 8-digit code from livedoc pair');return;}
  let metadata,execute;
  try{
    if(isOffice){
      await new Promise(resolve=>Office.onReady(resolve));
      metadata=officeMetadata(Office);
      execute=script=>officeExecute({Excel:globalThis.Excel,Word:globalThis.Word,PowerPoint:globalThis.PowerPoint},script,metadata.kind);
    }else{
      metadata=wpsMetadata(globalThis.Application || globalThis.wps);
      execute=script=>wpsExecute(globalThis.Application || globalThis.wps,script);
    }
  }catch(e){setStatus(`Host not ready: ${e.message}`);return;}
  // Pairing is a deliberate user action. This allows executing arbitrary trusted scripts.
  controller=new AbortController();$('connect').disabled=true;$('disconnect').disabled=false;
  try{await startBridge({code,metadata,execute,onStatus:setStatus,signal:controller.signal});}
  catch(e){setStatus(`Pairing failed: ${e.message}`);}
  finally{controller=null;$('connect').disabled=false;$('disconnect').disabled=true;}
}
$('connect').addEventListener('click',()=>{void attach();});
$('disconnect').addEventListener('click',()=>{controller?.abort();});
setStatus(isOffice?'Office bridge ready to pair':'WPS bridge ready to pair');
