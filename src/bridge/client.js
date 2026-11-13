// Browser-side protocol shared by WPS and Office task panes.
export async function request(path,payload,token,signal){
  const response=await fetch(path,{method:'POST',headers:{'Content-Type':'application/json',...(token?{Authorization:`Bearer ${token}`}:{})},body:JSON.stringify(payload),signal});
  if(response.status===204)return null;
  const obj=await response.json();
  if(!response.ok || !obj.ok){const e=new Error(obj.error?.message||`HTTP ${response.status}`);e.status=response.status;throw e;}
  return obj;
}
export async function startBridge({code,metadata,execute,onStatus=()=>{},signal}){
  const paired=await request('/v1/pair',{code,...metadata});
  const targetId=paired.targetId,token=paired.bridgeToken;
  onStatus(`Connected: ${metadata.name} (${targetId})`);
  const payload={targetId};
  while(!signal?.aborted){
    let job;
    try {const res=await request('/v1/bridge/next',payload,token,signal);job=res?.job;}catch(e){
      if(signal?.aborted)break;
      if(e.status===401 || e.status===403){onStatus(`Session expired: ${e.message}. Pair again.`);break;}
      onStatus(`Connection lost: ${e.message}. Retrying…`);
      await new Promise(r=>setTimeout(r,1500));continue;
    }
    if(!job)continue;
    onStatus(`Running job ${job.id.slice(0,8)} in ${metadata.name}`);
    let result;
    try{
      const value=await execute(job.code);
      // JSON round trip detects unserializable native host proxies and BigInt.
      const json=JSON.stringify(value===undefined?null:value);
      if(json.length>900_000)throw Error('Result exceeds 900 KB bridge limit');
      result={targetId,id:job.id,ok:true,result:JSON.parse(json)};
    }catch(e){result={targetId,id:job.id,ok:false,error:e?.message||String(e)};}
    try{await request('/v1/bridge/result',result,token);}
    catch(e){onStatus(`Result delivery failed: ${e.message}`);continue;}
    onStatus(`Connected: ${metadata.name} (${targetId})`);
  }
  if(signal?.aborted) {try{await request('/v1/bridge/bye',payload,token);}catch{}}
}
