#!/usr/bin/env node
import fs from 'node:fs/promises';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {createRuntime} from './runtime.js';
import {loadState, DEFAULT_PORT} from './settings.js';

const usage=`Live Document CLI (Node.js 20+)
Usage:
  livedoc start                 Start local runtime in background
  livedoc serve                 Run runtime in foreground
  livedoc stop                  Stop local runtime
  livedoc status                Runtime status
  livedoc pair                  Generate one-time pairing code for Office bridge
  livedoc targets               List active Office contexts
  livedoc exec --target ID --code 'return Application.Selection.Address'
  livedoc exec --target ID --file script.js [--timeout 30000]
  livedoc exec --target ID --stdin [--timeout 30000]

Config: LIVEDOC_PORT (default 18765), LIVEDOC_HOME, LIVEDOC_TLS_CERT and LIVEDOC_TLS_KEY.
Output: JSON on stdout; errors on stderr. Code is executed in the paired host, not in Node.
`;
const option=(args,key)=>{const i=args.indexOf(key);return i>=0?args[i+1]:undefined;};
async function readStdin(){let s='';for await(const c of process.stdin){s+=c;if(s.length>128*1024)throw Error('Input too large');}return s;}
async function connect(path,method='GET',data) {
  const s=await loadState();const url=`${s.https?'https':'http'}://127.0.0.1:${s.port}${path}`;
  const response=await fetch(url,{method,headers:{Authorization:`Bearer ${s.token}`,'Content-Type':'application/json'},body:data?JSON.stringify(data):undefined,signal:AbortSignal.timeout(135_000)});
  const result=await response.json();if(!response.ok||!result.ok){const e=new Error(result.error?.message||`HTTP ${response.status}`);e.code=result.error?.code;throw e;}return result;
}
async function main(){
  const [cmd,...args]=process.argv.slice(2);
  if(!cmd||['help','-h','--help'].includes(cmd)){process.stdout.write(usage);return;}
  if(cmd==='serve') {
    const cert=process.env.LIVEDOC_TLS_CERT,key=process.env.LIVEDOC_TLS_KEY;
    if(Boolean(cert)!==Boolean(key))throw Error('LIVEDOC_TLS_CERT and LIVEDOC_TLS_KEY must both be provided');
    const r=createRuntime({port:Number(process.env.LIVEDOC_PORT||DEFAULT_PORT),tls:cert?{cert,key}:undefined});
    const s=await r.start();console.log(JSON.stringify({ok:true,pid:process.pid,url:s.url}));
    const shutdown=()=>r.stop().then(()=>process.exit(0)).catch(e=>{console.error(e);process.exit(1);});
    process.once('SIGINT',shutdown);process.once('SIGTERM',shutdown);return;
  }
  if(cmd==='start') {
    try{const s=await connect('/v1/status');console.log(JSON.stringify({...s,alreadyRunning:true}));return;}catch{}
    const child=spawn(process.execPath,[fileURLToPath(new URL('./cli.js',import.meta.url)),'serve'],{detached:true,stdio:'ignore',env:process.env});child.unref();
    for(let i=0;i<40;i++) {await new Promise(r=>setTimeout(r,100));try{const s=await connect('/v1/status');console.log(JSON.stringify({...s,url:`${process.env.LIVEDOC_TLS_CERT?'https':'http'}://127.0.0.1:${s.port}`}));return;}catch{}}
    throw Error('Runtime did not start; run `livedoc serve` to inspect the error');
  }
  if(cmd==='stop'){console.log(JSON.stringify(await connect('/v1/stop','POST',{})));return;}
  if(cmd==='status'){console.log(JSON.stringify(await connect('/v1/status')));return;}
  if(cmd==='pair'){console.log(JSON.stringify(await connect('/v1/pair-code','POST',{})));return;}
  if(cmd==='targets'){console.log(JSON.stringify(await connect('/v1/targets')));return;}
  if(cmd==='exec') {
    const targetId=option(args,'--target'), codeFlag=option(args,'--code'), file=option(args,'--file'), stdin=args.includes('--stdin');
    const count=Number(codeFlag!==undefined)+Number(file!==undefined)+Number(stdin);
    if(!targetId||count!==1)throw Error('Provide --target and exactly one of --code, --file or --stdin');
    const code=codeFlag??(file?await fs.readFile(file,'utf8'):await readStdin());
    const timeoutMs=Number(option(args,'--timeout')||30_000);
    if(!Number.isFinite(timeoutMs)||timeoutMs<1000||timeoutMs>120000)throw Error('Timeout must be 1000..120000 ms');
    const result=await connect('/v1/execute','POST',{targetId,code,timeoutMs});console.log(JSON.stringify(result));return;
  }
  throw Error(`Unknown command ${cmd}. Run livedoc help.`);
}
main().catch(e=>{console.error(JSON.stringify({ok:false,error:{code:e.code||'CLI_ERROR',message:e.message}}));process.exitCode=1;});
