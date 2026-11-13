import http from 'node:http';
import https from 'node:https';
import fs from 'node:fs/promises';
import {randomBytes, randomInt, timingSafeEqual} from 'node:crypto';
import {readFile} from 'node:fs/promises';
import {DEFAULT_PORT, saveState, deleteStateIfOwned} from './settings.js';

const MAX_REQUEST = 256 * 1024;
const MAX_CODE = 128 * 1024;
const MAX_RESULT = 1024 * 1024;
const STALE_MS = 45_000;
const PAIR_TTL = 120_000;
const POLL_MS = 15_000;
const rand = n => randomBytes(n).toString('hex');
const equal = (a,b) => typeof a === 'string' && a.length === b.length && timingSafeEqual(Buffer.from(a), Buffer.from(b));
const fail = (status, message, code = 'ERROR') => Object.assign(new Error(message), {status, code});

async function body(req, limit = MAX_REQUEST) {
  const chunks=[]; let length=0;
  for await (const c of req) { length+=c.length; if (length>limit) throw fail(413,'Request too large','TOO_LARGE'); chunks.push(c); }
  try {return JSON.parse(Buffer.concat(chunks).toString('utf8'));} catch {throw fail(400,'Invalid JSON','BAD_JSON');}
}
function respond(res, code, data) {
  if(res.writableEnded || res.destroyed) return;
  res.writeHead(code, {'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});
  res.end(JSON.stringify(data));
}
function error(res, err) {respond(res,err.status||500,{ok:false,error:{code:err.code||'INTERNAL',message:err.status ? err.message : 'Internal runtime error'}});}
function auth(req, token) { if (!equal(req.headers.authorization?.replace(/^Bearer /,'')||'', token)) throw fail(401,'Unauthorized','UNAUTHORIZED'); }
function validateText(input,max,label) {if(typeof input!=='string'||!input.trim()||input.length>max) throw fail(400,`Invalid ${label}`,'VALIDATION');return input;}

export function createRuntime({port=DEFAULT_PORT, env=process.env, host='127.0.0.1', tls, writeState=true, onStop}={}) {
  const cliToken=rand(32), targets=new Map(), codes=new Map(), attempts=new Map();
  let server, stopping=false, address;
  function drop(t,reason='TARGET_DISCONNECTED') {
    targets.delete(t.id);
    for(const job of t.queue) job.finish(503,{ok:false,error:{code:reason,message:'Target disconnected'}});
    if(t.active) t.active.finish(503,{ok:false,error:{code:reason,message:'Target disconnected'}});
    for(const poll of t.waiters.splice(0)) poll(null);
  }
  function prune() {
    const now=Date.now();
    for(const t of targets.values()) if(now-t.lastSeen>STALE_MS && !t.active) drop(t);
    for(const [k,v] of codes) if(v<now) codes.delete(k);
  }
  function dispatch(t) {
    if(t.active || !t.queue.length || !t.waiters.length)return;
    t.active=t.queue.shift();
    const waiter=t.waiters.shift();waiter({id:t.active.id,code:t.active.code,timeoutMs:t.active.timeoutMs});
  }
  function targetAuth(req,data) {
    const t=targets.get(data.targetId);
    if(!t || !equal(req.headers.authorization?.replace(/^Bearer /,'')||'',t.bridgeToken)) throw fail(401,'Invalid bridge credentials','UNAUTHORIZED');
    t.lastSeen=Date.now();return t;
  }
  function resultFor(t,data) {
    const job=t.active;
    if(!job||job.id!==data.id)throw fail(409,'Unknown or expired job','JOB_MISMATCH');
    if(typeof data.ok!=='boolean')throw fail(400,'Missing result status','VALIDATION');
    t.active=null;
    if(data.ok) job.finish(200,{ok:true,result:data.result??null});
    else job.finish(200,{ok:false,error:{code:'BRIDGE_EXECUTION',message:String(data.error||'Execution failed').slice(0,4000)}});
    dispatch(t);
  }
  async function request(req,res) {
    const u=new URL(req.url,'http://127.0.0.1');
    const origin=req.headers.origin;
    const expectedPort=address?.port||port;
    const h=req.headers.host || '';
    if(![`127.0.0.1:${expectedPort}`,`localhost:${expectedPort}`].includes(h))throw fail(403,'Invalid host','BAD_HOST');
    const protocol=tls?'https':'http';
    const allowed=[`${protocol}://127.0.0.1:${expectedPort}`,`${protocol}://localhost:${expectedPort}`];
    if(origin && !allowed.includes(origin))throw fail(403,'Untrusted browser origin','BAD_ORIGIN');
    if(origin) {res.setHeader('Access-Control-Allow-Origin',origin);res.setHeader('Vary','Origin');res.setHeader('Access-Control-Allow-Headers','authorization, content-type');res.setHeader('Access-Control-Allow-Methods','GET, POST, OPTIONS');}
    if(req.method==='OPTIONS'){res.writeHead(204);res.end();return;}
    if(req.method==='GET'&&u.pathname.startsWith('/bridge/')) {
      const map={
        '/bridge/wps.html':['./bridge/panel.html','text/html; charset=utf-8'],
        '/bridge/office.html':['./bridge/office.html','text/html; charset=utf-8'],
        '/bridge/client.js':['./bridge/client.js','text/javascript; charset=utf-8'],
        '/bridge/wps.js':['./bridge/wps.js','text/javascript; charset=utf-8'],
        '/bridge/office.js':['./bridge/office.js','text/javascript; charset=utf-8'],
        '/bridge/panel.css':['./bridge/panel.css','text/css; charset=utf-8'],
        '/bridge/panel.js':['./bridge/panel.js','text/javascript; charset=utf-8'],
        '/bridge/icon.png':['./bridge/icon.png','image/png']};
      const spec=map[u.pathname];if(!spec)throw fail(404,'Not found','NOT_FOUND');
      const content=await readFile(new URL(spec[0],import.meta.url));
      res.writeHead(200,{'Content-Type':spec[1], 'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Content-Security-Policy': "default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com 'unsafe-eval'; style-src 'self'; connect-src 'self'; img-src 'self'"});
      res.end(content);return;
    }
    if(req.method==='GET'&&u.pathname.startsWith('/wps-addon/')) {
      const map={'/wps-addon/ribbon.xml':'../adapters/wps/ribbon.xml','/wps-addon/main.js':'../adapters/wps/main.js','/wps-addon/index.html':'../adapters/wps/index.html'};
      if(!map[u.pathname])throw fail(404,'Not found','NOT_FOUND');
      const content=await readFile(new URL(map[u.pathname],import.meta.url));
      res.writeHead(200,{'Content-Type':u.pathname.endsWith('.xml')?'application/xml':u.pathname.endsWith('.html')?'text/html':'text/javascript','Cache-Control':'no-store'});res.end(content);return;
    }
    if(req.method==='GET'&&u.pathname==='/v1/status') {auth(req,cliToken);prune();respond(res,200,{ok:true,pid:process.pid,port:expectedPort,targets:targets.size,version:'0.1.0'});return;}
    if(req.method==='POST'&&u.pathname==='/v1/stop') {auth(req,cliToken);respond(res,200,{ok:true});setImmediate(()=>stop());return;}
    if(req.method==='POST'&&u.pathname==='/v1/pair-code') {auth(req,cliToken);const code=String(randomInt(10_000_000,100_000_000));codes.set(code,Date.now()+PAIR_TTL);respond(res,200,{ok:true,code,expiresInSeconds:120});return;}
    if(req.method==='POST'&&u.pathname==='/v1/pair') {
      const data=await body(req);const who=req.socket.remoteAddress||'local';
      const a=attempts.get(who)||{count:0,since:Date.now()};if(Date.now()-a.since>60_000){a.count=0;a.since=Date.now();}
      a.count++;attempts.set(who,a);if(a.count>10)throw fail(429,'Too many pairing attempts','RATE_LIMIT');
      if(typeof data.code!=='string'||(codes.get(data.code)||0)<Date.now())throw fail(403,'Invalid or expired pairing code','BAD_PAIR_CODE');
      const hostType=validateText(data.host,20,'host');if(!['wps','office'].includes(hostType))throw fail(400,'Unknown host','VALIDATION');
      const kind=validateText(data.kind,32,'kind');if(!['spreadsheet','document','presentation'].includes(kind))throw fail(400,'Unknown document kind','VALIDATION');
      codes.delete(data.code);
      const id=rand(8),bridgeToken=rand(32);
      targets.set(id,{id,bridgeToken,host:hostType,kind,name:String(data.name||'Untitled').slice(0,200),lastSeen:Date.now(),queue:[],active:null,waiters:[]});
      respond(res,200,{ok:true,targetId:id,bridgeToken});return;
    }
    if(req.method==='GET'&&u.pathname==='/v1/targets') {auth(req,cliToken);prune();respond(res,200,{ok:true,targets:[...targets.values()].map(({id,host,kind,name,lastSeen})=>({id,host,kind,name,lastSeen:new Date(lastSeen).toISOString()}))});return;}
    if(req.method==='POST'&&u.pathname==='/v1/execute') {
      auth(req,cliToken);prune();const data=await body(req);
      const t=targets.get(data.targetId);if(!t)throw fail(404,'Target not found or disconnected','TARGET_NOT_FOUND');
      const code=validateText(data.code,MAX_CODE,'code');
      const timeoutMs=Math.min(120_000,Math.max(1000,Number(data.timeoutMs)||30_000));
      const id=rand(12);
      let done=false, timer;
      const finish=(status,value)=>{if(done)return;done=true;clearTimeout(timer);respond(res,status,value);};
      const job={id,code,timeoutMs,finish};
      timer=setTimeout(()=>{
        const wasActive=t.active===job;
        if(!wasActive)t.queue=t.queue.filter(j=>j!==job);
        finish(504,{ok:false,error:{code:'TIMEOUT',message:'Execution timed out (the host may still be running the script)'}});
        // Never dispatch a second write while the timed-out job may still run.
        if(wasActive)drop(t,'EXECUTION_TIMEOUT');else dispatch(t);
      },timeoutMs);
      // If caller disconnects before completion, remove this queued work.
      res.once('close',()=>{if(done)return;const wasActive=t.active===job;if(!wasActive)t.queue=t.queue.filter(j=>j!==job);done=true;clearTimeout(timer);if(wasActive)drop(t,'CALLER_DISCONNECTED');else dispatch(t);});
      t.queue.push(job);dispatch(t);return;
    }
    if(req.method==='POST'&&u.pathname==='/v1/bridge/next') {
      const data=await body(req);const t=targetAuth(req,data);if(t.waiters.length>=1)throw fail(409,'Too many pending polls','POLL_CONFLICT');
      let done=false;let timer;
      const release=(job)=>{if(done)return;done=true;clearTimeout(timer);t.waiters=t.waiters.filter(w=>w!==release);if(job)respond(res,200,{ok:true,job});else {res.writeHead(204,{'Cache-Control':'no-store'});res.end();}};
      timer=setTimeout(()=>release(null),POLL_MS);
      t.waiters.push(release);res.once('close',()=>{if(done)return;done=true;clearTimeout(timer);t.waiters=t.waiters.filter(w=>w!==release);});dispatch(t);return;
    }
    if(req.method==='POST'&&u.pathname==='/v1/bridge/result') {const data=await body(req,MAX_RESULT);const t=targetAuth(req,data);resultFor(t,data);respond(res,200,{ok:true});return;}
    if(req.method==='POST'&&u.pathname==='/v1/bridge/bye') {const data=await body(req);const t=targetAuth(req,data);drop(t);respond(res,200,{ok:true});return;}
    throw fail(404,'Not found','NOT_FOUND');
  }
  async function start() {
    if(server)throw new Error('Runtime already running');
    const handler=(req,res)=>{request(req,res).catch(e=>{if(!e.status)console.error('[livedoc runtime]',e);error(res,e);});};
    server=tls?https.createServer({key:await fs.readFile(tls.key),cert:await fs.readFile(tls.cert)},handler):http.createServer(handler);
    try {await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,host,()=>{server.off('error',reject);resolve();});});}
    catch(e){server=null;throw e;}
    address=server.address();
    if(writeState)await saveState({pid:process.pid,port:address.port,token:cliToken,https:!!tls},env);
    return {port:address.port,token:cliToken,url:`${tls?'https':'http'}://127.0.0.1:${address.port}`};
  }
  async function stop() {
    if(!server||stopping)return;
    stopping=true;
    for(const t of [...targets.values()])drop(t,'RUNTIME_STOPPING');
    server.closeAllConnections();await new Promise(resolve=>server.close(resolve));
    if(writeState)await deleteStateIfOwned(cliToken,env);
    server=null;onStop?.();
  }
  return {start,stop};
}
