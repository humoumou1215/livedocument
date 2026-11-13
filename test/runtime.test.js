import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {createRuntime} from '../src/runtime.js';
import {wpsMetadata,wpsExecute} from '../src/bridge/wps.js';
import {officeMetadata,officeExecute} from '../src/bridge/office.js';

async function fixture(t) {
  const dir=await mkdtemp(path.join(os.tmpdir(),'livedoc-test-'));
  const env={LIVEDOC_HOME:dir};
  const runtime=createRuntime({port:0,env});
  const {port,token}=await runtime.start();
  const base=`http://127.0.0.1:${port}`;
  const req=async(route,data,auth=token,extra={})=>{
    const r=await fetch(base+route,{method:data===undefined?'GET':'POST',headers:{...(auth?{Authorization:`Bearer ${auth}`}:{}),'Content-Type':'application/json',...extra},body:data===undefined?undefined:JSON.stringify(data)});
    const result=r.status===204?null:await r.json();return {status:r.status,data:result};
  };
  t.after(async()=>{await runtime.stop();await rm(dir,{recursive:true,force:true});});
  const pair=async(host='wps',kind='spreadsheet')=>{
    const c=await req('/v1/pair-code',{});
    assert.equal(c.status,200);
    const p=await req('/v1/pair',{code:c.data.code,host,kind,name:'example.xlsx'},null);
    assert.equal(p.status,200);return p.data;
  };
  return {req,pair,env,port,base};
}

test('unauthorized access, CSRF origin and pairing validation',async t=>{
  const {req}=await fixture(t);
  assert.equal((await req('/v1/targets',undefined,'bad')).status,401);
  assert.equal((await req('/v1/status',undefined,undefined,{Origin:'https://evil.example'})).status,403);
  assert.equal((await req('/v1/pair',{code:'00000000',host:'wps',kind:'spreadsheet'},null)).status,403);
  const code=(await req('/v1/pair-code',{})).data.code;
  assert.equal((await req('/v1/pair',{code,host:'wps',kind:'spreadsheet',name:'doc'},null)).status,200);
  assert.equal((await req('/v1/pair',{code,host:'wps',kind:'spreadsheet',name:'doc'},null)).status,403);
  assert.equal((await req('/v1/bridge/next',{targetId:'not-valid'},'bogus')).status,401);
});

test('full pairing -> target discovery -> execution -> result -> disconnect',async t=>{
  const {req,pair}=await fixture(t);
  const bridge=await pair(); const {targetId,bridgeToken}=bridge;
  const ts=(await req('/v1/targets')).data.targets;
  assert.equal(ts.length,1);assert.equal(ts[0].id,targetId);
  const pending=req('/v1/bridge/next',{targetId},bridgeToken);
  const executing=req('/v1/execute',{targetId,code:'return Application.Selection.Address',timeoutMs:3000});
  const message=await pending;assert.equal(message.status,200);
  assert.match(message.data.job.code,/Selection/);
  assert.equal((await req('/v1/bridge/result',{targetId,id:message.data.job.id,ok:true,result:'$A$1'},bridgeToken)).status,200);
  const result=await executing;assert.deepEqual(result.data,{ok:true,result:'$A$1'});
  assert.equal((await req('/v1/bridge/bye',{targetId},bridgeToken)).status,200);
  assert.equal((await req('/v1/targets')).data.targets.length,0);
});

test('queue is serialized and invalid bridge response cannot steal another job',async t=>{
  const {req,pair}=await fixture(t);const {targetId,bridgeToken}=await pair();
  const j1=req('/v1/execute',{targetId,code:'return 1',timeoutMs:4000});
  const j2=req('/v1/execute',{targetId,code:'return 2',timeoutMs:4000});
  const next1=(await req('/v1/bridge/next',{targetId},bridgeToken)).data.job;
  assert.equal(next1.code,'return 1');
  assert.equal((await req('/v1/bridge/result',{targetId,id:'wrong',ok:true,result:'bad'},bridgeToken)).status,409);
  const nextPending=req('/v1/bridge/next',{targetId},bridgeToken);
  await new Promise(r=>setTimeout(r,25));
  assert.equal((await req('/v1/bridge/result',{targetId,id:next1.id,ok:true,result:1},bridgeToken)).status,200);
  const next2=(await nextPending).data.job;assert.equal(next2.code,'return 2');
  await req('/v1/bridge/result',{targetId,id:next2.id,ok:false,error:'nope'},bridgeToken);
  assert.deepEqual((await j1).data,{ok:true,result:1});
  assert.equal((await j2).data.error.code,'BRIDGE_EXECUTION');
});

test('timed out execution disconnects target instead of dispatching overlap',async t=>{
  const {req,pair}=await fixture(t);const {targetId,bridgeToken}=await pair();
  const running=req('/v1/execute',{targetId,code:'return 1',timeoutMs:1000});
  const first=(await req('/v1/bridge/next',{targetId},bridgeToken)).data.job;
  assert.ok(first.id);
  const queued=req('/v1/execute',{targetId,code:'return 2',timeoutMs:4000});
  assert.equal((await running).status,504);
  assert.equal((await queued).status,503);
  assert.equal((await req('/v1/targets')).data.targets.length,0);
});

test('WPS adapter executes against native object',async()=>{
  const app={ActiveWorkbook:{Name:'data.xlsx'},Selection:{Address:'A1'}};
  assert.deepEqual(wpsMetadata(app),{host:'wps',kind:'spreadsheet',name:'data.xlsx'});
  assert.equal(await wpsExecute(app,'return Application.Selection.Address'),'A1');
  await wpsExecute(app,"Application.Selection.Address='B2'");assert.equal(app.Selection.Address,'B2');
  await assert.rejects(wpsExecute(app,'throw Error("failure")'),/failure/);
});

test('Office adapter passes native request context and syncs',async()=>{
  const ctx={syncCount:0,async sync(){this.syncCount++;},workbook:{selected:9}};
  const excel={run:async fn=>fn(ctx)};
  assert.deepEqual(officeMetadata({context:{host:'Excel',document:{url:'C:\\docs\\book.xlsx'}}}),{host:'office',kind:'spreadsheet',name:'book.xlsx'});
  assert.equal(await officeExecute({Excel:excel},'return context.workbook.selected','spreadsheet'),9);
  assert.equal(ctx.syncCount,1);
});

test('CLI status, targets and exec through real child process',async t=>{
  const {req,pair,env}=await fixture(t);const {targetId,bridgeToken}=await pair();
  async function run(args) {
    const output=[];const error=[];
    const child=spawn(process.execPath,[fileURLToPath(new URL('../src/cli.js',import.meta.url)),...args],{env:{...process.env,...env}});
    for await(const d of child.stdout)output.push(d);
    for await(const d of child.stderr)error.push(d);
    const code=await new Promise(resolve=>child.on('close',resolve));
    return {code,out:output.join(''),err:error.join('')};
  }
  assert.equal(JSON.parse((await run(['status'])).out).ok,true);
  assert.equal(JSON.parse((await run(['targets'])).out).targets[0].id,targetId);
  const poll=req('/v1/bridge/next',{targetId},bridgeToken);
  const client=run(['exec','--target',targetId,'--code','return 42']);
  const job=(await poll).data.job;
  await req('/v1/bridge/result',{targetId,id:job.id,ok:true,result:42},bridgeToken);
  assert.equal(JSON.parse((await client).out).result,42);
  const fail=await run(['exec','--target','bogus','--code','return 1']);
  assert.equal(fail.code,1);assert.match(fail.err,/TARGET_NOT_FOUND/);
});
