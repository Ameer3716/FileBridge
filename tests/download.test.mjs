import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
import ts from 'typescript';
import * as files from '../lib/files.mjs';

// Execute the real route with only the remote storage boundary replaced.
const source=await readFile(new URL('../app/api/file/route.ts',import.meta.url),'utf8');
const {outputText}=ts.transpileModule(source,{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2022}});
function route(storage){
 const exports={};
 vm.runInNewContext(outputText,{exports,Response,URL,require(name){
  if(name==='@vercel/blob')return storage;
  if(name==='@/lib/files.mjs')return files;
  throw Error(`Unexpected import: ${name}`);
 }});
 return exports;
}
const id='12345678-1234-1234-1234-123456789abc';
const request=()=>new Request(`https://files.example/api/file?id=${id}`);
for(const name of ['report.xlsx','Sales report.xlsx','Sales پاکستان.xlsx','100% complete.xlsx']){
 test(`download uses canonical storage URL and preserves bytes: ${name}`,async()=>{
  const pathname=files.makePath(id,123,'a'.repeat(64),name);
  const url=`https://example.private.blob.vercel-storage.com/${pathname.split('/').map(encodeURIComponent).join('/')}`;
  const bytes=new Uint8Array([0,1,127,128,255]);
  const handlers=route({
   async list(options){assert.equal(options.prefix,`files/${id}/`);return {blobs:[{pathname,url,size:bytes.length,uploadedAt:new Date('2026-10-27T20:00:00Z')}]};},
   async get(address,options){
    assert.equal(address,url);
    assert.equal(options.access,'private');
    assert.equal(options.useCache,false);
    return {statusCode:200,stream:new Response(bytes).body};
   },
  });
  const response=await handlers.GET(request());
  assert.equal(response.status,200,await (response.ok?Promise.resolve(''):response.clone().text()));
  assert.deepEqual(new Uint8Array(await response.arrayBuffer()),bytes);
  const filename=decodeURIComponent(response.headers.get('Content-Disposition').split("filename*=UTF-8''")[1]);
  assert.equal(filename,name);
 });
}
test('missing upload returns 404 without fetching storage contents',async()=>{
 const handlers=route({async list(){return {blobs:[]};},async get(){assert.fail('Unexpected download');}});
 assert.equal((await handlers.GET(request())).status,404);
});
test('deleted Blob contents return 404',async()=>{
 const pathname=files.makePath(id,0,'a'.repeat(64),'report.xlsx');
 const handlers=route({async list(){return {blobs:[{pathname,url:'https://example.private.blob.vercel-storage.com/file',size:1,uploadedAt:new Date()}]};},async get(){return null;}});
 assert.equal((await handlers.GET(request())).status,404);
});
