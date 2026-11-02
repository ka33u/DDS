import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createId,newMatrixProcess,copyMatrixProcess} from '../lib/drawing-matrix.ts';
import {makeProjectExcel,readProjectExcel} from '../lib/project-excel.ts';

test('secure context prefers native randomUUID with the correct receiver',t=>{
  const descriptor=Object.getOwnPropertyDescriptor(globalThis,'crypto')!;
  const fake={randomUUID(){assert.equal(this,fake);return 'native-uuid';},getRandomValues(){throw new Error('fallback should not run');}};
  Object.defineProperty(globalThis,'crypto',{configurable:true,value:fake});
  t.after(()=>Object.defineProperty(globalThis,'crypto',descriptor));
  assert.equal(createId(),'native-uuid');
});

test('LAN HTTP fallback sets UUID v4 version and variant bits',t=>{
  const descriptor=Object.getOwnPropertyDescriptor(globalThis,'crypto')!;
  Object.defineProperty(globalThis,'crypto',{configurable:true,value:{getRandomValues(bytes:Uint8Array){assert.equal(bytes.length,16);return bytes.fill(255);}}});
  t.after(()=>Object.defineProperty(globalThis,'crypto',descriptor));
  assert.equal(createId(),'ffffffff-ffff-4fff-bfff-ffffffffffff');
});

test('add, copy and Excel import work without randomUUID and preserve independent processes',async t=>{
  const descriptor=Object.getOwnPropertyDescriptor(globalThis,'crypto')!;
  const native=globalThis.crypto;
  Object.defineProperty(globalThis,'crypto',{configurable:true,value:{getRandomValues:native.getRandomValues.bind(native)}});
  t.after(()=>Object.defineProperty(globalThis,'crypto',descriptor));
  const ids=Array.from({length:100},()=>createId());
  assert.equal(new Set(ids).size,100);
  ids.forEach(id=>assert.match(id,/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/));
  const original=newMatrixProcess(2,1);
  Object.assign(original,{entry:'5.5',strength:'1100',speed:'8',dies:['5','4.5'],drums:['560','560']});
  const copy=copyMatrixProcess(original);
  assert.notEqual(copy.id,original.id);assert.equal(copy.name,'工艺 1 副本');
  copy.dies[0]='5.1';copy.drums[0]='600';assert.equal(original.dies[0],'5');assert.equal(original.drums[0],'560');
  const data={info:{projectName:'HTTP测试',customerName:'测试客户'},settings:{count:2,first:'1.4',other:'1.3'},processes:[original,copy]};
  const imported=await readProjectExcel(await makeProjectExcel(data));
  assert.equal(imported.processes.length,2);
  assert.equal(new Set([...data.processes,...imported.processes].map(p=>p.id)).size,4);
});
