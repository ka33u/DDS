import {test} from 'node:test';
import assert from 'node:assert/strict';
import {newGeneralConfig,applyProcessBatch} from '../lib/general-process.ts';
import {calculateDrawingRequirements,equalCompressionPasses} from '../lib/motor-selection.ts';
import {makeProcessCase,calculateProcessLoads,selectCommonMotors} from '../lib/process-selection.ts';
import {buildProcessExcel} from '../lib/excel-export.ts';

test('128-pass fine-wire calculations do not truncate or round small dies to zero',()=>{
  const c={...newGeneralConfig(),entryDiameter:.2,finalDiameter:.01,equipmentDrums:undefined,processStations:undefined,passes:equalCompressionPasses(128,.2,.01,200)};
  assert.equal(c.passes.length,128);assert.ok(c.passes.every(p=>p.diameter>0));
  assert.deepEqual(calculateProcessLoads([makeProcessCase(c)]).errors,[]);
  assert.equal(calculateDrawingRequirements(c).length,128);
});
test('different die counts use same physical device without combining unrelated devices',()=>{
  const c=newGeneralConfig();const a=makeProcessCase(c,'八模');
  const b=makeProcessCase(applyProcessBatch(c,'4 3 2','2 5 8','',''),'三模');
  const result=calculateProcessLoads([a,b],c);assert.deepEqual(result.errors,[]);
  assert.equal(result.points.filter(p=>p.caseName==='三模'&&p.kind==='生产').length,3);
  b.config.deviceName='另一台设备';assert.match(calculateProcessLoads([a,b],c).errors.join(),/不同设备/);
});
test('customer force, independently entered speeds, multiwire and mechanical efficiency are propagated',()=>{
  const c=applyProcessBatch({...newGeneralConfig(),forceModel:'measured',parallelWires:2,drawingEfficiency:.8},'4 2','1 8','100 200','60 120');
  const rows=calculateDrawingRequirements(c);assert.equal(rows[0].drawingForce,200);assert.equal(rows[0].drawingTorque,62.5);assert.equal(rows[1].lineSpeed,2);
  const p=calculateProcessLoads([makeProcessCase(c)],c);assert.deepEqual(p.errors,[]);assert.equal(p.points[0].torque,87.5);
  c.passes[0].measuredForce=undefined;assert.match(calculateProcessLoads([makeProcessCase(c)],c).errors.join(),/缺少客户拉拔力/);
});
test('invalid pasted data, missing stations and drum mismatches are not silently repaired',()=>{
  const c=newGeneralConfig();assert.throws(()=>applyProcessBatch(c,'4 bad 2','1 2 3','',''));
  assert.throws(()=>applyProcessBatch(c,'4 2','1 9','',''));
  const p=makeProcessCase(c);p.config.passes[0].drumDiameter=400;
  assert.match(calculateProcessLoads([p],c).errors.join(),/卷筒与设备不一致/);
  assert.equal(selectCommonMotors([p],c).motors.length,0);
});
test('material exponent and input model affect force, and export carries assumptions',()=>{
  const c=newGeneralConfig();const average=calculateDrawingRequirements(c)[0].drawingForce;
  c.forceModel='incoming';assert.ok(calculateDrawingRequirements(c)[0].drawingForce<average);
  c.strengthExponent=0;assert.equal(calculateDrawingRequirements(c)[3].strengthAfter,c.incomingStrength);
  const xml=buildProcessExcel(c,[makeProcessCase(c)],[]);assert.match(xml,/机械效率/);assert.match(xml,/固定工位卷筒/);assert.match(xml,/incoming/);
});
