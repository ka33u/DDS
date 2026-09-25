import { test } from 'node:test';
import assert from 'node:assert/strict';
import { recipes, motorReferences, loadRecipe } from '../lib/reference-data.ts';
import { defaultConfig, calculateDrawingRequirements } from '../lib/motor-selection.ts';
import { makeProcessCase, selectCommonMotors } from '../lib/process-selection.ts';
import { buildProcessExcel } from '../lib/excel-export.ts';

test('eight image recipes preserve units, die counts and explicit conflict',()=>{
  assert.equal(recipes.length,8); assert.equal(motorReferences.length,52);
  assert.equal(motorReferences.filter(m=>m.rawEfficiency>1).length,9);
  for(const r of recipes){const c=loadRecipe(defaultConfig,r);assert.equal(c.finalSpeed*60,r.design);assert.equal(c.passes.length,Number(r.id[2]));assert.equal(c.firstPassFactor,1.4);assert.equal(c.steadyPassFactor,1.3);}
  assert.equal(recipes[7].dies.at(-1),.65);assert.match(recipes[7].warning!,/0.63/);
  assert.deepEqual(motorReferences.filter(m=>m.sheet==='LZ7-350选型').map(m=>m.poles),[30,30,30,30,8,8,8]);
});
test('reference results are whole source tuples, cover all loads and do not fabricate length',()=>{
  for(const r of recipes){
    const c=loadRecipe(defaultConfig,r);const result=selectCommonMotors([makeProcessCase(c)],c);assert.deepEqual(result.errors,[]);
    for(const m of result.motors){assert.equal(m.length,null);if(m.ratedTorque!==null){assert.ok(motorReferences.some(x=>x.torque===m.ratedTorque&&x.rpm===m.ratedRpm&&x.frame===m.frame&&x.poles===m.poles));assert.ok(m.loads.every(p=>p.torque<=m.ratedTorque!&&p.rpm<=m.ratedRpm!));assert.ok(m.current!>0);}}
    assert.ok(selectCommonMotors([makeProcessCase(c)],{...c,maxOverallLength:1450}).motors.every(m=>m.ratedTorque===null));
  }
});
test('source rated speed below design is not excused by rounding, and export preserves anomalies',()=>{
  const c=loadRecipe(defaultConfig,recipes[0]);assert.ok(calculateDrawingRequirements(c).at(-1)!.designRpm>125);
  const xml=buildProcessExcel(c,[makeProcessCase(c)],selectCommonMotors([makeProcessCase(c)],c).motors);
  assert.match(xml,/附件历史参数/);assert.match(xml,/原表效率/);assert.match(xml,/0.63/);
});
