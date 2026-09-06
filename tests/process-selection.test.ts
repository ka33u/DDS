import { test } from 'node:test';
import assert from 'node:assert/strict';
import { calculateSelection, defaultConfig } from '../lib/motor-selection.ts';
import { makeProcessCase, selectCommonMotors } from '../lib/process-selection.ts';
import { buildProcessExcel } from '../lib/excel-export.ts';

test('one motor covers independent high torque and high speed recipes', () => {
  const strong = makeProcessCase({ ...defaultConfig, incomingStrength: 1600, finalSpeed: 4 }, '粗线高负载');
  const fast = makeProcessCase({ ...defaultConfig, incomingStrength: 600, finalSpeed: 12 }, '高速');
  const result = selectCommonMotors([strong, fast], { ...defaultConfig, maxOverallLength: 0 });
  assert.deepEqual(result.errors, []);
  for (const m of result.motors) {
    assert.match(m.torqueSource, /粗线高负载/);
    assert.match(m.speedSource, /高速/);
    assert.ok(m.loads.every(p => m.ratedTorque! >= p.torque && m.ratedRpm! >= p.rpm));
    assert.equal(m.ratedTorque, defaultConfig.torqueCatalog.find(t => t >= m.requiredTorque));
    assert.equal(m.ratedRpm, defaultConfig.speedCatalog.find(n => n >= m.requiredRpm));
  }
});
test('setup force and friction determine torque independently of drawing load', () => {
  const c = makeProcessCase(defaultConfig); c.jogForce = 20000; c.jogFactor = 1.5; c.additionalTorque = 100;
  const result = selectCommonMotors([c], { ...defaultConfig, maxOverallLength: 0 });
  assert.equal(result.motors[0].requiredTorque, 8500);
  assert.match(result.motors[0].torqueSource, /调模/);
});
test('physical station mapping does not combine unrelated positions', () => {
  const c = makeProcessCase(defaultConfig); c.stations = [2, 4, 6, 8, 10, 12];
  assert.deepEqual(selectCommonMotors([c], defaultConfig).motors.map(m => m.station), c.stations);
  c.stations[1] = 2;
  assert.ok(selectCommonMotors([c], defaultConfig).errors.length);
});
test('impossible catalogs, length, and different physical drums cannot produce approved model', () => {
  const c = makeProcessCase(defaultConfig);
  const noTorque = selectCommonMotors([c], { ...defaultConfig, torqueCatalog: [10] });
  assert.equal(noTorque.motors[0].model, '待确认');
  assert.equal(noTorque.motors[0].ratedTorque, null);
  assert.equal(selectCommonMotors([c], { ...defaultConfig, maxOverallLength: 100 }).motors[0].model, '待确认');
  const different = makeProcessCase(structuredClone(defaultConfig)); different.config.passes[0].drumDiameter = 600;
  assert.equal(selectCommonMotors([c, different], defaultConfig).motors[0].model, '待确认');
});
test('invalid process stops selection; disabled process does not influence envelope', () => {
  const c = makeProcessCase(defaultConfig); c.config.entryDiameter = 0;
  assert.equal(selectCommonMotors([c], defaultConfig).motors.length, 0);
  c.enabled = false;
  assert.deepEqual(selectCommonMotors([c, makeProcessCase(defaultConfig)], defaultConfig).errors, []);
});

test('export carries actual computed motor, process snapshots, and escaped customer names', () => {
  const c = makeProcessCase(defaultConfig, '客户 A & B <高速>');
  const { motors } = selectCommonMotors([c], defaultConfig);
  const xml = buildProcessExcel(defaultConfig, [c], motors);
  assert.equal((xml.match(/<Worksheet ss:Name=/g) || []).length, 5);
  assert.ok(xml.includes(motors[0].model));
  assert.ok(xml.includes('客户 A &amp; B &lt;高速&gt;'));
  assert.ok(xml.includes(String(motors[0].requiredRpm)));
  assert.ok(xml.includes('工艺输入'));
});

test('skip-head loads use destination radius and speed rather than the source drum', () => {
  const config = structuredClone(defaultConfig);
  config.skipHeadSpan = 1;
  config.passes[1].drumDiameter = 1120;
  const base = calculateSelection(config);
  const result = selectCommonMotors([makeProcessCase(config)], { ...config, maxOverallLength: 0 });
  const target = result.motors[1].loads.find(p => p.kind.startsWith('跳头'))!;
  assert.ok(target);
  assert.ok(Math.abs(target.torque - base[0].designTorque * 2) < 0.000001);
  assert.ok(Math.abs(result.motors[0].requiredRpm - base[1].lineSpeed * 60000 / (Math.PI * 560)) < 0.000001);
});
