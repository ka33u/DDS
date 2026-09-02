import assert from "node:assert/strict";
import test from "node:test";

import {
  applyPreset,
  calculateSelection,
  defaultConfig,
  equalCompressionPasses,
} from "../lib/motor-selection.ts";

void test("附件 LZ7 工况复算与原表关键转矩一致", () => {
  const config = {
    ...defaultConfig,
    preset: "custom" as const,
    firstPassFactor: 1.4,
    steadyPassFactor: 1.2,
    skipHeadSpan: 0,
  };
  const rows = calculateSelection(config);
  const expected = [3168.05, 2629.05, 2170.69, 1746.1, 1531.79, 1287.78];
  assert.equal(rows.length, 6);
  rows.forEach((row, index) => assert.ok(Math.abs(row.designTorque - expected[index]) < 0.08));
  assert.ok(Math.abs(rows[0].designRpm - 113.0023) < 0.01);
  assert.ok(Math.abs(rows[5].designRpm - 272.837) < 0.01);
});

void test("万能机覆盖前一道更大转矩并使用 1.4/1.2", () => {
  const config = applyPreset(defaultConfig, "universal");
  const rows = calculateSelection(config);
  assert.equal(config.firstPassFactor, 1.4);
  assert.equal(config.steadyPassFactor, 1.2);
  assert.equal(rows[1].reserveSourcePass, 1);
  assert.equal(rows[1].ratedTorque, 3200);
  assert.equal(rows[2].ratedTorque, 2650);
});

void test("均匀压缩助手保持首末边界并支持广道次数", () => {
  const passes = equalCompressionPasses(12, 8, 2, 600);
  assert.equal(passes.length, 12);
  assert.equal(passes.at(-1)?.diameter, 2);
  assert.ok(passes.every((pass, index) => index === 0 || pass.diameter < passes[index - 1].diameter));
});
