import { calculateSelection, type SelectionConfig } from './motor-selection.ts';

export interface ProcessCase {
  id: string;
  name: string;
  enabled: boolean;
  config: SelectionConfig;
  /** Physical station for each drawing pass; allows idle/skipped stations. */
  stations: number[];
  jogSpeed: number;
  jogFactor: number;
  jogForce: number;
  additionalTorque: number;
}
export interface LoadPoint {
  caseName: string; station: number; pass: number; kind: string;
  torque: number; rpm: number; force: number; drum: number;
}
export interface CommonMotor {
  station: number; requiredTorque: number; requiredRpm: number;
  torqueSource: string; speedSource: string; ratedTorque: number | null;
  ratedRpm: number | null; power: number | null; current: number | null;
  frame: number | null; poles: number | null; length: number | null;
  frequency: number | null; model: string; status: string; loads: LoadPoint[];
}
export const makeProcessCase = (config: SelectionConfig, name = '当前工艺'): ProcessCase => ({
  id: `${Date.now()}-${Math.random().toString(36).slice(2)}`, name, enabled: true,
  config: JSON.parse(JSON.stringify(config)), stations: config.passes.map((_, i) => i + 1),
  jogSpeed: 0.5, jogFactor: 1, jogForce: 0, additionalTorque: 0,
});

export function processErrors(item: ProcessCase): string[] {
  const c = item.config;
  const errors: string[] = [];
  for (const key of ['entryDiameter', 'finalSpeed', 'incomingStrength', 'strengthFactor', 'forceCoefficient', 'firstPassFactor', 'steadyPassFactor'] as const) {
    if (!Number.isFinite(c[key]) || c[key] <= 0) errors.push(`${key} 必须为正数`);
  }
  if (!c.passes.length || c.passes.length > 16) errors.push('应配置 1～16 道工艺');
  c.passes.forEach((p, i) => {
    if (!(p.diameter > 0 && p.diameter < (i ? c.passes[i - 1].diameter : c.entryDiameter))) errors.push(`第 ${i + 1} 道丝径必须小于前道且大于零`);
    if (!Number.isFinite(p.drumDiameter) || p.drumDiameter <= 0) errors.push(`第 ${i + 1} 道卷筒直径无效`);
  });
  if (item.stations.length !== c.passes.length || item.stations.some((s, i) => !Number.isInteger(s) || s < 1 || s > 16 || (i > 0 && s <= item.stations[i - 1]))) errors.push('工位映射必须与道次一一对应，且按 1～16 的递增整数填写');
  if ([item.jogSpeed, item.jogFactor, item.jogForce, item.additionalTorque].some(x => !Number.isFinite(x) || x < 0)) errors.push('调模速度、系数、拉力和附加转矩必须是非负数');
  if (!Number.isInteger(c.skipHeadSpan) || c.skipHeadSpan < 0) errors.push('跳头覆盖道数必须为非负整数');
  return errors.map(e => `${item.name}：${e}`);
}

/** Conservative constant-torque selection: every production and setup point is continuous.
 * No unverified overload/field-weakening capability is assumed. */
export function selectCommonMotors(cases: ProcessCase[], equipment: SelectionConfig) {
  const enabled = cases.filter(c => c.enabled);
  const errors = enabled.flatMap(processErrors);
  if (!enabled.length) errors.push('至少启用一项工艺');
  for (const key of ['voltage', 'efficiency', 'powerFactor', 'voltageUtilization'] as const) {
    if (!Number.isFinite(equipment[key]) || equipment[key] <= 0 || (key !== 'voltage' && equipment[key] > 1)) errors.push(`设备电气参数 ${key} 无效`);
  }
  if (!Number.isFinite(equipment.maxOverallLength) || equipment.maxOverallLength < 0) errors.push('允许总长必须为非负数');
  if ([equipment.lowSpeedThreshold, equipment.lowSpeedMinFrame].some(x => !Number.isFinite(x) || x < 0)) errors.push('低速阈值和低速最小中心高必须为非负数');
  const validCatalog = (list: number[]) => list.length > 0 && list.every(x => Number.isFinite(x) && x > 0);
  if (!validCatalog(equipment.torqueCatalog) || !validCatalog(equipment.speedCatalog)) errors.push('转矩和转速档位库必须包含有效正数');
  if (!equipment.frameRules.length || equipment.frameRules.some(f => ![f.frame, f.maxTorque, f.estimatedLength, f.poles].every(x => Number.isFinite(x) && x > 0) || f.poles % 2 !== 0)) errors.push('机座库必须包含正数参数，极数须为正偶数');
  if (errors.length) return { errors, motors: [] as CommonMotor[] };
  const points: LoadPoint[] = [];
  for (const item of enabled) {
    const rows = calculateSelection(item.config);
    rows.forEach((r, i) => {
      const station = item.stations[i];
      const previous = rows.slice(Math.max(0, i - item.config.skipHeadSpan), i + 1);
      // A transferred drawing load acts on the destination drum radius.
      // Comparing source torques would undersize a larger destination drum.
      const covered = previous.reduce((a, b) => a.drawingForce * a.factor > b.drawingForce * b.factor ? a : b);
      points.push({ caseName: item.name, station, pass: r.pass,
        kind: covered.pass === r.pass ? '生产' : `跳头覆盖第 ${covered.pass} 道`,
        torque: covered.drawingForce * covered.factor * r.drumDiameter / 2000 + item.additionalTorque,
        rpm: item.config.skipHeadSpan > 0 && i === 0 && rows[1] ? Math.max(r.lineSpeed, rows[1].lineSpeed) * 60000 / (Math.PI * r.drumDiameter) : r.designRpm,
        force: covered.drawingForce, drum: r.drumDiameter });
      if (item.jogSpeed > 0) points.push({ caseName: item.name, station, pass: r.pass, kind: '调模',
        torque: (item.jogForce > 0 ? item.jogForce * r.drumDiameter / 2000 : r.drawingTorque) * item.jogFactor + item.additionalTorque,
        rpm: item.jogSpeed * 60000 / (Math.PI * r.drumDiameter),
        force: item.jogForce || r.drawingForce, drum: r.drumDiameter });
    });
  }
  if (points.some(p => ![p.force, p.torque, p.rpm].every(Number.isFinite))) {
    return { errors: ['输入量级过大，计算超出有效数值范围，请检查丝径、拉力和车速单位。'], motors: [] as CommonMotor[] };
  }
  const pick = (value: number, catalog: number[]) => [...catalog].sort((a, b) => a - b).find(x => x >= value) ?? null;
  const motors: CommonMotor[] = [...new Set(points.map(p => p.station))].sort((a, b) => a - b).map(station => {
    const loads = points.filter(p => p.station === station);
    const t = loads.reduce((a, b) => a.torque >= b.torque ? a : b);
    const n = loads.reduce((a, b) => a.rpm >= b.rpm ? a : b);
    const ratedTorque = pick(t.torque, equipment.torqueCatalog);
    const ratedRpm = pick(n.rpm, equipment.speedCatalog);
    const candidates = equipment.frameRules.filter(f => ratedTorque !== null && f.maxTorque >= ratedTorque
      && (ratedRpm === null || ratedRpm > equipment.lowSpeedThreshold || f.frame >= equipment.lowSpeedMinFrame)
      && (!equipment.maxOverallLength || f.estimatedLength <= equipment.maxOverallLength)).sort((a, b) => a.frame - b.frame || a.estimatedLength - b.estimatedLength);
    const frame = candidates[0];
    const drumMismatch = loads.some(p => Math.abs(p.drum - loads[0].drum) > 0.001);
    const status = drumMismatch ? '同一工位卷筒不一致，请确认设备配置' : ratedTorque === null ? '转矩超出档位库，需定制' : ratedRpm === null ? '转速超出档位库，需定制' : !frame ? '无满足转矩和总长的机座' : '满足全部已录工况';
    const power = ratedTorque !== null && ratedRpm !== null ? ratedTorque * ratedRpm / 9550 : null;
    return { station, requiredTorque: t.torque, requiredRpm: n.rpm,
      torqueSource: `${t.caseName} · ${t.kind} · 第${t.pass}道`, speedSource: `${n.caseName} · ${n.kind} · 第${n.pass}道`,
      ratedTorque, ratedRpm, power, current: power === null ? null : power * 1000 / (Math.sqrt(3) * equipment.voltage * equipment.efficiency * equipment.powerFactor * equipment.voltageUtilization),
      frame: frame?.frame ?? null, poles: frame?.poles ?? null, length: frame?.estimatedLength ?? null,
      frequency: frame && ratedRpm !== null ? ratedRpm * frame.poles / 120 : null,
      model: frame && ratedTorque !== null && ratedRpm !== null && !drumMismatch ? `${equipment.series}${frame.frame}P${frame.poles}-${ratedTorque}-${ratedRpm}-${Math.round(power!)}` : '待确认', status, loads };
  });
  return { errors, motors };
}
