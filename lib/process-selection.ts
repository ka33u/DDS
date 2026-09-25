import { calculateDrawingRequirements, type SelectionConfig } from './motor-selection.ts';
import { motorReferences } from './reference-data.ts';

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
  reference?: string;
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
    if (p.measuredForce !== undefined && (!Number.isFinite(p.measuredForce) || p.measuredForce < 0)) errors.push(`第 ${i + 1} 道实测拉拔力必须为非负数`);
    if (p.safetyFactor !== undefined && (!Number.isFinite(p.safetyFactor) || p.safetyFactor <= 0)) errors.push(`第 ${i + 1} 道独立安全系数必须为正数`);
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
export function calculateProcessLoads(cases: ProcessCase[]) {
  const enabled = cases.filter(c => c.enabled);
  const errors = enabled.flatMap(processErrors);
  if (!enabled.length) errors.push('至少启用一项工艺');
  if (errors.length) return { errors, points: [] as LoadPoint[] };
  const points: LoadPoint[] = [];
  for (const item of enabled) {
    const rows = calculateDrawingRequirements(item.config);
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
    return { errors: ['输入量级过大，计算超出有效数值范围，请检查丝径、拉力和车速单位。'], points: [] as LoadPoint[] };
  }
  return { errors, points };
}

export function selectCommonMotors(cases: ProcessCase[], equipment: SelectionConfig) {
  const { errors, points } = calculateProcessLoads(cases);
  for (const key of ['voltage', 'efficiency', 'powerFactor', 'voltageUtilization'] as const) {
    if (!Number.isFinite(equipment[key]) || equipment[key] <= 0 || (key !== 'voltage' && equipment[key] > 1)) errors.push(`设备电气参数 ${key} 无效`);
  }
  if (!Number.isFinite(equipment.maxOverallLength) || equipment.maxOverallLength < 0) errors.push('允许总长必须为非负数');
  if ([equipment.lowSpeedThreshold, equipment.lowSpeedMinFrame, equipment.maxFrequency ?? 0].some(x => !Number.isFinite(x) || x < 0)) errors.push('低速规则与频率上限必须为非负数');
  if (equipment.selectionMode === 'reference') {
    if (errors.length) return {errors, motors: [] as CommonMotor[]};
    const motors: CommonMotor[] = [...new Set(points.map(p=>p.station))].sort((a,b)=>a-b).map(station=>{
      const loads=points.filter(p=>p.station===station);
      const t=loads.reduce((a,b)=>a.torque>=b.torque?a:b), n=loads.reduce((a,b)=>a.rpm>=b.rpm?a:b);
      const choice=equipment.stationChoices?.[station];
      const mismatch=loads.some(p=>Math.abs(p.drum-loads[0].drum)>.001);
      const candidates=motorReferences.filter(m=>m.torque>=t.torque && m.rpm>=n.rpm && (!choice?.frame || m.frame===choice.frame) && (!choice?.poles || m.poles===choice.poles) && (!equipment.maxFrequency || m.rpm*m.poles/120<=equipment.maxFrequency)).sort((a,b)=>a.torque*a.rpm-b.torque*b.rpm || a.frame-b.frame || a.id.localeCompare(b.id));
      const m=equipment.maxOverallLength || mismatch ? undefined : candidates[0];
      const power=m ? m.torque*m.rpm/9550 : null;
      return {station,loads,requiredTorque:t.torque,requiredRpm:n.rpm,torqueSource:`${t.caseName} · ${t.kind}`,speedSource:`${n.caseName} · ${n.kind}`,ratedTorque:m?.torque??null,ratedRpm:m?.rpm??null,frame:m?.frame??null,poles:m?.poles??null,length:null,power,current:power===null?null:power*1000/(Math.sqrt(3)*equipment.voltage*equipment.efficiency*equipment.powerFactor*equipment.voltageUtilization),frequency:m?m.rpm*m.poles/120:null,reference:m?`${m.sheet}!C${m.row}:M${m.row}`:undefined,model:m?`${equipment.series}${m.frame}P${m.poles}-${m.torque}-${m.rpm}`:'待定制',status:m?'参考参数满足扭矩与转速；低速持续能力、尺寸待厂家确认':mismatch?'同一工位卷筒不一致':equipment.maxOverallLength?'附件无尺寸数据，无法确认总长约束':'52 条参考参数中无同时满足项，需定制或调整约束'};
    });
    return {errors,motors};
  }
  const validCatalog = (list: number[]) => list.length > 0 && list.every(x => Number.isFinite(x) && x > 0);
  if (!validCatalog(equipment.torqueCatalog) || !validCatalog(equipment.speedCatalog)) errors.push('转矩和转速档位库必须包含有效正数');
  if (!equipment.frameRules.length || equipment.frameRules.some(f => ![f.frame, f.maxTorque, f.estimatedLength, f.poles].every(x => Number.isFinite(x) && x > 0) || f.poles % 2 !== 0)) errors.push('机座库必须包含正数参数，极数须为正偶数');
  if (errors.length) return { errors, motors: [] as CommonMotor[] };
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
    const choice = equipment.stationChoices?.[station];
    const poleFits = (p: number) => Number.isInteger(p) && p > 0 && p % 2 === 0 && ratedRpm !== null && (!equipment.maxFrequency || ratedRpm * p / 120 <= equipment.maxFrequency) && (!choice?.poles || choice.poles === p);
    const frame = choice?.frame ? candidates.find(f => f.frame === choice.frame) : candidates.find(f => (f.supportedPoles?.length ? f.supportedPoles : [f.poles]).some(poleFits));
    const supported = frame ? (frame.supportedPoles?.length ? frame.supportedPoles : [frame.poles]) : [];
    const allowedPoles = supported.filter(p => Number.isInteger(p) && p > 0 && p % 2 === 0 && ratedRpm !== null && (!equipment.maxFrequency || ratedRpm * p / 120 <= equipment.maxFrequency));
    const poles = choice?.poles ? allowedPoles.find(p => p === choice.poles) : allowedPoles.includes(frame?.poles ?? 0) ? frame!.poles : allowedPoles[0];
    const drumMismatch = loads.some(p => Math.abs(p.drum - loads[0].drum) > 0.001);
    const status = drumMismatch ? '同一工位卷筒不一致，请确认设备配置' : ratedTorque === null ? '转矩超出档位库，需定制' : ratedRpm === null ? '转速超出档位库，需定制' : !frame ? '机座不满足转矩、总长或指定条件' : !poles ? '极数不在产品库内或超出频率上限' : '满足已录工况及产品库约束';
    const power = ratedTorque !== null && ratedRpm !== null ? ratedTorque * ratedRpm / 9550 : null;
    return { station, requiredTorque: t.torque, requiredRpm: n.rpm,
      torqueSource: `${t.caseName} · ${t.kind} · 第${t.pass}道`, speedSource: `${n.caseName} · ${n.kind} · 第${n.pass}道`,
      ratedTorque, ratedRpm, power, current: power === null ? null : power * 1000 / (Math.sqrt(3) * equipment.voltage * equipment.efficiency * equipment.powerFactor * equipment.voltageUtilization),
      frame: frame?.frame ?? null, poles: poles ?? null, length: frame?.estimatedLength ?? null,
      frequency: poles && ratedRpm !== null ? ratedRpm * poles / 120 : null,
      model: frame && poles && ratedTorque !== null && ratedRpm !== null && !drumMismatch ? `${equipment.series}${frame.frame}P${poles}-${ratedTorque}-${ratedRpm}-${Math.round(power!)}` : '待确认', status, loads };
  });
  return { errors, motors };
}
