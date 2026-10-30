import { defaultConfig, equalCompressionPasses, type SelectionConfig } from './motor-selection.ts';
export function newGeneralConfig(): SelectionConfig {
  return {...structuredClone(defaultConfig),projectName:'新建拉丝设备方案',machineModel:'自定义直驱设备',deviceName:'设备1',material:'请填写实际材料',selectionMode:'custom',referenceRecipe:undefined,entryDiameter:5,finalDiameter:2,finalSpeed:5,incomingStrength:1000,maxOverallLength:0,lowSpeedThreshold:0,lowSpeedMinFrame:0,equipmentDrums:Array(8).fill(500),processStations:Array.from({length:8},(_,i)=>i+1),passes:equalCompressionPasses(8,5,2,500),forceModel:'average',strengthExponent:.5,drawingEfficiency:1,parallelWires:1};
}
export function parseNumbers(text: string): number[] {
  return text.trim().split(/[,，;；\s→]+/).filter(Boolean).map(Number);
}
/** Explicit application, never silently drops a malformed pasted value. */
export function applyProcessBatch(c: SelectionConfig, diesText: string, stationsText: string, forcesText: string, speedsText: string): SelectionConfig {
  const dies=parseNumbers(diesText), stations=parseNumbers(stationsText), forces=parseNumbers(forcesText), speeds=parseNumbers(speedsText);
  if (!dies.length || dies.length>128 || dies.some((d,i)=>!Number.isFinite(d)||d<=0||d>=(i?dies[i-1]:c.entryDiameter))) throw new Error('配模需填写 1～128 个严格递减的正数，不含进线直径。');
  if(stations.length!==dies.length || stations.some((s,i)=>!Number.isInteger(s)||s<1||s>128||(i>0&&s<=stations[i-1]))) throw new Error('工位编号须与配模数量相同，且为 1～128 的递增整数。');
  if(c.equipmentDrums?.length && stations.some(s=>s>c.equipmentDrums!.length)) throw new Error('所选工位超出当前设备工位数。');
  if(forces.length && (forces.length!==dies.length||forces.some(f=>!Number.isFinite(f)||f<0))) throw new Error('拉拔力须为空，或逐道填写相同数量的非负数。');
  if(speeds.length && (speeds.length!==dies.length||speeds.some(s=>!Number.isFinite(s)||s<=0))) throw new Error('独立线速须为空，或逐道填写相同数量的正数（m/min）。');
  return {...c,referenceRecipe:undefined,processStations:stations,finalDiameter:dies.at(-1)!,passes:dies.map((diameter,i)=>({diameter,drumDiameter:c.equipmentDrums?.[stations[i]-1]??c.passes[i]?.drumDiameter??500,measuredForce:forces.length?forces[i]:undefined,lineSpeed:speeds.length?speeds[i]/60:undefined}))};
}
