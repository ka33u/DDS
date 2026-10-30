import type { SelectionConfig } from './motor-selection.ts';

export const referenceWorkbook = '优玛乐直驱8种机型电机选型2026.9.15.xlsx';
export interface MotorReference {
  id: string; sheet: string; row: number; frame: number; rpm: number; torque: number; poles: number; rawEfficiency: number;
}
export interface Recipe {
  id: string; drum: number; entry: number; dies: number[]; design: number; normal: [number, number]; entryRange: string; exitRange: string; warning?: string;
}
export const recipes: Recipe[] = [
  {id:'LZ5-760',drum:760,entry:10,dies:[9.03,8.15,7.36,6.65,6],design:300,normal:[280,290],entryRange:'7.0–10.0',exitRange:'4.5–6.0'},
  {id:'LZ6-600',drum:600,entry:4.5,dies:[4.05,3.64,3.28,2.95,2.66,2.39],design:280,normal:[240,260],entryRange:'2.75–4.5',exitRange:'1.8–2.39'},
  {id:'LZ6-685',drum:685,entry:5.99,dies:[5.47,4.99,4.55,4.15,3.79,3.46],design:260,normal:[240,250],entryRange:'3.24–5.99',exitRange:'2.4–3.46'},
  {id:'LZ6-760',drum:760,entry:6.5,dies:[6.04,5.62,5.23,4.86,4.52,4.2],design:240,normal:[220,230],entryRange:'4.48–6.5',exitRange:'3.5–4.2'},
  {id:'LZ6-450',drum:450,entry:4.2,dies:[4,3.82,3.64,3.47,3.31,3.16],design:300,normal:[220,260],entryRange:'3.0–4.2',exitRange:'2.0–3.16'},
  {id:'LZ8-400',drum:400,entry:2.92,dies:[2.66,2.43,2.22,2.02,1.84,1.68,1.53,1.4],design:300,normal:[280,290],entryRange:'1.32–2.92',exitRange:'0.86–1.4'},
  {id:'LZ8-450',drum:450,entry:3.16,dies:[2.95,2.75,2.56,2.39,2.22,2.07,1.93,1.8],design:300,normal:[280,290],entryRange:'2.18–3.16',exitRange:'1.45–1.8'},
  {id:'LZ7-350',drum:350,entry:1.15,dies:[1.06,.98,.90,.83,.76,.70,.65],design:420,normal:[420,420],entryRange:'1.15',exitRange:'0.63',warning:'图片成品规格为 0.63 mm，配模末道为 0.65 mm。本模板按实际配模 0.65 mm 计算；若需 0.63 mm，请修改末道并重新校核。'},
];
// Each tuple is an indivisible historical selection, not a universal frame capability envelope.
const tuples: Record<string, number[][]> = {
  'LZ5-760':[[450,56,17600,80],[450,70,14100,80],[450,85,12000,80],[450,105,10300,80],[450,125,8900,80]],
  'LZ6-600':[[315,52,2850,64],[315,65,2300,64],[280,80,1920,32],[280,97,1660,32],[280,120,1400,32],[280,150,1220,32]],
  'LZ6-685':[[315,50,5380,64],[315,60,4380,64],[315,70,3830,64],[315,85,3330,64],[315,100,2880,64],[315,120,2520,64]],
  'LZ6-760':[[355,50,6350,64],[355,60,5230,64],[315,65,4700,64],[315,75,4250,64],[315,87,3800,64],[315,101,3420,64]],
  'LZ6-450':[[225,135,1300,32],[225,145,1080,32],[225,160,1030,32],[225,175,955,32],[225,195,890,32],[225,213,820,32]],
  'LZ8-400':[[225,67,740,32],[225,80,590,32],[225,95,520,32],[180,115,460,16],[180,140,395,16],[180,165,340,16],[180,200,300,16],[180,240,255,16]],
  'LZ8-450':[[225,80,850,32],[225,90,720,32],[225,105,650,32],[225,120,575,32],[225,140,535,32],[180,160,470,16],[180,185,420,16],[180,215,380,16]],
  'LZ7-350':[[160,145,95,30],[160,170,80,30],[160,200,70,30],[160,235,60,30],[160,280,55,8],[160,330,50,8],[160,385,40,8]],
};
export const motorReferences: MotorReference[] = Object.entries(tuples).flatMap(([name, rows]) => rows.map(([frame,rpm,torque,poles],i) => ({id:`${name}-${i+1}`,sheet:`${name}选型`,row:i+3,frame,rpm,torque,poles,rawEfficiency:i===5 || i===7 ? 1.9 : .9})));
export function loadRecipe(config: SelectionConfig, recipe: Recipe): SelectionConfig {
  return {...config,equipmentDrums:recipe.dies.map(()=>recipe.drum),processStations:recipe.dies.map((_,i)=>i+1),forceModel:'average',strengthExponent:.5,drawingEfficiency:1,parallelWires:1,referenceRecipe:recipe.id,selectionMode:'reference',projectName:`优玛乐 ${recipe.id}`,machineModel:recipe.id,material:'60–80 钢（图片强度上限 1250 MPa）',entryDiameter:recipe.entry,finalDiameter:recipe.dies.at(-1)!,finalSpeed:recipe.design/60,incomingStrength:1250,strengthFactor:1,forceCoefficient:.6,firstPassFactor:1.4,steadyPassFactor:1.3,preset:'standard',skipHeadSpan:0,stationChoices:{},maxOverallLength:0,passes:recipe.dies.map(diameter=>({diameter,drumDiameter:recipe.drum}))};
}
