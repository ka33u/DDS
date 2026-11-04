export interface MatrixProcess { id:string; name:string; entry:string; strength:string; speed:string; unit:'m/min'|'m/s'; dies:string[]; drums:string[] }
export const roundingOptions = [
  {id:'five', label:'逢5或10取整', description:'向上取到 5 的倍数，如 101.2 → 105。'},
  {id:'nearest', label:'四舍五入取整', description:'取最接近的整数，如 101.2 → 101，101.5 → 102。'},
  {id:'up', label:'向上取整', description:'向上取到整数，如 101.2 → 102。'},
] as const;
export type RoundingMode = (typeof roundingOptions)[number]['id'];
export function roundingOption(value:unknown) {
  return roundingOptions.find(option=>option.id===value) ?? roundingOptions[0];
}
export interface MatrixSettings { count:number; first:string; other:string; rounding?:RoundingMode }
export interface MatrixPoint { torque:number; rpm:number; force:number; speed:number; skipped:boolean }
export function createId():string {
  const crypto=globalThis.crypto;
  if(typeof crypto?.randomUUID==='function')return crypto.randomUUID();
  // LAN HTTP does not expose randomUUID, but supports cryptographic random bytes.
  const bytes=crypto.getRandomValues(new Uint8Array(16));
  bytes[6]=(bytes[6]&0x0f)|0x40;
  bytes[8]=(bytes[8]&0x3f)|0x80;
  const hex=Array.from(bytes,b=>b.toString(16).padStart(2,'0')).join('');
  return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
}
export function copyMatrixProcess(process:MatrixProcess):MatrixProcess {
  return {...structuredClone(process),id:createId(),name:process.name+' 副本'};
}
export function newMatrixProcess(count:number,index:number):MatrixProcess {
  return {id:createId(),name:`工艺 ${index}`,entry:'',strength:'',speed:'',unit:'m/s',dies:Array(count).fill(''),drums:Array(count).fill('')};
}
export function speedInMetresPerSecond(p:MatrixProcess):MatrixProcess {
  if(p.unit!=='m/min')return p;
  return {...p,unit:'m/s',speed:p.speed.trim()&&Number.isFinite(Number(p.speed))?String(Number((Number(p.speed)/60).toPrecision(15))):p.speed};
}
export function roundDemand(value:number, mode:RoundingMode='five') {
  if (mode==='nearest') return Math.round(value);
  if (mode==='up') return Math.ceil(value);
  return Math.ceil(value/5)*5;
}
export {roundDemand as roundTorque};
export function calculateMatrix(settings:MatrixSettings, processes:MatrixProcess[]) {
  const settingsErrors:string[]=[];
  if(!Number.isInteger(settings.count)||settings.count<1||settings.count>128) settingsErrors.push('总模数须为 1～128 的整数');
  if(![settings.first,settings.other].every(x=>Number.isFinite(Number(x))&&Number(x)>0))settingsErrors.push('扭矩系数须为正数');
  const rows=processes.map(p=>{
    const errors:string[]=[];
    for(const [name,v] of [['进线规格',p.entry],['进线抗拉强度',p.strength],['下线速度',p.speed]])if(!Number.isFinite(Number(v))||Number(v)<=0)errors.push(`${name}须为正数`);
    if(p.unit!=='m/min'&&p.unit!=='m/s')errors.push('速度单位无效');
    if(p.dies.length<settings.count||p.drums.length<settings.count)errors.push('道次输入不足');
    for(let i=0;i<settings.count;i++) {
      const d=Number(p.dies[i]),prev=Number(i?p.dies[i-1]:p.entry),drum=Number(p.drums[i]);
      if(!Number.isFinite(d)||d<=0||d>prev)errors.push(`第${i+1}道模具须大于0且不大于前道（相同即跳模）`);
      if(!Number.isFinite(drum)||drum<=0)errors.push(`第${i+1}道滚筒直径须为正数`);
    }
    const points:MatrixPoint[]=[];
    if(!errors.length&&!settingsErrors.length){
      const entry=Number(p.entry),sigma=Number(p.strength),end=Number(p.dies[settings.count-1]),endSpeed=Number(p.speed)*(p.unit==='m/s'?60:1);
      let torqueFactor = Number(settings.first);
      for(let i=0;i<settings.count;i++){
        const before=Number(i?p.dies[i-1]:p.entry),d=Number(p.dies[i]),D=Number(p.drums[i]);
        const skipped=d===before,compression=1-(d/before)**2;
        const avg=(sigma*Math.sqrt(entry/before)+sigma*Math.sqrt(entry/d))/2;
        const force=skipped?0:.6*before**2*avg*Math.sqrt(compression);
        const speed=endSpeed*(end/d)**2;
        points.push({force,speed,skipped,torque:force*D/2000*torqueFactor,rpm:speed*1000/(Math.PI*D)});
        if (!skipped) torqueFactor = Number(settings.other);
      }
      if(points.some(x=>![x.force,x.torque,x.speed,x.rpm].every(Number.isFinite))){errors.push('输入量级超出计算范围');points.length=0;}
    }
    return {id:p.id,name:p.name,errors,points};
  });
  const valid=processes.length>0&&!settingsErrors.length&&rows.every(r=>!r.errors.length);
  const rounding=roundingOption(settings.rounding);
  const maxima=valid?Array.from({length:settings.count},(_,i)=>{
    const t=rows.reduce((a,b)=>a.points[i].torque>=b.points[i].torque?a:b),n=rows.reduce((a,b)=>a.points[i].rpm>=b.points[i].rpm?a:b);
    return {torque:t.points[i].torque,roundedTorque:roundDemand(t.points[i].torque,rounding.id),rpm:n.points[i].rpm,roundedRpm:roundDemand(n.points[i].rpm,rounding.id),torqueSource:t.name,speedSource:n.name};
  }):[];
  return {rows,maxima,rounding,settingsErrors,valid,totalTorque:valid?maxima.reduce((sum,p)=>sum+p.torque,0):null,totalRoundedTorque:valid?maxima.reduce((sum,p)=>sum+p.roundedTorque,0):null};
}
export function matrixExcel(settings:MatrixSettings,processes:MatrixProcess[]) {
  const result=calculateMatrix(settings,processes);
  if(!result.valid)throw new Error('请先补齐所有工艺的有效参数');
  const escape=(x:unknown)=>String(x??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
  const sheet=(name:string,rows:(string|number)[][])=>`<Worksheet ss:Name="${name}"><Table>${rows.map((r,i)=>`<Row>${r.map(v=>`<Cell${i===0?' ss:StyleID="Header"':''}><Data ss:Type="${typeof v==='number'?'Number':'String'}">${escape(v)}</Data></Cell>`).join('')}</Row>`).join('')}</Table></Worksheet>`;
  const cols=Array.from({length:settings.count},(_,i)=>i+1);
  return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="Header"><Font ss:Bold="1"/><Interior ss:Color="#DCEEEE" ss:Pattern="Solid"/></Style></Styles>${sheet('取整综合最大值',[['道次','扭矩 N·m','转速 r/min'],...result.maxima.map((p,i)=>[`第${i+1}道`,p.roundedTorque,p.roundedRpm]),['取整扭矩总和 N·m',result.totalRoundedTorque!,'']])}${sheet('工艺输入',[['工艺','进线 mm','进线抗拉强度 MPa','下线速度','单位',...cols.flatMap(i=>[`${i}道模具 mm`,`${i}道滚筒 mm`])],...processes.map(speedInMetresPerSecond).map(p=>[p.name,Number(p.entry),Number(p.strength),Number(p.speed),p.unit,...cols.flatMap((_,i)=>[Number(p.dies[i]),Number(p.drums[i])])]),['首道系数',Number(settings.first),'其他道系数',Number(settings.other)],['取整规则',result.rounding.label+'。'+result.rounding.description+' 总和为各道取整扭矩之和。'],['说明','相同模具尺寸为跳模，理论拉拔扭矩为0，转速仍按线速计算。未含摩擦、加速、调模负载。'],['转速公式','v_i=下线速度(m/min)×(末道丝径/本道丝径)^2；n_i=1000×v_i/(π×滚筒直径mm)'],['拉力公式','F=0.6×进道丝径²×平均抗拉强度×sqrt(压缩率)；强度=进线强度×sqrt(进线丝径/本道丝径)；T=F×滚筒直径/2000×系数']])}</Workbook>`;
}
