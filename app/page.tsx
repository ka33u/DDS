'use client';
import {useEffect,useMemo,useState} from 'react';
import {Button} from '@/components/ui/button';
import {Input} from '@/components/ui/input';
import {Table,TableHeader,TableBody,TableRow,TableHead,TableCell} from '@/components/ui/table';
import {calculateMatrix,newMatrixProcess,matrixExcel,type MatrixProcess,type MatrixSettings} from '@/lib/drawing-matrix';
const key='dds-horizontal-matrix:v1';
const fmt=(v:number)=>v.toLocaleString('zh-CN',{maximumFractionDigits:2});
export default function Home(){
  const [settings,setSettings]=useState<MatrixSettings>({count:7,first:'1.4',other:'1.3'});
  const [count,setCount]=useState('7'),[processes,setProcesses]=useState<MatrixProcess[]>([]);
  const [ready,setReady]=useState(false),[message,setMessage]=useState('');
  useEffect(()=>{let saved:{settings:MatrixSettings;processes:MatrixProcess[]}|undefined;
    try{const raw=localStorage.getItem(key);if(raw){const d=JSON.parse(raw);if(!Number.isInteger(d.settings?.count)||d.settings.count<1||d.settings.count>128||!Array.isArray(d.processes)||!d.processes.every((p:MatrixProcess)=>typeof p.id==='string'&&Array.isArray(p.dies)&&Array.isArray(p.drums)))throw new Error();saved=d;}}
    catch{queueMicrotask(()=>setMessage('本机新表数据读取失败，请重新录入。旧版数据未修改。'));}
    queueMicrotask(()=>{if(saved){setSettings(saved.settings);setCount(String(saved.settings.count));setProcesses(saved.processes);}setReady(true);});
  },[]);
  useEffect(()=>{if(ready)try{localStorage.setItem(key,JSON.stringify({settings,processes}));}catch{queueMicrotask(()=>setMessage('本机保存失败，请导出结果。'));}},[ready,settings,processes]);
  const result=useMemo(()=>calculateMatrix(settings,processes),[settings,processes]);
  const cols=Array.from({length:settings.count},(_,i)=>i);
  const update=(id:string,patch:Partial<MatrixProcess>)=>setProcesses(old=>old.map(p=>p.id===id?{...p,...patch}:p));
  const numeric=(p:MatrixProcess,field:'entry'|'strength'|'speed')=><Input aria-label={p.name+' '+field} className="min-w-24" type="number" min="0" step="any" value={p[field]} onChange={e=>update(p.id,{[field]:e.target.value})}/>;
  const applyCount=()=>{const n=Number(count);if(!Number.isInteger(n)||n<1||n>128){setMessage('总模数须为1～128的整数。');return;}setSettings(s=>({...s,count:n}));setProcesses(old=>old.map(p=>({...p,dies:Array.from({length:Math.max(n,p.dies.length)},(_,i)=>p.dies[i]??''),drums:Array.from({length:Math.max(n,p.drums.length)},(_,i)=>p.drums[i]??p.drums.at(-1)??'')})));setMessage('模数已应用。减少模数只隐藏后续数据，重新增加可恢复；末个可见道次为下线道次。');};
  const download=()=>{const url=URL.createObjectURL(new Blob(['\uFEFF',matrixExcel(settings,processes)],{type:'application/vnd.ms-excel;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='拉丝多工艺_扭矩转速汇总.xls';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
  return <main className="min-h-screen bg-background text-foreground px-4 py-6 sm:px-8 space-y-6">
    <header className="flex flex-wrap items-center justify-between gap-3"><div><h1 className="text-2xl font-semibold">拉丝机 · 多工艺扭矩与转速</h1><p className="mt-1 text-sm text-muted-foreground">一行一套工艺，逐道计算，分别取最大扭矩与最大转速。</p></div><Button variant="outline" disabled={!result.valid} onClick={download}>导出 Excel</Button></header>
    <section className="rounded-xl border bg-card p-4 flex flex-wrap items-end gap-4">
      <label htmlFor="die-count" className="text-sm">总模数<Input id="die-count" className="w-24 mt-1" type="number" min="1" max="128" value={count} onChange={e=>setCount(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyCount();}}/></label><Button variant="outline" onClick={applyCount}>应用模数</Button>
      <label htmlFor="first-factor" className="text-sm">首道次扭矩系数<Input id="first-factor" className="w-32 mt-1" type="number" step="0.05" min="0" value={settings.first} onChange={e=>setSettings(s=>({...s,first:e.target.value}))}/></label>
      <label htmlFor="other-factor" className="text-sm">其他道次扭矩系数<Input id="other-factor" className="w-32 mt-1" type="number" step="0.05" min="0" value={settings.other} onChange={e=>setSettings(s=>({...s,other:e.target.value}))}/></label>
      <Button disabled={!ready} onClick={()=>setProcesses(old=>[...old,newMatrixProcess(settings.count,old.length+1)])}>＋ 添加拉拔工艺</Button><span className="text-sm text-muted-foreground">{processes.length} 套工艺 · {settings.count} 道 · 本机自动保存</span>
    </section>
    {Number(count)!==settings.count&&<p role="alert" className="text-sm text-amber-700">模数尚未应用，当前按 {settings.count} 道计算。</p>}
    {message&&<p role="status" className="text-sm text-primary">{message}</p>}
    <section className="space-y-3"><h2 className="font-semibold text-lg">工艺输入</h2><p className="text-sm text-muted-foreground">模具与上道尺寸相同即跳模，首道与进线比较。各道依次输入模具 / 滚筒直径，均为 mm；抗拉强度为进线强度。速度单位按钮可切换，数值同步换算。</p>
    <div className="rounded-xl border bg-card overflow-hidden"><Table><TableHeader><TableRow><TableHead className="sticky left-0 z-10 bg-card min-w-40">工艺名称</TableHead><TableHead>进线规格 mm</TableHead><TableHead>抗拉强度 MPa</TableHead><TableHead className="min-w-48">下线速度</TableHead>{cols.map(i=><TableHead key={i} className="min-w-44 border-l">第 {i+1} 道<br/><span className="font-normal">模具 mm / 滚筒 mm</span></TableHead>)}<TableHead className="min-w-52">操作</TableHead></TableRow></TableHeader><TableBody>
    {!processes.length&&<TableRow><TableCell colSpan={settings.count+5} className="h-28 text-center text-muted-foreground">先确认总模数和系数，再添加拉拔工艺。</TableCell></TableRow>}
    {processes.map(p=><TableRow key={p.id}><TableCell className="sticky left-0 z-10 bg-card"><Input aria-label="工艺名称" value={p.name} onChange={e=>update(p.id,{name:e.target.value})}/></TableCell><TableCell>{numeric(p,'entry')}</TableCell><TableCell>{numeric(p,'strength')}</TableCell><TableCell><div className="flex items-center gap-1">{numeric(p,'speed')}<Button size="sm" variant="outline" onClick={()=>{const unit=p.unit==='m/min'?'m/s':'m/min',v=Number(p.speed);update(p.id,{unit,speed:p.speed&&Number.isFinite(v)?String(Number((unit==='m/s'?v/60:v*60).toPrecision(12))):p.speed});}}>{p.unit}</Button></div></TableCell>{cols.map(i=><TableCell key={i} className="border-l"><div className="flex gap-1">{(['dies','drums'] as const).map(k=><Input key={k} className="w-20" aria-label={p.name+' 第'+(i+1)+'道'+(k==='dies'?'模具':'滚筒')} type="number" min="0" step="any" value={p[k][i]??''} onChange={e=>update(p.id,{[k]:Array.from({length:Math.max(settings.count,p[k].length)},(_,j)=>j===i?e.target.value:p[k][j]??'')})}/>)}</div>{p.dies[i]&&Number(p.dies[i])===Number(i?p.dies[i-1]:p.entry)&&<span className="text-xs text-amber-700">跳模</span>}</TableCell>)}<TableCell><div className="flex gap-1"><Button size="sm" variant="ghost" onClick={()=>update(p.id,{drums:p.drums.map((d,i)=>i<settings.count?p.drums[0]:d)})}>滚筒同首道</Button><Button size="sm" variant="ghost" onClick={()=>setProcesses(old=>[...old,{...structuredClone(p),id:crypto.randomUUID(),name:p.name+' 副本'}])}>复制</Button><Button size="sm" variant="ghost" onClick={()=>setProcesses(old=>old.filter(x=>x.id!==p.id))}>删除</Button></div></TableCell></TableRow>)}
    </TableBody></Table></div></section>
    <section className="space-y-3"><h2 className="font-semibold text-lg">计算结果</h2><p className="text-sm text-muted-foreground">左侧为扭矩 N·m，右侧为转速 r/min。分别取最大值，不上取额定档位、不叠加旧版跳头储备。</p>
    {!!result.settingsErrors.length&&<p role="alert" className="text-destructive text-sm">{result.settingsErrors.join('；')}</p>}
    <div className="rounded-xl border bg-card overflow-hidden"><Table><TableHeader><TableRow><TableHead className="sticky left-0 z-10 bg-card min-w-40">工艺</TableHead>{cols.map(i=><TableHead key={i} className="border-l min-w-48 text-center">第 {i+1} 道<br/>扭矩 N·m / 转速 r/min</TableHead>)}</TableRow></TableHeader><TableBody>
    {result.rows.map(r=><TableRow key={r.id}><TableCell className="sticky left-0 z-10 bg-card font-medium">{r.name}</TableCell>{r.errors.length||!r.points.length?<TableCell colSpan={settings.count} className="text-destructive whitespace-normal">{r.errors.join('；')||'请检查统一系数'}</TableCell>:r.points.map((p,i)=><TableCell key={i} className="border-l text-center tabular-nums">{fmt(p.torque)}<span className="mx-3 text-muted-foreground">/</span>{fmt(p.rpm)}{p.skipped&&<small className="block text-amber-700">跳模 · 理论拉拔扭矩为0</small>}</TableCell>)}</TableRow>)}
    <TableRow className="bg-primary/10 font-semibold"><TableCell className="sticky left-0 z-10 bg-primary text-primary-foreground">综合最大值</TableCell>{result.valid?result.maxima.map((p,i)=><TableCell key={i} className="border-l text-center tabular-nums">{fmt(p.torque)}<span className="mx-3">/</span>{fmt(p.rpm)}<span className="block text-xs font-normal mt-1">扭矩：{p.torqueSource}；转速：{p.speedSource}</span></TableCell>):<TableCell colSpan={settings.count} className="text-muted-foreground font-normal">{processes.length?'请补齐全部工艺后再综合，避免漏算最大需求。':'尚未添加工艺。'}</TableCell>}</TableRow>
    </TableBody></Table></div></section>
    <details className="rounded-xl border p-4 text-sm leading-7"><summary className="cursor-pointer font-medium">公式核对与跳模说明</summary><p>各道线速 vᵢ = 下线速度 × (末道丝径 / 本道丝径)²；转速 nᵢ = vᵢ × 1000 / (π × 滚筒直径)。v 用 m/min，滚筒用 mm；输入 m/s 时先乘60。直驱无减速比。</p><p>原表 5.5→5.05→4.55→4.13→3.8→3.5→3.25→3.25 mm，卷筒560 mm，下线8 m/s（480 m/min），各道转速为113.00、139.20、168.95、199.57、235.25、272.84、272.84 r/min，与原表 N17:N23 一致。</p><p>扭矩采用原表平均强度经验法：σᵢ=进线强度×√(进线丝径/本道丝径)，Fᵢ=0.6×进道丝径²×平均强度×√压缩率，Tᵢ=Fᵢ×滚筒直径/2000×系数。首道系数作用于物理第1道，其余道使用后道系数。原表后道1.2，新版按要求默认1.3，因此扭矩有相应差异。</p><p>跳模没有压缩，理论拉拔力和扭矩为0；滚筒继续带线时仍有转速。未含摩擦、张力、加速和调模附加载荷，不能将0 N·m当成电机不需实际扭矩。经验拉力模型须按材料验证。旧版草稿保留在原存储中，不自动导入新表。</p></details>
  </main>;
}
