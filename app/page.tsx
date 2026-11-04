'use client';
import {useEffect,useMemo,useState,useRef} from 'react';
import {Button} from '@/components/ui/button';
import {Input} from '@/components/ui/input';
import {Table,TableHeader,TableBody,TableRow,TableHead,TableCell} from '@/components/ui/table';
import {calculateMatrix,newMatrixProcess,copyMatrixProcess,speedInMetresPerSecond,type MatrixProcess,type MatrixSettings} from '@/lib/drawing-matrix';
import {MotorSelection} from '@/components/motor-selection';
import {workbookFilename} from '@/lib/export-filename';
import type {MotorChoice} from '@/lib/motor-spec';
import type {ProjectInfo,ProjectData} from '@/lib/project-excel';
const key='dds-horizontal-matrix:v1';
const fmt=(v:number)=>v.toLocaleString('zh-CN',{maximumFractionDigits:2});
export default function Home(){
  const [settings,setSettings]=useState<MatrixSettings>({count:7,first:'1.4',other:'1.3'});
  const [count,setCount]=useState('7'),[processes,setProcesses]=useState<MatrixProcess[]>([]);
  const [info,setInfo]=useState<ProjectInfo>({projectName:'',customerName:''});
  const [motorChoices,setMotorChoices]=useState<MotorChoice[]>([]);
  const [busy,setBusy]=useState(false),[importError,setImportError]=useState(''),[undo,setUndo]=useState<(ProjectData & {motorChoices:MotorChoice[]})|null>(null);
  const templateDialog=useRef<HTMLDialogElement>(null);
  const [templateCount,setTemplateCount]=useState('');
  const fileInput=useRef<HTMLInputElement>(null);
  const [ready,setReady]=useState(false),[message,setMessage]=useState('');
  useEffect(()=>{let saved:{settings:MatrixSettings;processes:MatrixProcess[];info?:ProjectInfo;motorChoices?:MotorChoice[]}|undefined;
    try{const raw=localStorage.getItem(key);if(raw){const d=JSON.parse(raw);if(!Number.isInteger(d.settings?.count)||d.settings.count<1||d.settings.count>128||!Array.isArray(d.processes)||!d.processes.every((p:MatrixProcess)=>typeof p.id==='string'&&Array.isArray(p.dies)&&Array.isArray(p.drums)))throw new Error();saved=d;}}
    catch{queueMicrotask(()=>setMessage('本机新表数据读取失败，请重新录入。旧版数据未修改。'));}
    queueMicrotask(()=>{if(saved){if(Array.isArray(saved.motorChoices))setMotorChoices(saved.motorChoices.map(c=>({height:typeof c?.height==='string'?c.height:'',poles:typeof c?.poles==='string'?c.poles:''})));setInfo({projectName:typeof saved.info?.projectName==='string'?saved.info.projectName:'',customerName:typeof saved.info?.customerName==='string'?saved.info.customerName:''});setSettings(saved.settings);setCount(String(saved.settings.count));setProcesses(saved.processes.map(speedInMetresPerSecond));}setReady(true);});
  },[]);
  useEffect(()=>{if(ready)try{localStorage.setItem(key,JSON.stringify({settings,processes,info,motorChoices}));}catch{queueMicrotask(()=>setMessage('本机保存失败，请导出结果。'));}},[ready,settings,processes,info,motorChoices]);
  const result=useMemo(()=>calculateMatrix(settings,processes),[settings,processes]);
  const cols=Array.from({length:settings.count},(_,i)=>i);
  const update=(id:string,patch:Partial<MatrixProcess>)=>setProcesses(old=>old.map(p=>p.id===id?{...p,...patch}:p));
  const numeric=(p:MatrixProcess,field:'entry'|'strength'|'speed')=><Input aria-label={p.name+' '+field} className="" type="number" min="0" step="any" value={p[field]} onChange={e=>update(p.id,{[field]:e.target.value})}/>;
  const applyCount=()=>{const n=Number(count);if(!Number.isInteger(n)||n<1||n>128){setMessage('总模数须为1～128的整数。');return;}setSettings(s=>({...s,count:n}));setProcesses(old=>old.map(p=>({...p,dies:Array.from({length:Math.max(n,p.dies.length)},(_,i)=>p.dies[i]??''),drums:Array.from({length:Math.max(n,p.drums.length)},(_,i)=>p.drums[i]??p.drums.at(-1)??'')})));setMessage('模数已应用。减少模数只隐藏后续数据，重新增加可恢复；末个可见道次为下线道次。');};
  const download=async(template=false,selectedCount=settings.count)=>{
    if(!template&&Number(count)!==settings.count){setMessage('请先应用总模数，再下载表格。');return;}
    setBusy(true);setImportError('');
    try{
      const {makeProjectExcel}=await import('@/lib/project-excel');
      const buffer=await makeProjectExcel({info,settings:template?{...settings,count:selectedCount}:settings,processes},template);
      const url=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const a=document.createElement('a');a.href=url;
      a.download=workbookFilename(info,template?selectedCount:settings.count,template?'标准工艺输入模板':'直驱选型参数表');
      a.click();if(template)templateDialog.current?.close();setTimeout(()=>URL.revokeObjectURL(url),1000);
    }catch(e){setImportError(e instanceof Error?e.message:'表格生成失败，请重试。');}finally{setBusy(false);}
  };
  const importFile=async(file:File)=>{
    setBusy(true);setImportError('');setMessage('');
    try{
      if(!file.name.toLowerCase().endsWith('.xlsx'))throw new Error('请选择.xlsx格式的标准输入表。');
      if(file.size>20*1024*1024)throw new Error('文件不能超过20MB。');
      const {readProjectExcel}=await import('@/lib/project-excel');
      const data=await readProjectExcel(await file.arrayBuffer(),settings);
      setUndo(structuredClone({info,settings,processes,motorChoices}));setMotorChoices([]);setInfo(data.info);setSettings(current=>({...current,count:data.settings.count}));setCount(String(data.settings.count));setProcesses(data.processes);
      setMessage('已导入 '+data.processes.length+' 套工艺，'+data.settings.count+' 道；已替换当前项目，可撤销本次导入。');
    }catch(e){setImportError(e instanceof Error?e.message:'导入失败，请检查文件。');}finally{setBusy(false);if(fileInput.current)fileInput.current.value='';}
  };
  return <main className="matrix-workspace min-h-screen text-foreground">
    <dialog ref={templateDialog} className="m-auto w-[min(420px,calc(100vw-32px))] rounded-xl border bg-card p-6 text-foreground shadow-xl backdrop:bg-black/30" aria-labelledby="template-title">
      <form onSubmit={e=>{e.preventDefault();void download(true,Number(templateCount));}}>
        <h2 id="template-title" className="text-lg font-semibold">请先选择模板模数</h2>
        <p className="mt-2 text-sm text-muted-foreground">按设备总模数生成客户填写表，不会修改当前项目模数。</p>
        <label htmlFor="template-count" className="mt-4 block text-sm">模板总模数（1～128）</label>
        <Input id="template-count" autoFocus required type="number" min="1" max="128" step="1" placeholder="请输入模数，例如13" className="mt-2" value={templateCount} onChange={e=>setTemplateCount(e.target.value)}/>
        <div className="mt-5 flex justify-end gap-2"><Button type="button" variant="outline" disabled={busy} onClick={()=>templateDialog.current?.close()}>取消</Button><Button type="submit" disabled={busy||!Number.isInteger(Number(templateCount))||Number(templateCount)<1||Number(templateCount)>128}>{busy?'正在生成…':'生成并下载'}</Button></div>
      </form>
    </dialog>
    <header className="flex flex-wrap items-center justify-between gap-3"><div><h1 className="text-xl font-semibold tracking-tight">直驱选型</h1><p className="mt-1 text-sm text-muted-foreground">一行一套工艺，逐道计算，分别取最大扭矩与最大转速。</p></div><Button variant="outline" disabled={!result.valid||busy} onClick={()=>download(false)}>导出 Excel</Button></header>
    <section className="project-toolbar rounded-xl border bg-card p-3 flex flex-wrap items-end gap-3">
      <label className="text-sm" htmlFor="project-name">项目名<Input id="project-name" className="mt-1 w-56" placeholder="例如：13模直驱改造" value={info.projectName} onChange={e=>setInfo(v=>({...v,projectName:e.target.value}))}/></label>
      <label className="text-sm" htmlFor="customer-name">客户名称<Input id="customer-name" className="mt-1 w-56" placeholder="填写客户或公司名称" value={info.customerName} onChange={e=>setInfo(v=>({...v,customerName:e.target.value}))}/></label>
      <Button variant="outline" disabled={busy||!ready} onClick={()=>{setTemplateCount('');templateDialog.current?.showModal();}}>下载标准输入表</Button>
      <Button variant="outline" disabled={busy||!ready} onClick={()=>fileInput.current?.click()}>{busy?'正在处理…':'一键导入 Excel'}</Button>
      <input ref={fileInput} type="file" accept=".xlsx" className="hidden" aria-label="导入标准工艺输入表" onChange={e=>{const file=e.target.files?.[0];if(file)void importFile(file);}}/>
      {undo&&<Button variant="ghost" disabled={busy} onClick={()=>{setMotorChoices(undo.motorChoices);setInfo(undo.info);setSettings(undo.settings);setCount(String(undo.settings.count));setProcesses(undo.processes);setUndo(null);setMessage('已恢复导入前的项目。');setImportError('');}}>撤销导入</Button>}
      <p className="w-full text-xs text-muted-foreground">下载前先选择模板模数 → 客户填写项目、客户和全部工艺 → 保存为 .xlsx 后一键导入。支持继续添加工艺行；扭矩系数仅在系统设置，导入不修改系数。校验失败不会替换当前数据。</p>
      {importError&&<p role="alert" className="w-full whitespace-pre-wrap text-sm text-destructive">{importError}</p>}
    </section>
    <section className="matrix-toolbar rounded-xl border bg-card flex flex-wrap items-end gap-3">
      <label htmlFor="die-count" className="text-sm">总模数<Input id="die-count" className="w-24 mt-1" type="number" min="1" max="128" value={count} onChange={e=>setCount(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyCount();}}/></label><Button variant="outline" onClick={applyCount}>应用模数</Button>
      <label htmlFor="first-factor" className="text-sm">首道次扭矩系数<Input id="first-factor" className="w-32 mt-1" type="number" step="0.05" min="0" value={settings.first} onChange={e=>setSettings(s=>({...s,first:e.target.value}))}/></label>
      <label htmlFor="other-factor" className="text-sm">其他道次扭矩系数<Input id="other-factor" className="w-32 mt-1" type="number" step="0.05" min="0" value={settings.other} onChange={e=>setSettings(s=>({...s,other:e.target.value}))}/></label>
      <Button disabled={!ready} onClick={()=>setProcesses(old=>[...old,newMatrixProcess(settings.count,old.length+1)])}>＋ 添加拉拔工艺</Button><span className="text-sm text-muted-foreground">{processes.length} 套工艺 · {settings.count} 道 · 本机自动保存</span>
    </section>
    {Number(count)!==settings.count&&<p role="alert" className="text-sm text-amber-700">模数尚未应用，当前按 {settings.count} 道计算。</p>}
    {message&&<p role="status" className="text-sm text-primary">{message}</p>}
    <section className="matrix-section space-y-2"><h2 className="font-semibold text-lg">工艺输入</h2><p className="text-sm text-muted-foreground">模具与上道尺寸相同即跳模，首道与进线比较。每道上格为模具、下格为滚筒（mm）。抗拉强度为进线强度；下线速度统一为 m/s。</p>
    <div className="rounded-xl border bg-card overflow-hidden"><Table className="matrix-input-table" style={{minWidth:398+settings.count*64}}><colgroup><col style={{width:104}}/><col style={{width:64}}/><col style={{width:72}}/><col style={{width:82}}/>{cols.map(i=><col key={i}/>)}<col style={{width:76}}/></colgroup><TableHeader><TableRow><TableHead className="sticky left-0 z-10 bg-card ">工艺名称</TableHead><TableHead>进线<br/><small>mm</small></TableHead><TableHead>抗拉强度<br/><small>MPa</small></TableHead><TableHead className="">下线速度<br/><small>m/s</small></TableHead>{cols.map(i=><TableHead key={i} className=" border-l"><span className="pass-number">{String(i+1).padStart(2,'0')}</span><span className="pass-label">模具 / 滚筒</span></TableHead>)}<TableHead className="">操作</TableHead></TableRow></TableHeader><TableBody>
    {!processes.length&&<TableRow><TableCell colSpan={settings.count+5} className="h-28 text-center text-muted-foreground">先确认总模数和系数，再添加拉拔工艺。</TableCell></TableRow>}
    {processes.map(p=><TableRow key={p.id}><TableCell className="sticky left-0 z-10 bg-card"><Input aria-label="工艺名称" value={p.name} onChange={e=>update(p.id,{name:e.target.value})}/></TableCell><TableCell>{numeric(p,'entry')}</TableCell><TableCell>{numeric(p,'strength')}</TableCell><TableCell><div className="matrix-speed">{numeric(p,'speed')}</div></TableCell>{cols.map(i=><TableCell key={i} className="border-l"><div className="matrix-pair">{(['dies','drums'] as const).map(k=><Input key={k} className="pass-input" aria-label={p.name+' 第'+(i+1)+'道'+(k==='dies'?'模具':'滚筒')} type="number" min="0" step="any" value={p[k][i]??''} onChange={e=>update(p.id,{[k]:Array.from({length:Math.max(settings.count,p[k].length)},(_,j)=>j===i?e.target.value:p[k][j]??'')})}/>)}</div><span className="skip-slot">{p.dies[i]&&Number(p.dies[i])===Number(i?p.dies[i-1]:p.entry)&&<span className="skip-label">跳模</span>}</span></TableCell>)}<TableCell><div className="matrix-actions"><Button size="sm" variant="ghost" onClick={()=>update(p.id,{drums:p.drums.map((d,i)=>i<settings.count?p.drums[0]:d)})}>滚筒同首道</Button><Button size="sm" variant="ghost" onClick={()=>setProcesses(old=>[...old,copyMatrixProcess(p)])}>复制</Button><Button size="sm" variant="ghost" onClick={()=>setProcesses(old=>old.filter(x=>x.id!==p.id))}>删除</Button></div></TableCell></TableRow>)}
    </TableBody></Table></div></section>
    <section className="matrix-section space-y-2"><h2 className="font-semibold text-lg">计算结果</h2><p className="text-sm text-muted-foreground">每道上行为扭矩（N·m），下行为转速（r/min）；综合行直接显示取整后的最大值：扭矩和转速均向上取到5的倍数（尾数为0或5）；悬停来源可查看完整工艺名称。</p>
    {!!result.settingsErrors.length&&<p role="alert" className="text-destructive text-sm">{result.settingsErrors.join('；')}</p>}
    <div className="rounded-xl border bg-card overflow-hidden"><Table className="matrix-result-table" style={{minWidth:104+settings.count*72}}><colgroup><col style={{width:104}}/>{cols.map(i=><col key={i}/>)}</colgroup><TableHeader><TableRow><TableHead className="sticky left-0 z-10 bg-card ">工艺</TableHead>{cols.map(i=><TableHead key={i} className="border-l  text-center"><span className="pass-number">{String(i+1).padStart(2,'0')}</span><span className="pass-label">扭矩 N·m</span><span className="pass-label">转速 r/min</span></TableHead>)}</TableRow></TableHeader><TableBody>
    {result.rows.map(r=><TableRow key={r.id}><TableCell className="sticky left-0 z-10 bg-card font-medium">{r.name}</TableCell>{r.errors.length||!r.points.length?<TableCell colSpan={settings.count} className="text-destructive whitespace-normal">{r.errors.length?<details className="matrix-errors"><summary>待完善 {r.errors.length} 项 · {r.errors[0]}</summary><p>{r.errors.join('；')}</p></details>:'请检查统一系数'}</TableCell>:r.points.map((p,i)=><TableCell key={i} className="border-l text-center tabular-nums"><span className="matrix-torque result-torque">{fmt(p.torque)}{p.skipped&&<small className="skip-label result-skip" title="跳模：理论拉拔扭矩为0">跳模</small>}</span><span className="matrix-rpm">{fmt(p.rpm)}</span></TableCell>)}</TableRow>)}
    <TableRow className="bg-primary/10 font-semibold"><TableCell className="sticky left-0 z-10 bg-primary text-primary-foreground">取整综合最大值</TableCell>{result.valid?result.maxima.map((p,i)=><TableCell key={i} className="border-l text-center tabular-nums"><span className="matrix-torque">{fmt(p.roundedTorque)}</span><span className="matrix-rpm">{fmt(p.roundedRpm)}</span><span className="matrix-source" title={'扭矩：'+p.torqueSource}>扭：{p.torqueSource}</span><span className="matrix-source" title={'转速：'+p.speedSource}>速：{p.speedSource}</span></TableCell>):<TableCell colSpan={settings.count} className="text-muted-foreground font-normal">{processes.length?'请补齐全部工艺后再综合，避免漏算最大需求。':'尚未添加工艺。'}</TableCell>}</TableRow>
    </TableBody></Table></div><div className="torque-total"><span>取整扭矩总和 <strong>{result.valid?fmt(result.totalRoundedTorque!):'—'}</strong> N·m</span></div></section>
    <MotorSelection project={{info,settings,processes}} demands={Number(count)===settings.count?result.maxima:[]} choices={motorChoices} onChange={setMotorChoices}/>
    <details className="rounded-xl border p-4 text-sm leading-7"><summary className="cursor-pointer font-medium">公式核对与跳模说明</summary><p>各道线速 vᵢ = 下线速度 × (末道丝径 / 本道丝径)²；转速 nᵢ = vᵢ × 1000 / (π × 滚筒直径)。v 用 m/min，滚筒用 mm；输入 m/s 时先乘60。直驱无减速比。</p><p>原表 5.5→5.05→4.55→4.13→3.8→3.5→3.25→3.25 mm，卷筒560 mm，下线8 m/s（480 m/min），各道转速为113.00、139.20、168.95、199.57、235.25、272.84、272.84 r/min，与原表 N17:N23 一致。</p><p>扭矩采用原表平均强度经验法：σᵢ=进线强度×√(进线丝径/本道丝径)，Fᵢ=0.6×进道丝径²×平均强度×√压缩率，Tᵢ=Fᵢ×滚筒直径/2000×系数。每套工艺第一次实际减径用首道系数。此后每次减径都用其他道次系数，包括中途跳模后的减径。原表后道1.2，新版按要求默认1.3，因此扭矩有相应差异。</p><p>跳模没有压缩，理论拉拔力和扭矩为0；滚筒继续带线时仍有转速。未含摩擦、张力、加速和调模附加载荷，不能将0 N·m当成电机不需实际扭矩。经验拉力模型须按材料验证。旧版草稿保留在原存储中，不自动导入新表。</p></details>
  </main>;
}
