"use client";
import { useEffect, useMemo, useState } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card';
import { Table, TableHeader, TableBody, TableRow, TableHead, TableCell } from '@/components/ui/table';
import { calculateProcessLoads, makeProcessCase, selectCommonMotors, type ProcessCase } from '@/lib/process-selection';
import type { SelectionConfig } from '@/lib/motor-selection';
import { downloadProcessExcel } from '@/lib/excel-export';

const format = (x: number | null) => x === null ? '—' : x.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
export function ProcessWorkbench({ config, onLoad, onEdit, phase = 'demand' }: { config: SelectionConfig; onLoad: (config: SelectionConfig) => void; onEdit?: () => void; phase?: string }) {
  const [cases, setCases] = useState<ProcessCase[]>([]);
  const [ready, setReady] = useState(false);
  const [message, setMessage] = useState('');
  useEffect(() => {
    try {
      const saved = localStorage.getItem('direct-drive-processes:v2');
      if (saved) { const data = JSON.parse(saved); if (Array.isArray(data)) queueMicrotask(() => setCases(data)); }
    } catch { queueMicrotask(() => setMessage('未能读取本机工艺清单，请重新加入工艺。')); }
    queueMicrotask(() => setReady(true));
  }, []);
  useEffect(() => { if (ready) { try { localStorage.setItem('direct-drive-processes:v2', JSON.stringify(cases)); } catch { queueMicrotask(() => setMessage('本机存储不可用，请导出方案保留数据。')); } } }, [cases, ready]);
  // Empty collection is a live single-process selection; saved cases become a fixed comparison set.
  const deviceName = config.deviceName?.trim() || '设备1';
  const deviceCases = useMemo(() => cases.filter(c => (c.config.deviceName?.trim() || '设备1') === deviceName), [cases,deviceName]);
  const activeCases = useMemo(() => deviceCases.length ? deviceCases : [makeProcessCase(config)], [deviceCases, config]);
  const result = useMemo(() => selectCommonMotors(activeCases, config), [activeCases, config]);
  const demand = useMemo(() => calculateProcessLoads(activeCases, config), [activeCases,config]);
  const stations = [...new Set(demand.points.map(p => p.station))].sort((a,b) => a-b);
  const update = (id: string, patch: Partial<ProcessCase>) => setCases(old => old.map(c => c.id === id ? { ...c, ...patch } : c));
  const loadProcess = (item: ProcessCase) => onLoad({...config, ...JSON.parse(JSON.stringify(item.config)), equipmentDrums: config.equipmentDrums, frameRules: config.frameRules, torqueCatalog: config.torqueCatalog, speedCatalog: config.speedCatalog, selectionMode: config.selectionMode, processStations: item.stations.slice()});
  return <Card className={phase === 'input' ? 'hidden' : 'mb-6 border-primary/30'}>
    <CardHeader><CardTitle>{phase === 'demand' ? '第二步 · 各工位电机需求' : '第三步 · 机座号与极数细化'}</CardTitle><p className="text-sm leading-6 text-muted-foreground">{phase === 'demand' ? '先确定各工位需要的扭矩和转速；此处需求与机座、极数、电气参数无关。多工艺按实际工位分别取最大需求。' : '在工艺需求基础上向上取额定档位，再按转矩能力、总长及产品库选择机座与极数。极数不改变工艺转速，通过 f = n × 极数 / 120 校核频率。'}</p></CardHeader>
    <CardContent className="space-y-4">
      <div className="flex flex-wrap gap-2">{[...new Set(cases.map(c=>c.config.deviceName?.trim()||'设备1'))].map(name=><Button key={name} variant={name===deviceName?'default':'outline'} onClick={()=>{const first=cases.find(c=>(c.config.deviceName?.trim()||'设备1')===name)!;onLoad({...first.config,deviceName:name,processStations:first.stations.slice()});}}>设备：{name}</Button>)}</div>
      <div className={phase === 'demand' ? 'flex flex-wrap gap-2' : 'hidden'}>
        <Button onClick={() => { setCases(old => [...old, makeProcessCase(config, `工艺 ${old.length + 1} · ${config.entryDiameter}→${config.passes.at(-1)?.diameter} mm`)]); setMessage('当前输入已加入工艺清单；修改下方参数后可再加入另一项工艺。'); }}>将当前输入加入工艺清单</Button>
        <Button variant="outline" disabled={!demand.points.length} onClick={() => downloadProcessExcel(config, activeCases, result.motors)}>导出工艺需求及选型 Excel</Button>
        {deviceCases.length > 0 && <Button variant="outline" onClick={() => { setCases(old=>old.filter(c=>(c.config.deviceName?.trim()||'设备1')!==deviceName)); setMessage('已返回单工艺实时选型。'); }}>返回单工艺实时选型</Button>}
      </div>
      <p className="text-sm">当前设备：{deviceName}；该设备已存 {deviceCases.length} 套工艺，其他设备 {cases.length-deviceCases.length} 套未参与本次选型。输入草稿只在本设备没有已存工艺时实时参与。</p>
      <output className="text-sm text-primary">{message || (deviceCases.length ? '已保存的工艺使用独立参数；设备电气、档位和机座库共用下方设置。' : '当前为单工艺实时选型。要配置调模或对比多工艺，请先加入清单。')}</output>
      {phase === 'demand' && deviceCases.map(item => <details key={item.id} className="rounded-lg border p-3" open>
        <summary className="cursor-pointer text-base font-medium">{item.enabled ? '参与选型' : '已停用'} · {item.name} · {item.config.passes.length} 道 · {item.config.finalSpeed} m/s</summary>
        <div className="mt-3 grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
          <label htmlFor={`process-field-1-${item.id}`} className="text-sm sm:col-span-2">工艺名称<Input id={`process-field-1-${item.id}`} value={item.name} onChange={e => update(item.id, { name: e.target.value })}/></label>
          <label htmlFor={`group-${item.id}`} className="text-sm">所属设备<Input id={`group-${item.id}`} value={item.config.deviceName??'设备1'} onChange={e=>update(item.id,{config:{...item.config,deviceName:e.target.value}})}/></label>
          <label htmlFor={`process-field-2-${item.id}`} className="text-sm sm:col-span-2">实际工位编号（逐道，用逗号分隔）<Input id={`process-field-2-${item.id}`} key={item.stations.join(',')} defaultValue={item.stations.join(', ')} onBlur={e => update(item.id, { stations: e.target.value.split(/[,，\s]+/).filter(Boolean).map(Number) })}/></label>
          <label htmlFor={`process-field-3-${item.id}`} className="text-sm">调模线速 m/s（0 停用）<Input id={`process-field-3-${item.id}`} type="number" min="0" step="0.1" value={item.jogSpeed} onChange={e => update(item.id, { jogSpeed: Number(e.target.value) })}/></label>
          <label htmlFor={`process-field-4-${item.id}`} className="text-sm">调模转矩系数<Input id={`process-field-4-${item.id}`} type="number" min="0" step="0.1" value={item.jogFactor} onChange={e => update(item.id, { jogFactor: Number(e.target.value) })}/></label>
          <label htmlFor={`process-field-5-${item.id}`} className="text-sm">调模拉力 N（0 按拉拔力）<Input id={`process-field-5-${item.id}`} type="number" min="0" value={item.jogForce} onChange={e => update(item.id, { jogForce: Number(e.target.value) })}/></label>
          <label htmlFor={`process-field-6-${item.id}`} className="text-sm">附加阻力转矩 N·m<Input id={`process-field-6-${item.id}`} type="number" min="0" value={item.additionalTorque} onChange={e => update(item.id, { additionalTorque: Number(e.target.value) })}/></label>
        </div>
        <div className="mt-3 flex flex-wrap gap-2">
          <Button size="sm" variant="outline" onClick={() => {loadProcess(item);onEdit?.();}}>载入第一步编辑</Button>
          <Button size="sm" variant="outline" onClick={() => { update(item.id, { config: JSON.parse(JSON.stringify(config)), stations: config.processStations?.slice() ?? config.passes.map((_,i)=>i+1) }); setMessage(`已用下方输入更新 ${item.name}，请检查工位映射。`); }}>用当前输入更新此工艺</Button>
          <Button size="sm" variant="outline" onClick={() => update(item.id, { enabled: !item.enabled })}>{item.enabled ? '停用此工艺' : '启用此工艺'}</Button>
          <Button size="sm" variant="outline" onClick={()=>{const copy=makeProcessCase(item.config,item.name+' 副本');setCases(old=>[...old,{...item,...copy,stations:item.stations.slice(),jogSpeed:item.jogSpeed,jogFactor:item.jogFactor,jogForce:item.jogForce,additionalTorque:item.additionalTorque}]);}}>复制工艺</Button>
          <Button size="sm" variant="outline" onClick={()=>{if(!config.equipmentDrums?.length){setMessage('请先定义设备卷筒。');return;} update(item.id,{config:{...item.config,passes:item.config.passes.map((p,i)=>({...p,drumDiameter:config.equipmentDrums?.[item.stations[i]-1]??p.drumDiameter}))}});setMessage('已同步设备卷筒；超出设备的工位仍需修改映射。');}}>同步设备卷筒</Button>
          <Button size="sm" variant="ghost" onClick={() => setCases(old => old.filter(c => c.id !== item.id))}>移除</Button>
        </div>
      </details>)}
      {phase === 'demand' && <><div role="alert" className="text-sm text-destructive">{demand.errors.join('；')}</div><Table><TableHeader><TableRow>{['实际工位', '所需扭矩 N·m', '所需转速 r/min', '扭矩决定工况', '转速决定工况'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{stations.map(s => { const points = demand.points.filter(p => p.station === s); const t = points.reduce((a,b) => a.torque >= b.torque ? a:b); const n = points.reduce((a,b) => a.rpm >= b.rpm ? a:b); return <TableRow key={s}><TableCell>{s}</TableCell><TableCell>{format(t.torque)}</TableCell><TableCell>{format(n.rpm)}</TableCell><TableCell>{t.caseName} · {t.kind}</TableCell><TableCell>{n.caseName} · {n.kind}</TableCell></TableRow>; })}</TableBody></Table><details open><summary>逐道负载明细</summary><Table><TableHeader><TableRow>{['工艺', '工位', '用途', '拉拔力 N', '扭矩需求 N·m', '转速需求 r/min'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{demand.points.map((p,i) => <TableRow key={i}><TableCell>{p.caseName}</TableCell><TableCell>{p.station}</TableCell><TableCell>{p.kind}</TableCell><TableCell>{format(p.force)}</TableCell><TableCell>{format(p.torque)}</TableCell><TableCell>{format(p.rpm)}</TableCell></TableRow>)}</TableBody></Table></details></>}
      <div className={phase === 'motor' ? 'space-y-4' : 'hidden'}>
      <div className="flex flex-wrap gap-2"><Button variant={config.selectionMode==='reference'?'default':'outline'} onClick={()=>onLoad({...config,selectionMode:'reference'})}>历史参数对照（非通用产品库）</Button><Button variant={config.selectionMode!=='reference'?'default':'outline'} onClick={()=>onLoad({...config,selectionMode:'custom'})}>自定义档位及机座规则</Button></div>
      <p className="text-sm">{config.selectionMode==='reference'?'参考模式：整条匹配额定转矩、转速、机座和极数；优先选择额定功率较小项。总长未知，有长度限制时不判定合格。电流使用全局效率估算，不使用原表异常效率。下方机座规则在此模式不生效。':'自定义模式：以下能力边界需有厂家产品资料支持，组合型号只是定制建议。'}</p>
      <label htmlFor={`process-field-7-${'global'}`} className="block max-w-sm text-sm">允许最高电气频率 Hz（0 表示未设限）<Input id={`process-field-7-${'global'}`} type="number" min="0" value={config.maxFrequency ?? 0} onChange={e => onLoad({ ...config, maxFrequency: Number(e.target.value) })}/></label>
      <div className={config.selectionMode==='reference'?'hidden':'space-y-3'}><p className="text-sm">以下为自定义模式机座库，不是 52 条参考记录的通用能力边界。请按实际产品资料维护。</p>
      <Button variant="outline" onClick={()=>onLoad({...config,frameRules:[...config.frameRules,{frame:160,maxTorque:100,estimatedLength:500,poles:8}]})}>增加机座规则（填入厂家数据后使用）</Button>
      <Table><TableHeader><TableRow>{['中心高', '能力 N·m', '总长 mm', '默认极数', '可选极数（逗号分隔）','操作'].map(h=><TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{config.frameRules.map((f,i)=><TableRow key={i}>{(['frame','maxTorque','estimatedLength','poles'] as const).map(k=><TableCell key={k}><Input aria-label={`机座${i+1} ${k}`} type="number" value={f[k]} onChange={e=>onLoad({...config,frameRules:config.frameRules.map((r,j)=>j===i?{...r,[k]:Number(e.target.value)}:r)})}/></TableCell>)}<TableCell><Input aria-label={`机座${i+1}可选极数`} defaultValue={(f.supportedPoles ?? [f.poles]).join(',')} onBlur={e=>onLoad({...config,frameRules:config.frameRules.map((r,j)=>j===i?{...r,supportedPoles:e.target.value.split(/[,，\s]+/).filter(Boolean).map(Number)}:r)})}/></TableCell><TableCell><Button variant="ghost" onClick={()=>onLoad({...config,frameRules:config.frameRules.filter((_,j)=>j!==i)})}>移除规则</Button></TableCell></TableRow>)}</TableBody></Table>
      </div>
      <Table><TableHeader><TableRow><TableHead>工位</TableHead><TableHead>指定机座（留空自动）</TableHead><TableHead>指定极数（留空按产品库）</TableHead><TableHead>选定机座 / 极数</TableHead></TableRow></TableHeader><TableBody>{result.motors.map(m=><TableRow key={m.station}><TableCell>{m.station}</TableCell>{(['frame','poles'] as const).map(k=><TableCell key={k}><Input type="number" aria-label={`工位${m.station}指定${k}`} value={config.stationChoices?.[m.station]?.[k] ?? ''} onChange={e=>onLoad({...config,stationChoices:{...config.stationChoices,[m.station]:{...config.stationChoices?.[m.station],[k]:e.target.value===''?undefined:Number(e.target.value)}}})}/></TableCell>)}<TableCell>{format(m.frame)} / {format(m.poles)} 极</TableCell></TableRow>)}</TableBody></Table>
      <Button disabled={!result.motors.length} onClick={() => downloadProcessExcel(config, activeCases, result.motors)}>导出完整选型参数表</Button>
      {result.errors.length > 0 && <div role="alert" className="rounded-lg bg-destructive/10 p-3 text-sm">{result.errors.join('；')}</div>}
      <Table><TableHeader><TableRow>{['实际工位', '推荐型号 / 状态', '额定转矩 N·m', '额定转速 r/min', '功率 kW', '电流 A', '频率 Hz', '总长 mm', '转矩决定工况', '转速决定工况'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader>
        <TableBody>{result.motors.map(m => <TableRow key={m.station}><TableCell>{m.station}</TableCell><TableCell><p className="font-medium">{m.model}</p><p className="text-sm">{m.status}</p></TableCell><TableCell>{format(m.ratedTorque)}<p className="text-sm text-muted-foreground">需求 {format(m.requiredTorque)}</p></TableCell><TableCell>{format(m.ratedRpm)}<p className="text-sm text-muted-foreground">需求 {format(m.requiredRpm)}</p></TableCell><TableCell>{format(m.power)}</TableCell><TableCell>{format(m.current)}</TableCell><TableCell>{format(m.frequency)}</TableCell><TableCell>{format(m.length)}</TableCell><TableCell>{m.torqueSource}</TableCell><TableCell>{m.speedSource}</TableCell></TableRow>)}</TableBody>
      </Table>
      <details><summary className="cursor-pointer text-sm font-medium">查看每项工艺的负载点及覆盖余量</summary><Table><TableHeader><TableRow>{['工艺', '工位', '用途', '拉力 N', '所需转矩 N·m', '所需转速 r/min', '转矩余量', '转速余量'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{result.motors.flatMap(m => m.loads.map((p, i) => <TableRow key={`${m.station}-${i}`}><TableCell>{p.caseName}</TableCell><TableCell>{m.station}</TableCell><TableCell>{p.kind}</TableCell><TableCell>{format(p.force)}</TableCell><TableCell>{format(p.torque)}</TableCell><TableCell>{format(p.rpm)}</TableCell><TableCell>{m.ratedTorque === null || !p.torque ? '—' : `${format((m.ratedTorque / p.torque - 1) * 100)}%`}</TableCell><TableCell>{m.ratedRpm === null || !p.rpm ? '—' : `${format((m.ratedRpm / p.rpm - 1) * 100)}%`}</TableCell></TableRow>))}</TableBody></Table></details>
      </div><p className="text-sm leading-6 text-muted-foreground">调模线速指各工位卷筒表面速度。调模转矩按拉力×卷筒半径×调模系数，加附加阻力转矩；所有负载点均按连续恒转矩能力覆盖，不借用短时过载或弱磁能力。同一工位的卷筒尺寸应一致。型号为参数组合建议，实际产品能力以确认后的机座库为准。</p>
    </CardContent>
  </Card>;
}
