"use client";
import { useEffect, useMemo, useState } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card';
import { Table, TableHeader, TableBody, TableRow, TableHead, TableCell } from '@/components/ui/table';
import { makeProcessCase, selectCommonMotors, type ProcessCase } from '@/lib/process-selection';
import type { SelectionConfig } from '@/lib/motor-selection';
import { downloadProcessExcel } from '@/lib/excel-export';

const format = (x: number | null) => x === null ? '—' : x.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
export function ProcessWorkbench({ config, onLoad }: { config: SelectionConfig; onLoad: (config: SelectionConfig) => void }) {
  const [cases, setCases] = useState<ProcessCase[]>([]);
  const [ready, setReady] = useState(false);
  const [message, setMessage] = useState('');
  useEffect(() => {
    try {
      const saved = localStorage.getItem('direct-drive-processes:v2');
      if (saved) { const data = JSON.parse(saved); if (Array.isArray(data)) setCases(data); }
    } catch { setMessage('未能读取本机工艺清单，请重新加入工艺。'); }
    setReady(true);
  }, []);
  useEffect(() => { if (ready) { try { localStorage.setItem('direct-drive-processes:v2', JSON.stringify(cases)); } catch { setMessage('本机存储不可用，请导出方案保留数据。'); } } }, [cases, ready]);
  // Empty collection is a live single-process selection; saved cases become a fixed comparison set.
  const activeCases = useMemo(() => cases.length ? cases : [makeProcessCase(config)], [cases, config]);
  const result = useMemo(() => selectCommonMotors(activeCases, config), [activeCases, config]);
  const update = (id: string, patch: Partial<ProcessCase>) => setCases(old => old.map(c => c.id === id ? { ...c, ...patch } : c));
  return <Card className="mb-6 border-primary/30">
    <CardHeader><CardTitle>客户工艺兼容选型</CardTitle><p className="text-sm leading-6 text-muted-foreground">单工艺随下方输入实时计算。加入工艺后，按实际工位合并所有启用工艺的生产、跳头和调模要求。标准转矩、转速取满足全部要求的最小档位，机座取满足转矩与总长的最小中心高。</p></CardHeader>
    <CardContent className="space-y-4">
      <div className="flex flex-wrap gap-2">
        <Button onClick={() => { setCases(old => [...old, makeProcessCase(config, `工艺 ${old.length + 1} · ${config.entryDiameter}→${config.passes.at(-1)?.diameter} mm`)]); setMessage('当前输入已加入工艺清单；修改下方参数后可再加入另一项工艺。'); }}>将当前输入加入工艺清单</Button>
        <Button variant="outline" disabled={!result.motors.length} onClick={() => downloadProcessExcel(config, activeCases, result.motors)}>导出兼容选型 Excel</Button>
        {cases.length > 0 && <Button variant="outline" onClick={() => { setCases([]); setMessage('已返回单工艺实时选型。'); }}>返回单工艺实时选型</Button>}
      </div>
      <p role="status" className="text-sm text-primary">{message || (cases.length ? '已保存的工艺使用独立参数；设备电气、档位和机座库共用下方设置。' : '当前为单工艺实时选型。要配置调模或对比多工艺，请先加入清单。')}</p>
      {cases.map(item => <details key={item.id} className="rounded-lg border p-3" open>
        <summary className="cursor-pointer text-base font-medium">{item.enabled ? '参与选型' : '已停用'} · {item.name} · {item.config.passes.length} 道 · {item.config.finalSpeed} m/s</summary>
        <div className="mt-3 grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
          <label className="text-sm sm:col-span-2">工艺名称<Input value={item.name} onChange={e => update(item.id, { name: e.target.value })}/></label>
          <label className="text-sm sm:col-span-2">实际工位编号（逐道，用逗号分隔）<Input defaultValue={item.stations.join(', ')} onBlur={e => update(item.id, { stations: e.target.value.split(/[,，\s]+/).filter(Boolean).map(Number) })}/></label>
          <label className="text-sm">调模线速 m/s（0 停用）<Input type="number" min="0" step="0.1" value={item.jogSpeed} onChange={e => update(item.id, { jogSpeed: Number(e.target.value) })}/></label>
          <label className="text-sm">调模转矩系数<Input type="number" min="0" step="0.1" value={item.jogFactor} onChange={e => update(item.id, { jogFactor: Number(e.target.value) })}/></label>
          <label className="text-sm">调模拉力 N（0 按拉拔力）<Input type="number" min="0" value={item.jogForce} onChange={e => update(item.id, { jogForce: Number(e.target.value) })}/></label>
          <label className="text-sm">附加阻力转矩 N·m<Input type="number" min="0" value={item.additionalTorque} onChange={e => update(item.id, { additionalTorque: Number(e.target.value) })}/></label>
        </div>
        <div className="mt-3 flex flex-wrap gap-2">
          <Button size="sm" variant="outline" onClick={() => onLoad(JSON.parse(JSON.stringify(item.config)))}>载入下方编辑</Button>
          <Button size="sm" variant="outline" onClick={() => { update(item.id, { config: JSON.parse(JSON.stringify(config)) }); setMessage(`已用下方输入更新 ${item.name}，请检查工位映射。`); }}>用当前输入更新此工艺</Button>
          <Button size="sm" variant="outline" onClick={() => update(item.id, { enabled: !item.enabled })}>{item.enabled ? '停用此工艺' : '启用此工艺'}</Button>
          <Button size="sm" variant="ghost" onClick={() => setCases(old => old.filter(c => c.id !== item.id))}>移除</Button>
        </div>
      </details>)}
      {result.errors.length > 0 && <div role="alert" className="rounded-lg bg-destructive/10 p-3 text-sm">{result.errors.join('；')}</div>}
      <Table><TableHeader><TableRow>{['实际工位', '推荐型号 / 状态', '额定转矩 N·m', '额定转速 r/min', '功率 kW', '电流 A', '频率 Hz', '总长 mm', '转矩决定工况', '转速决定工况'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader>
        <TableBody>{result.motors.map(m => <TableRow key={m.station}><TableCell>{m.station}</TableCell><TableCell><p className="font-medium">{m.model}</p><p className="text-sm">{m.status}</p></TableCell><TableCell>{format(m.ratedTorque)}<p className="text-sm text-muted-foreground">需求 {format(m.requiredTorque)}</p></TableCell><TableCell>{format(m.ratedRpm)}<p className="text-sm text-muted-foreground">需求 {format(m.requiredRpm)}</p></TableCell><TableCell>{format(m.power)}</TableCell><TableCell>{format(m.current)}</TableCell><TableCell>{format(m.frequency)}</TableCell><TableCell>{format(m.length)}</TableCell><TableCell>{m.torqueSource}</TableCell><TableCell>{m.speedSource}</TableCell></TableRow>)}</TableBody>
      </Table>
      <details><summary className="cursor-pointer text-sm font-medium">查看每项工艺的负载点及覆盖余量</summary><Table><TableHeader><TableRow>{['工艺', '工位', '用途', '拉力 N', '所需转矩 N·m', '所需转速 r/min', '转矩余量', '转速余量'].map(h => <TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{result.motors.flatMap(m => m.loads.map((p, i) => <TableRow key={`${m.station}-${i}`}><TableCell>{p.caseName}</TableCell><TableCell>{m.station}</TableCell><TableCell>{p.kind}</TableCell><TableCell>{format(p.force)}</TableCell><TableCell>{format(p.torque)}</TableCell><TableCell>{format(p.rpm)}</TableCell><TableCell>{m.ratedTorque === null || !p.torque ? '—' : `${format((m.ratedTorque / p.torque - 1) * 100)}%`}</TableCell><TableCell>{m.ratedRpm === null || !p.rpm ? '—' : `${format((m.ratedRpm / p.rpm - 1) * 100)}%`}</TableCell></TableRow>))}</TableBody></Table></details>
      <p className="text-sm leading-6 text-muted-foreground">调模线速指各工位卷筒表面速度。调模转矩按拉力×卷筒半径×调模系数，加附加阻力转矩；所有负载点均按连续恒转矩能力覆盖，不借用短时过载或弱磁能力。同一工位的卷筒尺寸应一致。型号为参数组合建议，实际产品能力以确认后的机座库为准。</p>
    </CardContent>
  </Card>;
}
