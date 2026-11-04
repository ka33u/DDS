'use client';
import {useState} from 'react';
import {Button} from '@/components/ui/button';
import {defaults, heightOptions, poleOptions, motorSpec, newChoice, validChoice, type MotorChoice, type DriveDemand} from '@/lib/motor-spec';
import type {ProjectInfo} from '@/lib/project-excel';

interface Props {
  info:ProjectInfo;
  demands:DriveDemand[];
  choices:MotorChoice[];
  onChange:(choices:MotorChoice[])=>void;
}
const display=(value:number|null, digits=2)=>value===null?'—':value.toLocaleString('zh-CN',{maximumFractionDigits:digits});

export function MotorSelection({info,demands,choices,onChange}:Props){
  const [saving,setSaving]=useState(false);
  const [error,setError]=useState('');
  const complete=demands.length>0 && demands.every((_,i)=>validChoice(choices[i]??newChoice()));
  const setChoice=(i:number,key:keyof MotorChoice,value:string)=>{
    const next=choices.slice();next[i]={...(next[i]??newChoice()),[key]:value};onChange(next);
  };
  const fillFromFirst=()=>{
    const next=choices.slice();
    demands.forEach((_,i)=>{next[i]={height:next[i]?.height||choices[0].height,poles:next[i]?.poles||choices[0].poles};});
    onChange(next);
  };
  const download=async()=>{
    setSaving(true);setError('');
    try{
      const {motorSpecWorkbook}=await import('@/lib/motor-spec-excel');
      const buffer=await motorSpecWorkbook(info,demands,choices);
      const url=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link=document.createElement('a');link.href=url;
      const prefix=[info.projectName,info.customerName].filter(Boolean).join('_').replace(/[\\/:*?"<>|]/g,'_');
      link.download=(prefix?prefix+'_':'')+demands.length+'模永磁电机选型参数.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
    }catch(e){setError(e instanceof Error?e.message:'规格表生成失败');}finally{setSaving(false);}
  };
  return <section className="matrix-section space-y-2">
    <div className="flex flex-wrap items-center justify-between gap-2">
      <h2 className="font-semibold">电机规格选型</h2>
      <div className="flex gap-2">
        <Button variant="outline" size="sm" disabled={!demands.length||!validChoice(choices[0]??newChoice())} onClick={fillFromFirst}>用首道选择补齐空项</Button>
        <Button size="sm" disabled={!complete||saving} onClick={download}>{saving?'正在生成…':'导出电机规格（含公式）'}</Button>
      </div>
    </div>
    <p className="text-muted-foreground">各道自动带入取整转矩和转速，请选择中心高（mm）和极数，也可直接输入正整数中心高、正偶数极数。</p>
    {!demands.length?<p className="rounded-lg border bg-card p-4">补齐拉拔工艺后，将在这里显示各道电机规格。</p>:
      <div className="rounded-xl border bg-card overflow-x-auto"><table className="motor-spec-table">
        <thead><tr>{['道次','中心高 mm','极数','电机型号','转矩 N·m','转速 r/min','功率 kW','频率 Hz','参考效率','电流 A'].map(label=><th key={label}>{label}</th>)}</tr></thead>
        <tbody>{demands.map((d,i)=>{
          const choice=choices[i]??newChoice(),spec=motorSpec(d,choice);
          return <tr key={i}><td>{i+1}</td>
            <td><input aria-label={`第${i+1}道中心高`} list="motor-heights" type="number" min="1" step="1" placeholder="选择" value={choice.height} onChange={e=>setChoice(i,'height',e.target.value)}/></td>
            <td><input aria-label={`第${i+1}道极数`} list="motor-poles" type="number" min="2" step="2" placeholder="选择" value={choice.poles} onChange={e=>setChoice(i,'poles',e.target.value)}/></td>
            <td>{spec.model||'待选中心高与极数'}</td><td>{display(spec.torque)}</td><td>{display(spec.rpm)}</td>
            <td>{display(spec.power,0)}</td><td>{spec.frequency===null?'—':spec.frequency.toFixed(1)}</td><td>{spec.efficiency.value===null?'—':display(spec.efficiency.value*100,1)+'%'}<small>{spec.efficiency.note}</small></td><td>{display(spec.current,1)}</td></tr>;
        })}</tbody>
      </table></div>}
    <datalist id="motor-heights">{heightOptions.map(n=><option key={n} value={n}/>)}</datalist>
    <datalist id="motor-poles">{poleOptions.map(n=><option key={n} value={n}/>)}</datalist>
    <p className="text-muted-foreground">{defaults.series} · {defaults.voltage} VAC · cosφ {defaults.powerFactor} · IP55 / F / S1 / Y · 水冷。功率按整数 kW 显示，效率和电流按计算功率求取。效率采用 GB 30253-2024 一级能效参考值。导出后可在「电机选型」表中调整中心高、极数、转速和转矩。</p>
    {error&&<p role="alert" className="text-destructive">{error}</p>}
  </section>;
}
