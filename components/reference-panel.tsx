'use client';
import { recipes, loadRecipe, motorReferences, referenceWorkbook } from '@/lib/reference-data';
import { calculateDrawingRequirements, type SelectionConfig } from '@/lib/motor-selection';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Table,TableHeader,TableHead,TableBody,TableRow,TableCell } from '@/components/ui/table';

export function ReferencePanel({config,onChange}:{config:SelectionConfig;onChange:(c:SelectionConfig)=>void}) {
  const recipe=recipes.find(r=>r.id===config.referenceRecipe);
  const rows=calculateDrawingRequirements(config);
  return <section className="rounded-xl border bg-card p-5 space-y-4">
    <h2 className="font-semibold text-lg">客户工艺模板 · 优玛乐 8 种机型</h2>
    <p className="text-sm text-muted-foreground">点击载入设计最高车速及示例配模，不代表覆盖图片中的全部进出线组合。系数恢复为 1.4 / 1.3；强度暂按图片上限 1250 MPa，需按实际进线材料修改。已存工艺清单不会被替换。</p>
    <div className="flex flex-wrap gap-2">{recipes.map(r=><Button key={r.id} variant={recipe?.id===r.id?'default':'outline'} onClick={()=>onChange(loadRecipe(config,r))}>{r.id}</Button>)}</div>
    {recipe && <>
      <p className="text-sm">图片适用范围：进线 {recipe.entryRange} mm，出线 {recipe.exitRange} mm；设计 {recipe.design} m/min，正常 {recipe.normal.join('–')} m/min。载入时总长限制置为未设，附件没有外形尺寸。</p>
      {recipe.warning && <p role="alert" className="text-sm text-amber-700">{recipe.warning}</p>}
      <div className="flex flex-wrap items-end gap-3"><label htmlFor="reference-speed" className="text-sm">成品车速 m/min<Input id="reference-speed" type="number" min="0" step="1" value={Number((config.finalSpeed*60).toFixed(4))} onChange={e=>onChange({...config,finalSpeed:Number(e.target.value)/60})}/></label><Button variant="outline" onClick={()=>onChange({...config,finalSpeed:recipe.design/60})}>设计最高车速</Button><Button variant="outline" onClick={()=>onChange({...config,finalSpeed:recipe.normal[1]/60})}>正常车速上限</Button></div>
      <details><summary className="cursor-pointer text-sm font-medium">与原表逐道核对（当前编辑工艺，不含已存多工艺及调模）</summary>
      <Table><TableHeader><TableRow>{['道次','机座 / 极数','原表 N·m / rpm','需求 N·m / rpm','当前工艺校核'].map(h=><TableHead key={h}>{h}</TableHead>)}</TableRow></TableHeader><TableBody>{motorReferences.filter(m=>m.sheet===recipe.id+'选型').map((m,i)=>{const r=rows[i];return <TableRow key={m.id}><TableCell>{i+1}</TableCell><TableCell>{m.frame} / {m.poles}</TableCell><TableCell>{m.torque} / {m.rpm}</TableCell><TableCell>{r?`${r.designTorque.toFixed(1)} / ${r.designRpm.toFixed(2)}`:'无对应道次'}</TableCell><TableCell>{r?(m.torque>=r.designTorque?'扭矩满足':'扭矩不足')+'；'+(m.rpm>=r.designRpm?'转速满足':'转速不足'):'待核对'}</TableCell></TableRow>})}</TableBody></Table></details>
    </>}
    <details className="text-sm"><summary className="cursor-pointer">来源与异常处理</summary><p className="mt-2 leading-6">配模及速度来源：124.jpg。电机参数来源：{referenceWorkbook}，共 52 条历史选型。9 处效率原值为 1.9，均不用于电流计算；电流统一使用设备页可调效率。LZ7-350 前四道 30 极、后三道 8 极按原表保留。参考选择按同时满足扭矩及转速后额定功率最小排序，不代表成本最优或厂家定型确认。11模页为另一个工艺：采用进道强度法，系数为首道 1.5、中间 1.4、末道 1.2，与默认平均强度法不同，不混入这 8 个模板。</p></details>
  </section>;
}
