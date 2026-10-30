"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Activity,
  ChevronDown,
  ChevronRight,
  Download,
  Factory,
  Minus,
  Plus,
  RefreshCcw,
  Save,
  SlidersHorizontal,
  Sparkles,
  Waves,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { ProcessWorkbench } from "@/components/process-workbench";
import { GeneralProcessPanel } from "@/components/general-process-panel";
import { newGeneralConfig } from "@/lib/general-process";
import { ReferencePanel } from "@/components/reference-panel";
import {
  applyPreset,
  calculateDrawingRequirements,
  equalCompressionPasses,
  type MachinePreset,
  type SelectionConfig,
} from "@/lib/motor-selection";

const STORAGE_KEY = "direct-drive-selector:v1";
const number = (value: string) => Number(value) || 0;
const cloneDefaults = (): SelectionConfig => newGeneralConfig();
const parseCatalog = (value: string) => value.split(/[，,\s]+/).map(Number).filter((item) => Number.isFinite(item) && item > 0);

export default function Home() {
  const [config, setConfig] = useState<SelectionConfig>(() => cloneDefaults());
  const [phase, setPhase] = useState('input');
  const rows = useMemo(() => calculateDrawingRequirements(config), [config]);
  const configRef = useRef(config);
  const rowsRef = useRef(rows);

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const saved = JSON.parse(stored) as Partial<SelectionConfig>;
        const restored = { ...cloneDefaults(), ...saved, equipmentDrums:saved.equipmentDrums, processStations:saved.processStations ?? saved.passes?.map((_,i)=>i+1) };
        queueMicrotask(() => setConfig(restored));
      }
    } catch {
      // Ignore a damaged device-local draft and retain the verified example.
    }
  }, []);

  useEffect(() => {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
  }, [config]);

  useEffect(() => {
    configRef.current = config;
    rowsRef.current = rows;
  }, [config, rows]);

  useEffect(() => {
    const context = document.modelContext;
    if (!context?.registerTool) return;
    const lifecycle = new AbortController();
    const register = async () => {
      await context.registerTool({
        name: "configure_and_calculate_motor_selection",
        title: "配置并计算直驱电机选型",
        description: "更新可见选型工况并立即重新计算拉拔工艺与直驱电机规格。",
        inputSchema: {
          type: "object",
          properties: {
            projectName: { type: "string" },
            preset: { type: "string", enum: ["standard", "universal", "custom"] },
            entryDiameter: { type: "number", exclusiveMinimum: 0 },
            finalDiameter: { type: "number", exclusiveMinimum: 0 },
            finalSpeed: { type: "number", exclusiveMinimum: 0 },
            incomingStrength: { type: "number", exclusiveMinimum: 0 },
            firstPassFactor: { type: "number", exclusiveMinimum: 0 },
            steadyPassFactor: { type: "number", exclusiveMinimum: 0 },
            maxOverallLength: { type: "number", minimum: 0 },
            passDiameters: { type: "array", items: { type: "number", exclusiveMinimum: 0 }, minItems: 1, maxItems: 128 },
            drumDiameters: { type: "array", items: { type: "number", exclusiveMinimum: 0 }, minItems: 1, maxItems: 128 },
          },
          additionalProperties: false,
        },
        annotations: { readOnlyHint: false, untrustedContentHint: false },
        execute(raw) {
          const input = raw as Record<string, unknown>;
          let next = { ...configRef.current };
          if (typeof input.preset === "string" && ["standard", "universal", "custom"].includes(input.preset)) next = applyPreset(next, input.preset as MachinePreset);
          for (const key of ["projectName", "entryDiameter", "finalDiameter", "finalSpeed", "incomingStrength", "firstPassFactor", "steadyPassFactor", "maxOverallLength"] as const) {
            const value = input[key];
            if (typeof value === (key === "projectName" ? "string" : "number")) next = { ...next, [key]: value };
          }
          if (Array.isArray(input.passDiameters)) {
            const drums = Array.isArray(input.drumDiameters) ? input.drumDiameters : [];
            next.passes = input.passDiameters.map((diameter, index) => ({
              diameter: Number(diameter),
              drumDiameter: Number(drums[index] ?? next.passes[index]?.drumDiameter ?? next.passes.at(-1)?.drumDiameter ?? 560),
            }));
          }
          const result = calculateDrawingRequirements(next);
          setConfig(next);
          return { projectName: next.projectName, passes: result.length, maxRequiredTorque: Math.max(...result.map((row) => row.designTorque)), maxRequiredRpm: Math.max(...result.map((row) => row.designRpm)) };
        },
      }, { signal: lifecycle.signal });
      await context.registerTool({
        name: "read_motor_selection_result",
        title: "读取直驱电机选型结果",
        description: "读取当前编辑工艺的逐道扭矩和转速需求，不含多工艺合并或机座极数选型。",
        inputSchema: { type: "object", properties: {}, additionalProperties: false },
        annotations: { readOnlyHint: true, untrustedContentHint: false },
        execute() {
          return { projectName: configRef.current.projectName, preset: configRef.current.preset, rows: rowsRef.current.map((row) => ({ pass: row.pass, requiredTorque: row.designTorque, requiredRpm: row.designRpm, drawingForce: row.drawingForce })) };
        },
      }, { signal: lifecycle.signal });
    };
    void register().catch(() => undefined);
    return () => lifecycle.abort();
  }, []);

  const setPreset = (preset: MachinePreset) => setConfig((old) => applyPreset(old, preset));
  const update = <K extends keyof SelectionConfig>(key: K, value: SelectionConfig[K]) => setConfig((old) => ({ ...old, [key]: value }));
  const updateCustom = <K extends keyof SelectionConfig>(key: K, value: SelectionConfig[K]) => setConfig((old) => ({ ...old, preset: "custom", [key]: value }));
  const distribute = (count = config.passes.length) => setConfig(old => ({...old, processStations: Array.from({length:count},(_,i)=>i+1), passes: equalCompressionPasses(count, old.entryDiameter, old.finalDiameter, old.passes.at(-1)?.drumDiameter || 500).map((p,i)=>({...p,drumDiameter:old.equipmentDrums?.[i]??p.drumDiameter}))}));
  const reset = () => {
    window.localStorage.removeItem(STORAGE_KEY);
    setConfig(cloneDefaults());
  };

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="sticky top-0 z-30 border-b border-border/80 bg-background/92 backdrop-blur-xl">
        <div className="mx-auto flex min-h-16 max-w-[1760px] items-center gap-4 px-4 py-2 sm:px-6 lg:px-8">
          <div className="flex items-center gap-3">
            <span className="grid size-9 place-items-center rounded-xl bg-primary text-primary-foreground shadow-sm"><Waves className="size-5" /></span>
            <div><p className="font-heading text-sm font-semibold tracking-tight">拉丝机直驱选型</p><p className="text-[11px] text-muted-foreground">DirectDrive Selector · 工程版</p></div>
          </div>
          <div className="ml-auto flex items-center gap-2">
            <span className="hidden items-center gap-1.5 text-[11px] text-muted-foreground md:flex"><Save className="size-3.5" />自动保存到本机</span>
            <Button variant="outline" size="sm" onClick={reset}><RefreshCcw />新建输入草稿</Button>
            <Button size="sm" onClick={() => setPhase('motor')}><Download />查看并导出选型</Button>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1760px] px-4 py-6 sm:px-6 lg:px-8">
        <section className="mb-5 flex flex-col justify-between gap-4 lg:flex-row lg:items-end">
          <div>
            <div className="mb-2 flex items-center gap-2 text-xs font-medium text-primary"><span>工况定义</span><ChevronRight className="size-3.5"/><span>拉拔计算</span><ChevronRight className="size-3.5"/><span>电机初选</span></div>
            <h1 className="font-heading text-2xl font-semibold tracking-tight sm:text-3xl">通用拉丝设备 · 多工艺电机选型</h1>
            <p className="mt-2 max-w-3xl text-sm leading-6 text-muted-foreground">设备分组 → 配模与工位映射 → 多工艺负载汇总 → 电机参数；历史 Excel 只作校核参考，不限制设备类型。</p>
          </div>
          <div className="flex flex-wrap gap-1.5 rounded-xl border bg-card p-1.5 shadow-sm">
            {(["standard", "universal", "custom"] as MachinePreset[]).map((preset) => (
              <Button key={preset} variant={config.preset === preset ? "default" : "ghost"} size="sm" onClick={() => setPreset(preset)}>
                {preset === "standard" ? "标准专机 1.4 / 1.3" : preset === "universal" ? "万能机 1.4 / 1.2" : "完全自定义"}
              </Button>
            ))}
          </div>
        </section>

        <Tabs value={phase} onValueChange={setPhase} className="mb-5"><TabsList><TabsTrigger value="input">1 · 工艺输入</TabsTrigger><TabsTrigger value="demand">2 · 扭矩与转速需求</TabsTrigger><TabsTrigger value="motor">3 · 机座与极数</TabsTrigger></TabsList></Tabs>
        <details className="mb-5 rounded-lg border p-4 text-sm leading-7"><summary className="cursor-pointer font-medium">Excel 选型思路与系统对应关系</summary><p>工艺表 5.5~3.25：G列按截面积守恒计算车速，J列按平均抗拉强度计算拉拔力，L列计算卷筒转矩，N列计算转速，O列乘安全系数。先得到需求，再取电机额定档位。</p><p>电机选型表：C列中心高、F列额定转速、G列额定转矩、M列极数均为人工填写；E列按 T×n/9550 算功率，J列按 n×极数/120 算频率。原表没有通用机座与极数推导公式。</p><p>原表六道工艺另列了第七台电机，但没有对应六道表的第七道拉拔计算。系统按实际录入和工位映射计算，不自动补造第七道工艺。机座尺寸和能力初始库需用实际产品资料核准。</p></details>
        <ProcessWorkbench config={config} onLoad={setConfig} onEdit={()=>setPhase('input')} phase={phase} />
        <div className={phase === 'input' ? 'grid gap-5' : 'hidden'}>
          <GeneralProcessPanel config={config} onChange={setConfig} />
          <details><summary className="cursor-pointer text-sm text-muted-foreground">历史项目参考（选用，不作为通用模板）</summary><ReferencePanel config={config} onChange={setConfig} /></details>
          <aside className="grid items-start gap-5 lg:grid-cols-3">
            <Card className="shadow-sm">
              <CardHeader className="border-b"><CardTitle className="flex items-center gap-2"><Factory className="size-4 text-primary"/>基础工况</CardTitle><CardDescription>输入边界后即刻重新计算。</CardDescription></CardHeader>
              <CardContent className="grid grid-cols-2 gap-4 pt-1">
                <Field label="项目名称" wide><Input value={config.projectName} onChange={(e) => update("projectName", e.target.value)} /></Field>
                <Field label="设备型号"><Input value={config.machineModel} onChange={(e) => update("machineModel", e.target.value)} /></Field>
                <Field label="材料"><Input value={config.material} onChange={(e) => update("material", e.target.value)} /></Field>
                <Field label="进线直径" unit="mm"><Input type="number" min="0.01" step="0.01" value={config.entryDiameter} onChange={(e) => update("entryDiameter", number(e.target.value))} /></Field>
                <Field label="出线直径" unit="mm"><Input type="number" min="0.01" step="0.01" value={config.finalDiameter} onChange={(e) => update("finalDiameter", number(e.target.value))} /></Field>
                <Field label="成品线速" unit="m/min"><Input type="number" min="0.01" step="0.1" value={Number((config.finalSpeed*60).toFixed(6))} onChange={(e) => update("finalSpeed", number(e.target.value)/60)} /></Field>
                <Field label="进线强度" unit="MPa"><Input type="number" min="1" step="10" value={config.incomingStrength} onChange={(e) => update("incomingStrength", number(e.target.value))} /></Field>
                <Field label="最大允许总长" unit="mm"><Input type="number" min="0" step="10" value={config.maxOverallLength} onChange={(e) => update("maxOverallLength", number(e.target.value))} /></Field>
              </CardContent>
            </Card>

            <Card className="shadow-sm">
              <CardHeader className="border-b"><div className="flex items-start justify-between gap-3"><div><CardTitle className="flex items-center gap-2"><Activity className="size-4 text-primary"/>道次工艺</CardTitle><CardDescription>1～128 道自由配置；末道对应成品线速。</CardDescription></div><div className="flex gap-1"><Button variant="outline" size="icon-sm" aria-label="减少一道" disabled={config.passes.length <= 1} onClick={() => distribute(config.passes.length - 1)}><Minus /></Button><Button variant="outline" size="icon-sm" aria-label="增加一道" disabled={config.passes.length >= 128} onClick={() => distribute(config.passes.length + 1)}><Plus /></Button></div></div></CardHeader>
              <CardContent className="space-y-2 pt-1">
                <div className="mb-3 flex items-center justify-between rounded-lg bg-muted/65 px-3 py-2"><span className="text-xs text-muted-foreground">当前 <b className="text-foreground">{config.passes.length} 道</b></span><Button variant="ghost" size="xs" onClick={() => distribute()}><Sparkles />均匀分配压缩率</Button></div>
                <div className="grid grid-cols-[36px_1fr_1fr] gap-2 px-1 text-[11px] font-medium text-muted-foreground"><span>道次</span><span>丝径 mm</span><span>卷筒 mm</span></div>
                {config.passes.map((pass, index) => (
                  <div key={index} className="rounded-lg border p-2"><div className="grid grid-cols-[36px_1fr_1fr] items-center gap-2"><span className="grid size-7 place-items-center rounded-lg bg-muted text-sm font-semibold">{index + 1}</span><Input aria-label={`第${index + 1}道丝径 mm`} type="number" min="0.001" step="0.01" value={pass.diameter} onChange={(e) => update("passes", config.passes.map((item, i) => i === index ? { ...item, diameter: number(e.target.value) } : item))} /><Input aria-label={`第${index + 1}道卷筒 mm`} type="number" min="1" step="10" value={pass.drumDiameter} onChange={(e) => update("passes", config.passes.map((item, i) => i === index ? { ...item, drumDiameter: number(e.target.value) } : item))} /></div><div className="mt-2 grid grid-cols-2 gap-2"><label htmlFor={`force-${index}`} className="text-sm">实测拉拔力 N<Input id={`force-${index}`} type="number" min="0" placeholder="留空按公式计算" value={pass.measuredForce ?? ''} onChange={e => update('passes', config.passes.map((p, i) => i === index ? { ...p, measuredForce: e.target.value === '' ? undefined : Number(e.target.value) } : p))}/></label><label htmlFor={`factor-${index}`} className="text-sm">独立安全系数<Input id={`factor-${index}`} type="number" min="0.01" step="0.05" placeholder="留空使用全局系数" value={pass.safetyFactor ?? ''} onChange={e => update('passes', config.passes.map((p, i) => i === index ? { ...p, safetyFactor: e.target.value === '' ? undefined : Number(e.target.value) } : p))}/></label></div></div>
                ))}
              </CardContent>
            </Card>

            <details id="advanced-settings" className="engineering-details rounded-xl border bg-card shadow-sm" open>
              <summary><span><SlidersHorizontal className="size-4 text-primary"/><b>工程参数</b><small>系数、电气与校核边界</small></span><ChevronDown className="size-4"/></summary>
              <div className="grid grid-cols-2 gap-4 border-t p-4">
                <Field label="首道安全系数"><Input type="number" min="0.1" max="128" step="0.05" value={config.firstPassFactor} onChange={(e) => updateCustom("firstPassFactor", number(e.target.value))} /></Field>
                <Field label="后续道安全系数"><Input type="number" min="0.1" max="5" step="0.05" value={config.steadyPassFactor} onChange={(e) => updateCustom("steadyPassFactor", number(e.target.value))} /></Field>
                <Field label="跳头覆盖道数"><Input type="number" min="0" max="5" step="1" value={config.skipHeadSpan} onChange={(e) => updateCustom("skipHeadSpan", Math.round(number(e.target.value)))} /></Field>
                <Field label="增强系数 K"><Input type="number" min="0.5" max="2" step="0.01" value={config.strengthFactor} onChange={(e) => updateCustom("strengthFactor", number(e.target.value))} /></Field>
                <Field label="拉拔力系数"><Input type="number" min="0.1" max="2" step="0.01" value={config.forceCoefficient} onChange={(e) => updateCustom("forceCoefficient", number(e.target.value))} /></Field>
                <Field label="额定电压" unit="V"><Input type="number" min="1" step="10" value={config.voltage} onChange={(e) => update("voltage", number(e.target.value))} /></Field>
                <Field label="效率 η"><Input type="number" min="0.1" max="1" step="0.01" value={config.efficiency} onChange={(e) => update("efficiency", number(e.target.value))} /></Field>
                <Field label="功率因数"><Input type="number" min="0.1" max="1" step="0.01" value={config.powerFactor} onChange={(e) => update("powerFactor", number(e.target.value))} /></Field>
                <Field label="电压利用率"><Input type="number" min="0.1" max="1" step="0.01" value={config.voltageUtilization} onChange={(e) => update("voltageUtilization", number(e.target.value))} /></Field>
                <Field label="型号系列"><Input value={config.series} onChange={(e) => update("series", e.target.value.toUpperCase())} /></Field>
                <Field label="压缩率下限"><Input type="number" min="0" max="1" step="0.01" value={config.compressionMin} onChange={(e) => update("compressionMin", number(e.target.value))} /></Field>
                <Field label="压缩率上限"><Input type="number" min="0" max="1" step="0.01" value={config.compressionMax} onChange={(e) => update("compressionMax", number(e.target.value))} /></Field>
                <Field label="低速阈值" unit="r/min"><Input type="number" min="0" step="5" value={config.lowSpeedThreshold} onChange={(e) => update("lowSpeedThreshold", number(e.target.value))} /></Field>
                <Field label="低速最小中心高"><Input type="number" min="0" step="5" value={config.lowSpeedMinFrame} onChange={(e) => update("lowSpeedMinFrame", number(e.target.value))} /></Field>
                <Field label="标准转矩库" unit="N·m" wide><Input value={config.torqueCatalog.join(", ")} onChange={(e) => update("torqueCatalog", parseCatalog(e.target.value))} /></Field>
                <Field label="标准转速库" unit="r/min" wide><Input value={config.speedCatalog.join(", ")} onChange={(e) => update("speedCatalog", parseCatalog(e.target.value))} /></Field>
              </div>
            </details>
          </aside>

        </div>
        <div className="mt-5 flex justify-between"><Button variant="outline" disabled={phase==='input'} onClick={()=>setPhase(phase==='motor'?'demand':'input')}>上一步</Button><Button disabled={phase==='motor'} onClick={()=>setPhase(phase==='input'?'demand':'motor')}>{phase==='input'?'计算各道需求':phase==='demand'?'细化机座与极数':'已到选型参数表'}</Button></div>
      </div>
    </main>
  );
}

function Field({ label, unit, wide, children }: { label: string; unit?: string; wide?: boolean; children: ReactNode }) {
  return <label className={wide ? "col-span-2 space-y-1.5" : "space-y-1.5"}><span className="flex justify-between text-xs font-medium"><span>{label}</span>{unit && <span className="font-normal text-muted-foreground">{unit}</span>}</span>{children}</label>;
}
