"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Activity,
  BadgeCheck,
  Calculator,
  ChevronDown,
  ChevronRight,
  CircleGauge,
  Download,
  Factory,
  FileSpreadsheet,
  Gauge,
  Minus,
  Plus,
  RefreshCcw,
  Ruler,
  Save,
  ShieldAlert,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  Waves,
  Zap,
} from "lucide-react";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { downloadSelectionExcel } from "@/lib/excel-export";
import { ProcessWorkbench } from "@/components/process-workbench";
import {
  applyPreset,
  calculateSelection,
  defaultConfig,
  equalCompressionPasses,
  lz7Passes,
  selectionWarnings,
  type MachinePreset,
  type SelectionConfig,
  type SelectionRow,
} from "@/lib/motor-selection";

const STORAGE_KEY = "direct-drive-selector:v1";
const number = (value: string) => Number(value) || 0;
const fmt = (value: number, digits = 1) => value.toLocaleString("zh-CN", {
  minimumFractionDigits: digits,
  maximumFractionDigits: digits,
});
const percent = (value: number) => `${(value * 100).toFixed(2)}%`;
const cloneDefaults = (): SelectionConfig => JSON.parse(JSON.stringify(defaultConfig)) as SelectionConfig;
const parseCatalog = (value: string) => value.split(/[，,\s]+/).map(Number).filter((item) => Number.isFinite(item) && item > 0);

export default function Home() {
  const [config, setConfig] = useState<SelectionConfig>(() => cloneDefaults());
  const rows = useMemo(() => calculateSelection(config), [config]);
  const warnings = useMemo(() => selectionWarnings(config, rows), [config, rows]);
  const configRef = useRef(config);
  const rowsRef = useRef(rows);

  const maxTorque = Math.max(0, ...rows.map((row) => row.ratedTorque));
  const maxPower = Math.max(0, ...rows.map((row) => row.ratedPower));
  const totalPower = rows.reduce((sum, row) => sum + row.ratedPower, 0);
  const maxCurrent = Math.max(0, ...rows.map((row) => row.ratedCurrent));
  const lengthPass = rows.every((row) => row.lengthStatus !== "fail");

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const restored = { ...cloneDefaults(), ...(JSON.parse(stored) as Partial<SelectionConfig>) };
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
            passDiameters: { type: "array", items: { type: "number", exclusiveMinimum: 0 }, minItems: 1, maxItems: 16 },
            drumDiameters: { type: "array", items: { type: "number", exclusiveMinimum: 0 }, minItems: 1, maxItems: 16 },
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
          const result = calculateSelection(next);
          setConfig(next);
          return { projectName: next.projectName, passes: result.length, maxRatedTorque: Math.max(...result.map((row) => row.ratedTorque)), maxRatedPower: Math.max(...result.map((row) => row.ratedPower)) };
        },
      }, { signal: lifecycle.signal });
      await context.registerTool({
        name: "read_motor_selection_result",
        title: "读取直驱电机选型结果",
        description: "读取当前页面已经计算出的电机型号和关键额定参数，不修改工况。",
        inputSchema: { type: "object", properties: {}, additionalProperties: false },
        annotations: { readOnlyHint: true, untrustedContentHint: false },
        execute() {
          return { projectName: configRef.current.projectName, preset: configRef.current.preset, rows: rowsRef.current.map((row) => ({ pass: row.pass, motorModel: row.motorModel, ratedTorque: row.ratedTorque, ratedRpm: row.ratedRpm, ratedPower: Number(row.ratedPower.toFixed(2)), lengthStatus: row.lengthStatus })) };
        },
      }, { signal: lifecycle.signal });
    };
    void register().catch(() => undefined);
    return () => lifecycle.abort();
  }, []);

  const setPreset = (preset: MachinePreset) => setConfig((old) => applyPreset(old, preset));
  const update = <K extends keyof SelectionConfig>(key: K, value: SelectionConfig[K]) => setConfig((old) => ({ ...old, [key]: value }));
  const updateCustom = <K extends keyof SelectionConfig>(key: K, value: SelectionConfig[K]) => setConfig((old) => ({ ...old, preset: "custom", [key]: value }));
  const distribute = (count = config.passes.length) => update("passes", equalCompressionPasses(count, config.entryDiameter, config.finalDiameter, config.passes.at(-1)?.drumDiameter || 560));
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
            <Button variant="outline" size="sm" onClick={reset}><RefreshCcw />重置示例</Button>
            <Button size="sm" onClick={() => downloadSelectionExcel(config, rows, warnings)} disabled={!rows.length}><Download />导出 Excel</Button>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1760px] px-4 py-6 sm:px-6 lg:px-8">
        <section className="mb-5 flex flex-col justify-between gap-4 lg:flex-row lg:items-end">
          <div>
            <div className="mb-2 flex items-center gap-2 text-xs font-medium text-primary"><span>工况定义</span><ChevronRight className="size-3.5"/><span>拉拔计算</span><ChevronRight className="size-3.5"/><span>电机初选</span></div>
            <h1 className="font-heading text-2xl font-semibold tracking-tight sm:text-3xl">从拉拔工艺直接得到电机规格</h1>
            <p className="mt-2 max-w-3xl text-sm leading-6 text-muted-foreground">按附件模型联动计算压缩率、拉拔力、转矩、转速与功率；工艺、系数、标准档位和尺寸库均可调整。</p>
          </div>
          <div className="flex flex-wrap gap-1.5 rounded-xl border bg-card p-1.5 shadow-sm">
            {(["standard", "universal", "custom"] as MachinePreset[]).map((preset) => (
              <Button key={preset} variant={config.preset === preset ? "default" : "ghost"} size="sm" onClick={() => setPreset(preset)}>
                {preset === "standard" ? "标准专机 1.4 / 1.3" : preset === "universal" ? "万能机 1.4 / 1.2" : "完全自定义"}
              </Button>
            ))}
          </div>
        </section>

        <ProcessWorkbench config={config} onLoad={setConfig} />
        <div className="grid gap-5 xl:grid-cols-[410px_minmax(0,1fr)]">
          <aside className="space-y-5">
            <Card className="shadow-sm">
              <CardHeader className="border-b"><CardTitle className="flex items-center gap-2"><Factory className="size-4 text-primary"/>基础工况</CardTitle><CardDescription>输入边界后即刻重新计算。</CardDescription></CardHeader>
              <CardContent className="grid grid-cols-2 gap-4 pt-1">
                <Field label="项目名称" wide><Input value={config.projectName} onChange={(e) => update("projectName", e.target.value)} /></Field>
                <Field label="设备型号"><Input value={config.machineModel} onChange={(e) => update("machineModel", e.target.value)} /></Field>
                <Field label="材料"><Input value={config.material} onChange={(e) => update("material", e.target.value)} /></Field>
                <Field label="进线直径" unit="mm"><Input type="number" min="0.01" step="0.01" value={config.entryDiameter} onChange={(e) => update("entryDiameter", number(e.target.value))} /></Field>
                <Field label="出线直径" unit="mm"><Input type="number" min="0.01" step="0.01" value={config.finalDiameter} onChange={(e) => update("finalDiameter", number(e.target.value))} /></Field>
                <Field label="成品线速" unit="m/s"><Input type="number" min="0.01" step="0.1" value={config.finalSpeed} onChange={(e) => update("finalSpeed", number(e.target.value))} /></Field>
                <Field label="进线强度" unit="MPa"><Input type="number" min="1" step="10" value={config.incomingStrength} onChange={(e) => update("incomingStrength", number(e.target.value))} /></Field>
                <Field label="最大允许总长" unit="mm"><Input type="number" min="0" step="10" value={config.maxOverallLength} onChange={(e) => update("maxOverallLength", number(e.target.value))} /></Field>
              </CardContent>
            </Card>

            <Card className="shadow-sm">
              <CardHeader className="border-b"><div className="flex items-start justify-between gap-3"><div><CardTitle className="flex items-center gap-2"><Activity className="size-4 text-primary"/>道次工艺</CardTitle><CardDescription>1～16 道自由配置；末道对应成品线速。</CardDescription></div><div className="flex gap-1"><Button variant="outline" size="icon-sm" aria-label="减少一道" disabled={config.passes.length <= 1} onClick={() => distribute(config.passes.length - 1)}><Minus /></Button><Button variant="outline" size="icon-sm" aria-label="增加一道" disabled={config.passes.length >= 16} onClick={() => distribute(config.passes.length + 1)}><Plus /></Button></div></div></CardHeader>
              <CardContent className="space-y-2 pt-1">
                <div className="mb-3 flex items-center justify-between rounded-lg bg-muted/65 px-3 py-2"><span className="text-xs text-muted-foreground">当前 <b className="text-foreground">{config.passes.length} 道</b></span><Button variant="ghost" size="xs" onClick={() => distribute()}><Sparkles />均匀分配压缩率</Button></div>
                <div className="grid grid-cols-[36px_1fr_1fr] gap-2 px-1 text-[11px] font-medium text-muted-foreground"><span>道次</span><span>丝径 mm</span><span>卷筒 mm</span></div>
                {config.passes.map((pass, index) => (
                  <div key={index} className="grid grid-cols-[36px_1fr_1fr] items-center gap-2"><span className="grid size-7 place-items-center rounded-lg bg-muted text-xs font-semibold">{index + 1}</span><Input type="number" min="0.001" step="0.01" value={pass.diameter} onChange={(e) => update("passes", config.passes.map((item, i) => i === index ? { ...item, diameter: number(e.target.value) } : item))} /><Input type="number" min="1" step="10" value={pass.drumDiameter} onChange={(e) => update("passes", config.passes.map((item, i) => i === index ? { ...item, drumDiameter: number(e.target.value) } : item))} /></div>
                ))}
                <Button variant="outline" className="mt-2 w-full" size="sm" onClick={() => update("passes", lz7Passes.map((item) => ({ ...item })))}><RefreshCcw />恢复附件 LZ7 六道次</Button>
              </CardContent>
            </Card>

            <details id="advanced-settings" className="engineering-details rounded-xl border bg-card shadow-sm" open>
              <summary><span><SlidersHorizontal className="size-4 text-primary"/><b>工程参数</b><small>系数、电气与校核边界</small></span><ChevronDown className="size-4"/></summary>
              <div className="grid grid-cols-2 gap-4 border-t p-4">
                <Field label="首道安全系数"><Input type="number" min="0.1" max="5" step="0.05" value={config.firstPassFactor} onChange={(e) => updateCustom("firstPassFactor", number(e.target.value))} /></Field>
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

          <section className="min-w-0 space-y-5">
            <div className="grid gap-3 sm:grid-cols-2 2xl:grid-cols-5">
              <Metric icon={<CircleGauge/>} label="最大额定转矩" value={`${fmt(maxTorque, 0)} N·m`} note={`覆盖 ${rows.length} 道工况`} />
              <Metric icon={<Zap/>} label="最大额定功率" value={`${fmt(maxPower, 1)} kW`} note={`装机合计 ${fmt(totalPower, 1)} kW`} />
              <Metric icon={<Gauge/>} label="最大额定电流" value={`${fmt(maxCurrent, 1)} A`} note={`${config.voltage} V · η ${(config.efficiency * 100).toFixed(0)}%`} />
              <Metric icon={<Ruler/>} label="尺寸预检" value={lengthPass ? "通过" : "有超限"} note={`允许总长 ${fmt(config.maxOverallLength, 0)} mm`} tone={lengthPass ? "green" : "amber"} />
              <Metric icon={<BadgeCheck/>} label="系数方案" value={`${config.firstPassFactor.toFixed(2)} / ${config.steadyPassFactor.toFixed(2)}`} note={config.preset === "universal" ? `万能跳头 ${config.skipHeadSpan} 道` : config.preset === "standard" ? "标准专机" : "自定义"} />
            </div>

            <div className={`status-strip ${warnings.length ? "warning" : "ok"}`}><span>{warnings.length ? <ShieldAlert/> : <ShieldCheck/>}</span><div><b>{warnings.length ? `${warnings.length} 项需要复核` : "当前工况通过自动校核"}</b><p>{warnings.length ? warnings.join("；") : "丝径递减、压缩率、标准档位与总长预检均未发现异常。"}</p></div></div>

            <Tabs defaultValue="selection" className="gap-4">
              <div className="flex flex-col justify-between gap-3 sm:flex-row sm:items-center">
                <TabsList><TabsTrigger value="selection"><FileSpreadsheet/>选型参数表</TabsTrigger><TabsTrigger value="calculation"><Calculator/>工艺计算明细</TabsTrigger><TabsTrigger value="frames"><Ruler/>机座尺寸库</TabsTrigger></TabsList>
                <Button onClick={() => downloadSelectionExcel(config, rows, warnings)} disabled={!rows.length}><Download/>导出当前方案</Button>
              </div>

              <TabsContent value="selection">
                <Card className="shadow-sm"><CardHeader className="border-b"><div className="flex flex-col justify-between gap-3 sm:flex-row sm:items-start"><div><CardTitle>最终选型参数表</CardTitle><CardDescription>{config.projectName} · {config.machineModel} · {config.passes.length} 道 · {config.preset === "universal" ? "万能机" : config.preset === "standard" ? "标准专机" : "自定义方案"}</CardDescription></div><Badge variant="outline" className="gap-1.5"><ShieldCheck/>型号按额定档位生成</Badge></div></CardHeader><CardContent className="px-0"><SelectionTable rows={rows} config={config}/></CardContent></Card>
              </TabsContent>

              <TabsContent value="calculation">
                <Card className="shadow-sm"><CardHeader className="border-b"><CardTitle>拉拔计算明细</CardTitle><CardDescription>每一步均保留输入、计算值和安全系数，便于工程复核。</CardDescription></CardHeader><CardContent className="px-0"><CalculationTable rows={rows}/></CardContent></Card>
              </TabsContent>

              <TabsContent value="frames">
                <Card className="shadow-sm"><CardHeader className="border-b"><CardTitle>机座与总长初始库</CardTitle><CardDescription>这些是可编辑的方案预检值；最终尺寸必须以厂家外形图为准。</CardDescription></CardHeader><CardContent className="px-0"><Table><TableHeader className="bg-muted/55"><TableRow><TableHead className="pl-4">中心高</TableHead><TableHead>最大匹配转矩 N·m</TableHead><TableHead>预计总长 mm</TableHead><TableHead>默认极数</TableHead><TableHead>预检状态</TableHead></TableRow></TableHeader><TableBody>{config.frameRules.map((rule, index) => <TableRow key={`${rule.frame}-${index}`}><TableCell className="pl-4"><Input className="w-24" type="number" value={rule.frame} onChange={(e) => update("frameRules", config.frameRules.map((item, i) => i === index ? { ...item, frame: number(e.target.value) } : item))}/></TableCell><TableCell><Input className="w-36" type="number" value={rule.maxTorque} onChange={(e) => update("frameRules", config.frameRules.map((item, i) => i === index ? { ...item, maxTorque: number(e.target.value) } : item))}/></TableCell><TableCell><Input className="w-32" type="number" value={rule.estimatedLength} onChange={(e) => update("frameRules", config.frameRules.map((item, i) => i === index ? { ...item, estimatedLength: number(e.target.value) } : item))}/></TableCell><TableCell><Input className="w-24" type="number" value={rule.poles} onChange={(e) => update("frameRules", config.frameRules.map((item, i) => i === index ? { ...item, poles: number(e.target.value) } : item))}/></TableCell><TableCell><Badge variant={!config.maxOverallLength || rule.estimatedLength <= config.maxOverallLength ? "secondary" : "destructive"}>{!config.maxOverallLength ? "未校核" : rule.estimatedLength <= config.maxOverallLength ? "可布置" : "超总长"}</Badge></TableCell></TableRow>)}</TableBody></Table></CardContent></Card>
              </TabsContent>
            </Tabs>

            <div className="grid gap-4 lg:grid-cols-2">
              <Card size="sm"><CardHeader><CardTitle className="flex items-center gap-2"><Calculator className="size-4 text-primary"/>计算依据</CardTitle></CardHeader><CardContent className="space-y-2 text-xs leading-6 text-muted-foreground"><p><b className="text-foreground">速度：</b>按钢丝截面积守恒逐道换算。</p><p><b className="text-foreground">强度：</b>σ = σ₀ × √(d₀ / d) × K。</p><p><b className="text-foreground">拉拔力：</b>0.6 × d前² × 平均强度 × √压缩率。</p><p><b className="text-foreground">功率：</b>P = T × n / 9550。</p></CardContent></Card>
              <Card size="sm"><CardHeader><CardTitle className="flex items-center gap-2"><ShieldAlert className="size-4 text-amber-600"/>工程边界</CardTitle></CardHeader><CardContent className="space-y-2 text-xs leading-6 text-muted-foreground"><p>标准档位、机座能力和预计总长均可编辑；内置值用于方案初选，不替代温升、过载、弱磁、轴承、冷却和结构强度校核。</p><p>导出的 Excel 同时包含选型表、计算明细、参数设定和机座尺寸库，便于审签追溯。</p></CardContent></Card>
            </div>
          </section>
        </div>
      </div>
    </main>
  );
}

function Field({ label, unit, wide, children }: { label: string; unit?: string; wide?: boolean; children: ReactNode }) {
  return <label className={wide ? "col-span-2 space-y-1.5" : "space-y-1.5"}><span className="flex justify-between text-xs font-medium"><span>{label}</span>{unit && <span className="font-normal text-muted-foreground">{unit}</span>}</span>{children}</label>;
}

function Metric({ icon, label, value, note, tone = "blue" }: { icon: ReactNode; label: string; value: string; note: string; tone?: "blue" | "green" | "amber" }) {
  return <Card size="sm" className="shadow-sm"><CardContent className="flex items-start gap-3"><span className={`metric-icon metric-${tone}`}>{icon}</span><div className="min-w-0"><p className="text-xs text-muted-foreground">{label}</p><p className="mt-1 truncate font-heading text-xl font-semibold tracking-tight">{value}</p><p className="mt-1 text-[11px] text-muted-foreground">{note}</p></div></CardContent></Card>;
}

function SelectionTable({ rows, config }: { rows: SelectionRow[]; config: SelectionConfig }) {
  return <Table><TableHeader className="bg-muted/55"><TableRow><TableHead className="pl-4">道次</TableHead><TableHead>电机型号</TableHead><TableHead>机座</TableHead><TableHead>额定功率</TableHead><TableHead>额定转速</TableHead><TableHead>额定转矩</TableHead><TableHead>电流</TableHead><TableHead>频率</TableHead><TableHead>极数</TableHead><TableHead>冷却</TableHead><TableHead>预计总长</TableHead><TableHead>跳头覆盖</TableHead></TableRow></TableHeader><TableBody>{rows.map((row) => <TableRow key={row.pass}><TableCell className="pl-4 font-semibold">{row.pass}</TableCell><TableCell><div className="font-mono text-xs font-semibold text-primary">{row.motorModel}</div><div className="mt-1 text-[11px] text-muted-foreground">{fmt(row.diameter, 2)} mm · {fmt(row.lineSpeed, 2)} m/s</div></TableCell><TableCell>{row.frame}</TableCell><TableCell>{fmt(row.ratedPower, 1)} kW</TableCell><TableCell>{fmt(row.ratedRpm, 0)} r/min</TableCell><TableCell>{fmt(row.ratedTorque, 0)} N·m</TableCell><TableCell>{fmt(row.ratedCurrent, 1)} A</TableCell><TableCell>{fmt(row.frequency, 1)} Hz</TableCell><TableCell>{row.poles}</TableCell><TableCell><span className="text-xs">{config.cooling}</span><div className="text-[10px] text-muted-foreground">{config.protection} · {config.insulation}</div></TableCell><TableCell><Badge variant={row.lengthStatus === "fail" ? "destructive" : "secondary"}>{row.estimatedLength} mm</Badge></TableCell><TableCell>{row.reserveSourcePass === row.pass ? <span className="text-xs text-muted-foreground">本工位</span> : <Badge variant="outline">覆盖第 {row.reserveSourcePass} 道</Badge>}</TableCell></TableRow>)}</TableBody></Table>;
}

function CalculationTable({ rows }: { rows: SelectionRow[] }) {
  return <Table><TableHeader className="bg-muted/55"><TableRow><TableHead className="pl-4">道次</TableHead><TableHead>丝径前→后</TableHead><TableHead>卷筒</TableHead><TableHead>抗拉强度前→后</TableHead><TableHead>速度</TableHead><TableHead>压缩率</TableHead><TableHead>拉拔力</TableHead><TableHead>拉拔转矩</TableHead><TableHead>系数</TableHead><TableHead>设计转矩</TableHead><TableHead>设计转速</TableHead></TableRow></TableHeader><TableBody>{rows.map((row) => <TableRow key={row.pass}><TableCell className="pl-4 font-semibold">{row.pass}</TableCell><TableCell>{fmt(row.incomingDiameter, 2)} → {fmt(row.diameter, 2)} mm</TableCell><TableCell>{fmt(row.drumDiameter, 0)} mm</TableCell><TableCell>{fmt(row.strengthBefore, 1)} → {fmt(row.strengthAfter, 1)} MPa</TableCell><TableCell>{fmt(row.lineSpeed, 2)} m/s</TableCell><TableCell><Badge variant={row.compression > 0 && row.compression <= .35 ? "secondary" : "destructive"}>{percent(row.compression)}</Badge></TableCell><TableCell>{fmt(row.drawingForce, 1)} N</TableCell><TableCell>{fmt(row.drawingTorque, 1)} N·m</TableCell><TableCell>{row.factor.toFixed(2)}</TableCell><TableCell className="font-semibold text-primary">{fmt(row.designTorque, 1)} N·m</TableCell><TableCell>{fmt(row.designRpm, 1)} r/min</TableCell></TableRow>)}</TableBody></Table>;
}
