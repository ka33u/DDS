import type { SelectionConfig, SelectionRow } from "./motor-selection";
import type { ProcessCase, CommonMotor } from './process-selection';

type Cell = string | number | null;

export function buildProcessExcel(config: SelectionConfig, cases: ProcessCase[], motors: CommonMotor[]) {
  return workbook([
    sheet('兼容选型参数表', [
      ['项目', '工位', '型号', '状态', '转矩 N·m', '转速 r/min', '功率 kW', '电流 A', '电压 V', '频率 Hz', '中心高', '极数', '总长 mm', '转矩决定工况', '转速决定工况'],
      ...motors.map(m => [config.projectName, m.station, m.model, m.status, m.ratedTorque, m.ratedRpm, m.power, m.current, config.voltage, m.frequency, m.frame, m.poles, m.length, m.torqueSource, m.speedSource]),
    ], [150, 50, 220, 220, 90, 90, 90, 90, 70, 80, 70, 60, 80, 240, 240]),
    sheet('工况覆盖明细', [
      ['工艺', '工位', '道次', '用途', '拉力 N', '卷筒 mm', '需求转矩 N·m', '需求转速 r/min', '转矩余量 N·m', '转速余量 r/min'],
      ...motors.flatMap(m => m.loads.map(p => [p.caseName, p.station, p.pass, p.kind, p.force, p.drum, p.torque, p.rpm, m.ratedTorque === null ? null : m.ratedTorque - p.torque, m.ratedRpm === null ? null : m.ratedRpm - p.rpm])),
    ]),
    sheet('工艺输入', [
      ['工艺', '启用', '进径 mm', '成品车速 m/s', '进线强度 MPa', '首道系数', '后道系数', 'K', '拉力系数', '跳头跨度', '调模 m/s', '调模系数', '调模拉力 N', '附加转矩 N·m', '工位', '丝径 mm', '卷筒 mm'],
      ...cases.flatMap(c => c.config.passes.map((p, i) => [c.name, c.enabled ? '是' : '否', c.config.entryDiameter, c.config.finalSpeed, c.config.incomingStrength, c.config.firstPassFactor, c.config.steadyPassFactor, c.config.strengthFactor, c.config.forceCoefficient, c.config.skipHeadSpan, c.jogSpeed, c.jogFactor, c.jogForce, c.additionalTorque, c.stations[i], p.diameter, p.drumDiameter])),
    ]),
    sheet('设备选型依据', [
      ['参数', '数值'], ['转矩档位 N·m', config.torqueCatalog.join(', ')], ['转速档位 r/min', config.speedCatalog.join(', ')], ['允许总长 mm', config.maxOverallLength], ['效率', config.efficiency], ['功率因数', config.powerFactor], ['电压利用率', config.voltageUtilization], ['低速阈值', config.lowSpeedThreshold], ['低速最小中心高', config.lowSpeedMinFrame],
      ['选择原则', '生产、跳头和调模逐工位取最大转矩和最大转速；向上选最小标准档位；选择满足转矩和长度的最小中心高。均按连续恒转矩，不假设弱磁或短时过载。'],
      ['调模拉力为零', '使用该道计算拉拔力；调模转矩再乘调模系数并加附加阻力转矩'],
      ['型号与尺寸', '参数组合建议；尺寸为可编辑初始库，按厂家外形图确认'],
    ], [160, 500]),
    sheet('机座库', [['中心高', '转矩能力 N·m', '预计总长 mm', '极数'], ...config.frameRules.map(f => [f.frame, f.maxTorque, f.estimatedLength, f.poles])]),
  ]);
}

export function downloadProcessExcel(config: SelectionConfig, cases: ProcessCase[], motors: CommonMotor[]) {
  const url = URL.createObjectURL(new Blob(['\uFEFF', buildProcessExcel(config, cases, motors)], { type: 'application/vnd.ms-excel;charset=utf-8' }));
  const a = document.createElement('a'); a.href = url; a.download = `${safeName(config.projectName)}_兼容选型.xls`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const escapeXml = (value: Cell) => String(value ?? "")
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&apos;");

const safeName = (value: string) => value.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "-").slice(0, 60);

function cell(value: Cell, style = "Body") {
  const numeric = typeof value === "number" && Number.isFinite(value);
  return `<Cell ss:StyleID="${style}"><Data ss:Type="${numeric ? "Number" : "String"}">${escapeXml(value)}</Data></Cell>`;
}

function sheet(name: string, rows: Cell[][], widths: number[] = []) {
  const columns = widths.map((width) => `<Column ss:AutoFitWidth="0" ss:Width="${width}"/>`).join("");
  return `<Worksheet ss:Name="${escapeXml(name.slice(0, 31))}"><Table>${columns}${rows.map((row, index) => `<Row>${row.map((value) => cell(value, index === 0 ? "Header" : "Body")).join("")}</Row>`).join("")}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>`;
}

function workbook(sheets: string[]) {
  return `<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>拉丝机直驱电机选型系统</Author><Created>${new Date().toISOString()}</Created></DocumentProperties>
<Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Microsoft YaHei" ss:Size="10"/><Borders/><Interior/><NumberFormat/><Protection/></Style>
  <Style ss:ID="Header"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Microsoft YaHei" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0F5F68" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0A454C"/></Borders></Style>
  <Style ss:ID="Body"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Microsoft YaHei" ss:Size="10"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#DDE5E6"/></Borders></Style>
</Styles>${sheets.join("")}</Workbook>`;
}

export function downloadSelectionExcel(config: SelectionConfig, rows: SelectionRow[], warnings: string[]) {
  const selectionRows: Cell[][] = [
    ["道次", "系列", "中心高", "电机型号", "额定功率(kW)", "额定转速(r/min)", "额定转矩(N·m)", "额定电压(VAC)", "额定电流(A)", "额定频率(Hz)", "效率", "功率因数", "极数", "防护等级", "绝缘等级", "工作制", "接法", "安装方式", "冷却方式", "预计总长(mm)", "总长校核", "跳头覆盖来源"],
    ...rows.map((row) => [row.pass, config.series, row.frame, row.motorModel, row.ratedPower, row.ratedRpm, row.ratedTorque, config.voltage, row.ratedCurrent, row.frequency, config.efficiency, config.powerFactor, row.poles, config.protection, config.insulation, config.duty, config.connection, config.mounting, config.cooling, row.estimatedLength, row.lengthStatus === "pass" ? "通过" : row.lengthStatus === "fail" ? "超限" : "未检查", row.reserveSourcePass === row.pass ? "本工位" : `第${row.reserveSourcePass}道`]),
  ];

  const calculationRows: Cell[][] = [
    ["道次", "拉拔前丝径(mm)", "拉拔后丝径(mm)", "卷筒直径(mm)", "抗拉强度前(MPa)", "抗拉强度后(MPa)", "拉拔速度(m/s)", "压缩率", "拉拔力(N)", "拉拔转矩(N·m)", "安全系数", "设计转矩(N·m)", "设计转速(r/min)"],
    ...rows.map((row) => [row.pass, row.incomingDiameter, row.diameter, row.drumDiameter, row.strengthBefore, row.strengthAfter, row.lineSpeed, row.compression, row.drawingForce, row.drawingTorque, row.factor, row.designTorque, row.designRpm]),
  ];

  const parameterRows: Cell[][] = [
    ["参数", "设定值", "说明"],
    ["项目名称", config.projectName, "用户输入"],
    ["设备型号", config.machineModel, "用户输入"],
    ["材料", config.material, "用户输入"],
    ["机型方案", config.preset === "standard" ? "标准专机" : config.preset === "universal" ? "万能跳头机" : "自定义", "决定系数和跳头覆盖"],
    ["进线直径(mm)", config.entryDiameter, "工况输入"],
    ["出线直径(mm)", config.finalDiameter, "工况输入"],
    ["成品线速(m/s)", config.finalSpeed, "末道线速度"],
    ["进线抗拉强度(MPa)", config.incomingStrength, "工况输入"],
    ["首道安全系数", config.firstPassFactor, "默认1.4"],
    ["后续道安全系数", config.steadyPassFactor, "标准专机默认1.3，万能机默认1.2"],
    ["跳头覆盖道数", config.skipHeadSpan, "0为专机；万能机默认1"],
    ["增强系数", config.strengthFactor, "附件经验范围0.95～1.05"],
    ["拉拔力系数", config.forceCoefficient, "附件模型取0.6"],
    ["最大允许总长(mm)", config.maxOverallLength, "0表示不校核"],
    ["压缩率校核范围", `${config.compressionMin * 100}%～${config.compressionMax * 100}%`, "可自定义"],
    ["标准转矩库(N·m)", config.torqueCatalog.join(", "), "可自定义"],
    ["标准转速库(r/min)", config.speedCatalog.join(", "), "可自定义"],
    ["冷却要求", `${config.cooling}；${config.waterPressure}；${config.inletTemperature}；${config.waterQuality}`, "工程要求"],
    ["校核提示", warnings.join("；") || "当前无阻断提示", "计算结果"],
    ["重要说明", "机座预计总长为可编辑初始库，仅用于方案预检，最终必须按厂家外形图核准。", "尺寸责任边界"],
    ["计算依据", "拉拔速度按面积守恒；强度按σ=σ0×√(d0/d)×K；拉拔力按0.6×d前²×平均强度×√q；转矩=力×卷筒直径/2000；功率=转矩×转速/9550。", "由附件公式归纳"],
  ];

  const frameRows: Cell[][] = [
    ["中心高", "最大匹配转矩(N·m)", "预计总长(mm)", "默认极数", "说明"],
    ...config.frameRules.map((rule) => [rule.frame, rule.maxTorque, rule.estimatedLength, rule.poles, "自定义初始库，最终按图样核准"]),
  ];

  const content = workbook([
    sheet("选型参数表", selectionRows, [42, 52, 58, 210, 78, 82, 86, 78, 76, 76, 54, 62, 42, 62, 62, 48, 45, 68, 62, 78, 64, 80]),
    sheet("拉拔计算", calculationRows, [42, 82, 82, 82, 90, 90, 78, 64, 84, 92, 64, 92, 86]),
    sheet("参数设定", parameterRows, [145, 340, 250]),
    sheet("机座尺寸库", frameRows, [75, 120, 100, 75, 220]),
  ]);
  const blob = new Blob(["\uFEFF", content], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${safeName(config.projectName || "直驱电机选型")}_选型参数表.xls`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
