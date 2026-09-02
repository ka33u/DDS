import type { SelectionConfig, SelectionRow } from "./motor-selection";

type Cell = string | number | null;

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
