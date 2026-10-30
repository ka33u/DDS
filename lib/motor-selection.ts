export type MachinePreset = "standard" | "universal" | "custom";

export interface DrawingPass {
  lineSpeed?: number;
  diameter: number;
  drumDiameter: number;
  measuredForce?: number;
  safetyFactor?: number;
}

export interface FrameRule {
  frame: number;
  maxTorque: number;
  estimatedLength: number;
  poles: number;
  supportedPoles?: number[];
}

export interface SelectionConfig {
  deviceName?: string;
  equipmentDrums?: number[];
  processStations?: number[];
  forceModel?: 'average' | 'incoming' | 'measured';
  strengthExponent?: number;
  drawingEfficiency?: number;
  parallelWires?: number;
  selectionMode?: 'custom' | 'reference';
  referenceRecipe?: string;
  maxFrequency?: number;
  stationChoices?: Record<number, { frame?: number; poles?: number }>;
  projectName: string;
  machineModel: string;
  material: string;
  preset: MachinePreset;
  entryDiameter: number;
  finalDiameter: number;
  finalSpeed: number;
  incomingStrength: number;
  strengthFactor: number;
  forceCoefficient: number;
  firstPassFactor: number;
  steadyPassFactor: number;
  skipHeadSpan: number;
  series: string;
  voltage: number;
  efficiency: number;
  powerFactor: number;
  voltageUtilization: number;
  protection: string;
  insulation: string;
  duty: string;
  cooling: string;
  connection: string;
  mounting: string;
  waterPressure: string;
  inletTemperature: string;
  waterQuality: string;
  maxOverallLength: number;
  compressionMin: number;
  compressionMax: number;
  catalogTolerance: number;
  lowSpeedThreshold: number;
  lowSpeedMinFrame: number;
  torqueCatalog: number[];
  speedCatalog: number[];
  frameRules: FrameRule[];
  passes: DrawingPass[];
}

export interface SelectionRow {
  pass: number;
  incomingDiameter: number;
  diameter: number;
  drumDiameter: number;
  compression: number;
  strengthBefore: number;
  strengthAfter: number;
  lineSpeed: number;
  drawingForce: number;
  drawingTorque: number;
  factor: number;
  designTorque: number;
  designRpm: number;
  ratedTorque: number;
  ratedRpm: number;
  ratedPower: number;
  ratedCurrent: number;
  frequency: number;
  frame: number;
  poles: number;
  estimatedLength: number;
  lengthStatus: "pass" | "fail" | "unchecked";
  motorModel: string;
  reserveSourcePass: number;
}

export const torqueCatalogDefault = [
  315, 400, 500, 630, 800, 1000, 1300, 1530, 1750, 2170, 2650, 3200,
  4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000,
];

export const speedCatalogDefault = [
  45, 60, 75, 90, 115, 140, 170, 200, 235, 275, 315, 355, 400, 450, 500,
  560, 630, 710, 800, 900, 1000,
];

export const frameRulesDefault: FrameRule[] = [
  { frame: 180, maxTorque: 400, estimatedLength: 760, poles: 16 },
  { frame: 200, maxTorque: 630, estimatedLength: 830, poles: 16 },
  { frame: 225, maxTorque: 800, estimatedLength: 910, poles: 16 },
  { frame: 250, maxTorque: 1300, estimatedLength: 1020, poles: 16 },
  { frame: 280, maxTorque: 2200, estimatedLength: 1180, poles: 16 },
  { frame: 315, maxTorque: 4000, estimatedLength: 1380, poles: 32 },
  { frame: 355, maxTorque: 6300, estimatedLength: 1580, poles: 32 },
  { frame: 400, maxTorque: 10000, estimatedLength: 1820, poles: 40 },
];

export const lz7Passes: DrawingPass[] = [
  { diameter: 5.05, drumDiameter: 560 },
  { diameter: 4.55, drumDiameter: 560 },
  { diameter: 4.13, drumDiameter: 560 },
  { diameter: 3.8, drumDiameter: 560 },
  { diameter: 3.5, drumDiameter: 560 },
  { diameter: 3.25, drumDiameter: 560 },
];

export const defaultConfig: SelectionConfig = {
  projectName: "河南恒星 LZ7-560 改造",
  machineModel: "LZ7/560",
  material: "钢丝",
  preset: "standard",
  entryDiameter: 5.5,
  finalDiameter: 3.25,
  finalSpeed: 8,
  incomingStrength: 1100,
  strengthFactor: 1,
  forceCoefficient: 0.6,
  firstPassFactor: 1.4,
  steadyPassFactor: 1.3,
  skipHeadSpan: 0,
  series: "YTDD",
  voltage: 380,
  efficiency: 0.9,
  powerFactor: 0.95,
  voltageUtilization: 0.92,
  protection: "IP55",
  insulation: "F",
  duty: "S1",
  cooling: "水冷",
  connection: "Y",
  mounting: "IMB3",
  waterPressure: "0.1～0.2 MPa",
  inletTemperature: "≤35 ℃",
  waterQuality: "防冻去离子水（水+乙二醇）",
  maxOverallLength: 1450,
  compressionMin: 0.05,
  compressionMax: 0.35,
  catalogTolerance: 0.005,
  lowSpeedThreshold: 200,
  lowSpeedMinFrame: 315,
  torqueCatalog: torqueCatalogDefault,
  speedCatalog: speedCatalogDefault,
  frameRules: frameRulesDefault,
  passes: lz7Passes,
};

const finite = (value: number, fallback = 0) =>
  Number.isFinite(value) ? value : fallback;

function catalogPick(required: number, catalog: number[], tolerance: number) {
  const sorted = [...catalog].filter((value) => value > 0).sort((a, b) => a - b);
  if (!sorted.length) return Math.ceil(required);
  return sorted.find((value) => required <= value * (1 + tolerance)) ?? Math.ceil(required);
}

function selectFrame(ratedTorque: number, ratedRpm: number, config: SelectionConfig) {
  const rules = config.frameRules;
  const sorted = [...rules].sort((a, b) => a.frame - b.frame);
  const minimumFrame = ratedRpm <= config.lowSpeedThreshold ? config.lowSpeedMinFrame : 0;
  return (
    sorted.find((rule) => rule.maxTorque * (1 + config.catalogTolerance) >= ratedTorque && rule.frame >= minimumFrame) ??
    sorted.at(-1) ?? { frame: 0, maxTorque: 0, estimatedLength: 0, poles: 16 }
  );
}

export function applyPreset(config: SelectionConfig, preset: MachinePreset) {
  if (preset === "standard") {
    return { ...config, preset, firstPassFactor: 1.4, steadyPassFactor: 1.3, skipHeadSpan: 0 };
  }
  if (preset === "universal") {
    return { ...config, preset, firstPassFactor: 1.4, steadyPassFactor: 1.2, skipHeadSpan: 1 };
  }
  return { ...config, preset };
}

export function equalCompressionPasses(count: number, entryDiameter: number, finalDiameter: number, drumDiameter: number) {
  const safeCount = Math.max(1, Math.min(128, Math.round(count)));
  const ratio = Math.pow(finalDiameter / entryDiameter, 1 / safeCount);
  return Array.from({ length: safeCount }, (_, index) => ({
    diameter: Number((entryDiameter * Math.pow(ratio, index + 1)).toPrecision(12)),
    drumDiameter,
  }));
}

export function calculateDrawingRequirements(config: SelectionConfig) {
  const activePasses = config.passes.slice(0, 128);
  if (!activePasses.length) return [];
  const actualFinalDiameter = activePasses.at(-1)!.diameter;

  return activePasses.map((pass, index) => {
    const incomingDiameter = index === 0 ? config.entryDiameter : activePasses[index - 1].diameter;
    const compression = 1 - Math.pow(pass.diameter / incomingDiameter, 2);
    const exponent = config.strengthExponent ?? 0.5;
    const strengthBefore = index === 0 ? config.incomingStrength : config.incomingStrength * Math.pow(config.entryDiameter / incomingDiameter, exponent) * config.strengthFactor;
    const strengthAfter = config.incomingStrength * Math.pow(config.entryDiameter / pass.diameter, exponent) * config.strengthFactor;
    const lineSpeed = pass.lineSpeed ?? config.finalSpeed * Math.pow(actualFinalDiameter / pass.diameter, 2);
    const calculatedForce = compression > 0
      ? config.forceCoefficient * Math.pow(incomingDiameter, 2) * (config.forceModel === 'incoming' ? strengthBefore : (strengthBefore + strengthAfter) / 2) * Math.sqrt(compression)
      : 0;
    const drawingForce = (pass.measuredForce ?? (config.forceModel === 'measured' ? NaN : calculatedForce)) * (config.parallelWires ?? 1);
    const drawingTorque = (drawingForce * pass.drumDiameter) / (2000 * (config.drawingEfficiency ?? 1));
    const factor = pass.safetyFactor ?? (index === 0 ? config.firstPassFactor : config.steadyPassFactor);
    const designTorque = drawingTorque * factor;
    const designRpm = (lineSpeed / (Math.PI * (pass.drumDiameter / 1000))) * 60;
    return { pass: index + 1, incomingDiameter, diameter: pass.diameter, drumDiameter: pass.drumDiameter, compression, strengthBefore, strengthAfter, lineSpeed, drawingForce, drawingTorque, factor, designTorque, designRpm };
  });

}

export function calculateSelection(config: SelectionConfig): SelectionRow[] {
  const baseRows = calculateDrawingRequirements(config);
  return baseRows.map((base, index) => {
    const reserveStart = Math.max(0, index - Math.max(0, config.skipHeadSpan));
    const reserve = baseRows.slice(reserveStart, index + 1).reduce((largest, row) => row.designTorque > largest.designTorque ? row : largest);
    const rpmRequired = config.skipHeadSpan > 0 && index === 0 && baseRows[1] ? Math.max(base.designRpm, baseRows[1].designRpm) : base.designRpm;
    const ratedTorque = catalogPick(reserve.designTorque, config.torqueCatalog, config.catalogTolerance);
    const ratedRpm = catalogPick(rpmRequired, config.speedCatalog, 0);
    const frame = selectFrame(ratedTorque, ratedRpm, config);
    const ratedPower = (ratedTorque * ratedRpm) / 9550;
    const ratedCurrent = (ratedPower * 1000) / (Math.sqrt(3) * config.voltage * config.voltageUtilization * config.efficiency * config.powerFactor);
    const frequency = (ratedRpm * frame.poles) / 120;
    const lengthStatus = !config.maxOverallLength ? "unchecked" : frame.estimatedLength <= config.maxOverallLength ? "pass" : "fail";
    return {
      ...base,
      ratedTorque: finite(ratedTorque), ratedRpm: finite(ratedRpm), ratedPower: finite(ratedPower),
      ratedCurrent: finite(ratedCurrent), frequency: finite(frequency), frame: frame.frame, poles: frame.poles,
      estimatedLength: frame.estimatedLength, lengthStatus, reserveSourcePass: reserve.pass,
      motorModel: `${config.series}${frame.frame}P${frame.poles}-${ratedTorque}-${ratedRpm}-${Math.round(ratedPower)}`,
    };
  });
}

export function selectionWarnings(config: SelectionConfig, rows: SelectionRow[]) {
  const warnings: string[] = [];
  if (!rows.length) return ["请至少录入一个有效拉拔道次。"];
  if (Math.abs(rows.at(-1)!.diameter - config.finalDiameter) > 0.001) warnings.push("末道丝径与工况中的出线直径不一致，速度按道次表末径计算。");
  const badCompression = rows.filter((row) => row.compression < config.compressionMin || row.compression > config.compressionMax);
  if (badCompression.length) warnings.push(`第 ${badCompression.map((row) => row.pass).join("、")} 道压缩率超出 ${(config.compressionMin * 100).toFixed(0)}%～${(config.compressionMax * 100).toFixed(0)}% 自定义校核范围。`);
  const lengthFailures = rows.filter((row) => row.lengthStatus === "fail");
  if (lengthFailures.length) warnings.push(`第 ${lengthFailures.map((row) => row.pass).join("、")} 道预计总长超过设备允许值。`);
  if (rows.some((row) => row.ratedTorque > Math.max(...config.torqueCatalog))) warnings.push("存在超出当前标准转矩库的规格，请扩展转矩档位或复核工况。");
  return warnings;
}
