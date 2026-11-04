import type ExcelJS from 'exceljs';
import {calculateMatrix, speedInMetresPerSecond} from './drawing-matrix.ts';
import type {ProjectData} from './project-excel.ts';

type Calculation = ReturnType<typeof calculateMatrix>;
const sheetName = '拉拔工艺';
function sectionHeading(row:ExcelJS.Row) {
  row.height=29;
  row.eachCell(cell=>{
    cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0F5F68'}};
    cell.font={name:'等线',size:11,bold:true,color:{argb:'FFFFFFFF'}};
    cell.alignment={vertical:'middle',wrapText:true};
  });
}
function highlightTotal(row:ExcelJS.Row) {
  row.height=23;
  row.eachCell(cell=>{
    cell.font={bold:true,color:{argb:'FF174D53'}};
    cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE4F1EF'}};
  });
}

// This sheet records the validated inputs and results at export time. Other
// sheets can reference the numeric totals without exposing drawing formulas.
export function addDrawingProcessSheet(wb:ExcelJS.Workbook, data:ProjectData, result:Calculation) {
  if (!result.valid) throw new Error('请先补齐全部工艺参数。');
  const sheet=wb.addWorksheet(sheetName);
  const passes=Array.from({length:data.settings.count},(_,i)=>i);
  const title=[data.info.projectName,data.info.customerName,sheetName].filter(Boolean).join(' · ');
  sheet.mergeCells('A1:F1');sheet.getCell('A1').value=title;sectionHeading(sheet.getRow(1));
  sheet.addRow(['总模数',passes.length,'首道扭矩系数',Number(data.settings.first),'其他道扭矩系数',Number(data.settings.other)]);
  sheet.addRow(['工艺输入（每行一套）','以下为导出时的工艺输入及计算结果']);
  sectionHeading(sheet.addRow(['工艺名称','进线规格 mm','进线抗拉强度 MPa','下线速度 m/s',
    ...passes.flatMap(i=>[`第${i+1}道模具 mm`,`第${i+1}道滚筒 mm`])]));
  for (const original of data.processes) {
    const p=speedInMetresPerSecond(original);
    const row=sheet.addRow([p.name,Number(p.entry),Number(p.strength),Number(p.speed),
      ...passes.flatMap(i=>[Number(p.dies[i]),Number(p.drums[i])])]);
    row.height=23;
  }
  sheet.addRow([]);
  function addResults(metric:'torque'|'rpm') {
    const torque=metric==='torque';
    sectionHeading(sheet.addRow([torque?'所需扭矩 N·m':'所需转速 r/min',...passes.map(i=>`第${i+1}道`)]));
    for(const process of result.rows){
      const row=sheet.addRow([process.name,...process.points.map(point=>point[metric])]);
      for(const i of passes)row.getCell(i+2).numFmt=torque && process.points[i].skipped?'0.00;-0.00;"0（跳模）"':'0.00';
    }
    const rounded=sheet.addRow([torque?'取整最大扭矩 N·m':'取整最大转速 r/min',
      ...result.maxima.map(p=>torque?p.roundedTorque:p.roundedRpm)]);
    highlightTotal(rounded);
    sheet.addRow(['决定工艺',...result.maxima.map(p=>torque?p.torqueSource:p.speedSource)]);
    return passes.map(i=>`'${sheetName}'!${rounded.getCell(i+2).address}`);
  }
  const torque=addResults('torque');
  const rpm=addResults('rpm');
  highlightTotal(sheet.addRow(['取整扭矩总和 N·m',result.totalRoundedTorque!]));
  sheet.columns.forEach((c,i)=>{c.width=i===0?26:17;});
  sheet.views=[{state:'frozen',xSplit:1,ySplit:4}];
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
  return passes.map(i=>({torque:torque[i],rpm:rpm[i]}));
}
