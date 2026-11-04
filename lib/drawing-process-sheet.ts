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

// The specification export owns its inputs here. In the importable project
// workbook these cells link to 工艺输入 so there is only one input per value.
export function addDrawingProcessSheet(wb:ExcelJS.Workbook, data:ProjectData, result:Calculation, inputSheet?:ExcelJS.Worksheet) {
  if(!result.valid)throw new Error('请先补齐全部工艺参数。');
  const sheet=wb.addWorksheet(sheetName);
  const passes=Array.from({length:data.settings.count},(_,i)=>i);
  const column=(index:number)=>sheet.getColumn(index).letter;
  const die=(i:number)=>column(5+2*i);
  const drum=(i:number)=>column(6+2*i);
  const resultCol=(i:number)=>column(2+i);
  const checkCol=column(5+2*passes.length);
  const title=[data.info.projectName,data.info.customerName,sheetName].filter(Boolean).join(' · ');
  sheet.mergeCells('A1:F1');sheet.getCell('A1').value=title;sectionHeading(sheet.getRow(1));
  sheet.addRow(['总模数',passes.length,'首道扭矩系数',Number(data.settings.first),'其他道扭矩系数',Number(data.settings.other)]);
  sheet.addRow(['工艺输入（每行一套）']);
  const header=sheet.addRow(['工艺名称','进线规格 mm','进线抗拉强度 MPa','下线速度 m/s',
    ...passes.flatMap(i=>[`第${i+1}道模具 mm`,`第${i+1}道滚筒 mm`]),'输入检查']);
  sectionHeading(header);
  const inputRows=data.processes.map((original,processIndex)=>{
    const p=speedInMetresPerSecond(original);
    const values=[p.name,Number(p.entry),Number(p.strength),Number(p.speed),
      ...passes.flatMap(i=>[Number(p.dies[i]),Number(p.drums[i])])];
    const row=sheet.addRow(inputSheet ? values.map((value,j)=>({
      formula:`'${inputSheet.name}'!${column(j+1)}${processIndex+2}`,result:value,
    })) : values);
    const n=row.number, numericRange=`B${n}:${drum(passes.length-1)}${n}`;
    const checks=passes.map(i=>`${die(i)}${n}<=${i?die(i-1):'B'}${n}`);
    row.getCell(checkCol).value={
      formula:`IF(AND(COUNT(${numericRange})=${values.length-1},MIN(${numericRange})>0,COUNT($D$2,$F$2)=2,MIN($D$2,$F$2)>0,${checks.join(',')}),"有效","请检查输入")`,
      result:'有效',
    };
    row.height=23;
    if(!inputSheet)row.eachCell(cell=>{
      if(cell.col!==row.getCell(checkCol).col)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF8E7'}};
    });
    return n;
  });
  sheet.addRow([]);
  function addResults(metric:'torque'|'rpm') {
    const torque=metric==='torque';
    const header=sheet.addRow([torque?'所需扭矩 N·m':'所需转速 r/min',...passes.map(i=>`第${i+1}道`)]);
    sectionHeading(header);
    for(const [j,process] of result.rows.entries()){
      const n=inputRows[j],entry=`$B${n}`,strength=`$C${n}`,speed=`$D${n}`;
      const end=`${die(passes.length-1)}${n}`;
      const values=passes.map(i=>{
        const before=i?`${die(i-1)}${n}`:entry, after=`${die(i)}${n}`, diameter=`${drum(i)}${n}`;
        const avg=`${strength}*(SQRT(${entry}/${before})+SQRT(${entry}/${after}))/2`;
        // With non-increasing diameters, before equals entry until the first reduction.
        const force=`0.6*${before}^2*(${avg})*SQRT(1-(${after}/${before})^2)`;
        const expression=torque
          ? `IF(${after}=${before},0,${force}*${diameter}/2000*IF(${before}=${entry},$D$2,$F$2))`
          : `${speed}*60*(${end}/${after})^2*1000/(PI()*${diameter})`;
        return {formula:`IF($${checkCol}${n}="有效",${expression},NA())`,result:process.points[i][metric]};
      });
      const row=sheet.addRow([{formula:`A${n}`,result:process.name},...values]);
      for(const i of passes)row.getCell(i+2).numFmt=torque?'0.00;-0.00;"0（跳模）"':'0.00';
    }
    const first=header.number+1,last=first+result.rows.length-1;
    const range=(i:number)=>`${resultCol(i)}${first}:${resultCol(i)}${last}`;
    const rounded=sheet.addRow([torque?'取整最大扭矩 N·m':'取整最大转速 r/min',...result.maxima.map((p,i)=>({
      formula:`ROUNDUP(MAX(${range(i)})/5,0)*5`,result:torque?p.roundedTorque:p.roundedRpm,
    }))]);
    highlightTotal(rounded);
    sheet.addRow(['决定工艺',...result.maxima.map((p,i)=>({
      formula:`INDEX($A$${first}:$A$${last},MATCH(MAX(${range(i)}),${range(i)},0))`,result:torque?p.torqueSource:p.speedSource,
    }))]);
    return passes.map(i=>sheet.getCell(rounded.number,i+2).address);
  }
  const torque=addResults('torque');
  const rpm=addResults('rpm');
  highlightTotal(sheet.addRow(['取整扭矩总和 N·m',{formula:`SUM(${torque[0]}:${torque[torque.length-1]})`,result:result.totalRoundedTorque!}]));
  sheet.columns.forEach((c,i)=>{c.width=i===0?26:17;});
  sheet.views=[{state:'frozen',xSplit:1,ySplit:4}];
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
  return passes.map(i=>({torque:`'${sheetName}'!${torque[i]}`,rpm:`'${sheetName}'!${rpm[i]}`}));
}
