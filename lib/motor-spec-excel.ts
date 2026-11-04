import ExcelJS from 'exceljs';
import {defaults, motorSpec, validChoice, type DriveDemand, type MotorChoice} from './motor-spec.ts';
import {percentGrid, powers, speeds} from './efficiency-table.ts';
import type {ProjectInfo} from './project-excel.ts';

const titles=['道次','系列','中心高','电机型号','额定功率\n(kW)','额定转速\n(r/min)','额定转矩\n(Nm)','额定电压\n(VAC)','额定电流\n(A)','额定频率\n(Hz)','效率\n%','功率因数\ncosΦ','极数','防护等级\nIP','绝缘等级\nIns','工作制','接法\nConn','冷却方式','进水压力\nMPa','进水口温度\n℃','水质要求'];
const widths=[7,9,10,35,12,12,13,13,12,12,10,12,8,12,10,10,10,12,16,16,27];
const line={style:'thin' as const,color:{argb:'FFBFCBD1'}};
function heading(row:ExcelJS.Row) {
  row.height=32;
  row.eachCell(cell=>{
    cell.font={name:'楷体',size:11,bold:true,color:{argb:'FF3333FF'}};
    cell.alignment={vertical:'middle',horizontal:'center',wrapText:true};
    cell.border={left:line,right:line,top:line,bottom:line};
  });
}
function formula(expression:string, result?:number|string|null):ExcelJS.CellFormulaValue {
  return result==null ? {formula:expression} : {formula:expression,result};
}

// Keep the lookup table, interpolation steps, and title inputs to the right of
// the printable A:U specification. All formulas work offline on this sheet.
function addSupportColumns(sheet:ExcelJS.Worksheet, count:number, info:ProjectInfo) {
  const set=(address:string,value:ExcelJS.CellValue)=>{sheet.getCell(address).value=value;};
  set('AY1','项目名'); set('AZ1',info.projectName.trim());
  set('AY2','客户名称'); set('AZ2',info.customerName.trim());
  set('AY3','电压系数'); set('AZ3',defaults.voltageFactor);
  set('AY4','效率表'); set('AZ4','旧效率查询工具一级能效参考表；1.1 kW / 45 r/min 沿用59.9%。低于45 r/min为外推。');
  set('AI1','kW / r/min');
  speeds.forEach((rpm,i)=>{sheet.getCell(1,36+i).value=rpm;});
  powers.forEach((kw,i)=>{
    sheet.getCell(i+2,35).value=kw;
    percentGrid[i].forEach((v,j)=>{sheet.getCell(i+2,36+j).value=v;});
  });
  const powerAxis='$AI$2:$AI$29',speedAxis='$AJ$1:$AW$1',grid='$AJ$2:$AW$29';
  const powerAt=(index:string)=>`INDEX(${powerAxis},${index})`;
  const speedAt=(index:string)=>`INDEX(${speedAxis},1,${index})`;
  const labels=['功率起点','功率下标','功率上标','功率权重','转速起点','转速下标','转速上标','转速权重','低功率效率 %','高功率效率 %'];
  labels.forEach((label,i)=>{sheet.getCell(2,23+i).value=label;});
  for (let r=3;r<count+3;r++){
    const power=`E${r}`, speed=`F${r}`;
    const cell=(col:string,expression:string)=>set(col+r,formula(expression));
    cell('W',`IFERROR(MATCH(${power},${powerAxis},1),1)`);
    const nextPower=`MIN(28,W${r}+1)`;
    cell('X',`IF(AND(${power}<=280,ABS(${power}-${powerAt(nextPower)})<0.001),${nextPower},W${r})`);
    cell('Y',`IF(OR(${power}>=315,AND(${power}<=280,ABS(${power}-${powerAt('X'+r)})<0.001)),X${r},MIN(28,X${r}+1))`);
    cell('Z',`IF(Y${r}=X${r},0,(${power}-${powerAt('X'+r)})/(${powerAt('Y'+r)}-${powerAt('X'+r)}))`);
    cell('AA',`IFERROR(MATCH(${speed},${speedAxis},1),1)`);
    const nextSpeed=`MIN(14,AA${r}+1)`;
    cell('AB',`IF(${speed}>600,IF(${speed}<=900,11,IF(${speed}<=1200,12,IF(${speed}<=1800,13,14))),IF(AND(${speed}>=45,${speed}<=500,ABS(${speed}-${speedAt(nextSpeed)})<0.001),${nextSpeed},AA${r}))`);
    cell('AC',`IF(OR(${speed}>600,AND(${speed}>=45,${speed}<=500,ABS(${speed}-${speedAt('AB'+r)})<0.001)),AB${r},MIN(14,AB${r}+1))`);
    cell('AD',`IF(AC${r}=AB${r},0,(${speed}-${speedAt('AB'+r)})/(${speedAt('AC'+r)}-${speedAt('AB'+r)}))`);
    for(const [col,rowIndex] of [['AE','X'],['AF','Y']]){
      const left=`INDEX(${grid},${rowIndex}${r},AB${r})`,right=`INDEX(${grid},${rowIndex}${r},AC${r})`;
      cell(col,`${left}+(${right}-${left})*AD${r}`);
    }
  }
  for(let col=22;col<=52;col++)sheet.getColumn(col).hidden=true;
}

export async function motorSpecWorkbook(info:ProjectInfo,demands:DriveDemand[],choices:MotorChoice[]) {
  if(!demands.length || demands.some((_,i)=>!choices[i]||!validChoice(choices[i])))throw new Error('请逐道选择有效中心高及偶数极数。');
  const wb=new ExcelJS.Workbook();wb.creator='直驱选型';wb.calcProperties.fullCalcOnLoad=true;
  const sheet=wb.addWorksheet('电机选型');
  sheet.mergeCells('A1:U1');
  const title=[info.projectName.trim(),info.customerName.trim()].filter(Boolean).join(' · ')||'电机选型';
  sheet.getCell('A1').value=formula('IF($AZ$1&$AZ$2="","电机选型",$AZ$1&IF(AND($AZ$1<>"",$AZ$2<>"")," · ","")&$AZ$2)',title);
  heading(sheet.getRow(1));sheet.getRow(1).height=30.75;
  sheet.addRow(titles);heading(sheet.getRow(2));
  sheet.columns.forEach((column,i)=>{column.width=widths[i];});
  sheet.views=[{state:'frozen',xSplit:1,ySplit:2}];
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,printTitlesRow:'1:2',printArea:`A1:U${demands.length+2}`,paperSize:9};
  demands.forEach((d,i)=>{
    const r=i+3, choice=choices[i], spec=motorSpec(d,choice);
    const row=sheet.addRow([i+1,defaults.series,Number(choice.height),
      formula(`IF(OR(C${r}="",M${r}=""),"",B${r}&C${r}&"P"&M${r}&"-"&G${r}&"-"&F${r}&"-"&ROUND(E${r},0))`,spec.model),
      formula(`F${r}*G${r}/9550`,spec.power),
      spec.rpm,spec.torque,defaults.voltage,
      formula(`IFERROR(IF(K${r}="","",E${r}*1000/1.732/(H${r}*$AZ$3)/K${r}/L${r}),"")`,spec.current),
      formula(`IF(M${r}="","",F${r}*M${r}/120)`,spec.frequency),
      formula(`IF(OR(F${r}<=0,F${r}>6000,E${r}<0.55,E${r}>1250),"",ROUND(AE${r}+(AF${r}-AE${r})*Z${r},1)/100)`,spec.efficiency.value),
      defaults.powerFactor,Number(choice.poles),defaults.protection,defaults.insulation,defaults.duty,defaults.connection,defaults.cooling,defaults.waterPressure,defaults.waterTemp,defaults.waterQuality]);
    row.height=30;row.eachCell(cell=>{
      cell.font={name:'Times New Roman',size:10,bold:true,color:{argb:'FF3333FF'}};
      cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};cell.border={left:line,right:line,top:line,bottom:line};
    });
    for(const col of ['C','F','G','M'])sheet.getCell(col+r).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF5D9'}};
    sheet.getCell('C'+r).dataValidation={type:'whole',operator:'greaterThan',formulae:[0],showErrorMessage:true,error:'中心高须为正整数。'};
    sheet.getCell('M'+r).dataValidation={type:'custom',formulae:[`AND(ISNUMBER(M${r}),M${r}>=2,MOD(M${r},2)=0)`],showErrorMessage:true,error:'极数须为正偶数。'};
    for(const [col,format] of Object.entries({E:'0',F:'0',G:'0',I:'0.0',J:'0.0',K:'0.0%',L:'0.00'}))sheet.getCell(col+r).numFmt=format;
  });
  addSupportColumns(sheet, demands.length, info);
  return await wb.xlsx.writeBuffer() as unknown as ArrayBuffer;
}
