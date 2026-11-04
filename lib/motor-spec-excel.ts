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

// Ordinary INDEX/MATCH formulas keep the file editable in desktop Excel and WPS.
function addEfficiencySheets(wb:ExcelJS.Workbook,count:number) {
  const grid=wb.addWorksheet('效率基表');
  grid.addRow(['kW / r/min',...speeds]);
  powers.forEach((power,i)=>grid.addRow([power,...percentGrid[i]]));
  heading(grid.getRow(1));grid.getColumn(1).width=17;
  grid.addRow([]);
  grid.addRow(['数据来源','原效率查询工具 motor-efficiency-，9799465；表中为效率百分数。']);
  grid.addRow(['口径','GB 30253-2024 一级能效参考，1.1 kW / 45 r/min 沿用旧工具的59.9%。']);
  grid.addRow(['转速分段','600～900、900～1200、1200～1800、1800～6000 各段按旧工具取值，低于45外推。']);
  grid.addRow(['表格说明','表末315 kW行适用于315～1250 kW。查询范围以外不提供效率和电流。']);
  const helper=wb.addWorksheet('效率插值');
  helper.addRow(['道次','功率 kW','转速 r/min','功率起点','功率下标','功率上标','功率权重','转速起点','转速下标','转速上标','转速权重','低功率效率 %','高功率效率 %','参考效率','查询说明']);
  heading(helper.getRow(1));helper.columns.forEach(c=>{c.width=17;});helper.getColumn(15).width=40;
  const px="'效率基表'!$A$2:$A$29",nx="'效率基表'!$B$1:$O$1",table="'效率基表'!$B$2:$O$29";
  for(let i=0;i<count;i++){
    const r=i+2,s=i+3;
    const p=(index:string)=>`INDEX(${px},${index})`;
    const n=(index:string)=>`INDEX(${nx},1,${index})`;
    const val=(row:string,col:string)=>`INDEX(${table},${row}${r},${col}${r})`;
    const lerp=(row:string)=>`${val(row,'I')}+(${val(row,'J')}-${val(row,'I')})*K${r}`;
    helper.addRow([i+1,
      formula(`'电机选型'!E${s}`),formula(`'电机选型'!F${s}`),
      formula(`IFERROR(MATCH(B${r},${px},1),1)`),
      formula(`IF(AND(B${r}<=280,ABS(B${r}-${p('MIN(28,D'+r+'+1)')})<0.001),MIN(28,D${r}+1),D${r})`),
      formula(`IF(OR(B${r}>=315,AND(B${r}<=280,ABS(B${r}-${p('E'+r)})<0.001)),E${r},MIN(28,E${r}+1))`),
      formula(`IF(F${r}=E${r},0,(B${r}-${p('E'+r)})/(${p('F'+r)}-${p('E'+r)}))`),
      formula(`IFERROR(MATCH(C${r},${nx},1),1)`),
      formula(`IF(C${r}>600,IF(C${r}<=900,11,IF(C${r}<=1200,12,IF(C${r}<=1800,13,14))),IF(AND(C${r}>=45,C${r}<=500,ABS(C${r}-${n('MIN(14,H'+r+'+1)')})<0.001),H${r}+1,H${r}))`),
      formula(`IF(OR(C${r}>600,AND(C${r}>=45,C${r}<=500,ABS(C${r}-${n('I'+r)})<0.001)),I${r},MIN(14,I${r}+1))`),
      formula(`IF(J${r}=I${r},0,(C${r}-${n('I'+r)})/(${n('J'+r)}-${n('I'+r)}))`),
      formula(lerp('E')),formula(lerp('F')),
      formula(`IF(OR(C${r}<=0,C${r}>6000,B${r}<0.55,B${r}>1250),"",ROUND(L${r}+(M${r}-L${r})*G${r},1)/100)`),
      formula(`IF(N${r}="","超出查询范围",IF(C${r}<45,"低于45 r/min，外推参考","表值 / 插值"))`),
    ]);
  }
  helper.getColumn(14).numFmt='0.0%';
}

export async function motorSpecWorkbook(info:ProjectInfo,demands:DriveDemand[],choices:MotorChoice[],title?:string) {
  if(!demands.length || demands.some((_,i)=>!choices[i]||!validChoice(choices[i])))throw new Error('请逐道选择有效中心高及偶数极数。');
  const wb=new ExcelJS.Workbook();wb.creator='直驱选型';wb.calcProperties.fullCalcOnLoad=true;
  const sheet=wb.addWorksheet('电机选型');
  sheet.mergeCells('A1:U1');sheet.getCell('A1').value=title||[info.customerName,info.projectName,`${demands.length}模单头电机参数`].filter(Boolean).join(' · ');
  heading(sheet.getRow(1));sheet.getRow(1).height=30.75;
  sheet.addRow(titles);heading(sheet.getRow(2));
  sheet.columns.forEach((column,i)=>{column.width=widths[i];});
  sheet.views=[{state:'frozen',xSplit:1,ySplit:2}];
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,printTitlesRow:'1:2',paperSize:9};
  const source=wb.addWorksheet('道次需求');
  source.addRows([['道次','取整扭矩 N·m','取整转速 r/min'],...demands.map((d,i)=>[i+1,d.roundedTorque,d.roundedRpm])]);
  heading(source.getRow(1));source.columns.forEach(c=>{c.width=20;});
  const preset=wb.addWorksheet('规格预设');
  preset.addRows([['项目','值'],['系列',defaults.series],['额定电压 VAC',defaults.voltage],['电压系数（样表）',defaults.voltageFactor],['功率因数',defaults.powerFactor],
    ['防护等级',defaults.protection],['绝缘等级',defaults.insulation],['工作制',defaults.duty],['接法',defaults.connection],['冷却方式',defaults.cooling],['进水压力 MPa',defaults.waterPressure],['进水口温度 ℃',defaults.waterTemp],['水质要求',defaults.waterQuality],['项目名',info.projectName],['客户名',info.customerName],
    ['修改规格','在电机选型的中心高、极数列填写数值。在道次需求中改转矩或转速，在本表改通用预设。'],
    ['效率口径','采用既有效率工具的一级能效参考值。低于45 r/min为外推，范围外效率、电流留空，原因见效率插值。']]);
  heading(preset.getRow(1));preset.getColumn(1).width=24;preset.getColumn(2).width=86;
  const presetRef=(r:number,value:number|string)=>formula(`'规格预设'!$B$${r}`,value);
  demands.forEach((d,i)=>{
    const r=i+3, ref=i+2, choice=choices[i], spec=motorSpec(d,choice);
    const row=sheet.addRow([i+1,presetRef(2,defaults.series),Number(choice.height),
      formula(`IF(OR(C${r}="",M${r}=""),"",B${r}&C${r}&"P"&M${r}&"-"&G${r}&"-"&F${r}&"-"&ROUND(E${r},0))`,spec.model),
      formula(`F${r}*G${r}/9550`,spec.power),
      formula(`'道次需求'!C${ref}`,spec.rpm),formula(`'道次需求'!B${ref}`,spec.torque),presetRef(3,defaults.voltage),
      formula(`IFERROR(IF(K${r}="","",E${r}*1000/1.732/(H${r}*'规格预设'!$B$4)/K${r}/L${r}),"")`,spec.current),
      formula(`IF(M${r}="","",F${r}*M${r}/120)`,spec.frequency),
      formula(`IF('效率插值'!N${ref}="","",'效率插值'!N${ref})`,spec.efficiency.value),
      presetRef(5,defaults.powerFactor),Number(choice.poles),presetRef(6,defaults.protection),presetRef(7,defaults.insulation),presetRef(8,defaults.duty),presetRef(9,defaults.connection),presetRef(10,defaults.cooling),presetRef(11,defaults.waterPressure),presetRef(12,defaults.waterTemp),presetRef(13,defaults.waterQuality)]);
    row.height=30;row.eachCell(cell=>{
      cell.font={name:'Times New Roman',size:10,bold:true,color:{argb:'FF3333FF'}};
      cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};cell.border={left:line,right:line,top:line,bottom:line};
    });
    for(const col of ['C','M'])sheet.getCell(col+r).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF5D9'}};
    sheet.getCell('C'+r).dataValidation={type:'whole',operator:'greaterThan',formulae:[0],showErrorMessage:true,error:'中心高须为正整数。'};
    sheet.getCell('M'+r).dataValidation={type:'custom',formulae:[`AND(ISNUMBER(M${r}),M${r}>=2,MOD(M${r},2)=0)`],showErrorMessage:true,error:'极数须为正偶数。'};
    for(const [col,format] of Object.entries({E:'0.00',F:'0',G:'0',I:'0.0',J:'0.0',K:'0.0%',L:'0.00'}))sheet.getCell(col+r).numFmt=format;
  });
  addEfficiencySheets(wb,demands.length);
  return await wb.xlsx.writeBuffer() as unknown as ArrayBuffer;
}
