import ExcelJS from 'exceljs';
import {addDrawingProcessSheet} from './drawing-process-sheet.ts';
import {calculateMatrix,newMatrixProcess,speedInMetresPerSecond,type MatrixSettings,type MatrixProcess} from './drawing-matrix.ts';

export interface ProjectInfo { projectName:string; customerName:string }
export interface ProjectData { info:ProjectInfo; settings:MatrixSettings; processes:MatrixProcess[] }
const headers=(count:number)=>['工艺名称','进线规格 mm','抗拉强度 MPa','下线速度 m/s',...Array.from({length:count},(_,i)=>[`第${i+1}道模具 mm`,`第${i+1}道滚筒 mm`]).flat()];
function styleSheet(sheet:ExcelJS.Worksheet,headerRow=1){
  sheet.views=[{state:'frozen',xSplit:1,ySplit:headerRow}];
  sheet.eachRow((row,i)=>{row.height=i===headerRow?32:24;row.eachCell(cell=>{cell.font={name:'等线',size:11,color:{argb:i===headerRow?'FFFFFFFF':'FF243D45'}};cell.alignment={vertical:'middle',wrapText:true};if(i===headerRow)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0F5F68'}};});});
  sheet.columns.forEach((c,i)=>{c.width=i===0?25:18;});
}
function addInputs(wb:ExcelJS.Workbook,data:ProjectData,blank:boolean){
  const info=wb.addWorksheet('项目信息');
  info.addRows([['直驱选型 · 输入表','DDS-INPUT-V2'],['项目名',data.info.projectName],['客户名称',data.info.customerName],['总模数',data.settings.count]]);
  styleSheet(info);info.getColumn(1).width=30;info.getColumn(2).width=48;
  for(let r=2;r<=4;r++)info.getCell(r,2).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF3D5'}};
  const input=wb.addWorksheet('工艺输入');input.addRow(headers(data.settings.count));
  if(blank){for(let i=0;i<5;i++)input.addRow(Array(4+data.settings.count*2).fill(''));}
  else for(const p of data.processes.map(speedInMetresPerSecond))input.addRow([p.name,Number(p.entry),Number(p.strength),Number(p.speed),...Array.from({length:data.settings.count},(_,i)=>[Number(p.dies[i]),Number(p.drums[i])]).flat()]);
  styleSheet(input);input.autoFilter={from:{row:1,column:1},to:{row:Math.max(2,input.rowCount),column:4+data.settings.count*2}};
  for(let r=2;r<=input.rowCount;r++)for(let c=1;c<=4+data.settings.count*2;c++){const cell=input.getCell(r,c);cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF9EC'}};if(c>1){cell.numFmt='0.########';cell.dataValidation={type:'decimal',operator:'greaterThan',formulae:[0],allowBlank:true,showErrorMessage:true,errorTitle:'请输入正数',error:'尺寸、强度和速度必须大于0。'};}}
  const help=wb.addWorksheet('填写说明');help.addRows([
    ['填写项','要求'],['填写顺序','先填写项目信息，再逐行填写工艺输入；黄色区域为填写区。'],
    ['项目与客户','项目名、客户名称必须填写。总模数须与工艺输入中的道次数量一致；改变模数后建议重新下载模板。'],
    ['一行一套工艺','可复制空白行继续增加工艺；不要修改表头、工作表名称或插入备注行。完全空白的行会忽略。'],
    ['输入单位','线径、模具和滚筒均为mm；抗拉强度为进线强度MPa；下线速度为m/s。填写纯数值，不要在数值后附加单位。'],
    ['跳模','本道模具尺寸与上道相同即跳模，首道与进线规格比较。跳模也要填写滚筒直径。'],
    ['完整性','每套工艺必须填写全部道次；模具不得大于前一道尺寸。请填写数值，不使用公式或合并单元格。'],
    ['导入','保存为.xlsx后点击系统的一键导入。全部校验通过才替换当前项目；导入后可撤销。'],
    ['结果','每道分别取各工艺的最大扭矩和最大转速，再按系统所选方式取整。总扭矩为各道取整扭矩之和。'],
    ['计算边界','经验拉拔力模型需按材料验证；跳模理论拉拔扭矩为0，未含摩擦、加速、张力及调模附加载荷。']
  ]);styleSheet(help);help.getColumn(2).width=105;help.eachRow((row,i)=>{if(i>1)row.height=42;});
}
export async function makeProjectExcel(data:ProjectData,template=false):Promise<ArrayBuffer>{
  if(!Number.isInteger(data.settings.count)||data.settings.count<1||data.settings.count>128)throw new Error('请先选择1～128的整数模数。');
  const wb=new ExcelJS.Workbook();wb.creator='直驱选型';
  if(!template){
    const result=calculateMatrix(data.settings,data.processes);if(!result.valid)throw new Error('请先补齐全部工艺参数再导出。');
    const sheet=wb.addWorksheet('取整综合最大值');
    addInputs(wb,data,false);
    const refs=addDrawingProcessSheet(wb,data,result);
    sheet.addRows([['直驱选型 · 选型参数表'],['项目名',data.info.projectName],['客户名称',data.info.customerName],['道次','扭矩 N·m','转速 r/min'],...result.maxima.map((p,i)=>[`第${i+1}道`,{formula:refs[i].torque,result:p.roundedTorque},{formula:refs[i].rpm,result:p.roundedRpm}]),['取整扭矩总和 N·m',{formula:`SUM(B5:B${4+data.settings.count})`,result:result.totalRoundedTorque!}]]);
    sheet.mergeCells('A1:C1');sheet.mergeCells('B2:C2');sheet.mergeCells('B3:C3');
    styleSheet(sheet,4);sheet.getRow(1).height=34;sheet.getRow(1).font={name:'等线',size:16,bold:true,color:{argb:'FF0F5F68'}};sheet.getColumn(2).width=28;sheet.getColumn(3).width=22;
  } else addInputs(wb,data,true);
  return await wb.xlsx.writeBuffer() as unknown as ArrayBuffer;
}
export async function readProjectExcel(buffer:ArrayBuffer,coefficients:Pick<MatrixSettings,'first'|'other'>={first:'1.4',other:'1.3'}):Promise<ProjectData>{
  if(buffer.byteLength>20*1024*1024)throw new Error('文件不能超过20MB。');
  const wb=new ExcelJS.Workbook();try{await wb.xlsx.load(buffer);}catch{throw new Error('无法读取Excel，请使用下载的标准模板并保存为.xlsx格式。');}
  const info=wb.getWorksheet('项目信息'),input=wb.getWorksheet('工艺输入');
  if(!info||!input)throw new Error('缺少“项目信息”或“工艺输入”工作表，请使用标准模板。');
  const text=(cell:ExcelJS.Cell)=>{const v=cell.value;if(v===null||v===undefined)return '';if(typeof v==='string'||typeof v==='number')return String(v).trim();throw new Error(`${cell.worksheet.name}!${cell.address} 请填写文本或数值，不支持公式、日期或其他对象。`);};
  if(!['DDS-INPUT-V1','DDS-INPUT-V2'].includes(text(info.getCell('B1'))))throw new Error('模板版本不匹配，请重新下载标准模板。');
  ['项目名','客户名称','总模数'].forEach((label,i)=>{if(text(info.getCell(i+2,1))!==label)throw new Error(`项目信息!A${i+2} 应为“${label}”，请勿修改模板字段。`);});
  const project={projectName:text(info.getCell('B2')),customerName:text(info.getCell('B3'))};
  if(!project.projectName||!project.customerName)throw new Error('请填写项目信息中的项目名（B2）和客户名称（B3）。');
  const settings={count:Number(text(info.getCell('B4'))),first:coefficients.first,other:coefficients.other};
  const settingErrors=calculateMatrix(settings,[]).settingsErrors;if(settingErrors.length)throw new Error('项目信息：'+settingErrors.join('；'));
  const expected=headers(settings.count);
  expected.forEach((h,i)=>{if(text(input.getCell(1,i+1))!==h)throw new Error(`工艺输入!${input.getCell(1,i+1).address} 应为“${h}”，请检查表头和总模数。`);});
  if(input.columnCount>expected.length){for(let c=expected.length+1;c<=input.columnCount;c++)if(text(input.getCell(1,c)))throw new Error('工艺输入的道次数量与总模数不一致。');}
  const processes:MatrixProcess[]=[],errors:string[]=[];
  input.eachRow((row,r)=>{if(r===1)return;const values=Array.from({length:Math.max(expected.length,input.columnCount)},(_,c)=>text(row.getCell(c+1)));if(values.every(v=>!v))return;
    if(values.slice(expected.length).some(Boolean))errors.push(`工艺输入第${r}行：模板列以外存在数据。`);
    const p=newMatrixProcess(settings.count,processes.length+1);p.name=values[0];p.entry=values[1];p.strength=values[2];p.speed=values[3];p.dies=Array.from({length:settings.count},(_,i)=>values[4+i*2]);p.drums=Array.from({length:settings.count},(_,i)=>values[5+i*2]);
    if(!p.name)errors.push(`工艺输入第${r}行：缺少工艺名称。`);
    calculateMatrix(settings,[p]).rows[0].errors.forEach(e=>errors.push(`工艺输入第${r}行（${p.name||'未命名'}）：${e}`));processes.push(p);
  });
  if(!processes.length)throw new Error('工艺输入中没有已填写的工艺。');
  if(errors.length)throw new Error(errors.slice(0,12).join('\n')+(errors.length>12?`\n另有${errors.length-12}项，请修正后重新导入。`:''));
  return {info:project,settings,processes};
}
