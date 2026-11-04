import {test} from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {makeProjectExcel,readProjectExcel,type ProjectData} from '../lib/project-excel.ts';
const data:ProjectData={info:{projectName:'13模改造 A&B',customerName:'客户甲'},settings:{count:2,first:'1.4',other:'1.3'},processes:[{id:'a',name:'工艺甲',entry:'5.5',strength:'1100',speed:'8',unit:'m/s',dies:['5','5'],drums:['560','560']}]};
async function book(buffer:ArrayBuffer){const wb=new ExcelJS.Workbook();await wb.xlsx.load(buffer);return wb;}
test('result workbook contains customer and project, integer results and reimportable input',async()=>{
 const buffer=await makeProjectExcel(data);const wb=await book(buffer);const sheet=wb.getWorksheet('取整综合最大值')!;
 assert.equal(sheet.getCell('B2').value,data.info.projectName);assert.equal(sheet.getCell('B3').value,data.info.customerName);
 assert.deepEqual([1,2,3].map(c=>sheet.getCell(4,c).value),['道次','扭矩 N·m','转速 r/min']);
 assert.equal(Number(sheet.getCell('B5').result)%5,0);assert.equal(sheet.getCell('B6').result,0);assert.equal(Number(sheet.getCell('C6').result)%5,0);
 assert.equal(wb.getWorksheet('拉拔工艺')!.getCell('B5').formula,"'工艺输入'!B2");
 const restored=await readProjectExcel(buffer);assert.deepEqual(restored.info,data.info);assert.deepEqual(restored.settings,data.settings);
 assert.deepEqual({...restored.processes[0],id:'a'},data.processes[0]);
});
test('customer can fill template, add rows and import all processes',async()=>{
 const wb=await book(await makeProjectExcel(data,true));assert.equal(wb.getWorksheet('取整综合最大值'),undefined);
 assert.equal(wb.getWorksheet('项目信息')!.rowCount,4);
 assert.equal(wb.getWorksheet('项目信息')!.getCell('B1').value,'DDS-INPUT-V2');
 const sheet=wb.getWorksheet('工艺输入')!;assert.equal(sheet.getCell('B2').value,'');
 sheet.getRow(2).values=['工艺1',5.5,1100,8,5,560,5,560];sheet.getRow(8).values=['工艺2',5.5,1000,10,5,560,4.5,560];
 const imported=await readProjectExcel(await wb.xlsx.writeBuffer() as unknown as ArrayBuffer);assert.equal(imported.processes.length,2);assert.equal(imported.processes[1].speed,'10');
});
test('legacy coefficients never override system settings and template count is independent',async()=>{
 const wb=await book(await makeProjectExcel(data));
 const info=wb.getWorksheet('项目信息')!;info.getCell('B1').value='DDS-INPUT-V1';info.getCell('A5').value='首道次扭矩系数';info.getCell('B5').value=99;info.getCell('A6').value='其他道次扭矩系数';info.getCell('B6').value=88;
 const imported=await readProjectExcel(await wb.xlsx.writeBuffer() as unknown as ArrayBuffer,{first:'1.6',other:'1.2'});
 assert.deepEqual(imported.settings,{count:2,first:'1.6',other:'1.2'});
 const template=await book(await makeProjectExcel({...data,settings:{count:13,first:'',other:''}},true));
 assert.equal(template.getWorksheet('项目信息')!.getCell('B4').value,13);
 assert.equal(template.getWorksheet('工艺输入')!.getRow(1).cellCount,30);
 await assert.rejects(()=>makeProjectExcel({...data,settings:{...data.settings,count:0}},true));
});
test('invalid units, missing names, formulas, changed headers and partial rows reject the whole import',async()=>{
 for(const modify of [
  (w:ExcelJS.Workbook)=>{w.getWorksheet('项目信息')!.getCell('B2').value='';},
  (w:ExcelJS.Workbook)=>{w.getWorksheet('工艺输入')!.getCell('D2').value='8m/s';},
  (w:ExcelJS.Workbook)=>{w.getWorksheet('工艺输入')!.getCell('F2').value='';},
  (w:ExcelJS.Workbook)=>{w.getWorksheet('工艺输入')!.getCell('E2').value=6;},
  (w:ExcelJS.Workbook)=>{w.getWorksheet('工艺输入')!.getCell('D2').value={formula:'4+4',result:8};},
  (w:ExcelJS.Workbook)=>{w.getWorksheet('工艺输入')!.getCell('D1').value='下线速度 m/min';}
 ]){const wb=await book(await makeProjectExcel(data));modify(wb);
 const buffer=await wb.xlsx.writeBuffer();await assert.rejects(()=>readProjectExcel(buffer as unknown as ArrayBuffer));}
});
