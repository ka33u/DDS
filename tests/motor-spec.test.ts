import {test} from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {motorSpec, efficiencyAt, validChoice} from '../lib/motor-spec.ts';
import {motorSpecWorkbook} from '../lib/motor-spec-excel.ts';

test('specification uses template formulas and lookup-tool efficiency',()=>{
  const m=motorSpec({roundedRpm:72,roundedTorque:6180},{height:'355',poles:'64'});
  assert.equal(m.model,'YTDD355P64-6180-72-47');
  assert.equal(m.power,6180*72/9550);
  assert.equal(m.frequency,38.4);
  assert.equal(m.efficiency.value,0.888);
  assert.ok(Math.abs(m.current!-91.21404536621516)<1e-9);
  assert.equal(validChoice({height:'355',poles:'3'}),false);
  assert.equal(validChoice({height:'',poles:'64'}),false);
});
test('existing efficiency tool reference values, bands and unavailable results',()=>{
  const reference=[[1500,15,94.9],[450,13,92.1],[52.5,0.925,60.4],[75,297.5,93.4],
    [600,11,92.8],[600.001,11,92.8],[6000,1250,96.9],[30,11,73.1],[45,1.1,59.9]];
  for(const [rpm,power,expected] of reference)assert.ok(Math.abs(efficiencyAt(rpm,power).value!*100-expected)<1e-10);
  assert.match(efficiencyAt(30,11).note,/外推/);
  for(const [rpm,power] of [[0,11],[6001,11],[45,0],[45,1251]])assert.equal(efficiencyAt(rpm,power).value,null);
});
test('spec workbook keeps 21 ordered fields, references and cached results for arbitrary pass counts',async()=>{
  const demands=Array.from({length:13},(_,i)=>({roundedTorque:i===12?0:3170,roundedRpm:115}));
  const choices=demands.map(()=>({height:'315',poles:'32'}));
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(await motorSpecWorkbook({projectName:'项目甲',customerName:'客户甲'},demands,choices));
  assert.deepEqual(wb.worksheets.map(sheet=>sheet.name),['电机选型']);
  const spec=wb.getWorksheet('电机选型')!;
  assert.equal(spec.getCell('A1').result,'项目甲 · 客户甲');
  assert.equal(spec.getCell('AZ1').value,'项目甲');
  assert.equal(spec.getCell('AZ2').value,'客户甲');
  assert.equal(spec.pageSetup.printArea,'A1:U15');
  assert.equal(spec.getColumn('U').hidden,false);
  assert.equal(spec.getColumn('W').hidden,true);
  assert.equal(spec.getColumn('AZ').hidden,true);
  assert.equal(spec.getCell('C3').value,315);assert.equal(spec.getCell('M3').value,32);
  assert.equal(spec.getCell('F15').value,115);
  assert.equal(spec.getCell('G15').value,0);
  assert.equal(spec.getCell('E3').formula,'F3*G3/9550');
  assert.equal(spec.getCell('E3').numFmt,'0');
  assert.equal(spec.getCell('J3').numFmt,'0.0');
  assert.equal(spec.getCell('AJ4').value,59.9);
  assert.match(spec.getCell('K3').formula!,/ROUND/);
  spec.eachRow(row=>row.eachCell(cell=>{if(cell.formula)assert.doesNotMatch(cell.formula,/!|#REF!/);}));
  await assert.rejects(()=>motorSpecWorkbook({projectName:'',customerName:''},demands,[]));
});
