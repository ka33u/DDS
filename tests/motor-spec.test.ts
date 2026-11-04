import {test} from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {motorSpec, efficiencyAt, validChoice, heightOptions, poleOptions} from '../lib/motor-spec.ts';
import {calculateMatrix} from '../lib/drawing-matrix.ts';
import type {ProjectData} from '../lib/project-excel.ts';
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
test('spec uses the nearest whole kW for the efficiency lookup and its range check',()=>{
  const choice={height:'355',poles:'64'};
  const spec=motorSpec({roundedRpm:105,roundedTorque:3940},choice);
  assert.equal(spec.efficiency.value,0.907); // 43.319 kW → 43 kW at 105 r/min
  assert.ok(Math.abs(spec.current!-83.02940874679133)<1e-9);
  assert.equal(motorSpec({roundedRpm:125,roundedTorque:40},choice).efficiency.value,efficiencyAt(125,1).value);
  assert.equal(motorSpec({roundedRpm:125,roundedTorque:35},choice).efficiency.value,null);
});
test('spec workbook keeps 21 ordered fields, references and cached results for arbitrary pass counts',async()=>{
  const project:ProjectData={info:{projectName:'项目甲',customerName:'客户甲'},settings:{count:13,first:'1.4',other:'1.3'},processes:[
    {id:'a',name:'头道跳模',entry:'5.5',strength:'1100',speed:'480',unit:'m/min',dies:['5.5',...Array(12).fill('5')],drums:Array(13).fill('560')},
    {id:'b',name:'高速',entry:'5.5',strength:'900',speed:'11',unit:'m/s',dies:Array(13).fill('5'),drums:Array(13).fill('560')},
  ]};
  const result=calculateMatrix(project.settings,project.processes);
  const choices=result.maxima.map(()=>({height:'368',poles:'32'}));
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(await motorSpecWorkbook(project,choices));
  assert.deepEqual(wb.worksheets.map(sheet=>sheet.name),['电机选型','拉拔工艺']);
  const spec=wb.getWorksheet('电机选型')!;
  assert.equal(spec.getCell('A1').result,'项目甲 · 客户甲');
  assert.equal(spec.getCell('AZ1').value,'项目甲');
  assert.equal(spec.getCell('AZ2').value,'客户甲');
  assert.equal(spec.pageSetup.printArea,'A1:U15');
  assert.equal(spec.getColumn('U').hidden,false);
  assert.equal(spec.getColumn('W').hidden,true);
  assert.equal(spec.getColumn('AZ').hidden,true);
  assert.equal(spec.getCell('C3').value,368);assert.equal(spec.getCell('M3').value,32);
  const drawing=wb.getWorksheet('拉拔工艺')!;
  assert.equal(drawing.getCell('A5').value,'头道跳模');assert.equal(drawing.getCell('D5').value,8);
  assert.equal(drawing.getCell('A6').value,'高速');assert.equal(drawing.getCell('D6').value,11);
  for(const [i,p] of result.maxima.entries()){
    for(const [col,expected] of [['F',p.roundedRpm],['G',p.roundedTorque]] as const){
      const c=spec.getCell(col+(i+3));
      assert.equal(c.result,expected);
      assert.match(c.formula!,/^'拉拔工艺'!/);
      assert.equal(drawing.getCell(c.formula!.split('!')[1]).result,expected);
    }
  }
  assert.equal(spec.getCell('E3').formula,'F3*G3/9550');
  assert.equal(spec.getCell('V3').formula,'ROUND(E3,0)');
  assert.match(spec.getCell('K3').formula!,/V3<0.55,V3>1250/);
  assert.equal(spec.getCell('E3').numFmt,'0');
  assert.equal(spec.getCell('J3').numFmt,'0.0');
  assert.equal(spec.getCell('AJ4').value,59.9);
  assert.match(spec.getCell('K3').formula!,/ROUND/);
  spec.eachRow(row=>row.eachCell(cell=>{if(cell.formula)assert.doesNotMatch(cell.formula,/#REF!/);}));
  await assert.rejects(()=>motorSpecWorkbook(project,[]));
});
test('offered frame sizes and pole counts support the new choices and exclude retired sizes',()=>{
  for(const height of [290,368,445,520,590,710]){
    assert.ok(heightOptions.includes(height));assert.ok(validChoice({height:String(height),poles:'32'}));
  }
  for(const height of [80,90,100,112,120,132]){
    assert.ok(!heightOptions.includes(height));assert.ok(!validChoice({height:String(height),poles:'32'}));
  }
  for(const poles of [4,6,120]){
    assert.ok(!poleOptions.includes(poles));assert.ok(!validChoice({height:'368',poles:String(poles)}));
  }
  assert.ok(validChoice({height:'375',poles:'18'}));
});
