import {test} from 'node:test';
import assert from 'node:assert/strict';
import {calculateMatrix,matrixExcel,roundTorque,speedInMetresPerSecond,newMatrixProcess,type MatrixProcess} from '../lib/drawing-matrix.ts';
const fixture:MatrixProcess={id:'reference',name:'原表校核',entry:'5.5',strength:'1100',speed:'8',unit:'m/s',dies:['5.05','4.55','4.13','3.8','3.5','3.25','3.25'],drums:Array(7).fill('560')};
const settings={count:7,first:'1.4',other:'1.3'};
test('all seven speeds reproduce Excel N17:N23, including unchanged final die',()=>{
  const r=calculateMatrix(settings,[fixture]);assert.equal(r.valid,true);
  [113.00230530282877,139.20257413285307,168.95457503915662,199.57349660563645,235.25235028452164,272.837045300392,272.837045300392].forEach((v,i)=>assert.ok(Math.abs(r.rows[0].points[i].rpm-v)<1e-9));
  assert.equal(r.rows[0].points[6].torque,0);
  assert.ok(Math.abs(r.rows[0].points[0].torque-3168.048974291362)<1e-8);
  assert.ok(Math.abs(r.rows[0].points[1].torque-2190.872617102834*1.3)<1e-8);
});
test('m/s and m/min inputs produce identical physical speeds',()=>{
  assert.deepEqual(calculateMatrix(settings,[fixture]).maxima,calculateMatrix(settings,[{...fixture,speed:'480',unit:'m/min'}]).maxima);
});
test('maximum torque and speed are independent across process rows',()=>{
  const strong={...fixture,name:'高扭矩',strength:'2000',speed:'4'},fast={...fixture,id:'fast',name:'高车速',strength:'500',speed:'16'};
  const r=calculateMatrix(settings,[strong,fast]);assert.equal(r.maxima[0].torqueSource,'高扭矩');assert.equal(r.maxima[0].speedSource,'高车速');
});
test('first and middle skipped passes give zero drawing torque but appropriate speed',()=>{
  const p={...fixture,dies:['5.5','4.55','4.55','3.8','3.5','3.25','3.25'],drums:['400','560','800','560','560','560','560']};
  const r=calculateMatrix(settings,[p]);assert.equal(r.valid,true);assert.equal(r.rows[0].points[0].torque,0);assert.equal(r.rows[0].points[2].torque,0);assert.ok(r.rows[0].points[0].rpm>0);assert.ok(Math.abs(r.rows[0].points[2].rpm/r.rows[0].points[1].rpm-560/800)<1e-12);
});
test('each process uses the first factor only on its first actual reduction',()=>{
  const processes=[
    {...fixture,name:'第二模开拔',dies:['5.5','5','5','4.5']},
    {...fixture,name:'第三模开拔',dies:['5.5','5.5','5','4.5']},
    {...fixture,name:'全部跳模',dies:['5.5','5.5','5.5','5.5']},
  ];
  const result=calculateMatrix({count:4,first:'1.6',other:'1.2'},processes);
  assert.ok(result.valid);
  const factors=[[0,1.6,0,1.2],[0,0,1.6,1.2],[0,0,0,0]];
  for(const [r,row] of result.rows.entries()){
    for(const [i,point] of row.points.entries()){
      assert.equal(point.torque,point.force*560/2000*factors[r][i]);
      assert.equal(point.skipped,factors[r][i]===0);
      assert.ok(point.rpm>0);
    }
  }
  assert.equal(result.maxima[2].torqueSource,'第三模开拔');
  assert.equal(result.maxima[2].roundedTorque,roundTorque(result.rows[1].points[2].torque));
});
test('incomplete or increasing dies block combined maxima instead of disappearing',()=>{
  for(const patch of [{speed:''},{dies:['6',...fixture.dies.slice(1)]},{drums:['0',...fixture.drums.slice(1)]}]){const r=calculateMatrix(settings,[fixture,{...fixture,id:'bad',...patch}]);assert.equal(r.valid,false);assert.deepEqual(r.maxima,[]);}
  assert.equal(calculateMatrix(settings,[]).valid,false);
});
test('export separates torque and speed into columns with one row per pass',()=>{
  const xml=matrixExcel(settings,[{...fixture,name:'A&B<工艺>'}]);assert.match(xml,/A&amp;B&lt;工艺&gt;/);assert.match(xml,/取整综合最大值/);assert.match(xml,/第|7道转速/);assert.match(xml,/480|8/);
});

test('rounding always rounds upward to a multiple of five and totals each pass',()=>{
  [0,100,101,105,105.01,106,110].forEach((v,i)=>assert.equal(roundTorque(v),[0,100,105,105,110,110,110][i]));
  const r=calculateMatrix(settings,[fixture]);
  assert.deepEqual(r.maxima.map(p=>p.roundedTorque),[3170,2850,2355,1895,1660,1400,0]);
  assert.equal(r.totalRoundedTorque,13330);
  assert.deepEqual(r.maxima.map(p=>p.roundedRpm),[115,140,170,200,240,275,275]);
  const summary=matrixExcel(settings,[fixture]).match(/<Worksheet ss:Name="取整综合最大值">[\s\S]*?<\/Worksheet>/)?.[0]??'';
  const rows=[...summary.matchAll(/<Row>([\s\S]*?)<\/Row>/g)].map(m=>[...m[1].matchAll(/<Data ss:Type="(?:String|Number)">([^<]*)<\/Data>/g)].map(c=>c[1]));
  assert.deepEqual(rows[0],['道次','扭矩 N·m','转速 r/min']);
  assert.deepEqual(rows.slice(1,8),[
    ['第1道','3170','115'],['第2道','2850','140'],['第3道','2355','170'],
    ['第4道','1895','200'],['第5道','1660','240'],['第6道','1400','275'],['第7道','0','275']
  ]);
  assert.deepEqual(rows[8],['取整扭矩总和 N·m','13330','']);
  assert.equal(rows.length,9);
  assert.match(summary,/>3170</);assert.match(summary,/>115</);assert.doesNotMatch(summary,/>113\./);
  assert.equal(calculateMatrix(settings,[]).totalRoundedTorque,null);
  assert.match(matrixExcel(settings,[fixture]),/取整扭矩总和 N·m/);
});
test('new and saved processes use m/s without changing their physical speed',()=>{
  assert.equal(newMatrixProcess(7,1).unit,'m/s');
  const converted=speedInMetresPerSecond({...fixture,unit:'m/min',speed:'480'});
  assert.equal(converted.speed,'8');assert.equal(converted.unit,'m/s');
  assert.deepEqual(calculateMatrix(settings,[converted]),calculateMatrix(settings,[fixture]));
  assert.deepEqual(speedInMetresPerSecond(converted),converted);
  assert.equal(speedInMetresPerSecond({...fixture,unit:'m/min',speed:''}).speed,'');
});
