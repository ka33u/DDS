import {test} from 'node:test';
import assert from 'node:assert/strict';
import {calculateMatrix,matrixExcel,type MatrixProcess} from '../lib/drawing-matrix.ts';
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
test('incomplete or increasing dies block combined maxima instead of disappearing',()=>{
  for(const patch of [{speed:''},{dies:['6',...fixture.dies.slice(1)]},{drums:['0',...fixture.drums.slice(1)]}]){const r=calculateMatrix(settings,[fixture,{...fixture,id:'bad',...patch}]);assert.equal(r.valid,false);assert.deepEqual(r.maxima,[]);}
  assert.equal(calculateMatrix(settings,[]).valid,false);
});
test('export includes horizontal inputs, per-process results and independently combined maxima',()=>{
  const xml=matrixExcel(settings,[{...fixture,name:'A&B<工艺>'}]);assert.match(xml,/A&amp;B&lt;工艺&gt;/);assert.match(xml,/综合最大值/);assert.match(xml,/第|7道转速/);assert.match(xml,/480|8/);
});
