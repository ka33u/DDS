import {percentGrid, powers, speeds} from './efficiency-table.ts';

export interface DriveDemand { roundedTorque: number; roundedRpm: number }
export interface MotorChoice { height: string; poles: string }
export const defaults = {
  series: 'YTDD', voltage: 380, voltageFactor: 0.92, powerFactor: 0.95,
  protection: 'IP55', insulation: 'F', duty: 'S1', connection: 'Y',
  cooling: '水冷', waterPressure: '0.1~0.2', waterTemp: '≤35', waterQuality: '防冻去离子水（水+乙二醇）',
};
export const heightOptions = [80,90,100,112,132,160,180,200,225,250,280,315,355,400,450,500,560,630];
export const poleOptions = [4,6,8,10,12,14,16,20,24,28,32,36,40,48,56,64,72,80,96,120];
export const newChoice = ():MotorChoice => ({height:'', poles:''});
export function validChoice(choice:MotorChoice) {
  const height=Number(choice.height), poles=Number(choice.poles);
  return Number.isSafeInteger(height) && height>0 && Number.isSafeInteger(poles) && poles>=2 && poles%2===0;
}

interface Interval { low:number; high:number; weight:number }
function at(index:number):Interval { return {low:index,high:index,weight:0}; }
function interval(value:number, axis:number[], snap:boolean):Interval {
  if (snap) {
    const exact=axis.findIndex(x=>Math.abs(x-value)<0.001);
    if (exact>=0) return at(exact);
  }
  const upper=axis.findIndex(x=>x>value);
  if (upper<0) return at(axis.length-1);
  const high=Math.max(1,upper), low=high-1;
  return {low,high,weight:(value-axis[low])/(axis[high]-axis[low])};
}
export function efficiencyAt(rpm:number, power:number) {
  if (!(rpm>0 && rpm<=6000)) return {value:null, note:'转速超出查询范围（0 < n ≤ 6000 r/min）'};
  if (!(power>=0.55 && power<=1250)) return {value:null, note:'功率超出查询范围（0.55～1250 kW）'};
  const p = power>=315 ? at(powers.length-1) : interval(power,powers,power<=280);
  const n = rpm>600 ? at(rpm<=900?10:rpm<=1200?11:rpm<=1800?12:13) : interval(rpm,speeds,rpm>=45 && rpm<=500);
  const acrossSpeed=(row:number)=>{
    const a=percentGrid[row][n.low], b=percentGrid[row][n.high];
    return a+(b-a)*n.weight;
  };
  const lo=acrossSpeed(p.low), hi=acrossSpeed(p.high);
  const value=Math.round((lo+(hi-lo)*p.weight)*10)/1000;
  return {value,note:rpm<45?'低于45 r/min，外推参考':p.low===p.high && n.low===n.high?'表值':'插值'};
}

export function motorSpec(demand:DriveDemand, choice:MotorChoice) {
  const torque=demand.roundedTorque, rpm=demand.roundedRpm;
  const power=torque*rpm/9550;
  const ratedPowerKW = Math.round(power);
  const efficiency=efficiencyAt(rpm,ratedPowerKW);
  const poles=Number(choice.poles);
  const frequency=Number.isSafeInteger(poles)&&poles>=2&&poles%2===0 ? rpm*poles/120 : null;
  const current=efficiency.value===null ? null :
    power*1000/1.732/(defaults.voltage*defaults.voltageFactor)/efficiency.value/defaults.powerFactor;
  const model=validChoice(choice) ? `${defaults.series}${Number(choice.height)}P${poles}-${torque}-${rpm}-${ratedPowerKW}` : '';
  return {torque,rpm,power,efficiency,current,frequency,model};
}
