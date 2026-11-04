import type {ProjectInfo} from './project-excel.ts';

export function workbookFilename(info:ProjectInfo, passCount:number, kind:'标准工艺输入模板'|'直驱选型参数表'|'永磁电机选型参数') {
  const clean=(name:string,fallback:string)=>name.trim().replace(/[\\/:*?"<>|\u0000-\u001f]/g,'_') || fallback;
  return [clean(info.projectName,'未命名项目'),clean(info.customerName,'未填客户'),`${passCount}模`,kind].join('_')+'.xlsx';
}
