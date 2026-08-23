import { WorkbookModel, type CellStyle, type WorkbookSnapshot } from '@react-sheets/core-model';

export type WorkbookTemplateId = 'blank' | 'template' | 'pivot' | 'project-plan' | 'budget';

export interface WorkbookTemplateDefinition {
  id: WorkbookTemplateId;
  name: string;
  description: string;
  create: (unitId: string, name?: string) => WorkbookSnapshot;
}

const HEADER_STYLE: CellStyle = {
  bold: true,
  textColor: '#FFFFFF',
  background: '#107C41',
  horizontalAlignment: 'center',
};

const SUBTOTAL_STYLE: CellStyle = {
  bold: true,
  background: '#EAF4EE',
};

function createWorkbook(unitId: string, name: string): WorkbookModel {
  if (!unitId.trim()) throw new Error('Workbook unitId is required');
  return new WorkbookModel(unitId, name);
}

function writeRows(workbook: WorkbookModel, rows: readonly (readonly (string | number | boolean | null)[])[], startRow = 0): void {
  const sheet = workbook.getSheet(workbook.primarySheetId);
  rows.forEach((row, rowOffset) => {
    row.forEach((value, column) => {
      sheet.cells.set(startRow + rowOffset, column, { value });
    });
  });
}

function styleRow(workbook: WorkbookModel, row: number, width: number, style: CellStyle): void {
  const sheet = workbook.getSheet(workbook.primarySheetId);
  for (let column = 0; column < width; column += 1) {
    const current = sheet.cells.get(row, column);
    if (current) sheet.cells.set(row, column, { ...current, style: { ...style } });
  }
}

function blank(unitId: string, name = '空白工作簿'): WorkbookSnapshot {
  return createWorkbook(unitId, name).snapshot();
}

function template(unitId: string, name = '会议记录模板'): WorkbookSnapshot {
  const workbook = createWorkbook(unitId, name);
  writeRows(workbook, [
    ['会议主题', '负责人', '会议日期', '状态', '备注'],
    ['产品评审', '项目组', '2026-08-24', '未开始', ''],
    ['研发同步', '研发部', '2026-08-25', '进行中', '准备风险清单'],
    ['周例会', '全体成员', '2026-08-26', '已完成', ''],
  ]);
  styleRow(workbook, 0, 5, HEADER_STYLE);
  return workbook.snapshot();
}

function pivot(unitId: string, name = '销售数据透视表模板'): WorkbookSnapshot {
  const workbook = createWorkbook(unitId, name);
  writeRows(workbook, [
    ['日期', '区域', '产品', '销售员', '数量', '金额'],
    ['2026-08-01', '华东', '标准版', '张敏', 12, 12000],
    ['2026-08-02', '华南', '专业版', '李强', 8, 16000],
    ['2026-08-03', '华东', '专业版', '王芳', 10, 20000],
    ['2026-08-04', '华北', '标准版', '赵磊', 15, 15000],
  ]);
  styleRow(workbook, 0, 6, HEADER_STYLE);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.columnWidths[0] = 110;
  sheet.columnWidths[1] = 90;
  sheet.columnWidths[2] = 110;
  sheet.columnWidths[3] = 100;
  sheet.columnWidths[4] = 80;
  sheet.columnWidths[5] = 100;
  return workbook.snapshot();
}

function projectPlan(unitId: string, name = '项目计划模板'): WorkbookSnapshot {
  const workbook = createWorkbook(unitId, name);
  writeRows(workbook, [
    ['任务名称', '负责人', '开始日期', '结束日期', '进度', '状态'],
    ['需求分析', '产品经理', '2026-08-24', '2026-08-27', 0.8, '进行中'],
    ['交互设计', '设计师', '2026-08-28', '2026-09-02', 0.2, '未开始'],
    ['开发实现', '研发团队', '2026-09-03', '2026-09-14', 0, '未开始'],
    ['验收发布', '项目经理', '2026-09-15', '2026-09-18', 0, '未开始'],
  ]);
  styleRow(workbook, 0, 6, HEADER_STYLE);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  [0.8, 0.2, 0, 0].forEach((value, index) => {
    const cell = sheet.cells.get(index + 1, 4);
    if (cell) sheet.cells.set(index + 1, 4, { ...cell, numberFormat: '0%' });
  });
  return workbook.snapshot();
}

function budget(unitId: string, name = '预算模板'): WorkbookSnapshot {
  const workbook = createWorkbook(unitId, name);
  writeRows(workbook, [
    ['预算科目', '预算金额', '实际金额', '差异', '负责人'],
    ['人员成本', 120000, 115000, 5000, '人力资源'],
    ['软件服务', 30000, 28000, 2000, '信息技术'],
    ['市场推广', 50000, 56000, -6000, '市场部'],
    ['合计', 200000, 199000, 1000, '财务部'],
  ]);
  styleRow(workbook, 0, 5, HEADER_STYLE);
  styleRow(workbook, 4, 5, SUBTOTAL_STYLE);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  for (let row = 1; row <= 4; row += 1) {
    for (const column of [1, 2, 3]) {
      const cell = sheet.cells.get(row, column);
      if (cell) sheet.cells.set(row, column, { ...cell, numberFormat: '#,##0.00' });
    }
  }
  return workbook.snapshot();
}

const DEFINITIONS: readonly WorkbookTemplateDefinition[] = [
  { id: 'blank', name: '空白工作簿', description: '从空白网格开始编辑', create: blank },
  { id: 'template', name: '从模板创建', description: '使用常用会议记录模板', create: template },
  { id: 'pivot', name: '数据透视表模板', description: '整理销售明细并分析汇总', create: pivot },
  { id: 'project-plan', name: '项目计划模板', description: '跟踪任务、进度和负责人', create: projectPlan },
  { id: 'budget', name: '预算模板', description: '比较预算、实际与差异', create: budget },
];

export function listWorkbookTemplates(): readonly WorkbookTemplateDefinition[] {
  return DEFINITIONS;
}

export function getWorkbookTemplate(id: WorkbookTemplateId): WorkbookTemplateDefinition {
  const definition = DEFINITIONS.find((entry) => entry.id === id);
  if (!definition) throw new Error(`Unknown workbook template: ${id}`);
  return definition;
}

export function createTemplateSnapshot(
  templateId: WorkbookTemplateId,
  unitId: string,
  name?: string,
): WorkbookSnapshot {
  return getWorkbookTemplate(templateId).create(unitId, name);
}

export function createWorkbookUnitId(prefix = 'wb'): string {
  const randomUuid = globalThis.crypto?.randomUUID?.();
  if (randomUuid) return `${prefix}-${randomUuid}`;
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}
