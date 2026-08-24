import { WorkbookModel, type CellStyle, type WorkbookSnapshot } from '@react-sheets/core-model';

export type WorkbookTemplateId = 'blank' | 'template' | 'pivot' | 'project-plan' | 'budget' | 'designer-demo';

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
  sheet.columnWidthsPx[0] = 110;
  sheet.columnWidthsPx[1] = 90;
  sheet.columnWidthsPx[2] = 110;
  sheet.columnWidthsPx[3] = 100;
  sheet.columnWidthsPx[4] = 80;
  sheet.columnWidthsPx[5] = 100;
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

const DEMO_CARD_STYLE: CellStyle = {
  background: '#f2f2f2',
  verticalAlignment: 'top',
  wrapText: true,
  padding: 1,
};
const DEMO_HEADING_STYLE: CellStyle = {
  bold: true,
  fontSizePx: 17,
  textColor: '#785d6f',
  background: '#f2f2f2',
  padding: 1,
};
const DEMO_LINK_STYLE: CellStyle = {
  underline: true,
  textColor: '#4f8fbd',
  background: '#e3f1ff',
  horizontalAlignment: 'left',
  padding: 23,
};
const DEMO_NOTE_STYLE: CellStyle = {
  textColor: '#9a9a9a',
  fontSizePx: 15,
  horizontalAlignment: 'right',
  verticalAlignment: 'middle',
};
const DEMO_SCENE_STYLE: CellStyle = {
  bold: false,
  fontSizePx: 22,
  textColor: '#626262',
  verticalAlignment: 'middle',
  padding: 0,
};
const DEMO_SCENE_DETAIL_STYLE: CellStyle = {
  fontSizePx: 13,
  textColor: '#8a8a8a',
  verticalAlignment: 'middle',
  padding: 0,
};

function designerDemo(unitId: string, name = 'SpreadJS Designer Demo'): WorkbookSnapshot {
  const workbook = createWorkbook(unitId, name);
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.name = '目录索引';
  sheet.rowHeightsPx[0] = 18;
  sheet.rowHeightsPx[1] = 70;
  sheet.rowHeightsPx[2] = 30;
  sheet.rowHeightsPx[3] = 32;
  sheet.rowHeightsPx[4] = 26;
  sheet.rowHeightsPx[5] = 26;
  sheet.rowHeightsPx[6] = 26;
  sheet.rowHeightsPx[7] = 20;
  sheet.rowHeightsPx[8] = 32;
  sheet.rowHeightsPx[9] = 26;
  sheet.rowHeightsPx[10] = 26;
  sheet.rowHeightsPx[11] = 26;
  sheet.rowHeightsPx[12] = 22;
  sheet.rowHeightsPx[13] = 25;
  sheet.rowHeightsPx[14] = 22;
  sheet.rowHeightsPx[15] = 32;
  sheet.rowHeightsPx[16] = 24;
  for (let column = 0; column < 21; column += 1) sheet.columnWidthsPx[column] = 60;
  sheet.columnWidthsPx[0] = 30;

  sheet.cells.set(1, 1, {
    value: '此表格编辑器基于葡萄城 SpreadJS 实现，实现在浏览器中编辑 Excel 表格的全新体验',
    style: { bold: true, fontSizePx: 32, textColor: '#3d3c41', verticalAlignment: 'middle', padding: 0 },
  });
  sheet.merges.push({
    range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 20 },
    anchor: { row: 1, column: 1 },
  });
  sheet.cells.set(2, 1, { value: '快速体验 SpreadJS 的强大功能，如：', style: { bold: true, fontSizePx: 16, textColor: '#3a4b42', padding: 0 } });
  sheet.merges.push({
    range: { sheetId: sheet.id, startRow: 2, endRow: 2, startColumn: 1, endColumn: 20 },
    anchor: { row: 2, column: 1 },
  });

  const cards = [
    ['强大的公式计算引擎', '兼容并支持超过500种以上的标准Excel公式函数，包含求和、财务、逻辑、文本、日期时间、查找引用以及数据函数等，同时也支持自定义及异步函数'],
    ['丰富的单元格表现', '支持标准单元格的各种设置，如样式、字体、格式、方向、填充、边框等，类型也包含下拉列表、按钮、日期等类型，用户还可以自定义单元格类型'],
    ['强大的数据透视表', '支持数据透视表，数据透视表支持排序、筛选等多种功能，可以按不同维度分析数据，并且支持多种主题。'],
    ['数据验证与条件格式', '支持常用的数据验证，如：列表、整数、日期以及文本长度，以及显示特殊单元格的条件格式'],
    ['与 Excel 兼容的图表', '支持的类型包括柱状图、折线图、饼状图、面积图、条形图、XY散点图、股票图、组合图、雷达图、旭日图以及树状图。'],
    ['文件导入及导出', '可以在线上传本地Excel/CSV文件，查看在浏览器中的展示效果，也可以测试导出和打印功能。'],
  ] as const;
  const cardOrigins = [[3, 1], [3, 7], [3, 13], [8, 1], [8, 7], [8, 13]] as const;
  cards.forEach(([heading, description], index) => {
    const [row, column] = cardOrigins[index]!;
    for (let currentRow = row; currentRow <= row + 3; currentRow += 1) {
      for (let currentColumn = column; currentColumn <= column + 4; currentColumn += 1) {
        sheet.cells.set(currentRow, currentColumn, { value: null, style: { ...DEMO_CARD_STYLE } });
      }
    }
    sheet.cells.set(row, column, { value: heading, style: { ...DEMO_HEADING_STYLE } });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column + 2 },
      anchor: { row, column },
    });
    sheet.cells.set(row, column + 3, { value: '查看示例 >>', style: { ...DEMO_LINK_STYLE } });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column + 3, endColumn: column + 4 },
      anchor: { row, column: column + 3 },
    });
    sheet.cells.set(row + 1, column, { value: description, style: { ...DEMO_CARD_STYLE, fontSizePx: 13 } });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: row + 1, endRow: row + 3, startColumn: column, endColumn: column + 4 },
      anchor: { row: row + 1, column },
    });
  });
  sheet.cells.set(13, 1, { value: 'SpreadJS 三大应用场景及典型案例介绍', style: { bold: true, fontSizePx: 18, textColor: '#3d3c41' } });
  sheet.merges.push({
    range: { sheetId: sheet.id, startRow: 13, endRow: 13, startColumn: 1, endColumn: 8 },
    anchor: { row: 13, column: 1 },
  });

  const notes = [
    [13, 12, '本版本为西安葡萄城 SpreadJS 表格控件产品试用版，未取得再分发授权。'],
    [14, 12, '如需获得正式授权，请致电 400-657-6008 或发送邮件到 info.xa@grapecity.com'],
  ] as const;
  for (const [row, column, value] of notes) {
    sheet.cells.set(row, column, { value, style: { ...DEMO_NOTE_STYLE } });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: 20 },
      anchor: { row, column },
    });
  }

  const scenes = [
    [15, 3, 8, '数据填报', DEMO_SCENE_STYLE],
    [15, 11, 17, '类 Excel 报表设计', DEMO_SCENE_STYLE],
    [15, 20, 20, '表格', DEMO_SCENE_STYLE],
    [16, 3, 8, '插件：数据图表、数据透视表、甘特图、报表', DEMO_SCENE_DETAIL_STYLE],
    [16, 11, 20, '插件：数据图表、数据透视表、甘特图、报表、AI - 有效提升办公效率', DEMO_SCENE_DETAIL_STYLE],
  ] as const;
  for (const [row, startColumn, endColumn, value, style] of scenes) {
    sheet.cells.set(row, startColumn, { value, style: { ...style } });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: row, endRow: row, startColumn, endColumn },
      anchor: { row, column: startColumn },
    });
  }

  const tabs = ['500+ 公式函数支持', '丰富的单元格表现', '强大的数据透视表', '数据验证与条件格式', '与 Excel 兼容的图表', '文件导入及导出'];
  tabs.forEach((tab, index) => workbook.addSheet(`designer-demo-${index + 1}`, tab, 1000, 26));
  return workbook.snapshot();
}

const DEFINITIONS: readonly WorkbookTemplateDefinition[] = [
  { id: 'blank', name: '空白工作簿', description: '从空白网格开始编辑', create: blank },
  { id: 'template', name: '从模板创建', description: '使用常用会议记录模板', create: template },
  { id: 'pivot', name: '数据透视表模板', description: '整理销售明细并分析汇总', create: pivot },
  { id: 'project-plan', name: '项目计划模板', description: '跟踪任务、进度和负责人', create: projectPlan },
  { id: 'budget', name: '预算模板', description: '比较预算、实际与差异', create: budget },
  { id: 'designer-demo', name: 'Designer Demo', description: '复刻 SpreadJS Designer 视觉验收工作簿', create: designerDemo },
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
