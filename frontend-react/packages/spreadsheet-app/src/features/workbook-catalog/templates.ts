import type { CellStyle } from '@react-sheets/core-model';
import type { KernelCell, KernelSheetManifest } from '@react-sheets/protocol';

export type WorkbookTemplateId = 'blank' | 'template' | 'pivot' | 'project-plan' | 'budget' | 'designer-demo';

export interface WorkbookTemplateCreateMutation {
  readonly id: 'cell.set';
  readonly sheetId: string;
  readonly params: {
    readonly sheetId: string;
    readonly row: number;
    readonly column: number;
    readonly value: KernelCell;
  };
}

/** A template is a cloud creation intent. It never materializes a browser workbook snapshot. */
export interface WorkbookTemplateCreatePlan {
  readonly unitId: string;
  readonly name: string;
  readonly sheets: readonly KernelSheetManifest[];
  readonly initialMutations: readonly WorkbookTemplateCreateMutation[];
}

export interface WorkbookTemplateDefinition {
  readonly id: WorkbookTemplateId;
  readonly name: string;
  readonly description: string;
  readonly create: (unitId: string, name?: string) => WorkbookTemplateCreatePlan;
}

interface MutableTemplateCell {
  readonly row: number;
  readonly column: number;
  cell: KernelCell;
}

interface MutableTemplateSheet {
  readonly sheetId: string;
  name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly metadata: Record<string, unknown>;
  readonly cells: Map<string, MutableTemplateCell>;
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

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

function createSheet(
  sheetId = 'sheet-1',
  name = 'Sheet1',
  rowCount = MAX_ROWS,
  columnCount = MAX_COLUMNS,
): MutableTemplateSheet {
  return { sheetId, name, rowCount, columnCount, metadata: {}, cells: new Map() };
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function setCell(sheet: MutableTemplateSheet, row: number, column: number, cell: KernelCell): void {
  if (!Number.isSafeInteger(row) || row < 0 || row >= sheet.rowCount
    || !Number.isSafeInteger(column) || column < 0 || column >= sheet.columnCount) {
    throw new Error(`Template cell is outside ${sheet.sheetId}: ${row}:${column}`);
  }
  sheet.cells.set(cellKey(row, column), { row, column, cell: structuredClone(cell) });
}

function writeRows(
  sheet: MutableTemplateSheet,
  rows: readonly (readonly (string | number | boolean | null)[])[],
  startRow = 0,
): void {
  rows.forEach((values, rowOffset) => values.forEach((value, column) => {
    setCell(sheet, startRow + rowOffset, column, { value });
  }));
}

function styleRow(sheet: MutableTemplateSheet, row: number, width: number, style: CellStyle): void {
  for (let column = 0; column < width; column += 1) {
    const entry = sheet.cells.get(cellKey(row, column));
    if (entry) entry.cell = { ...entry.cell, style: structuredClone(style) };
  }
}

function setNumberFormat(sheet: MutableTemplateSheet, row: number, column: number, numberFormat: string): void {
  const entry = sheet.cells.get(cellKey(row, column));
  if (entry) entry.cell = { ...entry.cell, numberFormat };
}

function merge(sheet: MutableTemplateSheet, startRow: number, endRow: number, startColumn: number, endColumn: number): void {
  const merges = (sheet.metadata.merges ??= []) as Array<Record<string, unknown>>;
  merges.push({
    range: { sheetId: sheet.sheetId, startRow, endRow, startColumn, endColumn },
    anchor: { row: startRow, column: startColumn },
  });
}

function setDimension(sheet: MutableTemplateSheet, axis: 'row' | 'column', index: number, pixels: number): void {
  const key = axis === 'row' ? 'rowHeightsPx' : 'columnWidthsPx';
  const values = (sheet.metadata[key] ??= {}) as Record<number, number>;
  values[index] = pixels;
}

function finalize(unitId: string, name: string, sheets: readonly MutableTemplateSheet[]): WorkbookTemplateCreatePlan {
  const normalizedUnitId = unitId.trim();
  const normalizedName = name.trim();
  if (!normalizedUnitId) throw new Error('Workbook unitId is required');
  if (!normalizedName) throw new Error('Workbook name is required');
  if (sheets.length === 0) throw new Error('Workbook template requires at least one worksheet');
  const initialMutations: WorkbookTemplateCreateMutation[] = [];
  for (const sheet of sheets) {
    for (const entry of sheet.cells.values()) {
      initialMutations.push({
        id: 'cell.set',
        sheetId: sheet.sheetId,
        params: {
          sheetId: sheet.sheetId,
          row: entry.row,
          column: entry.column,
          value: structuredClone(entry.cell),
        },
      });
    }
  }
  return {
    unitId: normalizedUnitId,
    name: normalizedName,
    sheets: sheets.map((sheet) => ({
      sheetId: sheet.sheetId,
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      metadata: structuredClone(sheet.metadata),
    })),
    initialMutations,
  };
}

function blank(unitId: string, name = '空白工作簿'): WorkbookTemplateCreatePlan {
  return finalize(unitId, name, [createSheet()]);
}

function template(unitId: string, name = '会议记录模板'): WorkbookTemplateCreatePlan {
  const sheet = createSheet();
  writeRows(sheet, [
    ['会议主题', '负责人', '会议日期', '状态', '备注'],
    ['产品评审', '项目组', '2026-08-24', '未开始', ''],
    ['研发同步', '研发部', '2026-08-25', '进行中', '准备风险清单'],
    ['周例会', '全体成员', '2026-08-26', '已完成', ''],
  ]);
  styleRow(sheet, 0, 5, HEADER_STYLE);
  return finalize(unitId, name, [sheet]);
}

function pivot(unitId: string, name = '销售数据透视表模板'): WorkbookTemplateCreatePlan {
  const sheet = createSheet();
  writeRows(sheet, [
    ['日期', '区域', '产品', '销售员', '数量', '金额'],
    ['2026-08-01', '华东', '标准版', '张敏', 12, 12000],
    ['2026-08-02', '华南', '专业版', '李强', 8, 16000],
    ['2026-08-03', '华东', '专业版', '王芳', 10, 20000],
    ['2026-08-04', '华北', '标准版', '赵磊', 15, 15000],
  ]);
  styleRow(sheet, 0, 6, HEADER_STYLE);
  [110, 90, 110, 100, 80, 100].forEach((pixels, column) => setDimension(sheet, 'column', column, pixels));
  return finalize(unitId, name, [sheet]);
}

function projectPlan(unitId: string, name = '项目计划模板'): WorkbookTemplateCreatePlan {
  const sheet = createSheet();
  writeRows(sheet, [
    ['任务名称', '负责人', '开始日期', '结束日期', '进度', '状态'],
    ['需求分析', '产品经理', '2026-08-24', '2026-08-27', 0.8, '进行中'],
    ['交互设计', '设计师', '2026-08-28', '2026-09-02', 0.2, '未开始'],
    ['开发实现', '研发团队', '2026-09-03', '2026-09-14', 0, '未开始'],
    ['验收发布', '项目经理', '2026-09-15', '2026-09-18', 0, '未开始'],
  ]);
  styleRow(sheet, 0, 6, HEADER_STYLE);
  for (let row = 1; row <= 4; row += 1) setNumberFormat(sheet, row, 4, '0%');
  return finalize(unitId, name, [sheet]);
}

function budget(unitId: string, name = '预算模板'): WorkbookTemplateCreatePlan {
  const sheet = createSheet();
  writeRows(sheet, [
    ['预算科目', '预算金额', '实际金额', '差异', '负责人'],
    ['人员成本', 120000, 115000, 5000, '人力资源'],
    ['软件服务', 30000, 28000, 2000, '信息技术'],
    ['市场推广', 50000, 56000, -6000, '市场部'],
    ['合计', 200000, 199000, 1000, '财务部'],
  ]);
  styleRow(sheet, 0, 5, HEADER_STYLE);
  styleRow(sheet, 4, 5, SUBTOTAL_STYLE);
  for (let row = 1; row <= 4; row += 1) {
    for (const column of [1, 2, 3]) setNumberFormat(sheet, row, column, '#,##0.00');
  }
  return finalize(unitId, name, [sheet]);
}

function designerDemo(unitId: string, name = 'SpreadJS Designer Demo'): WorkbookTemplateCreatePlan {
  const sheet = createSheet('sheet-1', '目录索引');
  [18, 70, 30, 32, 26, 26, 26, 20, 32, 26, 26, 26, 22, 25, 22, 32, 24]
    .forEach((pixels, row) => setDimension(sheet, 'row', row, pixels));
  for (let column = 0; column < 21; column += 1) setDimension(sheet, 'column', column, 60);
  setDimension(sheet, 'column', 0, 30);

  setCell(sheet, 1, 1, {
    value: '此表格编辑器基于葡萄城 SpreadJS 实现，实现在浏览器中编辑 Excel 表格的全新体验',
    style: { bold: true, fontSizePx: 32, textColor: '#3d3c41', verticalAlignment: 'middle', padding: 0 },
  });
  merge(sheet, 1, 1, 1, 20);
  setCell(sheet, 2, 1, { value: '快速体验 SpreadJS 的强大功能，如：', style: { bold: true, fontSizePx: 16, textColor: '#3a4b42', padding: 0 } });
  merge(sheet, 2, 2, 1, 20);

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
        setCell(sheet, currentRow, currentColumn, { value: null, style: { ...DEMO_CARD_STYLE } });
      }
    }
    setCell(sheet, row, column, { value: heading, style: { ...DEMO_HEADING_STYLE } });
    merge(sheet, row, row, column, column + 2);
    setCell(sheet, row, column + 3, { value: '查看示例 >>', style: { ...DEMO_LINK_STYLE } });
    merge(sheet, row, row, column + 3, column + 4);
    setCell(sheet, row + 1, column, { value: description, style: { ...DEMO_CARD_STYLE, fontSizePx: 13 } });
    merge(sheet, row + 1, row + 3, column, column + 4);
  });
  setCell(sheet, 13, 1, { value: 'SpreadJS 三大应用场景及典型案例介绍', style: { bold: true, fontSizePx: 18, textColor: '#3d3c41' } });
  merge(sheet, 13, 13, 1, 8);

  const notes = [
    [13, 12, '本版本为西安葡萄城 SpreadJS 表格控件产品试用版，未取得再分发授权。'],
    [14, 12, '如需获得正式授权，请致电 400-657-6008 或发送邮件到 info.xa@grapecity.com'],
  ] as const;
  for (const [row, column, value] of notes) {
    setCell(sheet, row, column, { value, style: { ...DEMO_NOTE_STYLE } });
    merge(sheet, row, row, column, 20);
  }

  const scenes = [
    [15, 3, 8, '数据填报', DEMO_SCENE_STYLE],
    [15, 11, 17, '类 Excel 报表设计', DEMO_SCENE_STYLE],
    [15, 20, 20, '表格', DEMO_SCENE_STYLE],
    [16, 3, 8, '插件：数据图表、数据透视表、甘特图、报表', DEMO_SCENE_DETAIL_STYLE],
    [16, 11, 20, '插件：数据图表、数据透视表、甘特图、报表、AI - 有效提升办公效率', DEMO_SCENE_DETAIL_STYLE],
  ] as const;
  for (const [row, startColumn, endColumn, value, style] of scenes) {
    setCell(sheet, row, startColumn, { value, style: { ...style } });
    merge(sheet, row, row, startColumn, endColumn);
  }

  const tabs = ['500+ 公式函数支持', '丰富的单元格表现', '强大的数据透视表', '数据验证与条件格式', '与 Excel 兼容的图表', '文件导入及导出'];
  const sheets = [sheet, ...tabs.map((tab, index) => createSheet(`designer-demo-${index + 1}`, tab, 1000, 26))];
  return finalize(unitId, name, sheets);
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

export function createTemplatePlan(
  templateId: WorkbookTemplateId,
  unitId: string,
  name?: string,
): WorkbookTemplateCreatePlan {
  return getWorkbookTemplate(templateId).create(unitId, name);
}

export function createWorkbookUnitId(prefix = 'wb'): string {
  const randomUuid = globalThis.crypto?.randomUUID?.();
  if (randomUuid) return `${prefix}-${randomUuid}`;
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}
