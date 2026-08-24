import type { Locale } from '../../i18n';

const messages = {
  'zh-CN': {
    createTitle: '创建数据透视表', chooseData: '请选择要分析的数据', chooseLocation: '选择放置数据透视表的位置', newWorksheet: '新工作表', existingWorksheet: '现有工作表', confirm: '确定', cancel: '取消',
    fieldsTitle: '数据透视表字段', addFields: '选择要添加到报表的字段', search: '搜索', total: '合计', dragFields: '在以下区域拖动字段:', filters: '筛选', columns: '列', rows: '行', values: '值', delayUpdate: '延迟布局更新', update: '更新', view: '视图',
    noFields: '没有可用字段', noMatches: '没有匹配字段', loading: '正在准备数据透视表字段。', error: '数据透视表字段加载失败。', empty: '请选择一个数据透视表以配置字段。', close: '关闭数据透视表字段', source: '数据源', target: '目标', configure: '配置', refresh: '刷新', compact: '紧凑', outline: '大纲', tabular: '表格', controlField: '控件字段', slicer: '切片器', timeline: '时间线', pivotChart: '数据透视图',
  },
  'en-US': {
    createTitle: 'Create PivotTable', chooseData: 'Choose the data that you want to analyze', chooseLocation: 'Choose where you want the PivotTable to be placed', newWorksheet: 'New Worksheet', existingWorksheet: 'Existing Worksheet', confirm: 'OK', cancel: 'Cancel',
    fieldsTitle: 'PivotTable Fields', addFields: 'Choose fields to add to report', search: 'Search', total: 'Total', dragFields: 'Drag fields between areas below:', filters: 'Filters', columns: 'Columns', rows: 'Rows', values: 'Values', delayUpdate: 'Defer Layout Update', update: 'Update', view: 'View',
    noFields: 'No fields are available', noMatches: 'No matching fields', loading: 'Preparing PivotTable fields.', error: 'PivotTable fields could not be loaded.', empty: 'Select a PivotTable to configure its fields.', close: 'Close PivotTable fields', source: 'Source', target: 'Target', configure: 'Configure', refresh: 'Refresh', compact: 'Compact', outline: 'Outline', tabular: 'Tabular', controlField: 'Control field', slicer: 'Slicer', timeline: 'Timeline', pivotChart: 'Pivot Chart',
  },
} as const;

export type PivotMessageKey = keyof typeof messages['en-US'];
export function pivotText(locale: Locale, key: PivotMessageKey): string { return messages[locale][key]; }
