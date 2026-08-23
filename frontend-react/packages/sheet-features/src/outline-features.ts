import type { OutlineGroup, OutlineModel, RangeRef, WorksheetModel } from '@react-sheets/core-model';

export interface OutlineControl {
  axis: 'row' | 'column';
  index: number;
  level: number;
  collapsed: boolean;
  groupId: string;
}

export function resolveOutlineControls(sheet: WorksheetModel): OutlineControl[] {
  const groups = sheet.outline?.groups ?? [];
  return groups.map((group) => ({
    axis: group.axis,
    index: group.start,
    level: group.level,
    collapsed: group.collapsed,
    groupId: group.id,
  }));
}

export function nextOutlineLevel(
  sheet: WorksheetModel,
  axis: OutlineGroup['axis'],
  start: number,
  end: number,
): number {
  const groups = sheet.outline?.groups ?? [];
  let maxLevel = 0;
  for (const group of groups) {
    if (group.axis !== axis) continue;
    if (group.end < start || group.start > end) continue;
    maxLevel = Math.max(maxLevel, group.level);
  }
  return Math.min(3, maxLevel + 1);
}

export function buildRowOutlineGroup(
  sheetId: string,
  range: RangeRef,
  sheet: WorksheetModel,
  id: string,
): OutlineGroup {
  return {
    id,
    axis: 'row',
    start: range.startRow,
    end: range.endRow,
    level: nextOutlineLevel(sheet, 'row', range.startRow, range.endRow),
    collapsed: false,
  };
}

export function buildColumnOutlineGroup(
  sheetId: string,
  range: RangeRef,
  sheet: WorksheetModel,
  id: string,
): OutlineGroup {
  void sheetId;
  return {
    id,
    axis: 'column',
    start: range.startColumn,
    end: range.endColumn,
    level: nextOutlineLevel(sheet, 'column', range.startColumn, range.endColumn),
    collapsed: false,
  };
}

export function groupsWithinRange(
  outline: OutlineModel | undefined,
  axis: OutlineGroup['axis'],
  range: RangeRef,
): OutlineGroup[] {
  return (outline?.groups ?? []).filter(
    (group) => group.axis === axis
      && group.start >= (axis === 'row' ? range.startRow : range.startColumn)
      && group.end <= (axis === 'row' ? range.endRow : range.endColumn),
  );
}

/** 命中行头大纲按钮区域（屏幕局部坐标，行头宽度内） */
export function hitOutlineControl(
  localX: number,
  localY: number,
  rowTop: number,
  rowHeight: number,
  controls: readonly OutlineControl[],
): OutlineControl | undefined {
  if (localX < 0 || localX > 46) return undefined;
  if (localY < rowTop || localY > rowTop + rowHeight) return undefined;
  for (const control of controls) {
    if (control.axis !== 'row') continue;
    const buttonLeft = 4 + (control.level - 1) * 10;
    const buttonRight = buttonLeft + 10;
    if (localX >= buttonLeft && localX <= buttonRight) return control;
  }
  return undefined;
}
