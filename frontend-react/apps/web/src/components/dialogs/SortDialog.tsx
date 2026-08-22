import React, { useState } from "react";
import { Box, Button, Inline, Stack, Text } from "@react-sheets/ui-system";

export interface SortCriterionInput {
  colIdx: number;
  ascending: boolean;
}

export interface SortDialogProps {
  open: boolean;
  columns: string[];
  onClose: () => void;
  onSort: (criteria: Array<{ colIdx: number; ascending: boolean }>, hasHeader: boolean) => void;
}

/** 多列排序对话框 */
export function SortDialog({ open, columns, onClose, onSort }: SortDialogProps): React.ReactElement | null {
  const [criteria, setCriteria] = useState<Array<{ colIdx: number; ascending: boolean }>>([{ colIdx: 0, ascending: true }]);
  const [hasHeader, setHasHeader] = useState(false);

  if (!open) return null;

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-28">
      <Box className="w-[26rem] rounded-xl border border-slate-200 bg-white p-4 shadow-2xl">
        <Stack gap="sm">
          <Text size="sm" weight="semibold">Sort range</Text>
          {criteria.map((criterion, index) => (
            <Inline key={index} gap="sm" className="items-center">
              <select
                aria-label={"Sort column " + (index + 1)}
                className="flex-1 rounded border border-slate-200 px-2 py-1 text-xs"
                value={criterion.colIdx}
                onChange={(event) => {
                  const next = [...criteria];
                  next[index] = { ...criterion, colIdx: Number(event.target.value) };
                  setCriteria(next);
                }}
              >
                {columns.map((label, idx) => (
                  <option key={label + idx} value={idx}>{label}</option>
                ))}
              </select>
              <select
                aria-label={"Sort order " + (index + 1)}
                className="w-28 rounded border border-slate-200 px-2 py-1 text-xs"
                value={criterion.ascending ? "asc" : "desc"}
                onChange={(event) => {
                  const next = [...criteria];
                  next[index] = { ...criterion, ascending: event.target.value === "asc" };
                  setCriteria(next);
                }}
              >
                <option value="asc">A → Z</option>
                <option value="desc">Z → A</option>
              </select>
              {criteria.length > 1 ? (
                <Button size="sm" variant="ghost" onClick={() => setCriteria(criteria.filter((_, i) => i !== index))}>
                  ✕
                </Button>
              ) : null}
            </Inline>
          ))}
          <Button
            size="sm"
            variant="ghost"
            className="self-start"
            onClick={() => setCriteria([...criteria, { colIdx: 0, ascending: true }])}
          >
            + Add sort column
          </Button>
          <label className="flex items-center gap-1.5 text-xs text-slate-600">
            <input checked={hasHeader} type="checkbox" onChange={(event) => setHasHeader(event.target.checked)} />
            Data has header row
          </label>
          <Inline gap="sm" className="justify-end">
            <Button size="sm" variant="ghost" onClick={onClose}>Cancel</Button>
            <Button
              size="sm"
              variant="primary"
              onClick={() => {
                onClose();
                onSort(criteria, hasHeader);
              }}
            >
              Sort
            </Button>
          </Inline>
        </Stack>
      </Box>
    </Box>
  );
}