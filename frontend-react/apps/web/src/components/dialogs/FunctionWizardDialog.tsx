import React, { useState } from 'react';
import { Box, Button, Dialog, Inline, ScrollArea, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { BUILTIN_FUNCTIONS } from '@react-sheets/formula-engine';

export interface FunctionWizardDialogProps {
  open: boolean;
  onClose: () => void;
  onInsertFormula: (formula: string) => void;
}

const FUNCTION_DESCRIPTIONS: Record<string, { category: string; syntax: string; description: string }> = {
  SUM: { category: 'Math', syntax: 'SUM(number1, [number2], ...)', description: 'Adds all numbers in a range of cells.' },
  AVERAGE: { category: 'Statistical', syntax: 'AVERAGE(number1, [number2], ...)', description: 'Returns the average (arithmetic mean) of its arguments.' },
  COUNT: { category: 'Statistical', syntax: 'COUNT(value1, [value2], ...)', description: 'Counts how many numbers are in the list of arguments.' },
  COUNTA: { category: 'Statistical', syntax: 'COUNTA(value1, [value2], ...)', description: 'Counts how many values are in the list of arguments.' },
  MIN: { category: 'Statistical', syntax: 'MIN(number1, [number2], ...)', description: 'Returns the smallest number in a set of values.' },
  MAX: { category: 'Statistical', syntax: 'MAX(number1, [number2], ...)', description: 'Returns the largest number in a set of values.' },
  IF: { category: 'Logical', syntax: 'IF(logical_test, [value_if_true], [value_if_false])', description: 'Specifies a logical test to perform.' },
  VLOOKUP: { category: 'Lookup', syntax: 'VLOOKUP(lookup_value, table_array, col_index, [range_lookup])', description: 'Looks for a value in the leftmost column of a table.' },
  INDEX: { category: 'Lookup', syntax: 'INDEX(array, row_num, [col_num])', description: 'Returns a value or reference of the cell at the intersection of a row and column.' },
  MATCH: { category: 'Lookup', syntax: 'MATCH(lookup_value, lookup_array, [match_type])', description: 'Returns the relative position of an item in an array.' },
  XLOOKUP: { category: 'Lookup', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])', description: 'Searches a range or an array, and returns an item corresponding to the first match it finds.' },
  CONCAT: { category: 'Text', syntax: 'CONCAT(text1, [text2], ...)', description: 'Combines the text from multiple ranges and/or strings.' },
  TEXTJOIN: { category: 'Text', syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)', description: 'Combines text from multiple ranges with a delimiter.' },
  DATE: { category: 'Date', syntax: 'DATE(year, month, day)', description: 'Returns the sequential serial number that represents a particular date.' },
  TODAY: { category: 'Date', syntax: 'TODAY()', description: 'Returns the serial number of the current date.' },
  NOW: { category: 'Date', syntax: 'NOW()', description: 'Returns the serial number of the current date and time.' },
};

export function FunctionWizardDialog({ open, onClose, onInsertFormula }: FunctionWizardDialogProps) {
  const [selectedFunction, setSelectedFunction] = useState('SUM');
  const [search, setSearch] = useState('');

  const fnList = Object.keys(BUILTIN_FUNCTIONS).filter((fn) =>
    fn.toLowerCase().includes(search.toLowerCase()),
  );

  const meta = FUNCTION_DESCRIPTIONS[selectedFunction] || {
    category: 'General',
    syntax: `${selectedFunction}(...)`,
    description: `Standard spreadsheet formula function ${selectedFunction}.`,
  };

  const handleInsert = () => {
    onInsertFormula(`=${selectedFunction}()`);
    onClose();
  };

  return (
    <Dialog
      open={open}
      title="Insert Function Wizard"
      description="Choose a spreadsheet calculation function to insert into active cell."
      onClose={onClose}
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="primary" onClick={handleInsert}>
            Insert ={selectedFunction}
          </Button>
        </>
      }
    >
      <Stack gap="md">
        <Box>
          <TextInput
            placeholder="Search functions..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            leadingIcon="search"
          />
        </Box>

        <Box className="grid grid-cols-2 gap-4">
          <ScrollArea className="h-56 rounded-lg border border-slate-200 bg-slate-50/50 p-1">
            {fnList.map((fn) => (
              <Button
                key={fn}
                onClick={() => setSelectedFunction(fn)}
                size="sm"
                variant="ghost"
                className={`w-full justify-start rounded px-2.5 py-1.5 text-left text-xs font-semibold transition-colors ${
                  selectedFunction === fn
                    ? 'bg-blue-600 text-white'
                    : 'text-slate-700 hover:bg-slate-200/70'
                }`}
              >
                {fn}
              </Button>
            ))}
          </ScrollArea>

          <Stack className="justify-between rounded-lg border border-slate-200 bg-white p-3.5">
            <Stack gap="xs">
              <Inline className="justify-between">
                <Text size="sm" weight="bold" className="text-slate-900">{selectedFunction}</Text>
                <Text size="xs" weight="semibold" className="rounded bg-blue-50 px-2 py-0.5 text-blue-700">
                  {meta.category}
                </Text>
              </Inline>
              <Box className="mt-2 rounded border border-slate-100 bg-slate-50 p-1.5 font-mono text-xs font-medium text-blue-600">
                {meta.syntax}
              </Box>
              <Text as="p" size="xs" className="mt-2 leading-relaxed text-slate-600">{meta.description}</Text>
            </Stack>
            <Text size="xs" tone="subtle">
              Double-click or press Insert to insert into the formula bar.
            </Text>
          </Stack>
        </Box>
      </Stack>
    </Dialog>
  );
}
