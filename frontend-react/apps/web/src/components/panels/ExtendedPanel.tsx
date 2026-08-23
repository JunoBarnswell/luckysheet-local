import React, { useState } from 'react';
import type { CapabilityDescriptor, PlatformCapability } from '@react-sheets/spreadsheet-app';
import { Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { parseRangeInput } from '../../domain/range-input';

export interface ExtendedPanelProps {
  capabilities: readonly CapabilityDescriptor[];
  lastWhatIfMessage?: string | null;
  canRunExtended: boolean;
  sheetId: string;
  onGoalSeek: (params: {
    setRow: number;
    setColumn: number;
    targetValue: number;
    changingRow: number;
    changingColumn: number;
  }) => void;
  onRunDataTable: (params: {
    inputMode: 'row' | 'column';
    inputCell: { row: number; column: number };
    tableRange: { startRow: number; startColumn: number; endRow: number; endColumn: number };
  }) => void;
  onRunScenario: (params: {
    name: string;
    changingCell: { row: number; column: number };
    changingValue: number;
    resultCell: { row: number; column: number };
  }) => void;
  onEvaluateCapability: (capability: PlatformCapability) => Promise<{ ok: boolean; message?: string }>;
  onClose?: () => void;
}

function parseCell(input: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(input.trim());
  if (!match) return null;
  const letters = match[1]!.toUpperCase();
  let column = 0;
  for (const char of letters) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

export function ExtendedPanel({
  canRunExtended,
  capabilities,
  lastWhatIfMessage,
  onClose,
  onEvaluateCapability,
  onGoalSeek,
  onRunDataTable,
  onRunScenario,
  sheetId,
}: ExtendedPanelProps) {
  const [setCell, setSetCell] = useState('A1');
  const [changingCell, setChangingCell] = useState('B1');
  const [targetValue, setTargetValue] = useState('100');
  const [dataTableInputCell, setDataTableInputCell] = useState('B1');
  const [dataTableRange, setDataTableRange] = useState('A1:D2');
  const [dataTableMode, setDataTableMode] = useState<'row' | 'column'>('column');
  const [scenarioName, setScenarioName] = useState('Best Case');
  const [scenarioChangingCell, setScenarioChangingCell] = useState('B1');
  const [scenarioChangingValue, setScenarioChangingValue] = useState('20');
  const [scenarioResultCell, setScenarioResultCell] = useState('A1');
  const [capabilityStatus, setCapabilityStatus] = useState<string | null>(null);

  const setStatus = (message: string) => setCapabilityStatus(message);

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Extended Capabilities</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <div>
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Goal Seek</Text>
            <Stack gap="sm">
              <TextInput
                value={setCell}
                onChange={(event) => setSetCell(event.target.value)}
                placeholder="Set cell (e.g. A1)"
                disabled={!canRunExtended}
              />
              <TextInput
                type="number"
                value={targetValue}
                onChange={(event) => setTargetValue(event.target.value)}
                placeholder="To value"
                disabled={!canRunExtended}
              />
              <TextInput
                value={changingCell}
                onChange={(event) => setChangingCell(event.target.value)}
                placeholder="By changing cell (e.g. B1)"
                disabled={!canRunExtended}
              />
              <Button
                size="sm"
                variant="primary"
                disabled={!canRunExtended}
                onClick={() => {
                  const set = parseCell(setCell);
                  const changing = parseCell(changingCell);
                  if (!set || !changing) {
                    setStatus('Invalid cell reference');
                    return;
                  }
                  onGoalSeek({
                    setRow: set.row,
                    setColumn: set.column,
                    targetValue: Number(targetValue),
                    changingRow: changing.row,
                    changingColumn: changing.column,
                  });
                }}
              >
                Run Goal Seek
              </Button>
            </Stack>
          </div>

          <div>
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Data Table</Text>
            <Stack gap="sm">
              <Inline gap="xs">
                <Button
                  size="xs"
                  variant={dataTableMode === 'column' ? 'primary' : 'ghost'}
                  disabled={!canRunExtended}
                  onClick={() => setDataTableMode('column')}
                >
                  Column Input
                </Button>
                <Button
                  size="xs"
                  variant={dataTableMode === 'row' ? 'primary' : 'ghost'}
                  disabled={!canRunExtended}
                  onClick={() => setDataTableMode('row')}
                >
                  Row Input
                </Button>
              </Inline>
              <TextInput
                value={dataTableInputCell}
                onChange={(event) => setDataTableInputCell(event.target.value)}
                placeholder="Input cell (e.g. B1)"
                disabled={!canRunExtended}
              />
              <TextInput
                value={dataTableRange}
                onChange={(event) => setDataTableRange(event.target.value)}
                placeholder="Table range (e.g. A1:D2)"
                disabled={!canRunExtended}
              />
              <Button
                size="sm"
                variant="secondary"
                disabled={!canRunExtended}
                onClick={() => {
                  const inputCell = parseCell(dataTableInputCell);
                  const tableRange = parseRangeInput(dataTableRange, sheetId);
                  if (!inputCell || !tableRange) {
                    setStatus('Invalid data table cell or range');
                    return;
                  }
                  onRunDataTable({
                    inputMode: dataTableMode,
                    inputCell,
                    tableRange,
                  });
                }}
              >
                Run Data Table
              </Button>
            </Stack>
          </div>

          <div>
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Scenario</Text>
            <Stack gap="sm">
              <TextInput
                value={scenarioName}
                onChange={(event) => setScenarioName(event.target.value)}
                placeholder="Scenario name"
                disabled={!canRunExtended}
              />
              <TextInput
                value={scenarioChangingCell}
                onChange={(event) => setScenarioChangingCell(event.target.value)}
                placeholder="Changing cell (e.g. B1)"
                disabled={!canRunExtended}
              />
              <TextInput
                type="number"
                value={scenarioChangingValue}
                onChange={(event) => setScenarioChangingValue(event.target.value)}
                placeholder="New value"
                disabled={!canRunExtended}
              />
              <TextInput
                value={scenarioResultCell}
                onChange={(event) => setScenarioResultCell(event.target.value)}
                placeholder="Result cell (e.g. A1)"
                disabled={!canRunExtended}
              />
              <Button
                size="sm"
                variant="secondary"
                disabled={!canRunExtended}
                onClick={() => {
                  const changing = parseCell(scenarioChangingCell);
                  const result = parseCell(scenarioResultCell);
                  if (!changing || !result) {
                    setStatus('Invalid scenario cell reference');
                    return;
                  }
                  onRunScenario({
                    name: scenarioName.trim() || 'Scenario',
                    changingCell: changing,
                    changingValue: Number(scenarioChangingValue),
                    resultCell: result,
                  });
                }}
              >
                Run Scenario
              </Button>
            </Stack>
          </div>

          {lastWhatIfMessage ? (
            <div className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-700">
              {lastWhatIfMessage}
            </div>
          ) : null}

          <div>
            <Text size="xs" weight="semibold" className="mb-2 text-slate-700">Platform Capabilities</Text>
            <Stack gap="sm">
              {capabilities.map((capability) => (
                <Panel key={capability.id} className="shadow-none">
                  <PanelBody className="p-3">
                    <Stack gap="xs">
                      <Text size="sm" weight="semibold">{capability.id}</Text>
                      <Text size="xs" tone="muted">
                        {capability.enabled ? 'Enabled' : capability.reason ?? 'Disabled'}
                      </Text>
                      {!capability.enabled ? (
                        <Button
                          size="xs"
                          variant="ghost"
                          onClick={() => {
                            void onEvaluateCapability(capability.id).then((result) => {
                              setStatus(result.message ?? (result.ok ? 'Can enable' : 'Cannot enable'));
                            });
                          }}
                        >
                          Evaluate
                        </Button>
                      ) : null}
                    </Stack>
                  </PanelBody>
                </Panel>
              ))}
            </Stack>
          </div>

          {capabilityStatus ? (
            <div className="rounded-lg border border-blue-100 bg-blue-50 px-3 py-2 text-xs text-blue-800">
              {capabilityStatus}
            </div>
          ) : null}

          <Text size="xs" tone="muted">
            GROUPBY / PIVOTBY matrix functions are available in the formula engine when the capability flag is enabled.
          </Text>
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
