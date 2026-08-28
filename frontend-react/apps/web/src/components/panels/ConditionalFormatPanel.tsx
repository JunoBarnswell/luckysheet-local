import React, { useState } from 'react';
import { Box, Button, CheckToggle, ColorPicker, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ConditionalFormatOperator, ConditionalFormatRule, ConditionalFormatType, RangeRef } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';
import { homeTemplate, homeText, resolveHomeLocale } from '../home/home-localization';

export interface ConditionalFormatPanelProps {
  sheetId: string;
  range: RangeRef;
  locale?: Locale;
  rules: ConditionalFormatRule[];
  onAddRule: (rule: ConditionalFormatRule) => void;
  onRemoveRule: (id: string) => void;
  onUpdateRule?: (id: string, patch: Partial<ConditionalFormatRule>) => void;
  onReorderRules?: (ruleIds: readonly string[]) => void;
  onClose?: () => void;
}

export function ConditionalFormatPanel({
  sheetId,
  range,
  locale,
  rules,
  onAddRule,
  onRemoveRule,
  onUpdateRule,
  onReorderRules,
  onClose,
}: ConditionalFormatPanelProps) {
  const activeLocale = resolveHomeLocale(locale);
  const [type, setType] = useState<ConditionalFormatType>('highlight');
  const [operator, setOperator] = useState<ConditionalFormatOperator>('greaterThan');
  const [value1, setValue1] = useState('50');
  const [bg, setBg] = useState('#dcfce7');
  const [color, setColor] = useState('#166534');
  const [stopIfTrue, setStopIfTrue] = useState(false);
  const [iconSet, setIconSet] = useState('threeTrafficLights1');
  const [topBottomPercent, setTopBottomPercent] = useState(false);
  const [editingRuleId, setEditingRuleId] = useState<string | null>(null);
  const [editingValue, setEditingValue] = useState('');

  const handleCreate = () => {
    const newRule: ConditionalFormatRule = {
      id: 'cf-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      ranges: [{ ...range, sheetId }],
      type,
      ...(type === 'colorScale' || type === 'dataBar' || type === 'iconSet' ? {} : { operator }),
      ...(type === 'topBottom' ? { topBottom: { direction: operator === 'bottom' ? 'bottom' as const : 'top' as const, rank: Math.max(1, Number(value1) || 10), percent: topBottomPercent } } : { value1 }),
      ...(type === 'iconSet' ? { iconSet, iconThresholds: [{ type: 'percent' as const, value: 0 }, { type: 'percent' as const, value: 33 }, { type: 'percent' as const, value: 67 }] } : {}),
      priority: rules.length + 1,
      stopIfTrue,
      style: {
        background: bg,
        textColor: color,
        bold: true,
      },
    };
    onAddRule(newRule);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">{homeText(activeLocale, 'conditionalFormatting')}</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              {homeText(activeLocale, 'formatType')}
            </Text>
            <Select
              value={type}
              onChange={(e) => { const next = e.target.value as ConditionalFormatType; setType(next); if (next === 'topBottom') setOperator('top'); else if (operator === 'top' || operator === 'bottom') setOperator('greaterThan'); }}
              sizeVariant="sm"
            >
              <option value="highlight">{homeText(activeLocale, 'highlightRules')}</option>
              <option value="colorScale">{homeText(activeLocale, 'colorScale')}</option>
              <option value="dataBar">{homeText(activeLocale, 'dataBars')}</option>
              <option value="iconSet">{homeText(activeLocale, 'iconSets')}</option>
              <option value="topBottom">{homeText(activeLocale, 'topBottomRules')}</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              {homeText(activeLocale, 'ruleCondition')}
            </Text>
            <Select
              value={operator}
              onChange={(e) => setOperator(e.target.value as ConditionalFormatOperator)}
              sizeVariant="sm"
            >
              <option value="greaterThan">Greater Than (&gt;)</option>
              <option value="lessThan">Less Than (&lt;)</option>
              <option value="equal">Equal To (=)</option>
              <option value="between">Between Range</option>
              <option value="containsText">Text Contains</option>
              <option value="notContainsText">Text Does Not Contain</option>
              <option value="duplicate">Duplicate Values</option>
              <option value="unique">Unique Values</option>
              <option value="formula">Use a formula</option>
              <option value="notEqual">Not Equal To</option>
              {type === 'topBottom' ? <><option value="top">Top</option><option value="bottom">Bottom</option></> : null}
            </Select>
          </Box>

          {type === 'iconSet' ? <Select aria-label="Icon set" sizeVariant="sm" value={iconSet} onChange={(event) => setIconSet(event.target.value)}><option value="threeTrafficLights1">三色交通灯</option><option value="threeArrows">三色箭头</option><option value="fourRatings">四级评级</option><option value="fiveQuarters">五级刻度</option></Select> : null}
          {type === 'topBottom' ? <CheckToggle checked={topBottomPercent} label="按百分比" onChange={(event) => setTopBottomPercent(event.target.checked)} /> : null}
          <CheckToggle checked={stopIfTrue} label={homeText(activeLocale, 'stopIfTrue')} onChange={(event) => setStopIfTrue(event.target.checked)} />

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              {homeText(activeLocale, 'thresholdValue')}
            </Text>
            <TextInput
              value={value1}
              onChange={(e) => setValue1(e.target.value)}
              placeholder="100"
            />
          </Box>

          <Box className="grid grid-cols-2 gap-2">
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                {homeText(activeLocale, 'cellFill')}
              </Text>
              <Stack gap="xs">
                <ColorPicker color={bg} onChange={setBg} />
                <TextInput value={bg} onChange={(e) => setBg(e.target.value)} className="h-8 text-xs font-mono" />
              </Stack>
            </Box>
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                {homeText(activeLocale, 'textColor')}
              </Text>
              <Stack gap="xs">
                <ColorPicker color={color} onChange={setColor} />
                <TextInput value={color} onChange={(e) => setColor(e.target.value)} className="h-8 text-xs font-mono" />
              </Stack>
            </Box>
          </Box>

          <Button variant="primary" size="sm" icon="plus" onClick={handleCreate}>
            {homeText(activeLocale, 'applyFormattingRule')}
          </Button>

          {rules.length > 0 ? (
            <Box className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                {homeTemplate(activeLocale, 'activeRules', { count: rules.length })}
              </Text>
              <Stack gap="xs">
                {[...rules].sort((left, right) => (left.priority ?? 0) - (right.priority ?? 0)).map((r, index, ordered) => (
                  <Box
                    key={r.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <Stack gap="none" className="min-w-0">
                      <Text size="xs" weight="medium" className="text-slate-800">
                        {r.priority ?? index + 1}. {r.operator ?? r.type} {r.value1}
                      </Text>
                      <Text size="xs" tone="subtle">{r.type} · {homeText(activeLocale, 'appliesTo')}: {r.ranges.map((entry) => `${entry.startRow + 1}:${entry.startColumn + 1}-${entry.endRow + 1}:${entry.endColumn + 1}`).join(', ')}</Text>
                      {editingRuleId === r.id ? <Stack gap="xs" className="mt-1"><TextInput value={editingValue} onChange={(event) => setEditingValue(event.target.value)} /><Button size="xs" variant="secondary" onClick={() => { const numeric = Number(editingValue); onUpdateRule?.(r.id, { value1: Number.isFinite(numeric) && editingValue.trim() !== '' ? numeric : editingValue }); setEditingRuleId(null); }}>{homeText(activeLocale, 'saveRule')}</Button></Stack> : null}
                    </Stack>
                    <Inline gap="none" className="shrink-0">
                      <Button variant="ghost" size="xs" aria-label={homeText(activeLocale, 'moveRuleUp')} icon="chevron-up" iconOnly disabled={index === 0 || !onReorderRules} onClick={() => { if (!onReorderRules) return; const ids = ordered.map((entry) => entry.id); [ids[index - 1], ids[index]] = [ids[index]!, ids[index - 1]!]; onReorderRules(ids); }} />
                      <Button variant="ghost" size="xs" aria-label={homeText(activeLocale, 'moveRuleDown')} icon="chevron-down" iconOnly disabled={index === ordered.length - 1 || !onReorderRules} onClick={() => { if (!onReorderRules) return; const ids = ordered.map((entry) => entry.id); [ids[index], ids[index + 1]] = [ids[index + 1]!, ids[index]!]; onReorderRules(ids); }} />
                      <Button variant="ghost" size="xs" aria-label={homeText(activeLocale, 'editRule')} icon="pencil" iconOnly disabled={!onUpdateRule} onClick={() => { setEditingRuleId(r.id); setEditingValue(String(r.value1 ?? '')); }} />
                      <Button variant="ghost" size="xs" icon="trash" iconOnly onClick={() => onRemoveRule(r.id)} className="text-rose-600 hover:bg-rose-50" />
                    </Inline>
                  </Box>
                ))}
              </Stack>
            </Box>
          ) : null}
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>
            {homeText(activeLocale, 'closePanel')}
          </Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
