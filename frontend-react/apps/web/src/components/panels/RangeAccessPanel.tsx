import React, { useEffect, useMemo, useState } from 'react';
import { Box, Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { EffectiveAccessRegion, RangeAccessGrant, RangeAccessLevel, RangeAccessRegion as ManagedRangeAccessRegion, RangeAccessRegionRequest } from '@react-sheets/protocol';
import { cellAddress, parseAddress } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../../i18n';

type PrincipalKind = 'subject' | 'group' | 'everyone';

const COPY = {
  'en-US': {
    title: 'Range access', accessRevision: (revision: number) => `Access revision ${revision}`,
    description: 'Range access controls who can read or edit workbook data. Worksheet protection remains a separate Excel setting.',
    effective: 'Your effective access', noRestrictions: 'No restricted ranges are assigned to you.',
    configured: (count: number) => `Configured regions (${count})`, noRules: 'No range access rules are configured.',
    newRegion: 'New region', modify: 'Modify region', add: 'Add region', edit: 'Edit', delete: 'Delete', cancel: 'Cancel', close: 'Close Panel',
    range: 'Worksheet range', sheetId: 'Worksheet ID', defaultAccess: 'Default access for other members', overrides: 'Member and group overrides',
    addOverride: 'Add override', create: 'Create region', save: 'Save changes', memberOrGroupId: 'Member or group ID',
    invalidRange: 'Enter a valid rectangular range such as C2:F100.', missingPrincipal: 'Enter a member or group ID.',
    loadError: 'Could not load range access regions', saveError: 'Could not save range access', deleteError: 'Could not delete range access',
    principal: { subject: 'Member', group: 'Group', everyone: 'Everyone' },
    access: { edit: 'Can edit', read: 'View only', hidden: 'No access' } as Record<RangeAccessLevel, string>,
  },
  'zh-CN': {
    title: '区域权限', accessRevision: (revision: number) => `权限版本 ${revision}`,
    description: '区域权限用于控制成员读取或编辑工作簿数据。工作表保护仍是独立的 Excel 设置。',
    effective: '你的有效权限', noRestrictions: '当前没有分配给你的受限区域。',
    configured: (count: number) => `已配置区域（${count}）`, noRules: '尚未配置区域权限。',
    newRegion: '新建区域', modify: '修改区域', add: '添加区域', edit: '编辑', delete: '删除', cancel: '取消', close: '关闭面板',
    range: '工作表区域', sheetId: '工作表 ID', defaultAccess: '其他成员的默认权限', overrides: '成员和群组例外权限',
    addOverride: '添加例外', create: '创建区域', save: '保存更改', memberOrGroupId: '成员或群组 ID',
    invalidRange: '请输入有效的矩形区域，例如 C2:F100。', missingPrincipal: '请输入成员或群组 ID。',
    loadError: '无法加载区域权限', saveError: '无法保存区域权限', deleteError: '无法删除区域权限',
    principal: { subject: '成员', group: '群组', everyone: '所有成员' },
    access: { edit: '可编辑', read: '仅查看', hidden: '禁止查看' } as Record<RangeAccessLevel, string>,
  },
} satisfies Record<Locale, {
  title: string; accessRevision: (revision: number) => string; description: string; effective: string; noRestrictions: string;
  configured: (count: number) => string; noRules: string; newRegion: string; modify: string; add: string; edit: string;
  delete: string; cancel: string; close: string; range: string; sheetId: string; defaultAccess: string; overrides: string;
  addOverride: string; create: string; save: string; memberOrGroupId: string; invalidRange: string; missingPrincipal: string;
  loadError: string; saveError: string; deleteError: string; principal: Record<PrincipalKind, string>; access: Record<RangeAccessLevel, string>;
}>;

export interface RangeAccessPanelProps {
  sheetId: string;
  selectedRange?: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  accessRevision: number;
  effectiveRegions: readonly EffectiveAccessRegion[];
  canManage: boolean;
  onList: () => Promise<ManagedRangeAccessRegion[]>;
  onCreate: (request: RangeAccessRegionRequest) => Promise<ManagedRangeAccessRegion>;
  onUpdate: (regionId: string, request: RangeAccessRegionRequest) => Promise<ManagedRangeAccessRegion>;
  onDelete: (regionId: string) => Promise<void>;
  onClose?: () => void;
}

function formatRange(range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): string {
  const start = cellAddress(range.startRow, range.startColumn);
  const end = cellAddress(range.endRow, range.endColumn);
  return start === end ? start : `${start}:${end}`;
}

function parseRange(value: string): Omit<RangeAccessRegionRequest, 'defaultAccess' | 'grants' | 'sheetId'>['range'] | null {
  const parts = value.trim().split(':');
  if (parts.length < 1 || parts.length > 2) return null;
  const start = parseAddress(parts[0] ?? '');
  const end = parseAddress(parts[1] ?? parts[0] ?? '');
  if (!start || !end || end.row < start.row || end.column < start.column) return null;
  return { sheetId: '', startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column };
}

function principalLabel(grant: RangeAccessGrant, locale: Locale): string {
  const copy = COPY[locale];
  if (grant.principal.kind === 'everyone') return copy.principal.everyone;
  return `${copy.principal[grant.principal.kind]}: ${grant.principal.id ?? ''}`;
}

export function RangeAccessPanel({ sheetId, selectedRange, accessRevision, effectiveRegions, canManage, onList, onCreate, onUpdate, onDelete, onClose, locale }: RangeAccessPanelProps & { locale: Locale }) {
  const copy = COPY[locale];
  const initialRange = selectedRange ? formatRange(selectedRange) : 'A1';
  const [regions, setRegions] = useState<ManagedRangeAccessRegion[]>([]);
  const [rangeText, setRangeText] = useState(initialRange);
  const [formSheetId, setFormSheetId] = useState(sheetId);
  const [defaultAccess, setDefaultAccess] = useState<RangeAccessLevel>('hidden');
  const [grants, setGrants] = useState<RangeAccessGrant[]>([]);
  const [principalKind, setPrincipalKind] = useState<PrincipalKind>('subject');
  const [principalId, setPrincipalId] = useState('');
  const [grantAccess, setGrantAccess] = useState<RangeAccessLevel>('edit');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const parsedRange = useMemo(() => parseRange(rangeText), [rangeText]);

  useEffect(() => {
    if (!editingId) {
      setRangeText(initialRange);
      setFormSheetId(sheetId);
    }
  }, [initialRange, sheetId, editingId]);

  useEffect(() => {
    let active = true;
    if (!canManage) return () => { active = false; };
    void onList().then((items) => { if (active) setRegions(items); }).catch((reason) => {
      if (active) setError(reason instanceof Error ? reason.message : copy.loadError);
    });
    return () => { active = false; };
  }, [accessRevision, canManage, onList]);

  const resetForm = () => {
    setEditingId(null);
    setDefaultAccess('hidden');
    setGrants([]);
    setPrincipalId('');
    setRangeText(initialRange);
    setFormSheetId(sheetId);
  };

  const editRegion = (region: ManagedRangeAccessRegion) => {
    setEditingId(region.id);
    setFormSheetId(region.sheetId);
    setRangeText(formatRange(region.range));
    setDefaultAccess(region.defaultAccess);
    setGrants(structuredClone(region.grants));
    setError('');
  };

  const addGrant = () => {
    if (principalKind !== 'everyone' && !principalId.trim()) {
      setError(copy.missingPrincipal);
      return;
    }
    const principal = principalKind === 'everyone' ? { kind: 'everyone' as const } : { kind: principalKind, id: principalId.trim() };
    setGrants((current) => [...current.filter((grant) => !(grant.principal.kind === principal.kind && grant.principal.id === principal.id)), { principal, access: grantAccess }]);
    setPrincipalId('');
    setError('');
  };

  const save = async () => {
    if (!parsedRange || !formSheetId) {
      setError(copy.invalidRange);
      return;
    }
    setBusy(true);
    setError('');
    try {
      const request: RangeAccessRegionRequest = {
        sheetId: formSheetId,
        range: { ...parsedRange, sheetId: formSheetId },
        defaultAccess,
        grants: structuredClone(grants),
      };
      const saved = editingId ? await onUpdate(editingId, request) : await onCreate(request);
      setRegions((current) => [...current.filter((region) => region.id !== saved.id), saved]);
      resetForm();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : copy.saveError);
    } finally {
      setBusy(false);
    }
  };

  const remove = async (regionId: string) => {
    setBusy(true);
    setError('');
    try {
      await onDelete(regionId);
      setRegions((current) => current.filter((region) => region.id !== regionId));
      if (editingId === regionId) resetForm();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : copy.deleteError);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <Inline gap="sm" className="w-full items-center justify-between">
          <PanelTitle size="sm">{copy.title}</PanelTitle>
          {canManage ? <Text size="xs" tone="subtle">{copy.accessRevision(accessRevision)}</Text> : null}
        </Inline>
      </PanelHeader>
      <PanelBody className="overflow-y-auto p-4">
        <Stack gap="md">
          <Text size="xs" tone="subtle">{copy.description}</Text>
          {!canManage ? (
            <Stack gap="xs">
              <Text size="xs" weight="semibold">{copy.effective}</Text>
              {effectiveRegions.length === 0 ? <Text size="xs" tone="subtle">{copy.noRestrictions}</Text> : effectiveRegions.map((region, index) => (
                <Box key={`${region.range.sheetId}-${index}`} className="rounded-lg border border-slate-200 bg-white p-2">
                  <Text size="xs" weight="medium">{region.range.sheetId}!{formatRange(region.range)}</Text>
                  <Text size="xs" tone="subtle">{copy.access[region.access]}</Text>
                </Box>
              ))}
            </Stack>
          ) : (
            <>
              <Stack gap="xs">
                <Inline gap="sm" className="items-center justify-between">
                  <Text size="xs" weight="semibold">{copy.configured(regions.length)}</Text>
                  <Button variant="ghost" size="xs" onClick={() => { resetForm(); setError(''); }}>{copy.newRegion}</Button>
                </Inline>
                {regions.length === 0 ? <Text size="xs" tone="subtle">{copy.noRules}</Text> : regions.map((region) => (
                  <Box key={region.id} className="rounded-lg border border-slate-200 bg-white p-2">
                    <Inline gap="sm" className="items-start justify-between">
                      <Stack gap="xs" className="min-w-0 flex-1">
                        <Text size="xs" weight="semibold">{region.sheetId}!{formatRange(region.range)} · {copy.access[region.defaultAccess]}</Text>
                        {region.grants.map((grant) => <Text key={`${grant.principal.kind}:${grant.principal.id ?? ''}`} size="xs" tone="subtle">{principalLabel(grant, locale)} · {copy.access[grant.access]}</Text>)}
                      </Stack>
                      <Inline gap="xs">
                        <Button variant="ghost" size="xs" onClick={() => editRegion(region)}>{copy.edit}</Button>
                        <Button variant="ghost" size="xs" disabled={busy} onClick={() => void remove(region.id)}>{copy.delete}</Button>
                      </Inline>
                    </Inline>
                  </Box>
                ))}
              </Stack>
              <Box className="border-t border-slate-200 pt-3">
                <Stack gap="sm">
                  <Text size="xs" weight="semibold">{editingId ? copy.modify : copy.add}</Text>
                  <TextInput aria-label={copy.range} value={rangeText} onChange={(event) => setRangeText(event.target.value)} placeholder="C2:F100" />
                  <TextInput aria-label={copy.sheetId} value={formSheetId} onChange={(event) => setFormSheetId(event.target.value)} placeholder={copy.sheetId} />
                  <Box>
                    <Text size="xs" className="mb-1">{copy.defaultAccess}</Text>
                    <Select sizeVariant="sm" value={defaultAccess} onChange={(event) => setDefaultAccess(event.target.value as RangeAccessLevel)}>
                      <option value="edit">{copy.access.edit}</option><option value="read">{copy.access.read}</option><option value="hidden">{copy.access.hidden}</option>
                    </Select>
                  </Box>
                  <Box className="rounded-lg border border-slate-100 p-2">
                    <Text size="xs" weight="medium" className="mb-2">{copy.overrides}</Text>
                    <Stack gap="xs">
                      <Select sizeVariant="sm" value={principalKind} onChange={(event) => setPrincipalKind(event.target.value as PrincipalKind)}>
                        <option value="subject">{copy.principal.subject}</option><option value="group">{copy.principal.group}</option><option value="everyone">{copy.principal.everyone}</option>
                      </Select>
                      {principalKind !== 'everyone' ? <TextInput aria-label={copy.memberOrGroupId} value={principalId} onChange={(event) => setPrincipalId(event.target.value)} placeholder={copy.memberOrGroupId} /> : null}
                      <Inline gap="xs">
                        <Select sizeVariant="sm" value={grantAccess} onChange={(event) => setGrantAccess(event.target.value as RangeAccessLevel)}>
                          <option value="edit">{copy.access.edit}</option><option value="read">{copy.access.read}</option><option value="hidden">{copy.access.hidden}</option>
                        </Select>
                        <Button variant="secondary" size="sm" onClick={addGrant}>{copy.addOverride}</Button>
                      </Inline>
                      {grants.map((grant, index) => (
                        <Inline key={`${grant.principal.kind}:${grant.principal.id ?? ''}`} gap="xs" className="items-center">
                          <Text size="xs" className="min-w-0 flex-1 truncate">{principalLabel(grant, locale)}</Text>
                          <Select sizeVariant="sm" value={grant.access} onChange={(event) => setGrants((current) => current.map((entry, entryIndex) => entryIndex === index ? { ...entry, access: event.target.value as RangeAccessLevel } : entry))}>
                            <option value="edit">{copy.access.edit}</option><option value="read">{copy.access.read}</option><option value="hidden">{copy.access.hidden}</option>
                          </Select>
                          <Button variant="ghost" size="xs" onClick={() => setGrants((current) => current.filter((_, entryIndex) => entryIndex !== index))}>{locale === 'zh-CN' ? '移除' : 'Remove'}</Button>
                        </Inline>
                      ))}
                    </Stack>
                  </Box>
                  {error ? <Text size="xs" className="text-rose-700">{error}</Text> : null}
                  <Inline gap="sm" className="justify-end">
                    {editingId ? <Button variant="ghost" size="sm" disabled={busy} onClick={resetForm}>{copy.cancel}</Button> : null}
                    <Button variant="primary" size="sm" disabled={busy || !parsedRange} onClick={() => void save()}>{busy ? (locale === 'zh-CN' ? '正在保存…' : 'Saving…') : editingId ? copy.save : copy.create}</Button>
                  </Inline>
                </Stack>
              </Box>
            </>
          )}
        </Stack>
      </PanelBody>
      {onClose ? <PanelFooter className="border-t border-slate-200 px-4 py-2"><Button variant="ghost" size="sm" onClick={onClose}>{copy.close}</Button></PanelFooter> : null}
    </Panel>
  );
}
