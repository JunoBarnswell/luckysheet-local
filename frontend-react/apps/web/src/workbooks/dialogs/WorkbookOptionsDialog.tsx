import { Button, CheckToggle, Dialog, LocationPicker, Select, Stack, Text, type LocationOption } from '@react-sheets/ui-system';

export interface WorkbookPreferences {
  defaultLocationId: string;
  autoSave: boolean;
  autoSync: boolean;
  offlineCache: boolean;
  xlsxCompatibility: 'balanced' | 'strict' | 'best-effort';
}

export interface WorkbookOptionsDialogProps {
  open: boolean;
  value: WorkbookPreferences;
  locationOptions: readonly LocationOption[];
  onChange: (value: WorkbookPreferences) => void;
  onClose: () => void;
  onSave: () => void;
}

export function WorkbookOptionsDialog({ open, value, locationOptions, onChange, onClose, onSave }: WorkbookOptionsDialogProps) {
  const patch = <K extends keyof WorkbookPreferences>(key: K, next: WorkbookPreferences[K]) => onChange({ ...value, [key]: next });
  return (
    <Dialog closeLabel="关闭文件中心选项" footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button onClick={onSave} size="sm" variant="brand">保存设置</Button></>} maxWidth="md" onClose={onClose} open={open} title="文件中心选项" testId="workbook-options-dialog">
      <Stack gap="lg">
        <Stack gap="xs"><Text as="label" htmlFor="options-default-location" size="sm" weight="medium">默认新建位置</Text><LocationPicker id="options-default-location" onChange={(event) => patch('defaultLocationId', event.target.value)} options={locationOptions} value={value.defaultLocationId} /></Stack>
        <Stack gap="sm"><Text size="sm" weight="medium">保存与缓存</Text><CheckToggle checked={value.autoSave} label="自动保存工作簿" onChange={(event) => patch('autoSave', event.target.checked)} /><CheckToggle checked={value.autoSync} label="在线时自动同步" onChange={(event) => patch('autoSync', event.target.checked)} /><CheckToggle checked={value.offlineCache} label="保留离线缓存" onChange={(event) => patch('offlineCache', event.target.checked)} /></Stack>
        <Stack gap="xs"><Text as="label" htmlFor="options-xlsx-compatibility" size="sm" weight="medium">XLSX 导入兼容级别</Text><Select id="options-xlsx-compatibility" onChange={(event) => patch('xlsxCompatibility', event.target.value as WorkbookPreferences['xlsxCompatibility'])} options={[{ value: 'balanced', label: '平衡（推荐）' }, { value: 'strict', label: '严格保留原始结构' }, { value: 'best-effort', label: '尽可能转换可编辑内容' }]} value={value.xlsxCompatibility} /></Stack>
      </Stack>
    </Dialog>
  );
}
