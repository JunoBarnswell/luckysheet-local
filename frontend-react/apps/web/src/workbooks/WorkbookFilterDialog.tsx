import { Button, CheckToggle, Dialog, Stack, Text } from '@react-sheets/ui-system';

export interface WorkbookFilterValues {
  favoritesOnly: boolean;
  localOnly: boolean;
  sharedOnly: boolean;
  needsSync: boolean;
}

export interface WorkbookFilterDialogProps {
  open: boolean;
  value: WorkbookFilterValues;
  onChange: (value: WorkbookFilterValues) => void;
  onClose: () => void;
  onApply: () => void;
}

export function WorkbookFilterDialog({ open, value, onChange, onClose, onApply }: WorkbookFilterDialogProps) {
  const set = (key: keyof WorkbookFilterValues, checked: boolean) => onChange({ ...value, [key]: checked });
  return (
    <Dialog
      closeLabel="关闭筛选"
      description="选择文件中心要显示的工作簿"
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button onClick={onApply} size="sm" variant="brand">应用筛选</Button></>}
      maxWidth="sm"
      onClose={onClose}
      open={open}
      title="筛选工作簿"
      testId="workbook-filter-dialog"
    >
      <Stack gap="md">
        <Text size="sm" weight="medium">文件条件</Text>
        <Stack gap="sm">
          <CheckToggle checked={value.favoritesOnly} label="仅显示星标文件" onChange={(event) => set('favoritesOnly', event.target.checked)} />
          <CheckToggle checked={value.localOnly} label="仅显示本地文件" onChange={(event) => set('localOnly', event.target.checked)} />
          <CheckToggle checked={value.sharedOnly} label="仅显示与我共享" onChange={(event) => set('sharedOnly', event.target.checked)} />
          <CheckToggle checked={value.needsSync} label="仅显示待同步或冲突文件" onChange={(event) => set('needsSync', event.target.checked)} />
        </Stack>
      </Stack>
    </Dialog>
  );
}
