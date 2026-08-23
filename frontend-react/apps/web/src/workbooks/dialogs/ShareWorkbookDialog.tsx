import { useEffect, useState } from 'react';
import { Button, Dialog, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { WorkbookRole } from '../types';

export interface ShareWorkbookDialogValue {
  subject: string;
  role: Exclude<WorkbookRole, 'owner'>;
}

export interface ShareWorkbookDialogProps {
  open: boolean;
  workbookName: string;
  onClose: () => void;
  onSubmit: (value: ShareWorkbookDialogValue) => void;
  submitting?: boolean;
}

export function ShareWorkbookDialog({ open, workbookName, onClose, onSubmit, submitting = false }: ShareWorkbookDialogProps) {
  const [subject, setSubject] = useState('');
  const [role, setRole] = useState<Exclude<WorkbookRole, 'owner'>>('viewer');
  useEffect(() => { if (open) { setSubject(''); setRole('viewer'); } }, [open]);
  const submit = () => { const next = subject.trim(); if (next && !submitting) onSubmit({ subject: next, role }); };
  return (
    <Dialog
      closeLabel="关闭共享设置"
      description={`管理 ${workbookName} 的访问权限`}
      footer={<><Button onClick={onClose} size="sm" variant="ghost">取消</Button><Button disabled={!subject.trim()} loading={submitting} onClick={submit} size="sm" variant="brand">发送邀请</Button></>}
      maxWidth="md"
      onClose={onClose}
      open={open}
      title="共享工作簿"
      testId="share-workbook-dialog"
    >
      <Stack gap="md">
        <Stack gap="xs"><Text as="label" htmlFor="share-subject" size="sm" weight="medium">成员邮箱或账号</Text><TextInput autoFocus id="share-subject" onChange={(event) => setSubject(event.target.value)} placeholder="name@example.com" value={subject} /></Stack>
        <Stack gap="xs"><Text as="label" htmlFor="share-role" size="sm" weight="medium">访问权限</Text><Select id="share-role" onChange={(event) => setRole(event.target.value as Exclude<WorkbookRole, 'owner'>)} options={[{ value: 'editor', label: '可编辑' }, { value: 'commenter', label: '可评论' }, { value: 'viewer', label: '只读' }]} value={role} /></Stack>
        <Text size="xs" tone="muted">权限会随空间成员权限和工作簿 ACL 取当前用户可获得的最高有效角色。</Text>
      </Stack>
    </Dialog>
  );
}
