import { Box, Button, Text, TextInput, DataTable } from '@react-sheets/ui-system';
import { useEffect, useState, type FormEvent } from 'react';
import { getAuthSession, type LocalUser } from './session';
import { navigate } from '../app-routing';

/** Account administration container; passwords never enter workbook state. */
export function AdminUsersPage() {
  const [users, setUsers] = useState<LocalUser[]>([]);
  const [username, setUsername] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [password, setPassword] = useState('');
  const [selected, setSelected] = useState<LocalUser | null>(null);
  const [resetPassword, setResetPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const auth = getAuthSession();
  async function load() {
    const response = await auth.request('/api/admin/users', 'GET');
    const result: LocalUser[] = await response.json();
    if (!Array.isArray(result)) throw new Error('用户列表契约无效');
    setUsers(result);
  }
  useEffect(() => { void load().catch(cause => setError(cause instanceof Error ? cause.message : '无法加载用户')); }, []);
  async function run(action: () => Promise<void>) {
    setBusy(true); setError('');
    try { await action(); await load(); }
    catch (cause) { setError(cause instanceof Error ? cause.message : '操作失败'); }
    finally { setBusy(false); }
  }
  function create(event: FormEvent<HTMLElement>) {
    event.preventDefault();
    void run(async () => {
      await auth.request('/api/admin/users', 'POST', { username, displayName, password, admin: false });
      setUsername(''); setDisplayName(''); setPassword('');
    });
  }
  return <main className="mx-auto max-w-4xl space-y-6 p-8">
    <Button onClick={() => navigate('/workbooks')}>← 返回文件中心</Button>
    <h1 className="text-2xl font-semibold">用户管理</h1>
    <p className="text-sm text-slate-600">复制用户 ID 可用于工作簿共享。禁用账号或重置密码将结束该账号的已有会话。</p>
    {error && <p role="alert" className="text-red-700">{error}</p>}
    <Box as="form" onSubmit={create} className="flex flex-wrap gap-3 rounded border p-4">
      <Text as="label">用户名<TextInput required className="block rounded border p-2" value={username} onChange={e => setUsername(e.target.value)} /></Text>
      <Text as="label">显示名称<TextInput required className="block rounded border p-2" value={displayName} onChange={e => setDisplayName(e.target.value)} /></Text>
      <Text as="label">初始密码<TextInput required type="password" minLength={12} autoComplete="new-password" className="block rounded border p-2" value={password} onChange={e => setPassword(e.target.value)} /></Text>
      <Button type="submit" disabled={busy} className="self-end rounded bg-emerald-700 p-2 text-white">创建用户</Button>
    </Box>
    <DataTable rows={users} rowKey={user => user.id} columns={[
      { key: 'name', header: '用户', render: user => <Box>{user.displayName}<Text as="small" className="block">{user.username}{user.admin ? ' · 管理员' : ''}</Text></Box> },
      { key: 'id', header: '用户 ID', render: user => <Text className="select-all text-xs">{user.id}</Text> },
      { key: 'enabled', header: '状态', render: user => user.enabled ? '启用' : '禁用' },
      { key: 'actions', header: '操作', render: user => <Box className="flex gap-3">
        <Button disabled={busy || user.id === auth.getSnapshot().subject} onClick={() => void run(async () => { await auth.request(`/api/admin/users/${encodeURIComponent(user.id)}`, 'PATCH', { enabled: !user.enabled }); })}>{user.enabled ? '禁用' : '启用'}</Button>
        <Button disabled={busy} onClick={() => { setSelected(user); setResetPassword(''); }}>重置密码</Button>
      </Box> },
    ]} />
    {selected && <Box as="form" className="space-x-3 rounded border p-4" onSubmit={event => { event.preventDefault(); void run(async () => { await auth.request(`/api/admin/users/${encodeURIComponent(selected.id)}/password`, 'POST', { password: resetPassword }); setSelected(null); setResetPassword(''); }); }}>
      <Text as="label">{selected.displayName} 的新密码 <TextInput required type="password" autoComplete="new-password" minLength={12} className="rounded border p-2" value={resetPassword} onChange={e => setResetPassword(e.target.value)} /></Text>
      <Button type="submit" disabled={busy}>确认重置</Button><Button type="button" onClick={() => setSelected(null)}>取消</Button>
    </Box>}
  </main>;
}
