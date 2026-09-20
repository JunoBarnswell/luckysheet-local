import { Box, Button, Text, TextInput, DataTable } from '@react-sheets/ui-system';
import { useState, type FormEvent } from 'react';
import { getAuthSession } from './session';
import type { AuthSnapshot } from './oidc';

export function LoginPage({ snapshot }: { snapshot: AuthSnapshot }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [token, setToken] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const bootstrap = snapshot.bootstrapRequired === true;
  async function submit(event: FormEvent<HTMLElement>) {
    event.preventDefault(); setBusy(true); setError('');
    try {
      if (bootstrap) await getAuthSession().bootstrap(token, username, password, displayName);
      else await getAuthSession().authenticate(username, password);
      setPassword(''); setToken('');
    } catch (cause) { setError(cause instanceof Error ? cause.message : '登录失败'); }
    finally { setBusy(false); }
  }
  return <main className="flex min-h-screen items-center justify-center bg-slate-50 p-6">
    <Box as="form" className="w-full max-w-md space-y-5 rounded-xl border bg-white p-8 shadow-sm" onSubmit={submit}>
      <Box><h1 className="text-2xl font-semibold">{bootstrap ? '初始化管理员' : '登录 React Sheets'}</h1>
        <p className="mt-2 text-sm text-slate-600">{bootstrap ? '在服务器本机输入数据目录中的一次性初始化凭据。' : '使用此服务器的账号打开和协作编辑工作簿。'}</p></Box>
      {bootstrap && <Text as="label" className="block">初始化凭据<TextInput required type="password" className="mt-1 w-full rounded border p-2" value={token} onChange={e => setToken(e.target.value)} autoComplete="off" /></Text>}
      <Text as="label" className="block">用户名<TextInput required className="mt-1 w-full rounded border p-2" value={username} onChange={e => setUsername(e.target.value)} autoComplete="username" /></Text>
      {bootstrap && <Text as="label" className="block">显示名称<TextInput required className="mt-1 w-full rounded border p-2" value={displayName} onChange={e => setDisplayName(e.target.value)} /></Text>}
      <Text as="label" className="block">密码<TextInput required type="password" minLength={bootstrap ? 12 : undefined} className="mt-1 w-full rounded border p-2" value={password} onChange={e => setPassword(e.target.value)} autoComplete={bootstrap ? 'new-password' : 'current-password'} /></Text>
      {error && <p role="alert" className="text-sm text-red-700">{error}</p>}
      <Button type="submit" disabled={busy} className="w-full rounded bg-emerald-700 p-2 text-white disabled:opacity-50">{busy ? '正在提交…' : bootstrap ? '创建管理员' : '登录'}</Button>
    </Box>
  </main>;
}
