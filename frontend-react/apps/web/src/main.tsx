import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { AuthProvider } from './auth/AuthProvider';
import { ApplicationServicesProvider } from './ApplicationServicesProvider';
import { registerOfflineShell } from './offline-shell';
import './styles.css';
import { initializeKernel, KernelInvocationError } from '@react-sheets/kernel-client';
import { Box, StatePanel } from '@react-sheets/ui-system';

const rootElement = document.getElementById('root');

if (!rootElement) throw new Error('React Sheets root element is missing');
const mountNode = rootElement;

registerOfflineShell();

function renderKernelFailure(cause: unknown): void {
  const error = cause instanceof KernelInvocationError ? cause : new Error(cause instanceof Error ? cause.message : 'Spreadsheet kernel initialization failed');
  const detail = cause instanceof KernelInvocationError ? `（错误码：${cause.code}；恢复：${cause.recovery}）` : '';
  createRoot(mountNode).render(
    <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8">
      <StatePanel kind="error" title="表格内核不可用" description={`${error.message}${detail}`} />
    </Box>,
  );
}

void initializeKernel()
  .then(() => {
    createRoot(mountNode).render(
      <StrictMode>
        <AuthProvider>
          <ApplicationServicesProvider>
            <App />
          </ApplicationServicesProvider>
        </AuthProvider>
      </StrictMode>,
    );
  })
  .catch(renderKernelFailure);
