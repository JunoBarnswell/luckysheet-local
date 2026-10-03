import { StrictMode } from 'react';
import { resolveShareToken } from '@react-sheets/spreadsheet-app';

// Capture and erase the route capability before any app requests or effects.
resolveShareToken();
import { createRoot } from 'react-dom/client';
import App from './App';
import { AuthProvider } from './auth/AuthProvider';
import { ApplicationServicesProvider } from './ApplicationServicesProvider';
import { registerOfflineShell } from './offline-shell';
import './styles.css';

const rootElement = document.getElementById('root');

if (!rootElement) throw new Error('React Sheets root element is missing');

registerOfflineShell();

createRoot(rootElement).render(
  <StrictMode>
    <AuthProvider>
      <ApplicationServicesProvider>
        <App />
      </ApplicationServicesProvider>
    </AuthProvider>
  </StrictMode>,
);
