import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { registerOfflineShell } from './offline-shell';
import './styles.css';

const rootElement = document.getElementById('root');

if (!rootElement) throw new Error('React Sheets root element is missing');

registerOfflineShell();

createRoot(rootElement).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
