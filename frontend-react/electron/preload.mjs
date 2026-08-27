import { contextBridge } from 'electron';
import { parseDesktopConfig } from './desktop-config.mjs';

contextBridge.exposeInMainWorld('reactSheetsDesktopConfig', parseDesktopConfig(process.argv));
