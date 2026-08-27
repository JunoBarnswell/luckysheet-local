const { contextBridge } = require('electron');

const collaborationArgumentPrefix = '--react-sheets-collaboration-url=';
const collaborationArgument = process.argv.find((value) => value.startsWith(collaborationArgumentPrefix));

if (!collaborationArgument) {
  throw new Error('Desktop runtime is missing its collaboration endpoint');
}

contextBridge.exposeInMainWorld('reactSheetsDesktopConfig', Object.freeze({
  collaborationUrl: collaborationArgument.slice(collaborationArgumentPrefix.length),
}));
