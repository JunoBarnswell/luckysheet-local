const COLLABORATION_ARGUMENT = '--react-sheets-collaboration-url=';

export function resolveBackendOrigin(value = 'http://127.0.0.1:8082') {
  const origin = new URL(value);
  if (!['http:', 'https:'].includes(origin.protocol) || origin.username || origin.password || origin.pathname !== '/' || origin.search || origin.hash) {
    throw new Error('REACT_SHEETS_API_ORIGIN must be an uncredentialed http(s) origin');
  }
  return origin;
}

export function resolveCollaborationUrl(origin) {
  const url = new URL('/ws', origin);
  url.protocol = origin.protocol === 'https:' ? 'wss:' : 'ws:';
  return url.toString();
}

export function collaborationArgument(url) {
  return `${COLLABORATION_ARGUMENT}${validateCollaborationUrl(url)}`;
}

export function parseDesktopConfig(argv) {
  const argument = argv.find((value) => value.startsWith(COLLABORATION_ARGUMENT));
  if (!argument) throw new Error('Desktop runtime is missing its collaboration endpoint');
  return Object.freeze({ collaborationUrl: validateCollaborationUrl(argument.slice(COLLABORATION_ARGUMENT.length)) });
}

function validateCollaborationUrl(value) {
  const url = new URL(value);
  if (!['ws:', 'wss:'].includes(url.protocol) || url.username || url.password || url.pathname !== '/ws' || url.search || url.hash) {
    throw new Error('Desktop collaboration endpoint is invalid');
  }
  return url.toString();
}
