import { WorkbookSession } from './workbook-session';

/** Command-level fixture only; it does not exercise or stand in for a server planner. */
export function createRemoteReadySessionFixture(): WorkbookSession {
  const session = new WorkbookSession();
  session['runtime'].localOnly = false;
  session['runtime'].remoteConnected = true;
  return session;
}
