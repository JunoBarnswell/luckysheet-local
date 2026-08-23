import {
  AUTOMATION_WORKER_PROTOCOL,
  consumeAutomationWorkerRequest,
  isAutomationWorkerCancel,
  type AutomationWorkerResult,
} from './automation-worker';

export interface AutomationWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: AutomationWorkerResult): void;
}

/** The only production worker entry; it never evaluates DSL on the host. */
export function installAutomationWorkerEntry(scope: AutomationWorkerScope): () => void {
  const previous = scope.onmessage;
  const cancelled = new Set<string>();
  scope.onmessage = (event) => {
    if (isAutomationWorkerCancel(event.data)) {
      cancelled.add(event.data.taskId);
      return;
    }
    const result = consumeAutomationWorkerRequest(event.data);
    if (cancelled.delete(result.taskId)) {
      scope.postMessage({
        protocol: AUTOMATION_WORKER_PROTOCOL,
        taskId: result.taskId,
        status: 'cancelled',
      });
      return;
    }
    scope.postMessage(result);
  };
  return () => {
    scope.onmessage = previous;
  };
}

installAutomationWorkerEntry(self as unknown as AutomationWorkerScope);
