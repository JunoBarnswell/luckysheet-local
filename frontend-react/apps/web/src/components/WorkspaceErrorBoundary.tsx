import { Component, type ErrorInfo, type ReactNode } from 'react';
import { Panel, StatePanel } from '@react-sheets/ui-system';

interface WorkspaceErrorBoundaryProps { children: ReactNode; }
interface WorkspaceErrorBoundaryState { error: Error | null; }

export class WorkspaceErrorBoundary extends Component<WorkspaceErrorBoundaryProps, WorkspaceErrorBoundaryState> {
  state: WorkspaceErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): WorkspaceErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error('Workspace runtime error', error, info.componentStack);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <Panel className="m-6 flex h-[calc(100vh-3rem)] items-center justify-center">
        <StatePanel
          kind="error"
          title="Workbook runtime error"
          description={this.state.error.message}
          actionLabel="Reload workbook"
          onAction={() => window.location.reload()}
        />
      </Panel>
    );
  }
}
