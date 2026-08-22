import { Box, Inline } from '@react-sheets/ui-system';
import { AppShell } from './components/AppShell';
import { FeatureSidebar } from './components/FeatureSidebar';
import { FormulaBar } from './components/FormulaBar';
import { Ribbon } from './components/Ribbon';
import { SheetCanvas } from './components/SheetCanvas';
import { SheetTabs } from './components/SheetTabs';
import { StatusBar } from './components/StatusBar';
import { getInitialWorkspacePhase, useWorkspaceState } from './state/workspace';

export default function App() {
  const { actions, state } = useWorkspaceState({ initialPhase: getInitialWorkspacePhase() });
  const isBusy = state.phase !== 'ready';

  return (
    <AppShell
      formulaBar={
        <FormulaBar
          cellName={state.activeCell}
          disabled={isBusy}
          formula={state.formulaDraft}
          onCancel={() => actions.notify('Formula edit cancelled')}
          onChange={actions.setFormulaDraft}
          onCommit={actions.commitFormula}
          phase={state.phase}
        />
      }
      isBusy={isBusy}
      notice={state.notice}
      onMenu={() => actions.notify('Workbook menu opened')}
      onShare={() => actions.notify('Share controls are ready')}
      ribbon={
        <Ribbon
          activeTab={state.ribbonTab}
          onAction={(action) => {
            if (action === 'undo') actions.undo();
            else if (action === 'redo') actions.redo();
            else actions.notify(`${action.replace('-', ' ')} command is not registered yet`);
          }}
          onTabChange={actions.setRibbonTab}
          phase={state.phase}
        />
      }
      saveState={state.saveState}
      sheetTabs={<SheetTabs activeSheetId={state.activeSheetId} disabled={isBusy} onAdd={actions.addSheet} onSelect={actions.selectSheet} sheets={state.sheets} />}
      statusBar={
        <StatusBar
          activeCell={state.activeCell}
          onOpenShortcuts={() => actions.notify('Keyboard shortcuts opened')}
          onZoomChange={actions.setZoom}
          phase={state.phase}
          saveState={state.saveState}
          sheetCount={state.sheets.length}
          zoom={state.zoom}
        />
      }
      title="Q3 Growth Planning"
      workspacePhase={state.phase}
    >
      <Inline gap="none" className="h-full min-h-0 w-full flex-nowrap">
        <Box className="min-w-0 flex-1">
          <SheetCanvas
            onCreateSheet={actions.addSheet}
            onMoveCell={actions.moveCell}
            onRetry={actions.retry}
            onSelectCell={actions.selectCell}
            phase={state.phase}
            selectedCell={state.activeCell}
            sheet={state.selectedSheet}
            zoom={state.zoom}
          />
        </Box>
        <FeatureSidebar activeCell={state.activeCell} activePanel={state.activePanel} onPanelChange={actions.setActivePanel} onRetry={actions.retry} phase={state.phase} sheet={state.selectedSheet} />
      </Inline>
    </AppShell>
  );
}
