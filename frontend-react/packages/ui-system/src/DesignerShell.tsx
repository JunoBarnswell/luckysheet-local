import type { ReactNode } from 'react';
import { Box } from './layout';
import type { DesignerShellProps } from './shell-types';
import { RIBBON_DENSITY_CLASSES } from './shell-types';

/**
 * Excel Designer shell geometry is a product contract, not a responsive guess:
 * ribbon 167px, formula bar 48px, workbook 843px, status bar 22px at 1920x1080.
 * Sheet tabs live inside the workbook region like the reference Designer.
 */
export function DesignerShell({ children, floatingOverlay, formulaBar, formulaBarVisible = true, ribbon, ribbonVisible = true, isBusy, sheetTabs, statusBar, workspacePhase }: DesignerShellProps): ReactNode {
  return (
    <Box
      as="main"
      aria-busy={isBusy}
      aria-label="Spreadsheet Designer"
      className="designer-shell flex h-screen min-h-[680px] min-w-[960px] flex-col overflow-hidden bg-white text-slate-800"
      data-testid="designer-shell"
      data-workspace-phase={workspacePhase}
      role="application"
    >
      <Box as="section" className={`${ribbonVisible ? RIBBON_DENSITY_CLASSES.shell : 'hidden'} shrink-0 overflow-hidden border-b border-[#e7e7e7] bg-white`} data-testid="designer-ribbon" aria-hidden={!ribbonVisible}>
        {ribbon}
      </Box>
      <Box as="section" className={`${formulaBarVisible ? 'h-[48px]' : 'hidden'} shrink-0 overflow-hidden border-b border-[#e7e7e7] bg-white`} data-testid="designer-formula-bar" aria-hidden={!formulaBarVisible}>
        {formulaBar}
      </Box>
      <Box as="section" className="flex min-h-0 flex-1 flex-col overflow-hidden bg-white" data-testid="designer-workspace">
        <Box className="relative min-h-0 min-w-0 flex-1 overflow-hidden">
          {children}
          {floatingOverlay}
        </Box>
        <Box as="nav" aria-label="Worksheets" className="h-[29px] shrink-0 overflow-hidden border-t border-[#d9d9d9] bg-white" data-testid="designer-sheet-tabs">
          {sheetTabs}
        </Box>
      </Box>
      <Box as="footer" className="h-[22px] shrink-0 overflow-hidden border-t border-[#7aa58d] bg-[#217345] text-white" data-testid="designer-status-bar">
        {statusBar}
      </Box>
    </Box>
  );
}

export type { DesignerShellProps } from './shell-types';
