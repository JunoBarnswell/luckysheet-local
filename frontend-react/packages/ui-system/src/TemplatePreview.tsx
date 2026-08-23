import { cn } from './cn';
import { Icon } from './Icon';
import { Box, Inline } from './layout';

export type TemplatePreviewKind = 'blank' | 'template' | 'import' | 'pivot' | 'project' | 'budget';

export interface TemplatePreviewProps {
  kind: TemplatePreviewKind;
  className?: string;
  compact?: boolean;
}

function GridPreview() {
  return (
    <Box className="relative grid h-[90px] w-[132px] grid-cols-3 grid-rows-4 overflow-hidden rounded-sm border border-slate-200 bg-white shadow-sm">
      {Array.from({ length: 12 }, (_, index) => <Box key={index} className={cn('border-b border-r border-slate-100', index < 3 && 'bg-brand-soft')} />)}
      <Box className="absolute mt-[22px] h-0.5 w-[42px] bg-brand" />
    </Box>
  );
}

function FilePreview({ kind }: { kind: 'template' | 'import' }) {
  return (
    <Box className="relative flex h-[102px] w-[76px] items-center justify-center rounded-md border border-slate-200 bg-white shadow-sm">
      <Box className="absolute right-0 top-0 h-5 w-5 rounded-bl-md border-b border-l border-slate-200 bg-slate-50" />
      <Box className={cn('flex h-9 w-9 items-center justify-center rounded-md text-white', kind === 'import' ? 'bg-brand' : 'bg-brand/80')}>
        <Icon name="file-spreadsheet" size="lg" />
      </Box>
      <Box className={cn('absolute -bottom-2 -right-2 flex h-8 w-8 items-center justify-center rounded-full border-2 border-white shadow-sm', kind === 'import' ? 'bg-brand text-white' : 'bg-white text-brand')}>
        <Icon name={kind === 'import' ? 'upload' : 'sparkles'} size="sm" />
      </Box>
    </Box>
  );
}

function PivotPreview() {
  return (
    <Box className="relative flex h-[102px] w-[86px] items-end justify-center gap-1 rounded-md border border-slate-200 bg-white px-3 pb-3 shadow-sm">
      <Box className="h-12 w-3 rounded-t bg-brand/30" />
      <Box className="h-20 w-3 rounded-t bg-brand/60" />
      <Box className="h-14 w-3 rounded-t bg-brand" />
      <Box className="absolute left-3 top-3 h-2 w-10 rounded bg-slate-100" />
    </Box>
  );
}

function ProjectPreview() {
  return (
    <Box className="flex h-[102px] w-[122px] flex-col justify-center gap-2 rounded-md border border-slate-200 bg-white px-3 shadow-sm">
      {[['w-9', 'w-14'], ['w-12', 'w-20'], ['w-8', 'w-16']].map(([label, bar], index) => (
        <Inline key={index} gap="xs" className="h-3">
          <Box className={cn('h-2 rounded bg-slate-200', label)} />
          <Box className={cn('h-2 rounded bg-brand', bar)} />
        </Inline>
      ))}
    </Box>
  );
}

function BudgetPreview() {
  return (
    <Box className="relative flex h-[102px] w-[105px] items-center justify-center rounded-md border border-slate-200 bg-white shadow-sm">
      <Box className="h-16 w-16 rounded-full border-[13px] border-brand/20 border-t-brand border-r-brand" />
      <Box className="absolute right-3 top-3 h-2.5 w-2.5 rounded-sm bg-slate-300" />
      <Box className="absolute right-3 top-8 h-2.5 w-5 rounded-sm bg-slate-200" />
      <Box className="absolute right-3 top-[52px] h-2.5 w-3 rounded-sm bg-brand" />
    </Box>
  );
}

export function TemplatePreview({ kind, className, compact = false }: TemplatePreviewProps) {
  return (
    <Box className={cn('relative flex h-[130px] items-center justify-center', compact && 'h-[92px]', className)}>
      {kind === 'blank' ? <GridPreview /> : null}
      {kind === 'template' || kind === 'import' ? <FilePreview kind={kind} /> : null}
      {kind === 'pivot' ? <PivotPreview /> : null}
      {kind === 'project' ? <ProjectPreview /> : null}
      {kind === 'budget' ? <BudgetPreview /> : null}
    </Box>
  );
}
