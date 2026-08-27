import type { ComponentProps, ReactNode } from 'react';
import type { CellAddress, FormulaDependency, FormulaValue } from '@react-sheets/formula-engine';
import { formatCellAddress, isFormulaError } from '@react-sheets/formula-engine';
import type {
  FormulaAuditArrow,
  FormulaAuditDirection,
  FormulaAuditEvaluationProjection,
  FormulaAuditError,
  FormulaAuditFormulaProjection,
  FormulaAuditProjection,
} from '@react-sheets/spreadsheet-app';
import {
  Box,
  Button,
  Icon,
  Inline,
  Panel,
  PanelBody,
  PanelHeader,
  PanelTitle,
  Stack,
  StatePanel,
  Text,
} from '@react-sheets/ui-system';

export type FormulaAuditTaskState = 'idle' | 'loading' | 'ready' | 'error';

export interface FormulaAuditSectionStates {
  precedents?: FormulaAuditTaskState;
  dependents?: FormulaAuditTaskState;
  formulas?: FormulaAuditTaskState;
  errors?: FormulaAuditTaskState;
  evaluation?: FormulaAuditTaskState;
}

export interface FormulaAuditPanelCallbacks {
  onShowPrecedents?: () => void;
  onShowDependents?: () => void;
  onRemoveArrows?: () => void;
  onSetShowFormulas?: (enabled: boolean) => void;
  onScanErrors?: () => void;
  onEvaluateFormula?: () => void;
  onRetry?: () => void;
}

export interface FormulaAuditPanelProps {
  projection?: FormulaAuditProjection | null;
  /** Overall panel state. A missing projection is treated as empty. */
  state?: 'empty' | 'loading' | 'error' | 'ready';
  errorMessage?: string;
  sectionStates?: FormulaAuditSectionStates;
  callbacks?: FormulaAuditPanelCallbacks;
  activeCell?: string;
  locale?: 'en-US' | 'zh-CN';
}

interface FormulaAuditCopy {
  title: string;
  subtitle: string;
  selectedCell: string;
  noActiveCell: string;
  traceTitle: string;
  precedents: string;
  dependents: string;
  removeArrows: string;
  noPrecedents: string;
  noDependents: string;
  runPrecedents: string;
  runDependents: string;
  showFormulas: string;
  hideFormulas: string;
  formulasTitle: string;
  noFormulas: string;
  showFormulasHint: string;
  errorsTitle: string;
  scanErrors: string;
  noErrors: string;
  scanErrorsHint: string;
  evaluationTitle: string;
  evaluateFormula: string;
  noEvaluation: string;
  evaluateHint: string;
  step: string;
  value: string;
  formula: string;
  target: string;
  formulaCell: string;
  retry: string;
  loading: string;
  error: string;
  errorCount: string;
  moreItems: string;
  noFormulaSelected: string;
  selectFormulaHint: string;
}

const COPY: Record<'en-US' | 'zh-CN', FormulaAuditCopy> = {
  'en-US': {
    title: 'Formula Auditing',
    subtitle: 'Inspect dependencies and calculation steps for the active cell.',
    selectedCell: 'Active cell',
    noActiveCell: 'No active cell',
    traceTitle: 'Dependency arrows',
    precedents: 'Show precedents',
    dependents: 'Show dependents',
    removeArrows: 'Remove arrows',
    noPrecedents: 'No precedents are connected to this cell.',
    noDependents: 'No dependents are connected to this cell.',
    runPrecedents: 'Run Show Precedents to inspect referenced cells.',
    runDependents: 'Run Show Dependents to inspect formulas using this cell.',
    showFormulas: 'Show formulas',
    hideFormulas: 'Hide formulas',
    formulasTitle: 'Authored formulas',
    noFormulas: 'No authored formulas were found in this workbook.',
    showFormulasHint: 'Enable Show formulas to list authored formulas.',
    errorsTitle: 'Error checking',
    scanErrors: 'Scan errors',
    noErrors: 'No formula errors are reported for the current scan.',
    scanErrorsHint: 'Run Error checking to scan the active worksheet.',
    evaluationTitle: 'Evaluate formula',
    evaluateFormula: 'Evaluate formula',
    noEvaluation: 'No evaluation trace is available.',
    evaluateHint: 'Select a formula cell and run Evaluate formula to inspect each step.',
    step: 'Step',
    value: 'Value',
    formula: 'Formula',
    target: 'Target',
    formulaCell: 'Formula cell',
    retry: 'Try again',
    loading: 'Loading formula audit data.',
    error: 'Formula audit data could not be loaded.',
    errorCount: '{count} error(s)',
    moreItems: 'Showing {shown} of {total}.',
    noFormulaSelected: 'No formula selected',
    selectFormulaHint: 'Select a formula cell to trace its relationships.',
  },
  'zh-CN': {
    title: '公式审核',
    subtitle: '检查当前单元格的依赖关系和计算步骤。',
    selectedCell: '当前单元格',
    noActiveCell: '没有当前单元格',
    traceTitle: '依赖箭头',
    precedents: '显示追踪引用单元格',
    dependents: '显示追踪从属单元格',
    removeArrows: '删除箭头',
    noPrecedents: '当前单元格没有引用单元格。',
    noDependents: '没有公式使用当前单元格。',
    runPrecedents: '运行“显示追踪引用单元格”查看被引用单元格。',
    runDependents: '运行“显示追踪从属单元格”查看使用当前单元格的公式。',
    showFormulas: '显示公式',
    hideFormulas: '隐藏公式',
    formulasTitle: '工作簿公式',
    noFormulas: '工作簿中没有找到公式。',
    showFormulasHint: '启用“显示公式”查看工作簿公式。',
    errorsTitle: '错误检查',
    scanErrors: '扫描错误',
    noErrors: '当前扫描没有发现公式错误。',
    scanErrorsHint: '运行“错误检查”扫描当前工作表。',
    evaluationTitle: '公式求值',
    evaluateFormula: '求值公式',
    noEvaluation: '没有可用的求值轨迹。',
    evaluateHint: '选择公式单元格并运行“求值公式”查看每一步。',
    step: '步骤',
    value: '值',
    formula: '公式',
    target: '目标',
    formulaCell: '公式单元格',
    retry: '重试',
    loading: '正在加载公式审核数据。',
    error: '公式审核数据加载失败。',
    errorCount: '{count} 个错误',
    moreItems: '显示 {shown}/{total}。',
    noFormulaSelected: '未选择公式',
    selectFormulaHint: '选择公式单元格以追踪其依赖关系。',
  },
};

const MAX_VISIBLE_ITEMS = 200;

export function formatFormulaAuditValue(value: FormulaValue): string {
  if (isFormulaError(value)) return `${value.code} ${value.message}`;
  if (Array.isArray(value)) {
    return value.map((row) => row.map((entry) => formatFormulaAuditValue(entry)).join(', ')).join('; ');
  }
  return value == null ? '' : String(value);
}

export function formatFormulaAuditAddress(address: CellAddress): string {
  try {
    return formatCellAddress(address, true);
  } catch {
    return `${address.sheetId}!R${address.row + 1}C${address.column + 1}`;
  }
}

export function formatFormulaAuditDependency(dependency: FormulaDependency): string {
  if (dependency.kind === 'cell') return formatFormulaAuditAddress(dependency.address);
  if (dependency.kind === 'range') return `${formatFormulaAuditAddress(dependency.start)}:${formatFormulaAuditAddress(dependency.end)}`;
  if (dependency.kind === 'name') return dependency.name;
  return dependency.reference.type;
}

export function resolveFormulaAuditPanelState(
  projection: FormulaAuditProjection | null | undefined,
  state?: FormulaAuditPanelProps['state'],
): NonNullable<FormulaAuditPanelProps['state']> {
  if (state) return state;
  return projection ? 'ready' : 'empty';
}

export function formatFormulaAuditCount(template: string, count: number): string {
  return template.replace('{count}', String(count));
}

function formatMoreItems(template: string, shown: number, total: number): string {
  return template.replace('{shown}', String(shown)).replace('{total}', String(total));
}

function stateFor(
  states: FormulaAuditSectionStates | undefined,
  key: keyof FormulaAuditSectionStates,
  hasData: boolean,
): FormulaAuditTaskState {
  return states?.[key] ?? (hasData ? 'ready' : 'idle');
}

function ActionButton({
  children,
  disabled,
  icon,
  loading,
  onClick,
  pressed,
  title,
}: {
  children: string;
  disabled?: boolean;
  icon: ComponentProps<typeof Icon>['name'];
  loading?: boolean;
  onClick?: () => void;
  pressed?: boolean;
  title?: string;
}) {
  return (
    <Button
      aria-pressed={pressed}
      disabled={disabled || !onClick}
      icon={icon}
      loading={loading}
      onClick={onClick}
      size="sm"
      title={title}
      variant="outline"
    >
      {children}
    </Button>
  );
}

function TaskState({
  state,
  emptyDescription,
  emptyTitle,
  errorMessage,
  loadingDescription,
  onRetry,
  retryLabel,
}: {
  state: FormulaAuditTaskState;
  emptyDescription: string;
  emptyTitle: string;
  errorMessage: string;
  loadingDescription: string;
  onRetry?: () => void;
  retryLabel: string;
}) {
  if (state === 'loading') {
    return <StatePanel kind="loading" description={loadingDescription} />;
  }
  if (state === 'error') {
    return <StatePanel actionLabel={onRetry ? retryLabel : undefined} description={errorMessage} kind="error" onAction={onRetry} />;
  }
  if (state === 'idle') {
    return <StatePanel kind="empty" title={emptyTitle} description={emptyDescription} />;
  }
  return null;
}

function Header({
  action,
  icon,
  title,
}: {
  action?: ReactNode;
  icon: ComponentProps<typeof Icon>['name'];
  title: string;
}) {
  return (
    <PanelHeader className="px-3 py-2.5">
      <Inline gap="sm" className="min-w-0">
        <Icon name={icon} size="sm" className="text-blue-600" />
        <PanelTitle as="h3" size="sm">{title}</PanelTitle>
      </Inline>
      {action}
    </PanelHeader>
  );
}

function ArrowList({ arrows, copy }: { arrows: readonly FormulaAuditArrow[]; copy: FormulaAuditCopy }) {
  const visible = arrows.slice(0, MAX_VISIBLE_ITEMS);
  return (
    <Stack gap="xs">
      <Box as="ul" aria-label={copy.target} className="divide-y divide-slate-100 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {visible.map((arrow) => (
          <Box as="li" key={arrow.id} className="px-2.5 py-2">
            <Inline gap="sm" className="items-start">
              <Icon name={arrow.direction === 'precedent' ? 'arrow-left' : 'arrow-right'} size="sm" className="mt-0.5 text-blue-600" />
              <Stack gap="none" className="min-w-0">
                <Text size="xs" weight="semibold" className="truncate">
                  {arrow.direction === 'precedent' ? formatFormulaAuditDependency(arrow.target) : formatFormulaAuditAddress(arrow.formulaCell)}
                </Text>
                <Text size="xs" tone="subtle" className="truncate">
                  {arrow.direction === 'precedent'
                    ? `${copy.formulaCell}: ${formatFormulaAuditAddress(arrow.formulaCell)}`
                    : `${copy.target}: ${formatFormulaAuditDependency(arrow.target)}`}
                </Text>
              </Stack>
            </Inline>
          </Box>
        ))}
      </Box>
      {arrows.length > visible.length ? <Text size="xs" tone="subtle">{formatMoreItems(copy.moreItems, visible.length, arrows.length)}</Text> : null}
    </Stack>
  );
}

function DependencySection({
  arrows,
  copy,
  direction,
  onRetry,
  state,
}: {
  arrows: readonly FormulaAuditArrow[];
  copy: FormulaAuditCopy;
  direction: FormulaAuditDirection;
  onRetry?: () => void;
  state: FormulaAuditTaskState;
}) {
  const relevant = arrows.filter((arrow) => arrow.direction === direction);
  const isPrecedent = direction === 'precedent';
  const actionTitle = isPrecedent ? copy.precedents : copy.dependents;
  const visibleState = state === 'ready' && relevant.length === 0 ? 'idle' : state;
  return (
    <Stack gap="sm">
      <TaskState
        state={visibleState}
        emptyDescription={isPrecedent ? copy.runPrecedents : copy.runDependents}
        emptyTitle={isPrecedent ? copy.noPrecedents : copy.noDependents}
        errorMessage={copy.error}
        loadingDescription={copy.loading}
        onRetry={onRetry}
        retryLabel={copy.retry}
      />
      {visibleState === 'ready' && relevant.length > 0 ? <ArrowList arrows={relevant} copy={copy} /> : null}
      <Text size="xs" tone="subtle">{actionTitle}: {relevant.length}</Text>
    </Stack>
  );
}

function FormulaList({ formulas, copy }: { formulas: readonly FormulaAuditFormulaProjection[]; copy: FormulaAuditCopy }) {
  const visible = formulas.slice(0, MAX_VISIBLE_ITEMS);
  return (
    <Stack gap="xs">
      <Box as="ul" aria-label={copy.formulasTitle} className="divide-y divide-slate-100 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {visible.map((entry) => (
          <Box as="li" key={formatFormulaAuditAddress(entry.address)} className="px-2.5 py-2">
            <Inline gap="sm" className="items-start">
              <Icon name="function" size="sm" className="mt-0.5 text-emerald-600" />
              <Stack gap="none" className="min-w-0">
                <Text size="xs" weight="semibold">{formatFormulaAuditAddress(entry.address)}</Text>
                <Text size="xs" className="break-all font-mono text-slate-600">{entry.formula}</Text>
                <Text size="xs" tone={isFormulaError(entry.value) ? 'danger' : 'subtle'} className="truncate">
                  {copy.value}: {formatFormulaAuditValue(entry.value)}
                </Text>
              </Stack>
            </Inline>
          </Box>
        ))}
      </Box>
      {formulas.length > visible.length ? <Text size="xs" tone="subtle">{formatMoreItems(copy.moreItems, visible.length, formulas.length)}</Text> : null}
    </Stack>
  );
}

function ErrorList({ errors, copy }: { errors: readonly FormulaAuditError[]; copy: FormulaAuditCopy }) {
  const visible = errors.slice(0, MAX_VISIBLE_ITEMS);
  return (
    <Stack gap="xs">
      <Text size="xs" tone="danger" weight="semibold">{formatFormulaAuditCount(copy.errorCount, errors.length)}</Text>
      <Box as="ul" aria-label={copy.errorsTitle} className="divide-y divide-rose-100 overflow-hidden rounded-lg border border-rose-200 bg-rose-50/40">
        {visible.map((entry) => (
          <Box as="li" key={`${formatFormulaAuditAddress(entry.address)}:${entry.position ?? ''}`} className="px-2.5 py-2">
            <Inline gap="sm" className="items-start">
              <Icon name="alert-circle" size="sm" className="mt-0.5 text-rose-600" />
              <Stack gap="none" className="min-w-0">
                <Inline gap="xs" className="flex-wrap">
                  <Text size="xs" weight="semibold" tone="danger">{entry.code}</Text>
                  <Text size="xs" tone="subtle">{formatFormulaAuditAddress(entry.address)}</Text>
                </Inline>
                <Text size="xs" tone="danger" className="break-words">{entry.message}</Text>
                <Text size="xs" tone="subtle" className="break-all font-mono">{entry.formula}</Text>
              </Stack>
            </Inline>
          </Box>
        ))}
      </Box>
      {errors.length > visible.length ? <Text size="xs" tone="subtle">{formatMoreItems(copy.moreItems, visible.length, errors.length)}</Text> : null}
    </Stack>
  );
}

function EvaluationTrace({ evaluation, copy }: { evaluation: FormulaAuditEvaluationProjection; copy: FormulaAuditCopy }) {
  const visible = evaluation.steps.slice(0, MAX_VISIBLE_ITEMS);
  return (
    <Stack gap="sm">
      <Stack gap="none">
        <Text size="xs" weight="semibold">{copy.formula}: {evaluation.formula}</Text>
        <Text size="xs" tone={isFormulaError(evaluation.value) ? 'danger' : 'accent'}>
          {copy.value}: {formatFormulaAuditValue(evaluation.value)}
        </Text>
      </Stack>
      <Box as="ul" aria-label={copy.evaluationTitle} className="divide-y divide-slate-100 overflow-hidden rounded-lg border border-slate-200 bg-white">
        {visible.map((step) => (
          <Box as="li" key={step.index} className="px-2.5 py-2">
            <Inline gap="sm" className="items-start">
              <Text size="xs" tone="subtle" className="w-10 shrink-0 font-mono">{copy.step} {step.index + 1}</Text>
              <Stack gap="none" className="min-w-0">
                <Text size="xs" className="break-words font-mono">{step.expression}</Text>
                <Text size="xs" tone={isFormulaError(step.value) ? 'danger' : 'subtle'} className="break-words">
                  {copy.value}: {formatFormulaAuditValue(step.value)}
                </Text>
              </Stack>
            </Inline>
          </Box>
        ))}
      </Box>
      {evaluation.steps.length > visible.length ? <Text size="xs" tone="subtle">{formatMoreItems(copy.moreItems, visible.length, evaluation.steps.length)}</Text> : null}
    </Stack>
  );
}

export function FormulaAuditPanel({
  activeCell,
  callbacks,
  errorMessage,
  locale = 'en-US',
  projection,
  sectionStates,
  state,
}: FormulaAuditPanelProps) {
  const copy = COPY[locale];
  const panelState = resolveFormulaAuditPanelState(projection, state);

  if (panelState === 'loading') {
    return <StatePanel kind="loading" title={copy.title} description={copy.loading} />;
  }
  if (panelState === 'error') {
    return <StatePanel actionLabel={callbacks?.onRetry ? copy.retry : undefined} description={errorMessage ?? copy.error} kind="error" onAction={callbacks?.onRetry} />;
  }
  if (panelState === 'empty' || !projection) {
    return <StatePanel kind="empty" title={copy.title} description={copy.selectFormulaHint} />;
  }

  const selectedAddress = projection.selectedCell ? formatFormulaAuditAddress(projection.selectedCell) : activeCell ?? copy.noActiveCell;
  const precedentsState = stateFor(sectionStates, 'precedents', projection.arrows.some((arrow) => arrow.direction === 'precedent'));
  const dependentsState = stateFor(sectionStates, 'dependents', projection.arrows.some((arrow) => arrow.direction === 'dependent'));
  const formulasState = stateFor(sectionStates, 'formulas', projection.showFormulas && projection.formulas.length > 0);
  const errorsState = stateFor(sectionStates, 'errors', projection.errors.length > 0);
  const evaluationState = stateFor(sectionStates, 'evaluation', projection.evaluation !== undefined);
  const visibleFormulasState = formulasState === 'ready' && projection.formulas.length === 0 ? 'idle' : formulasState;
  const visibleErrorsState = errorsState === 'ready' && projection.errors.length === 0 ? 'idle' : errorsState;
  const visibleEvaluationState = evaluationState === 'ready' && projection.evaluation === undefined ? 'idle' : evaluationState;
  const setShowFormulas = callbacks?.onSetShowFormulas;

  return (
    <Stack gap="md">
      <Panel tone="accent" className="overflow-hidden shadow-none">
        <PanelBody className="p-3">
          <Inline gap="sm" className="items-start">
            <Box className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white text-blue-600 shadow-xs">
              <Icon name="function" size="md" />
            </Box>
            <Stack gap="none" className="min-w-0">
              <PanelTitle as="h3" size="sm">{copy.title}</PanelTitle>
              <Text size="xs" tone="muted" className="mt-0.5">{copy.subtitle}</Text>
              <Text size="xs" tone="accent" weight="semibold" className="mt-2 truncate">{copy.selectedCell}: {selectedAddress}</Text>
            </Stack>
          </Inline>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <Header
          action={<ActionButton disabled={projection.arrows.length === 0} icon="x" onClick={callbacks?.onRemoveArrows}>{copy.removeArrows}</ActionButton>}
          icon="share"
          title={copy.traceTitle}
        />
        <PanelBody className="p-3">
          <Stack gap="md">
            <Inline gap="xs" className="flex-wrap">
              <ActionButton
                icon="arrow-left"
                loading={precedentsState === 'loading'}
                onClick={callbacks?.onShowPrecedents}
              >
                {copy.precedents}
              </ActionButton>
              <ActionButton
                icon="arrow-right"
                loading={dependentsState === 'loading'}
                onClick={callbacks?.onShowDependents}
              >
                {copy.dependents}
              </ActionButton>
            </Inline>
            <DependencySection arrows={projection.arrows} copy={copy} direction="precedent" onRetry={callbacks?.onShowPrecedents} state={precedentsState} />
            <DependencySection arrows={projection.arrows} copy={copy} direction="dependent" onRetry={callbacks?.onShowDependents} state={dependentsState} />
          </Stack>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <Header
          action={(
            <ActionButton
              icon="eye"
              onClick={setShowFormulas ? () => setShowFormulas(!projection.showFormulas) : undefined}
              pressed={projection.showFormulas}
            >
              {projection.showFormulas ? copy.hideFormulas : copy.showFormulas}
            </ActionButton>
          )}
          icon="function"
          title={copy.formulasTitle}
        />
        <PanelBody className="p-3">
          <TaskState
            state={visibleFormulasState}
            emptyDescription={copy.showFormulasHint}
            emptyTitle={copy.noFormulas}
            errorMessage={copy.error}
            loadingDescription={copy.loading}
            onRetry={setShowFormulas ? () => setShowFormulas(true) : undefined}
            retryLabel={copy.retry}
          />
          {visibleFormulasState === 'ready' ? <FormulaList copy={copy} formulas={projection.formulas} /> : null}
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <Header
          action={<ActionButton icon="check-circle" loading={visibleErrorsState === 'loading'} onClick={callbacks?.onScanErrors}>{copy.scanErrors}</ActionButton>}
          icon="alert-circle"
          title={copy.errorsTitle}
        />
        <PanelBody className="p-3">
          <TaskState
            state={visibleErrorsState}
            emptyDescription={copy.scanErrorsHint}
            emptyTitle={copy.noErrors}
            errorMessage={copy.error}
            loadingDescription={copy.loading}
            onRetry={callbacks?.onScanErrors}
            retryLabel={copy.retry}
          />
          {visibleErrorsState === 'ready' ? <ErrorList copy={copy} errors={projection.errors} /> : null}
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <Header
          action={<ActionButton icon="calculator" loading={visibleEvaluationState === 'loading'} onClick={callbacks?.onEvaluateFormula}>{copy.evaluateFormula}</ActionButton>}
          icon="calculator"
          title={copy.evaluationTitle}
        />
        <PanelBody className="p-3">
          <TaskState
            state={visibleEvaluationState}
            emptyDescription={copy.evaluateHint}
            emptyTitle={copy.noEvaluation}
            errorMessage={copy.error}
            loadingDescription={copy.loading}
            onRetry={callbacks?.onEvaluateFormula}
            retryLabel={copy.retry}
          />
          {visibleEvaluationState === 'ready' && projection.evaluation ? <EvaluationTrace copy={copy} evaluation={projection.evaluation} /> : null}
          {!projection.selectedCell ? <Text size="xs" tone="subtle" className="mt-2">{copy.noFormulaSelected}</Text> : null}
        </PanelBody>
      </Panel>

      <Text size="xs" tone="subtle" className="px-1">{copy.selectedCell}: {selectedAddress}</Text>
    </Stack>
  );
}
