import { Box, Button, Icon, Inline, Stack, TemplatePreview, Text } from '@react-sheets/ui-system';
import type { WorkbookTemplateDefinition, WorkbookTemplateKind } from './types';
import { workbookTemplates } from './types';

export interface CreateTemplateGridProps {
  templates?: readonly WorkbookTemplateDefinition[];
  onSelect: (kind: WorkbookTemplateKind) => void;
  onMoreTemplates?: () => void;
}

export function CreateTemplateGrid({ templates = workbookTemplates, onSelect, onMoreTemplates }: CreateTemplateGridProps) {
  return (
    <Stack gap="md" className="w-full">
      <Inline gap="sm" className="justify-between">
        <Inline gap="sm">
          <Icon name="chevron-down" size="sm" className="text-slate-900" />
          <Text className="text-[17px] font-semibold text-slate-900" weight="semibold">新建</Text>
        </Inline>
        {onMoreTemplates ? <Button icon="arrow-right" onClick={onMoreTemplates} size="sm" variant="ghost" className="text-brand-dark hover:bg-brand-soft">更多模板</Button> : null}
      </Inline>
      <Box className="grid grid-cols-1 gap-5 min-[620px]:grid-cols-2 min-[930px]:grid-cols-3 min-[1320px]:grid-cols-6 min-[1440px]:px-[18px]">
        {templates.map((template, index) => (
          <Button
            key={template.kind}
            aria-label={template.title}
            className={[
              'group relative h-[182px] w-full flex-col items-stretch justify-between overflow-hidden rounded-lg border border-slate-200 bg-white px-3 pb-3 pt-2 text-center shadow-hub-card transition-all hover:-translate-y-0.5 hover:border-brand hover:shadow-brand-sm focus-visible:ring-brand/40',
              index === 0 ? 'border-2 border-brand' : '',
            ].join(' ')}
            onClick={() => onSelect(template.kind)}
            size="md"
            variant="ghost"
          >
            <TemplatePreview kind={template.kind} className="h-[127px] w-full" />
            <Stack gap="none" className="items-center">
              <Text className="text-[14px] font-medium text-slate-800" weight="medium">{template.title}</Text>
              <Text className="mt-0.5 text-[11px] text-slate-400" size="xs">{template.description}</Text>
            </Stack>
            {index === 0 ? <Box className="absolute bottom-9 right-4 flex h-8 w-8 items-center justify-center rounded-full bg-brand text-white shadow-brand-sm"><Icon name="plus" size="md" /></Box> : null}
          </Button>
        ))}
      </Box>
    </Stack>
  );
}
