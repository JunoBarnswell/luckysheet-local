import type { WorkbookTemplateId } from '@react-sheets/spreadsheet-app';
export type { WorkbookTemplateId } from '@react-sheets/spreadsheet-app';
export interface WorkbookCreateOptions {
  readonly name: string;
  readonly template?: WorkbookTemplateId;
  readonly spaceId?: string;
  readonly folderId?: string;
}
