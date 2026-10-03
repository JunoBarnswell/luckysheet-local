import { protectionRuleValidationError, type ProtectionRule } from '@react-sheets/core-model';
import { SdkError } from '../error';
import type { Workbook } from './workbook';
import type { Worksheet } from './worksheet';
import { domainFor } from './object-domain';
import { immutableSnapshot } from './value';

export class WorksheetProtection {
  readonly #workbook: Workbook;
  readonly #sheet: Worksheet;
  constructor(workbook: Workbook, sheet: Worksheet) { this.#workbook = workbook; this.#sheet = sheet; Object.freeze(this); }
  list(): readonly Readonly<ProtectionRule>[] { return immutableSnapshot(domainFor(this.#workbook).sheet(this.#sheet.id).protectionRules); }
  async set(rule: ProtectionRule): Promise<void> {
    this.#sheet.name;
    const error = protectionRuleValidationError(rule);
    if (error !== undefined) return domainFor(this.#workbook).invalid('worksheet.protection.set', new Error(`PROTECTION_RULE_INVALID: ${error}`), { sheetId: this.#sheet.id });
    if (rule.scope === 'workbook') throw new SdkError('UNSUPPORTED_FEATURE', 'worksheet.protection.set', 'Workbook protection requires its workbook owner.', 'Use worksheet or range protection.', { object: { workbookId: this.#workbook.id, sheetId: this.#sheet.id } });
    await domainFor(this.#workbook).command('worksheet.protection.set', { commandId: 'sheet.protect.set', params: { sheetId: this.#sheet.id, rule: structuredClone(rule) } });
  }
  async remove(ruleId: string): Promise<void> {
    if (!this.list().some(rule => rule.id === ruleId)) return domainFor(this.#workbook).invalid('worksheet.protection.remove', new Error(`Protection rule does not exist: ${ruleId}`), { sheetId: this.#sheet.id });
    await domainFor(this.#workbook).command('worksheet.protection.remove', { commandId: 'sheet.protect.remove', params: { sheetId: this.#sheet.id, ruleId } });
  }
}
