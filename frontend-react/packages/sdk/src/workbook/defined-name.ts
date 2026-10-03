import type { DefinedNameModel } from '@react-sheets/core-model';
import type { Workbook } from './workbook';
import { domainFor } from './object-domain';
import { immutableSnapshot } from './value';

export class DefinedName {
  readonly #workbook: Workbook;
  readonly #identity: Pick<DefinedNameModel, 'name' | 'scope' | 'sheetId'>;
  constructor(workbook: Workbook, identity: Pick<DefinedNameModel, 'name' | 'scope' | 'sheetId'>) {
    this.#workbook = workbook; this.#identity = structuredClone(identity); Object.freeze(this);
  }
  snapshot(): Readonly<DefinedNameModel> {
    const model = domainFor(this.#workbook).names().find(item => item.name.toUpperCase() === this.#identity.name.toUpperCase()
      && item.scope === this.#identity.scope && item.sheetId === this.#identity.sheetId);
    if (!model) return domainFor(this.#workbook).invalid('definedName.read', new Error(`Defined name no longer exists: ${this.#identity.name}`));
    return immutableSnapshot(model);
  }
  async setFormula(formula: string): Promise<void> { await this.#workbook.names.define({ ...this.snapshot(), formula }); }
  async remove(): Promise<void> { const model = this.snapshot(); await this.#workbook.names.remove(model.name, model.scope, model.sheetId); }
}

export class DefinedNameCollection {
  readonly #workbook: Workbook;
  readonly #objects = new Map<string, DefinedName>();
  constructor(workbook: Workbook) { this.#workbook = workbook; Object.freeze(this); }
  list(): readonly DefinedName[] { return Object.freeze(domainFor(this.#workbook).names().map(model => this.byName(model.name, model.scope, model.sheetId))); }
  byName(name: string, scope: DefinedNameModel['scope'], sheetId?: string): DefinedName {
    const models = domainFor(this.#workbook).names();
    if (typeof name !== 'string' || (scope !== 'workbook' && scope !== 'sheet')
      || scope === 'workbook' && sheetId !== undefined || scope === 'sheet' && typeof sheetId !== 'string') {
      return domainFor(this.#workbook).invalid('definedNames.byName', new Error('An explicit name and canonical scope are required.'));
    }
    const model = models.find(item => item.name.toUpperCase() === name.toUpperCase() && item.scope === scope && item.sheetId === sheetId);
    if (!model) return domainFor(this.#workbook).invalid('definedNames.byName', new Error(`Defined name does not exist: ${name}`));
    const key = JSON.stringify([scope, sheetId ?? null, model.name.toUpperCase()]);
    let handle = this.#objects.get(key);
    if (!handle) { handle = new DefinedName(this.#workbook, { name: model.name, scope, ...(sheetId === undefined ? {} : { sheetId }) }); this.#objects.set(key, handle); }
    return handle;
  }
  async define(model: DefinedNameModel): Promise<DefinedName> {
    const intent = structuredClone(model);
    await domainFor(this.#workbook).command('definedNames.define', { commandId: 'workbook.name.set', params: intent });
    return this.byName(intent.name.trim(), intent.scope, intent.sheetId);
  }
  async remove(name: string, scope: DefinedNameModel['scope'], sheetId?: string): Promise<void> {
    const model = this.byName(name, scope, sheetId).snapshot();
    await domainFor(this.#workbook).command('definedNames.remove', { commandId: 'workbook.name.remove', params: { name: model.name, scope, ...(sheetId === undefined ? {} : { sheetId }) } });
  }
}
