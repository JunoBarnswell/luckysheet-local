import type { Workbook } from './workbook';
import type { ExternalLinkSnapshot } from './contract';
import { domainFor } from './object-domain';

export class WorkbookExternalLinks {
  readonly #workbook: Workbook;
  constructor(workbook: Workbook) { this.#workbook = workbook; Object.freeze(this); }
  bind(source: Workbook, token: string): Promise<void> { return domainFor(this.#workbook).bind(source, token); }
  refresh(): Promise<readonly ExternalLinkSnapshot[]> { return domainFor(this.#workbook).refresh(); }
}
