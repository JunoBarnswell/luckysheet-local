import type { CellData, WorkbookSnapshotV1 } from '@react-sheets/core-model';

export interface XlsxExportOptions {
  workbook: WorkbookSnapshotV1;
}

export function exportSnapshotToXlsxXml(snapshot: WorkbookSnapshotV1): Record<string, string> {
  const files: Record<string, string> = {};

  // 1. [Content_Types].xml
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStringItem+xml"/>
  ${snapshot.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n  ')}
</Types>`;

  // 2. _rels/.rels
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  // 3. xl/_rels/workbook.xml.rels
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  ${snapshot.sheets.map((_, i) => `<Relationship Id="rIdSheet${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n  ')}
</Relationships>`;

  // 4. xl/workbook.xml
  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${snapshot.sheets.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rIdSheet${i + 1}"/>`).join('\n    ')}
  </sheets>
</workbook>`;

  // 5. xl/styles.xml
  files['xl/styles.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;

  // 6. Shared strings & sheets
  const sharedStrings: string[] = [];
  const stringIndexMap = new Map<string, number>();

  function getSharedStringIndex(str: string): number {
    let idx = stringIndexMap.get(str);
    if (idx === undefined) {
      idx = sharedStrings.length;
      sharedStrings.push(str);
      stringIndexMap.set(str, idx);
    }
    return idx;
  }

  snapshot.sheets.forEach((sheet, idx) => {
    let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>`;

    const rowKeys = Object.keys(sheet.cells).map(Number).sort((a, b) => a - b);
    for (const r of rowKeys) {
      const colObj = sheet.cells[String(r)] ?? {};
      const colKeys = Object.keys(colObj).map(Number).sort((a, b) => a - b);
      if (colKeys.length === 0) continue;

      sheetXml += `\n    <row r="${r + 1}">`;
      for (const c of colKeys) {
        const cell = colObj[String(c)];
        if (!cell) continue;
        const cellRef = `${columnToLetter(c)}${r + 1}`;

        if (cell.formula) {
          const formulaClean = cell.formula.startsWith('=') ? cell.formula.slice(1) : cell.formula;
          sheetXml += `<c r="${cellRef}"><f>${escapeXml(formulaClean)}</f></c>`;
        } else if (typeof cell.value === 'number') {
          sheetXml += `<c r="${cellRef}"><v>${cell.value}</v></c>`;
        } else if (typeof cell.value === 'boolean') {
          sheetXml += `<c r="${cellRef}" t="b"><v>${cell.value ? 1 : 0}</v></c>`;
        } else if (cell.value !== null && cell.value !== undefined) {
          const strIdx = getSharedStringIndex(String(cell.value));
          sheetXml += `<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`;
        }
      }
      sheetXml += `</row>`;
    }

    sheetXml += `\n  </sheetData>`;

    // Merges
    if (sheet.merges.length > 0) {
      sheetXml += `\n  <mergeCells count="${sheet.merges.length}">`;
      for (const m of sheet.merges) {
        const ref = `${columnToLetter(m.range.startColumn)}${m.range.startRow + 1}:${columnToLetter(m.range.endColumn)}${m.range.endRow + 1}`;
        sheetXml += `\n    <mergeCell ref="${ref}"/>`;
      }
      sheetXml += `\n  </mergeCells>`;
    }

    sheetXml += `\n</worksheet>`;
    files[`xl/worksheets/sheet${idx + 1}.xml`] = sheetXml;
  });

  // 7. xl/sharedStrings.xml
  files['xl/sharedStrings.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
  ${sharedStrings.map((s) => `<si><t>${escapeXml(s)}</t></si>`).join('\n  ')}
</sst>`;

  return files;
}

export function parseXlsxXmlToSnapshot(files: Record<string, string>): WorkbookSnapshotV1 {
  const workbookXml = files['xl/workbook.xml'] ?? '';
  const sharedStringsXml = files['xl/sharedStrings.xml'] ?? '';

  // Extract shared strings
  const sharedStrings: string[] = [];
  const sstMatches = sharedStringsXml.matchAll(/<t[^>]*>(.*?)<\/t>/gs);
  for (const match of sstMatches) {
    sharedStrings.push(unescapeXml(match[1] ?? ''));
  }

  // Extract sheets list
  const sheets: WorkbookSnapshotV1['sheets'] = [];
  const sheetTagMatches = workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*>/g);

  let sheetIdx = 1;
  for (const match of sheetTagMatches) {
    const name = unescapeXml(match[1] ?? `Sheet${sheetIdx}`);
    const sheetFile = files[`xl/worksheets/sheet${sheetIdx}.xml`];
    const cells: Record<string, Record<string, CellData>> = {};

    if (sheetFile) {
      const rowMatches = sheetFile.matchAll(/<row[^>]*r="(\d+)"[^>]*>(.*?)<\/row>/gs);
      for (const rMatch of rowMatches) {
        const rowNum = parseInt(rMatch[1] ?? '1', 10) - 1;
        const rowContent = rMatch[2] ?? '';

        const cellMatches = rowContent.matchAll(/<c[^>]*r="([A-Z]+)(\d+)"(?:[^>]*t="([^"]+)")?[^>]*>(?:<f>(.*?)<\/f>)?(?:<v>(.*?)<\/v>)?<\/c>/gs);
        for (const cMatch of cellMatches) {
          const colLetter = cMatch[1] ?? 'A';
          const colIdx = letterToColumn(colLetter);
          const type = cMatch[3];
          const formula = cMatch[4];
          const rawVal = cMatch[5];

          let val: string | number | boolean | null = null;
          if (type === 's' && rawVal !== undefined) {
            val = sharedStrings[parseInt(rawVal, 10)] ?? '';
          } else if (type === 'b' && rawVal !== undefined) {
            val = rawVal === '1';
          } else if (rawVal !== undefined && rawVal !== '') {
            const num = Number(rawVal);
            val = Number.isNaN(num) ? rawVal : num;
          }

          cells[rowNum] ??= {};
          cells[rowNum][colIdx] = {
            value: val,
            formula: formula ? `=${unescapeXml(formula)}` : undefined,
          };
        }
      }
    }

    sheets.push({
      id: `sheet-${sheetIdx}`,
      name,
      rowCount: 1000,
      columnCount: 26,
      cells,
      merges: [],
      freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
      charts: [],
      pivots: [],
      shapes: [],
      sparklines: [],
    });

    sheetIdx++;
  }

  if (sheets.length === 0) {
    sheets.push({
      id: 'sheet-1',
      name: 'Sheet1',
      rowCount: 1000,
      columnCount: 26,
      cells: {},
      merges: [],
      freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
      charts: [],
      pivots: [],
      shapes: [],
      sparklines: [],
    });
  }

  return {
    schema: 'WorkbookSnapshotV1',
    unitId: 'imported-' + Math.random().toString(36).substring(2, 9),
    name: 'Imported Workbook',
    activeSheetId: sheets[0]!.id,
    sheets,
  };
}

function columnToLetter(colIndex: number): string {
  let temp = colIndex + 1;
  let letter = '';
  while (temp > 0) {
    const rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

function letterToColumn(letter: string): number {
  let column = 0;
  for (const char of letter.toUpperCase()) {
    column = column * 26 + char.charCodeAt(0) - 64;
  }
  return Math.max(0, column - 1);
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function unescapeXml(str: string): string {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
