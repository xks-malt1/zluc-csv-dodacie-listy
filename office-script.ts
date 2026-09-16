function main(workbook: ExcelScript.Workbook) {
  const prefix = "zasilky-balik";

  // Pôvodné stĺpce G, H, L, M, N — zhodné s $Odstranit v Zluc-Csv.ps1.
  const removeHeaders = [
    "Vytvorené",
    "Podané",
    "Note",
    "Hmotnosť",
    "Užív. hmotnosť"
  ];

  // Koľko riadkov sa prenáša naraz — Office Scripts má limit
  // veľkosti jednej požiadavky.
  const batchSize = 5000;

  const removeNorm = removeHeaders.map(normalize);
  const isRemoved = (header: string) =>
    removeNorm.indexOf(normalize(header)) >= 0;

  type SheetInfo = {
    sheet: ExcelScript.Worksheet;
    table: ExcelScript.Table;
    name: string;
    rowCount: number;
    headers: string[];
  };

  const sheetInfo: SheetInfo[] = [];

  // Nájde všetky zásielkové hárky.
  for (const sheet of workbook.getWorksheets()) {
    const name = sheet.getName();

    if (!name.toLowerCase().startsWith(prefix)) {
      continue;
    }

    const tables = sheet.getTables();

    if (tables.length !== 1) {
      throw new Error(
        `Hárok "${name}" musí obsahovať presne jednu tabuľku.`
      );
    }

    sheetInfo.push({
      sheet,
      table: tables[0],
      name,
      rowCount: tables[0].getRowCount(),
      headers: tables[0].getHeaderRowRange().getTexts()[0]
    });
  }

  if (sheetInfo.length === 0) {
    throw new Error(
      `Nenašiel sa žiadny hárok začínajúci názvom "${prefix}".`
    );
  }

  // ---------------------------------------------------------------
  // 1. Kontroly — ešte pred akoukoľvek zmenou zošita
  // ---------------------------------------------------------------

  // Stĺpce na odstránenie, ktoré nie sú v žiadnom hárku.
  const allHeaders: string[] = [];
  for (const item of sheetInfo) {
    for (const h of item.headers) {
      allHeaders.push(normalize(h));
    }
  }

  const notFound = removeHeaders.filter(
    h => allHeaders.indexOf(normalize(h)) < 0
  );

  if (notFound.length > 0) {
    console.log(
      `Upozornenie: tieto stĺpce sa nenašli v žiadnom hárku ` +
      `(možno už boli odstránené): ${notFound.join(", ")}`
    );
  }

  // Hárok s najväčším počtom zásielok.
  let targetIndex = 0;

  for (let i = 1; i < sheetInfo.length; i++) {
    if (sheetInfo[i].rowCount > sheetInfo[targetIndex].rowCount) {
      targetIndex = i;
    }
  }

  const target = sheetInfo[targetIndex];
  const sources = sheetInfo
    .filter(s => s !== target)
    .sort((a, b) => b.rowCount - a.rowCount);

  // Hlavičky po odstránení stĺpcov musia sedieť aj v poradí,
  // lebo riadky sa pripájajú podľa pozície.
  const keep = (headers: string[]) => headers.filter(h => !isRemoved(h));
  const targetKept = keep(target.headers);
  const problems: string[] = [];

  for (const source of sources) {
    const sourceKept = keep(source.headers);

    if (!arraysEqual(targetKept.map(normalize), sourceKept.map(normalize))) {
      problems.push(`"${source.name}": ${sourceKept.join(", ")}`);
    }
  }

  if (problems.length > 0) {
    throw new Error(
      `Hlavičky sa nezhodujú s cieľovým hárkom "${target.name}" ` +
      `(${targetKept.join(", ")}). Nezhodné hárky: ` +
      `${problems.join(" | ")}. Zošit nebol zmenený.`
    );
  }

  // ---------------------------------------------------------------
  // 2. Odstránenie stĺpcov
  // ---------------------------------------------------------------

  for (const item of sheetInfo) {
    for (const header of item.headers) {
      if (isRemoved(header)) {
        const column = item.table.getColumnByName(header);
        if (column) {
          column.delete();
        }
      }
    }
  }

  // Ak je iba jeden hárok, stačí odstránenie stĺpcov.
  if (sources.length === 0) {
    target.sheet.activate();

    console.log(
      `Hotovo: v hárku "${target.name}" boli odstránené stĺpce ` +
      `${removeHeaders.join(", ")}. Zlúčenie nebolo potrebné.`
    );

    return;
  }

  // ---------------------------------------------------------------
  // 3. Zlúčenie od najväčšieho po najmenší, po dávkach
  // ---------------------------------------------------------------

  const expectedRows = sheetInfo.reduce((sum, s) => sum + s.rowCount, 0);
  let appendedRows = 0;

  for (const source of sources) {
    const rows = source.rowCount;

    if (rows === 0) {
      continue;
    }

    const body = source.table.getRangeBetweenHeaderAndTotal();
    const cols = body.getColumnCount();

    for (let start = 0; start < rows; start += batchSize) {
      const count = Math.min(batchSize, rows - start);
      const values = body
        .getCell(start, 0)
        .getResizedRange(count - 1, cols - 1)
        .getValues();

      target.table.addRows(-1, values);
      appendedRows += count;
    }
  }

  // ---------------------------------------------------------------
  // 4. Overenie a vymazanie menších hárkov
  // ---------------------------------------------------------------

  const actualRows = target.table.getRowCount();

  if (actualRows !== expectedRows) {
    throw new Error(
      `Cieľová tabuľka má ${actualRows} riadkov namiesto ` +
      `${expectedRows}. Menšie hárky neboli vymazané.`
    );
  }

  for (const source of sources) {
    source.sheet.delete();
  }

  target.sheet.activate();

  console.log(
    `Hotovo: ${appendedRows} zásielok bolo pridaných do ` +
    `"${target.name}" (spolu ${actualRows}). Odstránené hárky: ` +
    sources.map(s => s.name).join(", ")
  );
}

// Zjednotí diakritiku, veľkosť písmen, medzery (aj nezlomiteľné) a BOM.
function normalize(text: string): string {
  return text
    .replace(/^\uFEFF/, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function arraysEqual(
  first: string[],
  second: string[]
): boolean {
  if (first.length !== second.length) {
    return false;
  }

  for (let i = 0; i < first.length; i++) {
    if (first[i] !== second[i]) {
      return false;
    }
  }

  return true;
}
