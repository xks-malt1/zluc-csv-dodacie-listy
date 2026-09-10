function main(workbook: ExcelScript.Workbook) {
  const prefix = "zasilky-balik";

  // Pôvodné stĺpce G, H, L, M, N.
  const removeHeaders = [
    "Vytvorené",
    "Podané",
    "Note",
    "Hmotnosť",
    "Užív. hmotnosť"
  ];

  type SheetInfo = {
    sheet: ExcelScript.Worksheet;
    table: ExcelScript.Table;
    name: string;
    rowCount: number;
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
      rowCount: tables[0].getRowCount()
    });
  }

  if (sheetInfo.length === 0) {
    throw new Error(
      'Nenašiel sa žiadny hárok začínajúci názvom "zasilky-balik".'
    );
  }

  // Odstráni požadované stĺpce podľa názvov hlavičiek.
  // Ak už boli odstránené, skript ich jednoducho preskočí.
  for (const item of sheetInfo) {
    for (const headerToRemove of removeHeaders) {
      const columns = item.table.getColumns();
      let columnToDelete: ExcelScript.TableColumn | undefined;

      for (const column of columns) {
        if (column.getName() === headerToRemove) {
          columnToDelete = column;
          break;
        }
      }

      if (columnToDelete) {
        columnToDelete.delete();
      }
    }
  }

  // Ak je iba jeden hárok, stačí odstránenie stĺpcov.
  if (sheetInfo.length === 1) {
    sheetInfo[0].sheet.activate();

    console.log(
      `Hotovo: v hárku "${sheetInfo[0].name}" boli odstránené ` +
      `stĺpce Vytvorené, Podané, Note, Hmotnosť a Užív. hmotnosť. ` +
      `Zlúčenie nebolo potrebné.`
    );

    return;
  }

  // Vyberie hárok s najväčším počtom zásielok.
  let targetIndex = 0;

  for (let i = 1; i < sheetInfo.length; i++) {
    if (sheetInfo[i].rowCount > sheetInfo[targetIndex].rowCount) {
      targetIndex = i;
    }
  }

  const target = sheetInfo[targetIndex];

  // Overí, že po odstránení stĺpcov majú tabuľky rovnaké hlavičky.
  const targetHeaders =
    target.table.getHeaderRowRange().getTexts()[0];

  for (let i = 0; i < sheetInfo.length; i++) {
    if (i === targetIndex) {
      continue;
    }

    const sourceHeaders =
      sheetInfo[i].table.getHeaderRowRange().getTexts()[0];

    if (!arraysEqual(targetHeaders, sourceHeaders)) {
      throw new Error(
        `Hlavičky v hárku "${sheetInfo[i].name}" ` +
        `sa nezhodujú s cieľovým hárkom "${target.name}".`
      );
    }
  }

  let appendedRows = 0;
  const deletedSheets: string[] = [];

  // Pridá zásielky z ostatných hárkov na koniec hlavnej tabuľky.
  for (let i = 0; i < sheetInfo.length; i++) {
    if (i === targetIndex) {
      continue;
    }

    const source = sheetInfo[i];
    const sourceRowCount = source.table.getRowCount();

    if (sourceRowCount > 0) {
      const values = source.table
        .getRangeBetweenHeaderAndTotal()
        .getValues();

      target.table.addRows(-1, values);
      appendedRows += values.length;
    }

    deletedSheets.push(source.name);
  }

  // Menšie hárky vymaže až po úspešnom prenose všetkých údajov.
  for (let i = 0; i < sheetInfo.length; i++) {
    if (i !== targetIndex) {
      sheetInfo[i].sheet.delete();
    }
  }

  target.sheet.activate();

  console.log(
    `Hotovo: ${appendedRows} zásielok bolo pridaných do ` +
    `"${target.name}". Odstránené hárky: ` +
    deletedSheets.join(", ")
  );
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
