/**
 * Import zásielok do hárku Data.
 *
 * Zdrojom sú hárky, ktorých názov začína niektorou z definovaných
 * predpôn — buď priamy export (zasilky-balik...), alebo výsledok
 * zlúčenia z nástroja csv-merge-ps (merged_DDMMYYYY).
 *
 * Ak názov hárku obsahuje dátum (merged_10092026), použije sa ako
 * dátum odoslania na zberné depo DPD. Inak sa berie dnešný dátum.
 */

function main(workbook: ExcelScript.Workbook): string {
    // Predpony zdrojových hárkov.
    const sourcePrefixes = ["zasilky-balik", "merged"];

    const data = workbook.getWorksheet("Data");

    if (!data) {
        throw new Error("Hárok Data neexistuje.");
    }

    // Nájde všetky zdrojové hárky.
    const sourceSheets = workbook.getWorksheets().filter(sheet => {
        const name = sheet.getName().toLowerCase();
        return sourcePrefixes.some(prefix => name.startsWith(prefix));
    });

    if (sourceSheets.length === 0) {
        throw new Error(
            "Nenašiel sa žiadny hárok s predponou " +
            sourcePrefixes.join(" alebo ") + "."
        );
    }

    const dataUsedRange = data.getUsedRange(true);
    const existingRows = dataUsedRange?.getValues() ?? [];

    // Prvý voľný riadok v hárku Data.
    let nextRow = dataUsedRange
        ? dataUsedRange.getRowIndex() + dataUsedRange.getRowCount()
        : 1;

    // Zoznam už evidovaných zásielok.
    const existingDPD = new Set<string>();
    const existingZCodes = new Set<string>();

    for (let row = 1; row < existingRows.length; row++) {
        const dpd = normalizeKey(existingRows[row][1]);
        const zCode = normalizeKey(existingRows[row][2]);

        if (dpd) existingDPD.add(dpd);
        if (zCode) existingZCodes.add(zCode);
    }

    // Záložný dátum, ak ho názov hárku neobsahuje.
    const todayDate = getCurrentExcelDate();

    let added = 0;
    let skipped = 0;

    // Prehľad, z ktorého hárku sa brali dáta a s akým dátumom.
    const report: string[] = [];

    for (const sourceSheet of sourceSheets) {
        const sheetName = sourceSheet.getName();
        const sourceRange = sourceSheet.getUsedRange(true);

        if (!sourceRange || sourceRange.getRowCount() < 2) {
            continue;
        }

        // Dátum z názvu hárku má prednosť pred dnešným.
        const nameDate = extractDateFromName(sheetName);
        const sendDate = nameDate ?? todayDate;

        let sheetAdded = 0;
        let sheetSkipped = 0;

        const sourceRows = sourceRange.getValues();

        // Prvý riadok obsahuje hlavičky.
        for (let row = 1; row < sourceRows.length; row++) {
            // Order No.
            const dpdNumber = String(sourceRows[row][0] ?? "").trim();

            // Podacie číslo.
            const zCode = formatZCode(
                String(sourceRows[row][1] ?? "")
            );

            // Vytvorené.
            const createdDate = convertToExcelDate(
                sourceRows[row][6]
            );

            if (!dpdNumber && !zCode) {
                continue;
            }

            const dpdKey = normalizeKey(dpdNumber);
            const zKey = normalizeKey(zCode);

            // Kontrola duplicity podľa DPD čísla alebo Z kódu.
            if (
                (dpdKey && existingDPD.has(dpdKey)) ||
                (zKey && existingZCodes.has(zKey))
            ) {
                skipped++;
                sheetSkipped++;
                continue;
            }

            // Dátum vytvorenia, DPD číslo a Z kód.
            const targetABC = data.getRangeByIndexes(
                nextRow,
                0,
                1,
                3
            );

            targetABC.setValues([[
                createdDate,
                dpdNumber,
                zCode
            ]]);

            targetABC.setNumberFormats([[
                "d.m.yyyy",
                "@",
                "@"
            ]]);

            // Dátum odoslania na zberné depo DPD.
            const targetG = data.getCell(nextRow, 6);
            targetG.setValue(sendDate);
            targetG.setNumberFormat("d.m.yyyy");

            if (dpdKey) existingDPD.add(dpdKey);
            if (zKey) existingZCodes.add(zKey);

            nextRow++;
            added++;
            sheetAdded++;
        }

        report.push(
            `  ${sheetName}: +${sheetAdded}, duplicity ${sheetSkipped}` +
            `, dátum ${formatExcelDate(sendDate)}` +
            (nameDate === null ? " (dnešný)" : " (z názvu)")
        );
    }

    return (
        `Import dokončený. Pridané zásielky: ${added}. ` +
        `Preskočené duplicity: ${skipped}.\n` +
        report.join("\n")
    );
}

/**
 * Vytiahne dátum z názvu hárku vo formáte DDMMYYYY.
 * Zachytí merged_10092026, merged-10092026 aj merged10092026.
 * Vráti null, ak názov dátum neobsahuje alebo je neplatný.
 */
function extractDateFromName(name: string): number | null {
    const match = name.match(/(\d{2})(\d{2})(\d{4})/);

    if (!match) return null;

    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);

    if (day < 1 || day > 31) return null;
    if (month < 1 || month > 12) return null;
    if (year < 2000 || year > 2100) return null;

    return dateToExcelSerial(year, month, day);
}

function formatExcelDate(serial: number): string {
    const ms = serial * 86400000 + Date.UTC(1899, 11, 30);
    const d = new Date(ms);

    return (
        d.getUTCDate() + "." +
        (d.getUTCMonth() + 1) + "." +
        d.getUTCFullYear()
    );
}

function normalizeKey(value: unknown): string {
    return String(value ?? "")
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, "");
}

function formatZCode(value: string): string {
    const cleaned = value
        .toUpperCase()
        .replace(/\s+/g, "");

    if (!cleaned) return "";

    const digits = cleaned.replace(/^Z/, "");

    if (/^\d{10}$/.test(digits)) {
        return `Z ${digits.substring(0, 5)} ${digits.substring(5)}`;
    }

    return `Z ${digits}`;
}

function convertToExcelDate(
    value: string | number | boolean
): string | number {
    // Excelový dátum už uložený ako číslo.
    if (typeof value === "number") {
        return Math.floor(value);
    }

    const text = String(value ?? "").trim();

    // Formát d.m.rrrr s voliteľným časom.
    const match = text.match(
        /^(\d{1,2})\.(\d{1,2})\.(\d{4})/
    );

    if (!match) return text;

    return dateToExcelSerial(
        Number(match[3]),
        Number(match[2]),
        Number(match[1])
    );
}

function getCurrentExcelDate(): number {
    const now = new Date();

    return dateToExcelSerial(
        now.getFullYear(),
        now.getMonth() + 1,
        now.getDate()
    );
}

function dateToExcelSerial(
    year: number,
    month: number,
    day: number
): number {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const date = Date.UTC(year, month - 1, day);

    return Math.floor(
        (date - excelEpoch) / 86400000
    );
}
