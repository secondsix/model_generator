import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const dir = process.argv[2];
const previewPath = process.argv[3];
const names = (await fs.readdir(dir)).filter((name) => name.endsWith(".xlsx")).sort();
let rows = 0;
let blankConfigs = 0;
let invalidConfigs = 0;
let formulaErrors = 0;
const typeCounts = {};
let firstWorkbook = null;

for (const name of names) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(dir, name)));
  firstWorkbook ??= wb;
  const sheet = wb.worksheets.getItemAt(0);
  const used = sheet.getUsedRange(true);
  const values = used.values;
  const headers = values[0].map((value) => String(value ?? "").trim());
  const typeCol = headers.indexOf("数据类型");
  const configCol = headers.indexOf("数据类型配置");
  if (typeCol < 0 || configCol < 0) throw new Error(`${name}: 缺少目标列`);
  for (const row of values.slice(1)) {
    if (!row.some((value) => value !== null && value !== "")) continue;
    rows += 1;
    const type = String(row[typeCol] ?? "");
    typeCounts[type] = (typeCounts[type] ?? 0) + 1;
    const config = String(row[configCol] ?? "").trim();
    if (!config) blankConfigs += 1;
    else {
      try { JSON.parse(config); } catch { invalidConfigs += 1; }
    }
  }
  const errors = await wb.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 20 },
    summary: "final formula error scan",
    maxChars: 3000,
  });
  formulaErrors += errors.ndjson.split("\n").filter((line) => line.includes('"kind":"match"')).length;
}

if (firstWorkbook) {
  const preview = await firstWorkbook.render({ sheetName: "Sheet0", range: "A1:J20", scale: 1.5, format: "png" });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
}

console.log(JSON.stringify({ files: names.length, rows, blankConfigs, invalidConfigs, formulaErrors, typeCounts, previewPath }));
