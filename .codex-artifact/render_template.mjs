import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = process.argv[2];
const out = process.argv[3];
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const preview = await wb.render({ sheetName: "Sheet0", range: "A1:J12", scale: 1.5, format: "png" });
await fs.writeFile(out, new Uint8Array(await preview.arrayBuffer()));
console.log(out);
