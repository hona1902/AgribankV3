import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("outputs/quiz-export-preview.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const inspection = await workbook.inspect({
  kind: "table",
  range: "Bo Cau Hoi!A1:D30",
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 4,
});
console.log(inspection.ndjson);
const preview = await workbook.render({
  sheetName: "Bo Cau Hoi",
  range: "A1:D30",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  "outputs/quiz-export-preview.png",
  new Uint8Array(await preview.arrayBuffer()),
);
