import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectRoot = path.resolve(import.meta.dirname, "..");
const outputDir = path.join(
  projectRoot,
  "src",
  "agribank_v3",
  "resources",
  "templates",
);
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const questions = workbook.worksheets.add("CauHoi");
questions.showGridLines = false;
questions.freezePanes.freezeRows(1);

const headers = [
  "ID",
  "STT",
  "Mã nghiệp vụ",
  "Nghiệp vụ",
  "Chuyên đề",
  "Câu hỏi",
  "Đáp án A",
  "Đáp án B",
  "Đáp án C",
  "Đáp án D",
  "Đáp án đúng",
  "Đáp án không đảo",
  "Nguồn",
];
const sample = [
  null,
  1,
  "KTC",
  "Kiến thức chung",
  "Ngân hàng điện tử",
  "Nội dung câu hỏi mẫu",
  "Nội dung đáp án A",
  "Nội dung đáp án B",
  "Nội dung đáp án C",
  "Nội dung đáp án D",
  "A",
  "C,D",
  "Văn bản hoặc nguồn tham khảo",
];
questions.getRange("A1:M2").values = [headers, sample];
questions.getRange("A1:M1").format = {
  fill: "#831F41",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  verticalAlignment: "center",
};
questions.getRange("A2:M2").format = {
  fill: "#FFF8FA",
  wrapText: true,
  verticalAlignment: "top",
};
questions.getRange("A1:M2").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9C8CF",
};
questions.getRange("A1:M1").format.rowHeight = 32;
questions.getRange("A2:M2").format.rowHeight = 52;
const widths = [10, 8, 16, 23, 23, 48, 30, 30, 30, 30, 15, 20, 38];
for (let index = 0; index < widths.length; index += 1) {
  questions.getRangeByIndexes(0, index, 2, 1).format.columnWidth = widths[index];
}
questions.getRange("K2:K1000").dataValidation = {
  rule: { type: "list", values: ["A", "B", "C", "D"] },
};
questions.tables.add("A1:M2", true, "BangCauHoi").style = "TableStyleMedium2";

const guide = workbook.worksheets.add("HuongDan");
guide.showGridLines = false;
guide.getRange("A1:F1").merge();
guide.getRange("A1").values = [["HƯỚNG DẪN NHẬP NGÂN HÀNG CÂU HỎI AGRIBANKV3"]];
guide.getRange("A1:F1").format = {
  fill: "#831F41",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
guide.getRange("A1:F1").format.rowHeight = 34;
guide.getRange("A3:B9").values = [
  ["Cột", "Cách nhập"],
  ["Mã nghiệp vụ", "Mã ngắn, duy nhất. Ví dụ: KTC, CNTT, TD."],
  ["Nghiệp vụ", "Nhóm nghiệp vụ cấp cao. Ví dụ: Kiến thức chung."],
  ["Chuyên đề", "Nhóm con của nghiệp vụ. Ví dụ: Ngân hàng điện tử."],
  ["Đáp án đúng", "Chỉ nhập A, B, C hoặc D."],
  ["Đáp án không đảo", "Nhập các ký tự cần giữ nguyên, cách nhau bằng dấu phẩy. Ví dụ: A,C."],
  ["ID / STT", "Không bắt buộc. Để trống ID khi thêm câu hỏi mới."],
];
guide.getRange("A3:B3").format = {
  fill: "#F2D9E2",
  font: { bold: true, color: "#4A1327" },
};
guide.getRange("A3:B9").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9C8CF",
};
guide.getRange("A3:A9").format.columnWidth = 24;
guide.getRange("B3:B9").format.columnWidth = 72;
guide.getRange("A3:B9").format.wrapText = true;
guide.getRange("A4:B9").format.rowHeight = 32;

const preview = await workbook.render({
  sheetName: "CauHoi",
  range: "A1:M2",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(projectRoot, "tools", "quiz-template-preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "AgribankV3-Mau-Cau-Hoi.xlsx"));

const inspection = await workbook.inspect({
  kind: "table",
  range: "CauHoi!A1:M2",
  include: "values,formulas",
  tableMaxRows: 3,
  tableMaxCols: 13,
});
console.log(inspection.ndjson);
