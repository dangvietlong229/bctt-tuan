import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

function parseArgs(argv) {
  const result = {};
  for (let i = 0; i < argv.length; i += 2) {
    const key = argv[i]?.replace(/^--/, "");
    result[key] = argv[i + 1];
  }
  return result;
}

const args = parseArgs(process.argv.slice(2));
for (const required of ["artifact-tool", "buy", "sell", "output"]) {
  if (!args[required]) throw new Error(`Missing --${required}`);
}

const modulePath = path.join(
  args["artifact-tool"],
  "@oai",
  "artifact-tool",
  "dist",
  "artifact_tool.mjs",
);
const { FileBlob, SpreadsheetFile, Workbook } = await import(pathToFileURL(modulePath).href);

async function readTopRows(sourcePath, side) {
  const source = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
  const sheet = source.worksheets.getItem("Sheet1");
  const rows = sheet.getRange("A6:G15").values;
  const header = side === "sell"
    ? ["Mã", "Giá", "Bán ròng (GT)", "Tỷ trọng bán của Khối Tự doanh"]
    : ["Mã", "Giá", "Mua ròng (GT)", "Tỷ trọng mua của Khối Tự doanh"];
  return [
    header,
    ...rows.map((row) => [row[0] ?? "", row[1] ?? null, row[5] ?? null, row[6] ?? null]),
  ];
}

function styleTopSheet(sheet, matrix) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange("A1:D11").values = matrix;
  sheet.getRange("A1:D1").format = {
    fill: "#17365D",
    font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#A9B8CA" },
    rowHeight: 34,
  };
  sheet.getRange("A2:D11").format = {
    font: { name: "Arial", size: 10, color: "#1F2937" },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
    rowHeight: 21,
  };
  for (let row = 2; row <= 11; row += 2) {
    sheet.getRange(`A${row}:D${row}`).format.fill = "#F4F7FB";
  }
  sheet.getRange("A2:A11").format.font = { name: "Arial", size: 10, bold: true, color: "#17365D" };
  sheet.getRange("A2:A11").format.horizontalAlignment = "center";
  sheet.getRange("B2:C11").format.horizontalAlignment = "right";
  sheet.getRange("D2:D11").format.horizontalAlignment = "right";
  sheet.getRange("B2:C11").setNumberFormat("#,##0");
  sheet.getRange("D2:D11").setNumberFormat("0.00%");
  sheet.getRange("A1:A11").format.columnWidth = 15;
  sheet.getRange("B1:B11").format.columnWidth = 14;
  sheet.getRange("C1:C11").format.columnWidth = 24;
  sheet.getRange("D1:D11").format.columnWidth = 38;
}

const sold = await readTopRows(args.sell, "sell");
const bought = await readTopRows(args.buy, "buy");
const workbook = Workbook.create();
styleTopSheet(workbook.worksheets.add("Top_10_Ban_Rong"), sold);
styleTopSheet(workbook.worksheets.add("Top_10_Mua_Rong"), bought);

const sourceSheet = workbook.worksheets.add("Nguon_Du_Lieu");
sourceSheet.showGridLines = false;
sourceSheet.getRange("A1:B3").values = [
  ["Nguồn", "Tệp FiinProX"],
  ["Top bán ròng 1 tuần", path.basename(args.sell)],
  ["Top mua ròng 1 tuần", path.basename(args.buy)],
];
sourceSheet.getRange("A1:B1").format = {
  fill: "#17365D",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
};
sourceSheet.getRange("A2:B3").format = {
  font: { name: "Arial", size: 10, color: "#1F2937" },
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
};
sourceSheet.getRange("A1:A3").format.columnWidth = 25;
sourceSheet.getRange("B1:B3").format.columnWidth = 78;

await fs.mkdir(path.dirname(args.output), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(args.output);

if (args.preview) {
  await fs.mkdir(args.preview, { recursive: true });
  for (const name of ["Top_10_Ban_Rong", "Top_10_Mua_Rong", "Nguon_Du_Lieu"]) {
    const preview = await workbook.render({ sheetName: name, autoCrop: "all", scale: 1.5, format: "png" });
    await fs.writeFile(path.join(args.preview, `${name}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
}

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
process.stdout.write(`${JSON.stringify({ output: args.output, errorScan }, null, 2)}\n`);
