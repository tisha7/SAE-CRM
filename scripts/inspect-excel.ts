import XLSX from "xlsx";
import fs from "fs";
import path from "path";

const sourceDir = "/workspaces/sae-crm-source";

function walk(dir: string): string[] {
  const results: string[] = [];

  for (const item of fs.readdirSync(dir)) {
    const fullPath = path.join(dir, item);
    const fullItemPath = path.join(dir, item);
    const stat = fs.statSync(fullItemPath);

    if (stat.isDirectory()) {
      results.push(...walk(fullItemPath));
    } else if (/\.(xlsx|xls)$/i.test(item)) {
      results.push(fullItemPath);
    }
  }

  return results;
}

const files = walk(sourceDir);

console.log("\n========================================");
console.log("SAE CRM — EXCEL SOURCE INVENTORY");
console.log("========================================");

for (const file of files) {
  console.log(`\n📘 FILE: ${path.basename(file)}`);
  console.log(`   PATH: ${file}`);

  const workbook = XLSX.readFile(file);

  console.log(`   SHEETS: ${workbook.SheetNames.length}`);

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const range = sheet["!ref"] || "EMPTY";

    const data = XLSX.utils.sheet_to_json<any[]>(sheet, {
      header: 1,
      defval: null,
      raw: true,
    });

    const columnCount = data.reduce(
      (max, row) => Math.max(max, row.length),
      0
    );

    const nonEmptyRows = data.filter((row) =>
      row.some((cell) => cell !== null && cell !== "")
    );

    console.log(`\n   ── SHEET: ${sheetName}`);
    console.log(`      Range: ${range}`);
    console.log(`      Rows: ${data.length}`);
    console.log(`      Columns: ${columnCount}`);
    console.log(`      Non-empty rows: ${nonEmptyRows.length}`);

    console.log("      First 5 non-empty rows:");

    nonEmptyRows.slice(0, 5).forEach((row, index) => {
      console.log(`        ${index + 1}: ${JSON.stringify(row)}`);
    });
  }
}

console.log("\n========================================");
console.log("END OF INVENTORY");
console.log("========================================\n");
