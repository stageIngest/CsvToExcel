/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

/// <reference types="office-js" />

/* global console, document, Excel, Office */
var UserInput;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";

    const browse = document.getElementById("browse")!;
    browse.addEventListener("click", () => {
      (document.getElementById("fileInput") as HTMLInputElement).click();
    });

    UserInput = document.getElementById("fileInput") as HTMLInputElement;
    UserInput.addEventListener("change", fileImport);
  }
});

let colCount;
let fileNo = 1;
let FileSelected

async function fileImport() {
  if (!UserInput || !UserInput.files) return;

  for (let i = 0; i < UserInput.files.length; i++) {
    FileSelected = UserInput.files[i];
    if (FileSelected) {
      await ReadFile(FileSelected); 
    }
  }
}


async function ReadFile(FileSelected: File): Promise<void> {
  return new Promise<void>((end) => {
    const Reader = new FileReader();
    Reader.readAsArrayBuffer(FileSelected);

    Reader.onload = async () => {
      try {
        const buffer = Reader.result as ArrayBuffer;
        if (!buffer) return end();

        const CSVText = new TextDecoder().decode(buffer).trim();
        if (!CSVText) return end();

        const rows = CSVText
          .split(/\r\n/)
          .filter(r => r.trim() !== "");

        if (!rows.length) return end();

        const CSVData: (string | number)[][] = rows.map((row, rowIndex) => {
          const isHeader = rowIndex === 0;

          if (row.includes(";")) {
            return row.split(";").map(cell => setCellAs(cell));
          }

          if (row.includes(",")) {
            const result: (string | number)[] = [];
            let current = "";
            let inQuotes = false;

            for (let i = 0; i < row.length; i++) {
              const char = row[i];
              if (char === '"') {
                inQuotes = !inQuotes;
                continue;
              }
              if (char === "," && !inQuotes) {
                result.push(processCell(current, isHeader));
                current = "";
              } else {
                current += char;
              }
            }
            result.push(processCell(current, isHeader));
            return result;
          }

          return [];
        });

        await writeExcel(CSVData);
        end();                    
      } catch (err) {
        console.error("Errore CSV:", err);
        end();
      }
    };

    Reader.onerror = () => end();
  });
}


function processCell(cell: string, isHeader: boolean): string | number {
  let trimmed = cell.trim();
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    trimmed = trimmed.slice(1, -1);
  }
  return setCellAs(trimmed);
}

function setCellAs(cell: string): string | number {
  const str = cell.trim();
  const numericStr = str.replace(/,/g, '.');

  if (numericStr.includes('.') && /\.\d+/.test(numericStr)) {
    return parseFloat(numericStr);
  }
  return str;
}

function SetNumericFormat(CSVData: (string | number)[][]): boolean[] {
  const colCount = CSVData[0].length;
  const numericCols: boolean[] = Array(colCount).fill(true);

  for (let col = 0; col < colCount; col++) {
    if (CSVData[0][col].toString().toLowerCase() == "matricola") {
      numericCols[col] = false;
    }
    else if (CSVData[0][col].toString().toLowerCase().includes('nr.')) {
      numericCols[col] = false;
    }
    for (let row = 1; row < CSVData.length; row++) {
      if (typeof CSVData[row][col] !== "number") {
        numericCols[col] = false;
        break;
      }
    }
  }
  return numericCols;
}


async function writeExcel(CSVData: (string | number)[][]) {
  if (!CSVData.length) return;

  colCount = CSVData[0].length;

  for (let i = 0; i < CSVData.length; i++) {
    if (CSVData[i].length < colCount) {
      CSVData[i].push("");
    } else if (CSVData[i].length > colCount) {
      const fixed = CSVData[i].slice(0, colCount - 1);
      const overflow = CSVData[i].slice(colCount - 1).join(",");
      fixed.push(overflow);
      CSVData[i] = fixed;
    }
  }

  await createNewExcel(CSVData);
}


async function createNewExcel(CSVData: (string | number)[][]) {
  await Excel.run(async (context) => {
    let worksheet;

    if (fileNo === 1) {
      worksheet = context.workbook.worksheets.getActiveWorksheet();
    } else {
      const workbook = Excel.createWorkbook();
      worksheet = context.workbook.worksheets.getFirst();
    }

    fileNo++;

    const rowCount = CSVData.length;
    const colCount = CSVData[0].length;

    const range = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
    range.values = CSVData;

    const isNumeric = SetNumericFormat(CSVData);
    for (let column = 0; column < colCount; column++) {
      if (isNumeric[column]) {
        range.getColumn(column).numberFormat = [['0.00']];
      }
    }

    range.format.autofitColumns();
    range.format.autofitRows();

    await context.sync();
  });
}
