/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference types="office-js" />

/* global console, document, Excel, Office */
let UserInput;
let fileNo = 1;

   Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;
    const run = document.getElementById("run") as HTMLButtonElement;
    
    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        console.log("File selezionato:", fileInput.files[0].name);
        document.getElementById("fileName")!.textContent =
          "Selected file: " + fileInput.files[0].name;
      }
    });

    UserInput = fileInput
    
    run.addEventListener("click", fileImport);
  }
});


let colCount;

var FileSelected: File;
var FileName : string;

//gestione nel caso di file multipli
async function fileImport() {
  if (!UserInput || !UserInput.files) return;

  for (let i = 0; i < UserInput.files.length; i++) {
    FileSelected = UserInput.files[i];
    if (FileSelected) {
      FileName = FileSelected.name;
      await ReadFile(FileSelected);
    }
  }
}

//legge il file e divide in righe e in celle
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
          .split(/\r?\n/)
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

//pulizia della stringa nella cella
function processCell(cell: string, isHeader: boolean): string | number {
  let trimmed = cell.trim();
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    trimmed = trimmed.slice(1, -1);
  }
  return setCellAs(trimmed);
}

//interpretazione del contenuto della cella 
function setCellAs(cell: string): string | number {
  const str = cell.trim();
  const numericStr = str.replace(/,/g, '.');

  if (numericStr.includes('.') && /\.\d+/.test(numericStr)) {
    return parseFloat(numericStr);
  }
  return str;
}

//formato di ogni colonna, controllando tutte le singole celle escludendo l'header
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

//gestione della scrittura su excel, con controllo dell'overflow rispetto alle colonne dell'header
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

//creazione del file excel vero
//vengono salvati i file in diversi fogli sullo stesso file excel
async function createNewExcel(CSVData: (string | number)[][]) {

  Excel.run(async (context) => {
    let worksheetAccessible: Excel.Worksheet;

    if (fileNo == 1) {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const CleanName = FileName.split(".")[0].substring(0, 31);
      worksheet.name = CleanName;
      worksheetAccessible = worksheet;
    } else {
      const worksheet = context.workbook.worksheets.add();
      const CleanName = FileName.split(".")[0].substring(0,31);
      worksheet.name = CleanName;
      worksheetAccessible = worksheet;
    }

    fileNo++;
    
    const rowCount = CSVData.length;
    const colCount = CSVData[0].length;

    const range = worksheetAccessible.getRangeByIndexes(0, 0, rowCount, colCount);
    range.values = CSVData;

    const isNumeric = SetNumericFormat(CSVData);
    for (let column = 0; column < colCount; column++) {
      if (isNumeric[column]) {
        range.getColumn(column).numberFormat = [["#,##0.00;[Red]-#,##0.00"]];
      }
    }

    const headerRange = range.getRow(0);
    headerRange.format.font.bold = true;

    range.format.autofitColumns();
    range.format.autofitRows();

    await context.sync();  
  });
}

