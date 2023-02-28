//import xlsx library to read the Excel file.
import * as XLSX from "xlsx";
import * as path from "path";

//Interface MergedCell that represents a merged cell, with s and e properties representing the start and end coordinates of the merged area
interface MergedCell {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

//function getMergedCells that takes a file path and returns an object mapping sheet names to arrays of MergedCell objects.
function getMergedCells(filePath: string): { [key: string]: MergedCell[] } {
  //parse the Excel file at the given path and get a workbook object.
  const workbook = XLSX.readFile(filePath);

  const mergedCells: { [key: string]: MergedCell[] } = {};

  //Iterate over each sheet in the workbook
  workbook.SheetNames.forEach((sheetName) => {
    //extract the worksheet object
    const worksheet = workbook.Sheets[sheetName];

    //check if the worksheet has any merged cells by looking at the !merges property. If it does, we iterate over the array of merged ranges
    if (worksheet["!merges"]) {
      const merges = worksheet["!merges"] as XLSX.Range[];
      const mergedCellsOnSheet: MergedCell[] = [];

      merges.forEach((merge) => {
        const { s, e } = merge;

        //For each merged range, we create a MergedCell object with the start and end coordinates and push it into an array of merged cells for the sheet
        const mergedCell: MergedCell = {
          s: { r: s.r, c: s.c },
          e: { r: e.r, c: e.c },
        };
        mergedCellsOnSheet.push(mergedCell);
      });

      //this adds the array of MergedCell objects to the mergedCells object, using the sheet name as the key
      mergedCells[sheetName] = mergedCellsOnSheet;
    }
  });

  return mergedCells;
}

//Program to use function

//process.argv to get the file path from the command line arguments.
const filePath = process.argv[2];

//path.resolve to get the absolute file path
const absoluteFilePath = path.resolve(filePath);

//getMergedCells call with the file path to store the resulting mergedCells object.
const mergedCells = getMergedCells(absoluteFilePath);

console.log(mergedCells);

