import ExcelJS, { Style } from "exceljs";

function calculateMaxRowWidth(_ignore: unknown) {
  // dummy implementation
  return 80;
}

async function createExcelSpreadSheet(filename: string) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("My Sheet");

  const headerStyle: Partial<Style> = {
    alignment: {
      horizontal: "left",
    },
    font: {
      name: "Verdana",
      size: 30,
    },
  };

  worksheet.columns = [
    { header: "Id", key: "id", width: 10, style: headerStyle },
    { header: "Name", key: "name", width: 32, style: headerStyle },
    { header: "D.O.B.", key: "dob", width: 15, style: headerStyle },
  ];

  let seenMaxCellWidth = 0;
  const rows = [
    { id: 1, name: "John Doe", dob: new Date(1970, 1, 1) },
    { id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) },
    { id: 3, name: "Jane Doe", dob: new Date(1999, 1, 7) },
    { id: 4, name: "Jane Doe", dob: new Date(2010, 1, 7) },
  ];

  for (const row of rows) {
    seenMaxCellWidth = calculateMaxRowWidth(row);
    worksheet.addRow(row).commit();
  }

  // formatting
  for (const column of worksheet.columns) {
    column.width = seenMaxCellWidth;
  }

  await workbook.xlsx.writeFile(filename);
}

const filename = `testing${new Date().getTime()}.xlsx`;

createExcelSpreadSheet(filename);
