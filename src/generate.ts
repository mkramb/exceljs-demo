import ExcelJS, { Style } from "exceljs";

async function createExcelSpreadSheet(filename: string) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");

    const headerStyle: Partial<Style> = {
        alignment: {
            horizontal: 'left'
        },
        font: {
            name: 'Verdana',
            size: 30
        }
    }

    worksheet.columns = [
        { header: "Id", key: "id", width: 10, style: headerStyle },
        { header: "Name", key: "name", width: 32, style: headerStyle },
        { header: "D.O.B.", key: "dob", width: 15, style: headerStyle },
    ];

    worksheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
    worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });
    worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1999, 1, 7) });
    worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(2010, 1, 7) });

    await workbook.xlsx.writeFile(filename);
}

const filename = `testing${new Date().getTime()}.xlsx`;

createExcelSpreadSheet(filename);