// excelFormatter.ts
export const updateOptions = async (context: Excel.RequestContext) => {
    const workbook = context.workbook;
    const sheet = workbook.worksheets.getActiveWorksheet();

    const cellE = sheet.getCell(30, 3); // Row and column are 0-based (83 = 84 - 1, 4 = E)
    cellE.formulas = [["0.20"]]; // Replace with your desired formula

    sheet.calculate(true);

    await context.sync();

    console.log("Data and formatting copied successfully.");
};
