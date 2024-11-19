export const fixSheet = async (context: Excel.RequestContext) => {
  const workbook = context.workbook;

  // Get the source and target sheets
  const sheet = workbook.worksheets.getActiveWorksheet();
  // Update the formula for cell (84, E)
  const cellE = sheet.getCell(83, 4); // Row and column are 0-based (83 = 84 - 1, 4 = E)
  cellE.formulas = [["=E82+E77"]]; // Replace with your desired formula

  // Update the formula for cell (84, F)
  const cellF = sheet.getCell(83, 5); // Row and column are 0-based (83 = 84 - 1, 5 = F)
  cellF.formulas = [["=F82+F77"]]; // Replace with your desired formula
  
  const cellG = sheet.getCell(83, 6); // Row and column are 0-based (83 = 84 - 1, 5 = F)
  cellG.formulas = [["=G82+G77"]]; // Replace with your desired formula
  
  const cellH = sheet.getCell(83, 7); // Row and column are 0-based (83 = 84 - 1, 5 = F)
  cellH.formulas = [["=H82+H77"]]; // Replace with your desired formula
  sheet.calculate(true);
  await context.sync();

  console.log("Data and formatting copied successfully.");
};
