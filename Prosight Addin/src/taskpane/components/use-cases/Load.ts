// excelFormatter.ts
export const loadDataFunction = async (context: Excel.RequestContext) => {
    const workbook = context.workbook;

    // Get the source and target sheets
    const sourceSheet = workbook.worksheets.getItem("Model1");
    const targetSheet = workbook.worksheets.getActiveWorksheet();

    // Get the range with data in the source sheet
    const sourceRange = sourceSheet.getUsedRange();
    sourceRange.load([
        "address",
        "values",
        "formulas",
        "rowCount",
        "columnCount",
        "format/fill",
        "format/font"
    ]);
    await context.sync();

    if (!sourceRange.address) {
        console.error("Source range is empty or invalid.");
        return;
    }

    console.log("Source range address:", sourceRange.address);

    // Define the target range using the same dimensions as the source range
    const rowCount = sourceRange.rowCount;
    const columnCount = sourceRange.columnCount;
    const targetRange = targetSheet.getRangeByIndexes(0, 0, rowCount, columnCount);

    // Copy values and formulas
    targetRange.values = sourceRange.values;
    targetRange.formulas = sourceRange.formulas;
    
    targetSheet.calculate(true);

    await context.sync();

    console.log("Data and formatting copied successfully.");
};
