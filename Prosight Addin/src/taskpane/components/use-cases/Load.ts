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

    // Iterate over each cell to copy formatting
    for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < columnCount; col++) {
            const sourceCell = sourceRange.getCell(row, col);
            const targetCell = targetRange.getCell(row, col);

            // Load cell-specific formatting properties
            sourceCell.format.fill.load("color");
            sourceCell.format.font.load(["bold", "color"]);
            await context.sync();

            // Apply cell-specific formatting
            if (sourceCell.format.fill.color) {
                targetCell.format.fill.color = sourceCell.format.fill.color;
            }
            targetCell.format.font.bold = sourceCell.format.font.bold;
            targetCell.format.font.color = sourceCell.format.font.color;
        }
    }

    await context.sync();

    console.log("Data and formatting copied successfully.");
};
