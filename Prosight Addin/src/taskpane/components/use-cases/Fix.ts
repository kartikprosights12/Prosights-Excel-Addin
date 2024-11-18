export const fixSheet = async (context: Excel.RequestContext) => {

    const workbook = context.workbook;

    // Get the source and target sheets
    const sourceSheet = workbook.worksheets.getItem("Model3");
    const targetSheet = workbook.worksheets.getItem("Sheet1");

    // Get the range with data in the source sheet
    const sourceRange = sourceSheet.getUsedRange();
    sourceRange.load([
        "address",
        "values",
        "formulas",
        "rowCount",
        "columnCount",
        "format/fill/color",
        "format/font/bold",
        "format/font/color"
    ]);
    await context.sync();

    if (!sourceRange.address) {
        console.error("Source range is empty or invalid.");
        return;
    }

    console.log("Source range address:", sourceRange.address);
    console.log("Source range fill color:", sourceRange.format.fill.color);

    // Define the target range using the same dimensions as the source range
    const rowCount = sourceRange.rowCount;
    const columnCount = sourceRange.columnCount;
    const targetRange = targetSheet.getRangeByIndexes(0, 0, rowCount, columnCount);

    // Copy data and formulas to the target sheet
    targetRange.values = sourceRange.values;
    targetRange.formulas = sourceRange.formulas;

    // Copy fill color if it exists
    if (sourceRange.format.fill.color) {
        targetRange.format.fill.color = sourceRange.format.fill.color;
    } else {
        console.log("No fill color found in the source range.");
    }

    // Copy font properties
    targetRange.format.font.bold = sourceRange.format.font.bold;
    targetRange.format.font.color = sourceRange.format.font.color;

    await context.sync();

    console.log("Data and formatting copied successfully.");
}
