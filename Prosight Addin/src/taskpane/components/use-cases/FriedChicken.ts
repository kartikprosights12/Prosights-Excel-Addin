
export const friedChicken = async (context: Excel.RequestContext) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Define the data
    const data = [
        ["Year", "Market Size in USD Bn"],
        ["2018", 4.5],
        ["2019", 5],
        ["2020", 6.2],
        ["2021", 6.2],
        ["2022", 6.5],
        ["2023", 6.7],
        ["2024", 6.85],
        ["2025", 7],
        ["2026", 7.2],
        ["2030", 9.5],
        ["2031", 10],
        ["2032", 10.52],
    ];

    // Insert data into a range
    const range = sheet.getRange("A1:B13"); // Adjust the range to fit the data
    range.values = data;

     // Load and sync the range address
    range.load("address");
    await context.sync();

      // Create a table from the data
      const table = sheet.tables.add(range.address, true); // true indicates headers are included
      table.name = "MarketSizeTable" + Math.random().toString(36).substring(2, 15);
  
      // Format the table
      table.columns.getItemAt(1).getRange().numberFormat = [["#,##0.00"]]; // Format the second column as numbers
  
      // Add a bar chart
      const chart = sheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.columns);
  
      // Position the chart
      chart.setPosition("D1", "L15"); // Adjust chart position as needed
  
      // Customize the chart
      chart.title.text = "Market Size Over the Years";
      chart.axes.categoryAxis.title.text = "Year";
      chart.axes.valueAxis.title.text = "Market Size (USD Bn)";

    await context.sync();

    console.log("Data and chart added successfully.");
}