// excelFormatter.ts

export const formatExcelSheet = async (context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Title and Operating Case
    const titleRange = sheet.getRange("B2");
    titleRange.values = [["LBO Model"]];
    titleRange.format.font.bold = true;
    titleRange.format.font.size = 14;

    const operatingCaseRange = sheet.getRange("B4:C4");
    operatingCaseRange.values = [["Operating Case"]];
    operatingCaseRange.merge(); // Merge cells
    operatingCaseRange.format.horizontalAlignment = "Left";

    sheet.getRange("D4").values = [["3"]];
    sheet.getRange("D4").format.font.bold = true;
    sheet.getRange("D4").format.font.color = "blue";

    // Formatting for the Operating Section
    const operatingHeaders = sheet.getRange("B7:B20");
    operatingHeaders.values = [
        ["Operating"],
        ["# New Restaurant Growth"],
        ["AUV / restaurant"],
        ["AUV % Growth"],
        ["% Gross Margin"],
        [""],
        ["Rent Expense"],
        ["Owners Base Salary"],
        ["Owners Revenue Share"],
        ["Growth Capex"],
        ["Maint Capex"],
        ["Operating Expenses"]
    ];
    operatingHeaders.format.font.bold = true;

    // Scenarios
    const scenariosRange = sheet.getRange("F6:H6");
    scenariosRange.values = [["Downside", "Base", "Upside"]];
    scenariosRange.format.fill.color = "#E6E6E6"; // Light gray
    scenariosRange.format.font.bold = true;

    // Populate scenario data
    const scenarioData = sheet.getRange("F8:H20");
    scenarioData.values = [
        ["2", "6", "10"],
        ["$1,500", "$1,500", "$1,500"],
        ["0.0%", "0.0%", "0.0%"],
        ["70.0%", "70.0%", "70.0%"],
        ["", "", ""],
        ["$75/rest.", "$75/rest.", "$75/rest."],
        ["$500/own.", "$500/own.", "$500/own."],
        ["0.5%", "0.5%", "0.5%"],
        ["$300/rest.", "$300/rest.", "$300/rest."],
        ["$10/rest.", "$10/rest.", "$10/rest."],
        ["$650/rest.", "$650/rest.", "$650/rest."]
    ];
    scenarioData.format.font.color = "blue";

    // Apply additional formatting (e.g., Operating Assumptions)
    const assumptionsHeader = sheet.getRange("B22");
    assumptionsHeader.values = [["x Operating Assumptions"]];
    assumptionsHeader.format.font.bold = true;

    const ongoingHeader = sheet.getRange("F22");
    ongoingHeader.values = [["Ongoing Assumptions"]];
    ongoingHeader.format.font.bold = true;

    const ongoingData = sheet.getRange("F24:H28");
    ongoingData.values = [
        ["Tax Rate", "40.0%"],
        ["Min Cash", "$2,000.0"],
        ["Depreciation", "1.0% of sales"],
        ["Starting restaurants", "100"],
        ["Opex", "$70,000.0"]
    ];

    // Global formatting
    sheet.getRange("B2:H28").format.horizontalAlignment = "Center";
    sheet.getRange("B2:H28").format.wrapText = true;

    await context.sync();
};
