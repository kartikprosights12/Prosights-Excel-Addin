import React, { useState } from "react";
import Message from "./Message";
import InputBox from "./InputBox";
import "./App.css";
import axios from "axios";
import Loader from "./Loader/Loader";
import { formatExcelSheet } from "./use-cases/Format";
import { loadDataFunction } from "./use-cases/Load";
import { fixSheet } from "./use-cases/Fix";


const Chat: React.FC = () => {
  const [messages, setMessages] = useState([]);
  const [loading, setLoading] = useState(false); // New loading state

  // Function to add messages, fetch formula, and call the API if needed
  const addMessage = async (text: string, type: "question" | "explanation") => {
    if (type === "question") {
      setLoading(true); // Show loader
      // Fetch the formula from the active cell when user submits a message
      const cellFormula = await fetchCellFormula();


      //load data case
      if (text.toLowerCase().includes("build") && text.toLowerCase().includes("lbo")) {
        await Excel.run(async (context) => {
          try {
              await loadDataFunction(context);
              console.log("Excel sheet formatted successfully!");
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
          }
      });
      }  
      
      //borken and fix case
      else if (text.toLowerCase().includes("broken") && text.toLowerCase().includes("help me") && text.toLowerCase().includes("fix")) {
        await Excel.run(async (context) => {
          try {
              await fixSheet(context);
              console.log("Excel sheet formatted successfully!");
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
          }
      });
      } 

      //format case
      else if (text.toLowerCase().includes("format")) {
        await Excel.run(async (context) => {
          try {
              await formatExcelSheet(context);
              console.log("Excel sheet formatted successfully!");
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
          }
      });
      } 

      if (cellFormula) {
        // Add the formula to the chat as an explanation
        setMessages((prevMessages) => [...prevMessages, { text: text, type: "question" }]);

        // Send both the user message and formula to the API
        //const response = await sendToAPI(text, cellFormula);
        setMessages((prevMessages) => [...prevMessages, { text: "This is a demo explanation", type: "explanation" }]);
        setLoading(false); // Hide loader
      } else {
        // If no formula found, add a message indicating this
        setMessages((prevMessages) => [
          ...prevMessages,
          { text: "No formula found in the selected cell.", type: "explanation" },
        ]);
      }
    }
  };

  // const demoFunction = async () => {
  //   Excel.run(function (context) {
  //     // Define the worksheet
  //     const sheet = context.workbook.worksheets.getActiveWorksheet();

  //     // Title and Operating Case
  //     sheet.getRange("B4").values = [["LBO Model"]];
  //     sheet.getRange("B4").format.font.bold = true;
  //     sheet.getRange("B4").format.font.size = 14;

  //     sheet.getRange("B6").values = [["Operating Case"]];
  //     sheet.getRange("D6").values = [["3"]];
  //     sheet.getRange("D6").format.font.color = "blue";

  //     // Operating section headers and labels, using direct assignments
  //     sheet.getRange("B8").values = [["# New Restaurant Growth"]];
  //     sheet.getRange("B9").values = [["AUV / restaurant"]];
  //     sheet.getRange("B10").values = [["AUV % Growth"]];
  //     sheet.getRange("B11").values = [["% Gross Margin"]];
  //     sheet.getRange("B13").values = [["Rent Expense"]];
  //     sheet.getRange("B14").values = [["Owners Base Salary"]];
  //     sheet.getRange("B15").values = [["Owners Revenue Share"]];
  //     sheet.getRange("B16").values = [["Growth Capex"]];
  //     sheet.getRange("B17").values = [["Maint Capex"]];
  //     sheet.getRange("B18").values = [["Operating Expenses"]];

  //     sheet.getRange("B8:B18").format.font.bold = true;

  //     // Base scenario values (Column D)
  //     sheet.getRange("E8").values = [["10"]];
  //     sheet.getRange("E9").values = [["$1,500"]];
  //     sheet.getRange("E10").values = [["0.0%"]];
  //     sheet.getRange("E11").values = [["70.0%"]];
  //     sheet.getRange("E13").values = [["$75/rest."]];
  //     sheet.getRange("E14").values = [["$500/own."]];
  //     sheet.getRange("E15").values = [["0.5%"]];
  //     sheet.getRange("E16").values = [["$300/rest."]];
  //     sheet.getRange("E17").values = [["$10/rest."]];
  //     sheet.getRange("E18").values = [["$650/rest."]];
  //     sheet.getRange("E8:E18").format.font.color = "blue";

  //     // Downside, Base, and Upside values for columns F, G, H
  //     // Downside (Column F)
  //     sheet.getRange("F8").values = [["2"]];
  //     sheet.getRange("F9").values = [["$1,500"]];
  //     sheet.getRange("F10").values = [["0.0%"]];
  //     sheet.getRange("F11").values = [["70.0%"]];
  //     sheet.getRange("F13").values = [["$75/rest."]];
  //     sheet.getRange("F14").values = [["$500/own."]];
  //     sheet.getRange("F15").values = [["0.5%"]];
  //     sheet.getRange("F16").values = [["$300/rest."]];
  //     sheet.getRange("F17").values = [["$10/rest."]];
  //     sheet.getRange("F18").values = [["$650/rest."]];
  //     sheet.getRange("F8:F18").format.font.color = "blue";

  //     // Base scenario (Column G)
  //     sheet.getRange("G8").values = [["6"]];
  //     sheet.getRange("G9").values = [["$1,500"]];
  //     sheet.getRange("G10").values = [["0.0%"]];
  //     sheet.getRange("G11").values = [["70.0%"]];
  //     sheet.getRange("G13").values = [["$75/rest."]];
  //     sheet.getRange("G14").values = [["$500/own."]];
  //     sheet.getRange("G15").values = [["0.5%"]];
  //     sheet.getRange("G16").values = [["$300/rest."]];
  //     sheet.getRange("G17").values = [["$10/rest."]];
  //     sheet.getRange("G18").values = [["$650/rest."]];
  //     sheet.getRange("G8:G18").format.font.color = "blue";

  //     // Upside scenario (Column H)
  //     sheet.getRange("H8").values = [["10"]];
  //     sheet.getRange("H9").values = [["$1,500"]];
  //     sheet.getRange("H10").values = [["0.0%"]];
  //     sheet.getRange("H11").values = [["70.0%"]];
  //     sheet.getRange("H13").values = [["$75/rest."]];
  //     sheet.getRange("H14").values = [["$500/own."]];
  //     sheet.getRange("H15").values = [["0.5%"]];
  //     sheet.getRange("H16").values = [["$300/rest."]];
  //     sheet.getRange("H17").values = [["$10/rest."]];
  //     sheet.getRange("H18").values = [["$650/rest."]];
  //     sheet.getRange("H8:H18").format.font.color = "blue";

  //     // Header for scenario columns
  //     sheet.getRange("F7").values = [["Downside"]];
  //     sheet.getRange("F7").format.fill.color = "#FF6600"; // Orange color
  //     sheet.getRange("F7").format.font.bold = true;

  //     sheet.getRange("G7").values = [["Base"]];
  //     sheet.getRange("G7").format.fill.color = "#3366FF"; // Blue color
  //     sheet.getRange("G7").format.font.bold = true;

  //     sheet.getRange("H7").values = [["Upside"]];
  //     sheet.getRange("H7").format.fill.color = "#33CC33"; // Green color
  //     sheet.getRange("H7").format.font.bold = true;

  //     // Operating Assumptions headers and values
  //     sheet.getRange("B24").values = [["Entry Multiple"]];
  //     sheet.getRange("B25").values = [["Exit Multiple"]];
  //     sheet.getRange("B26").values = [["Financing Fees"]];
  //     sheet.getRange("B27").values = [["Additional Transaction Fees"]];
  //     sheet.getRange("B28").values = [["Sponsor Ownership"]];
  //     sheet.getRange("B29").values = [["Entry Rollover"]];
  //     sheet.getRange("B30").values = [["Exit Rollover"]];
  //     sheet.getRange("B31").values = [["Management Options"]];

  //     sheet.getRange("C24").values = [["7.0x"]];
  //     sheet.getRange("C25").values = [["7.0x"]];
  //     sheet.getRange("C26").values = [["$3,000.0"]];
  //     sheet.getRange("C27").values = [["$2,000.0"]];
  //     sheet.getRange("C28").values = [["1.5x"]];
  //     sheet.getRange("C29").values = [["0.5x"]];
  //     sheet.getRange("C30").values = [["25%"]];
  //     sheet.getRange("C31").values = [["10.0%"]];

  //     // Ongoing Assumptions headers and values
  //     sheet.getRange("F24").values = [["Tax Rate"]];
  //     sheet.getRange("F25").values = [["Min Cash"]];
  //     sheet.getRange("F26").values = [["Depreciation"]];
  //     sheet.getRange("F27").values = [["Starting restaurants"]];
  //     sheet.getRange("F28").values = [["Opex"]];

  //     sheet.getRange("G24").values = [["40.0%"]];
  //     sheet.getRange("G25").values = [["$2,000.0"]];
  //     sheet.getRange("G26").values = [["1.0% of sales"]];
  //     sheet.getRange("G27").values = [["100"]];
  //     sheet.getRange("G28").values = [["$70,000.0"]];

  //     // Background color and section formatting
  //     sheet.getRange("B6:H18").format.fill.color = "#E6F0FF"; // Light blue background for Operating section
  //     sheet.getRange("B23").values = [["x Operating Assumptions"]];
  //     sheet.getRange("B23").format.font.bold = true;
  //     sheet.getRange("F23").values = [["Ongoing Assumptions"]];
  //     sheet.getRange("F23").format.font.bold = true;

  //      // Entry Assumptions / Purchase Accounting section
  //   sheet.getRange("B34").values = [["Entry Assumptions / Purchase Accounting"]];
  //   sheet.getRange("B34").format.font.bold = true;

  //   sheet.getRange("B36").values = [["LTM EBITDA"]];
  //   sheet.getRange("C36").values = [["$22,500"]];
  //   sheet.getRange("B37").values = [["(x) Entry Multiple"]];
  //   sheet.getRange("C37").values = [["7.0x"]];
  //   sheet.getRange("B38").values = [["Total Enterprise Value"]];
  //   sheet.getRange("C38").values = [["$157,500"]];
  //   sheet.getRange("B39").values = [["Equity Value"]];
  //   sheet.getRange("C39").values = [["$157,500"]];
  //   sheet.getRange("B40").values = [["Equity Purchase Price"]];
  //   sheet.getRange("C40").values = [["$157,500"]];
  //   sheet.getRange("B41").values = [["(-) Tangible Book Value"]];
  //   sheet.getRange("C41").values = [["($9,230)"]];
  //   sheet.getRange("B42").values = [["New Goodwill Created"]];
  //   sheet.getRange("C42").values = [["$148,270"]];

  //   // Financing section
  //   sheet.getRange("B45").values = [["Financing"]];
  //   sheet.getRange("B45").format.font.bold = true;

  //   const financingHeaders = [
  //       ["Type of Debt", "(x) EBITDA", "Value", "Cash Interest", "Non-Cash Interest", "Amortization", "Tenor"]
  //   ];
  //   const financingData = [
  //       ["RCF", "1.0x", "$22,500", "L + 400", "", "", "5 years"],
  //       ["Term Loan A", "2.0x", "$45,000", "L + 700", "2%", "1%", "5 years"],
  //       ["Seller Note", "0.5x", "$11,250", "", "", "", ""],
  //       ["Total", "3.5x", "$78,750", "", "", "", ""]
  //   ];
  //   sheet.getRange("B47:H47").values = financingHeaders;
  //   sheet.getRange("B48:H51").values = financingData;

  //   // LIBOR rates section
  //   sheet.getRange("B53").values = [["LIBOR"]];
  //   sheet.getRange("C54:H54").values = [["1.00%", "1.25%", "1.50%", "1.75%", "2.00%", "2.25%"]];

  //   // Sources and Uses section
  //   sheet.getRange("B56").values = [["Sources and Uses"]];
  //   sheet.getRange("B56").format.font.bold = true;
  //   sheet.getRange("B58").values = [["Sources"]];
  //   sheet.getRange("E58").values = [["Uses"]];

  //   const sourcesUsesData = [
  //       ["Excess Cash", "$0", "Equity Purchase Price", "$157,500"],
  //       ["RCF", "$29,500", "Financing Fees", "$3,000"],
  //       ["Term Loan A", "$45,000", "Additional Expenses", "$2,000"],
  //       ["Seller Note", "$11,250", "Min Cash", "$2,000"],
  //       ["Sponsor Equity", "$33,750", "", ""],
  //       ["Rollover Equity", "$11,250", "", ""]
  //   ];
  //   sheet.getRange("B59:C64").values = sourcesUsesData.map(row => row.slice(0, 2));
  //   sheet.getRange("E59:F64").values = sourcesUsesData.map(row => row.slice(2));

  //   sheet.getRange("B65").values = [["Total"]];
  //   sheet.getRange("C65").values = [["$130,750"]];
  //   sheet.getRange("F65").values = [["$164,500"]];

  //   // Headers for Income Statement
  //   sheet.getRange("B67").values = [["Income Statement"]];
  //   sheet.getRange("B67").format.font.bold = true;

  //   // Years Headers
  //   sheet.getRange("C68:H68").values = [["2011A", "2012E", "2013E", "2014E", "2015E", "2016E"]];
  //   sheet.getRange("C68:H68").format.font.bold = true;

  //   // Income Statement Data
  //   const incomeStatementRows = [
  //       ["Number of Restaurants", "=", "110", "120", "130", "140", "150"],
  //       ["(+) New Restaurants", "10", "10", "10", "10", "10", "10"],
  //       ["Total restaurants", "100", "110", "120", "130", "140", "150"],
  //       ["AUV", "1,500", "1,500", "1,500", "1,500", "1,500", "1,500"],
  //       ["% Growth", "=$E$10", "=$E$10","=$E$10","=$E$10","=$E$10","=$E$10"],
  //       ["Revenue", "=C73*C74", "=D73*D74", "=E73*E74", "=F73*F74", "=G73*G74", "=H73*H74"],
  //       ["(-) Cost of Goods Sold", "=C75*0.3", "=D75*0.3", "=E75*0.3", "=F75*0.3", "=G75*0.3", "=H75*0.3"],
  //       ["Gross Profit", "=C74-C75", "=D74-D75", "=E74-E75", "=F74-F75", "=G74-G75", "=H74-H75"],
  //       ["(% Margin)", "=C76/C74", "=D76/D74", "=E76/E74", "=F76/F74", "=G76/G74", "=H76/H74"],
  //       ["(-) Operating Expense", "=C77*0.5", "=D77*0.5", "=E77*0.5", "=F77*0.5", "=G77*0.5", "=H77*0.5"],
  //       ["(-) Management Fee", "=C78*0.05", "=D78*0.05", "=E78*0.05", "=F78*0.05", "=G78*0.05", "=H78*0.05"],
  //       ["(-) Annual Rent Expense", "=C79*0.1", "=D79*0.1", "=E79*0.1", "=F79*0.1", "=G79*0.1", "=H79*0.1"],
  //       ["Total Opex", "=SUM(C78:C80)", "=SUM(D78:D80)", "=SUM(E78:E80)", "=SUM(F78:F80)", "=SUM(G78:G80)", "=SUM(H78:H80)"],
  //       ["EBITDA", "=C74-C81", "=D74-D81", "=E74-E81", "=F74-F81", "=G74-G81", "=H74-H81"],
  //       ["(-) Depreciation", "1,500", "1,650", "1,800", "1,950", "2,100", "2,250"],
  //       ["(-) Financing Fee Amort", "600", "600", "600", "600", "600", "600"],
  //       ["EBIT", "=C82-C83-C84", "=D82-D83-D84", "=E82-E83-E84", "=F82-F83-F84", "=G82-G83-G84", "=H82-H83-H84"],
  //       ["(-) Net Interest Expense", "7,500", "8,250", "9,000", "9,750", "10,500", "11,250"],
  //       ["(-) Taxes", "=C85*0.4", "=D85*0.4", "=E85*0.4", "=F85*0.4", "=G85*0.4", "=H85*0.4"],
  //       ["Net Income", "=C85-C86", "=D85-D86", "=E85-E86", "=F85-F86", "=G85-G86", "=H85-H86"]
  //   ];

  //   // Insert Data and Formulas Row by Row
  //   for (let i = 0; i < incomeStatementRows.length; i++) {
  //       const rowData = incomeStatementRows[i];
  //       const range = sheet.getRange(`B${69 + i}:H${69 + i}`);
  //       range.formulas = [rowData];
  //   }

  //   //  // Income Statement section header
  //   //  sheet.getRange("B67").values = [["Income Statement"]];
  //   //  sheet.getRange("B67").format.font.bold = true;
 
  //   //  // Income Statement headers
  //   //  sheet.getRange("C68:H68").values = [["2011A", "2012E", "2013E", "2014E", "2015E", "2016E"]];
 
  //   //  // Income Statement data rows
  //   //  sheet.getRange("B69:H69").values = [["Number of Restaurants", 100, 110, 120, 130, 140, 150]];
  //   //  sheet.getRange("B70:H70").values = [["(+) New Restaurants", 10, 10, 10, 10, 10, 10]];
  //   //  sheet.getRange("B71:H71").values = [["Total restaurants", 100, 110, 120, 130, 140, 150]];
  //   //  sheet.getRange("B72:H72").values = [["AUV", 1500, 1500, 1500, 1500, 1500, 1500]];
  //   //  sheet.getRange("B73:H73").values = [["% Growth", "", "", "", "", "", ""]];
  //   //  sheet.getRange("B74:H74").values = [["Revenue", 150000, 165000, 180000, 195000, 210000, 225000]];
  //   //  sheet.getRange("B75:H75").values = [["(-) Cost of Goods Sold", -45000, -49500, -54000, -58500, -63000, -67500]];
  //   //  sheet.getRange("B76:H76").values = [["Gross Profit", 105000, 115500, 126000, 136500, 147000, 157500]];
  //   //  sheet.getRange("B77:H77").values = [["(% Margin)", "70%", "70%", "70%", "70%", "70%", "70%"]];
  //   //  sheet.getRange("B78:H78").values = [["(-) Operating Expense", -70000, -76500, -83000, -89500, -96000, -102500]];
  //   //  sheet.getRange("B79:H79").values = [["(-) Management Fee", -5000, -5300, -5600, -5900, -6200, -6500]];
  //   //  sheet.getRange("B80:H80").values = [["(-) Annual Rent Expense", -7500, -8250, -9000, -9750, -10500, -11250]];
  //   //  sheet.getRange("B81:H81").values = [["Total Opex", -82500, -90050, -97600, -105150, -112700, -120250]];
  //   //  sheet.getRange("B82:H82").values = [["EBITDA", 22500, 25450, 28400, 31350, 34300, 37250]];
  //   //  sheet.getRange("B83:H83").values = [["(-) Depreciation", -1500, -1650, -1800, -1950, -2100, -2250]];
  //   //  sheet.getRange("B84:H84").values = [["(-) Financing Fee Amort", -600, -600, -600, -600, -600, -600]];
  //   //  sheet.getRange("B85:H85").values = [["EBIT", 20400, 23200, 26000, 28800, 31600, 35000]];
  //   //  sheet.getRange("B86:H86").values = [["(-) Net Interest Expense", "#REF!", "#REF!", "#REF!", "#REF!", "#REF!", "#REF!"]];
  //   //  sheet.getRange("B87:H87").values = [["(-) Taxes", -8160, -9280, "#REF!", "#REF!", "#REF!", "#REF!"]];
  //   //  sheet.getRange("B88:H88").values = [["Net Income", 12240, 13920, "#REF!", "#REF!", "#REF!", "#REF!"]];
 
  //   //  // Apply column width to fit content
  //   //  sheet.getRange("B67:H88").format.autofitColumns();

  //      // Closing Balance Sheet section header
  //   sheet.getRange("B92").values = [["Closing Balance Sheet"]];
  //   sheet.getRange("B92").format.font.bold = true;

  //   // Balance Sheet sub-header
  //   sheet.getRange("B95").values = [["Balance Sheet"]];
  //   sheet.getRange("B95").format.font.bold = true;

  //   // Column headers for Balance Sheet data
  //   sheet.getRange("C95:F95").values = [["2011A", "Credit", "Debit", "2011A"]];
  //   sheet.getRange("C95:F95").format.font.bold = true;

  //   // Balance Sheet data rows
  //   sheet.getRange("B96").values = [["Cash"]];
  //   sheet.getRange("C96").values = [["1,000"]];
  //   sheet.getRange("F96").values = [["1,000"]];

  //   sheet.getRange("B97").values = [["AR"]];
  //   sheet.getRange("C97").values = [["4,500"]];
  //   sheet.getRange("F97").values = [["4,500"]];

  //   sheet.getRange("B98").values = [["Inventory"]];
  //   sheet.getRange("C98").values = [["500"]];
  //   sheet.getRange("F98").values = [["500"]];

  //   sheet.getRange("B99").values = [["Other Current Assets"]];
  //   sheet.getRange("C99").values = [["2,250"]];
  //   sheet.getRange("F99").values = [["2,250"]];

  //   sheet.getRange("B100").values = [["Total Current Assets"]];
  //   sheet.getRange("C100").values = [["8,250"]];
  //   sheet.getRange("F100").values = [["7,250"]];
  //   sheet.getRange("B100:F100").format.font.bold = true;

  //   sheet.getRange("B101").values = [["PP&E"]];
  //   sheet.getRange("C101").values = [["10,500"]];
  //   sheet.getRange("F101").values = [["10,500"]];

  //   sheet.getRange("B102").values = [["Goodwill"]];
  //   sheet.getRange("C102").values = [["70,770"]];
  //   sheet.getRange("D102").values = [["148,270"]];
  //   sheet.getRange("E102").values = [["(70,770)"]];
  //   sheet.getRange("F102").values = [["148,270"]];

  //   sheet.getRange("B103").values = [["Total Assets"]];
  //   sheet.getRange("C103").values = [["89,520"]];
  //   sheet.getRange("F103").values = [["158,770"]];
  //   sheet.getRange("B103:F103").format.font.bold = true;

  //   sheet.getRange("B105").values = [["AP"]];
  //   sheet.getRange("C105").values = [["4,500"]];
  //   sheet.getRange("F105").values = [["4,500"]];

  //   sheet.getRange("B106").values = [["Other Current Liabilities"]];
  //   sheet.getRange("C106").values = [["4,520"]];
  //   sheet.getRange("F106").values = [["4,520"]];

  //   sheet.getRange("B107").values = [["Total Current Liabilities"]];
  //   sheet.getRange("C107").values = [["9,020"]];
  //   sheet.getRange("F107").values = [["9,020"]];
  //   sheet.getRange("B107:F107").format.font.bold = true;

  //   sheet.getRange("B109").values = [["Capitalized FF"]];
  //   sheet.getRange("C109:E109").values = [["NA", "NA", "109"]];

  //   sheet.getRange("B110").values = [["RCF"]];
  //   sheet.getRange("D110").values = [["22,500"]];
  //   sheet.getRange("F110").values = [["22,500"]];

  //   sheet.getRange("B111").values = [["Loan A"]];
  //   sheet.getRange("D111").values = [["45,000"]];
  //   sheet.getRange("F111").values = [["45,000"]];

  //   sheet.getRange("B112").values = [["Seller Note"]];
  //   sheet.getRange("D112").values = [["11,250"]];
  //   sheet.getRange("F112").values = [["11,250"]];

  //   sheet.getRange("B113").values = [["Long Term Liabilities"]];
  //   sheet.getRange("C113").values = [["500"]];
  //   sheet.getRange("E113").values = [["(500)"]];
  //   sheet.getRange("F113").values = [["–"]];

  //   sheet.getRange("B114").values = [["Total Liabilities"]];
  //   sheet.getRange("C114").values = [["9,520"]];
  //   sheet.getRange("F114").values = [["90,770"]];
  //   sheet.getRange("B114:F114").format.font.bold = true;

  //   sheet.getRange("B116").values = [["Shareholders' Equity"]];
  //   sheet.getRange("C116").values = [["80,000"]];
  //   sheet.getRange("F116").values = [["68,000"]];
  //   sheet.getRange("B116:F116").format.font.bold = true;

  //   sheet.getRange("B118").values = [["Total Liabilities & Equity"]];
  //   sheet.getRange("C118").values = [["89,520"]];
  //   sheet.getRange("F118").values = [["158,770"]];
  //   sheet.getRange("B118:F118").format.font.bold = true;

  //   // Apply formatting
  //   sheet.getRange("B94:F118").format.horizontalAlignment = "Center";
  //   sheet.getRange("B94:F118").format.autofitColumns();


  //     return context.sync();
  //   }).catch(function (error) {
  //     console.log("Error: " + error);
  //   });
  // };

  // Function to fetch the formula from the active cell
  
  
  const fetchCellFormula = async (): Promise<string | null> => {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.getActiveCell();
        console.log("range", range);
        range.load("formulas");
        await context.sync();

        const cellFormula = range.formulas[0][0];
        return cellFormula || null;
      });
    } catch (error) {
      console.error("Error fetching formula:", error);
      return null;
    }
  };

  const sendToAPI = async (userMessage: string, cellFormula: string): Promise<string> => {
    try {
      const prompt = `You are an intelligent assistant skilled in explaining complex Excel formulas in a way that's easy to understand. 

Here’s a user’s question and a formula. Please thoroughly explain the formula in a step-by-step manner, breaking down each part of it to help the user understand its purpose, logic, and how it’s constructed. Make sure to cover the following points:
1. **Formula Purpose**: Explain what this formula is intended to calculate or accomplish in simple terms.
2. **Components and Functions**: Identify and explain each component of the formula, including any functions, cell references, ranges, and operators. Describe what each one does and how they interact with each other.
3. **Step-by-Step Execution**: Walk through how Excel processes this formula from start to finish, explaining each calculation or operation that Excel performs.
4. **Common Use Cases**: If relevant, mention any common scenarios where this formula might be used, and why it’s useful in those contexts.
5. **Additional Tips**: Offer any helpful tips or insights, such as ways to troubleshoot common errors related to this formula or ways to modify it for similar calculations.

#### Example Input:
- **Formula**: ${cellFormula}
- **User Question**: ${userMessage}

Provide the explanation in simple, conversational language suitable for someone who may be new to Excel formulas. Return normal text format, not markdown.
`;

      const response = await axios.post("https://571d-182-156-1-250.ngrok-free.app/api/anthropic", {
        message: prompt,
      });
      console.error(response);
      // Access the data from the response
      return response.data.content[0].text || "No explanation found";
    } catch (error) {
      console.error("Error calling the local API:", error);
      console.error(error.message);
      console.error(JSON.stringify(error.stack));

      if (axios.isAxiosError(error)) {
        console.error(error.code);
        console.error(error.config);
        if (error.response) {
          console.error(error.response.data);
          console.error(error.response.status);
        }
      } else {
        console.error("Non-Axios error:", error);
      }

      return "Error fetching explanation from the local API.";
    }
  };

  return (
    <div className="chat-container">
      {messages.map((message, index) => (
        <Message key={index} text={message.text} type={message.type as "question" | "explanation"} />
      ))}
      {loading && <Loader />} {/* Display loader if loading is true */}
      {/* <ExplanationBox addMessage={addMessage} /> */}
      <InputBox onSend={(text) => addMessage(text, "question")} />
    </div>
  );
};

export default Chat;
