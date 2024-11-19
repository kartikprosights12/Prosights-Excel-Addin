import React, { useState } from "react";
import ChatInput from "./ChatInput";
import ChatMessages, { ChatMessageType } from "./ChatMessages";
import "./App.css";
import { formatExcelSheet } from "./use-cases/Format";
import { loadDataFunction } from "./use-cases/Load";
import { fixSheet } from "./use-cases/Fix";


const ChatWindow: React.FC = () => {
  const [chatMessages, setChatMessages] = useState([]);
  // const [loading, setLoading] = useState(false); // New loading state

  // Function to add messages, fetch formula, and call the API if needed
  const handleChatSend = async (text: ChatMessageType) => {
     // setLoading(true); // Show loader
      // Fetch the formula from the active cell when user submits a message
      let responseMessage = "Try again!"; // Default response


      //load data case
      if (text.inputMessage.toLowerCase().includes("build") && text.inputMessage.toLowerCase().includes("lbo")) {
        await Excel.run(async (context) => {
          try {
              await loadDataFunction(context);
              responseMessage = "LBO built successfully!";
              console.log("Excel sheet formatted successfully!");
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
          }
      });
      }  

      //format data case
      else if (text.inputMessage.toLowerCase().includes("format")) {
        await Excel.run(async (context) => {
          try {
              await formatExcelSheet(context);
              console.log("Excel sheet formatted successfully!");
              responseMessage = "LBO built successfully!";
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
              responseMessage = "Try again! Error formatting data";
          }
      });
      }  
      
      //borken and fix case
      else if (text.inputMessage.toLowerCase().includes("broken") && text.inputMessage.toLowerCase().includes("fix")) {
        await Excel.run(async (context) => {
          try {
              await fixSheet(context);
              console.log("Excel sheet formatted successfully!");
              responseMessage = "Fixed successfully!";
          } catch (error) {
              console.log(error);
              console.error("Error formatting Excel sheet:", error);
              responseMessage = "Try again! Error fixing the issue";
          }
      });
      } 

       //calculate sheet
        else if (text.inputMessage.toLowerCase().includes("calculate")) {
          await Excel.run(async (context) => {
            try {
              // Get the active worksheet
              const sheet =context.workbook.worksheets.getActiveWorksheet();

              // Recalculate the current worksheet
              sheet.calculate(true);
                console.log("calculation!");
            } catch (error) {
                console.log(error);
                console.error("Error formatting Excel sheet:", error);
            }
        });
        } 

        setChatMessages((prevMessages) => [
          ...prevMessages,
          { ...text, responseMessage },
        ]);
  };


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

//   const sendToAPI = async (userMessage: string, cellFormula: string): Promise<string> => {
//     try {
//       const prompt = `You are an intelligent assistant skilled in explaining complex Excel formulas in a way that's easy to understand. 

// Here’s a user’s question and a formula. Please thoroughly explain the formula in a step-by-step manner, breaking down each part of it to help the user understand its purpose, logic, and how it’s constructed. Make sure to cover the following points:
// 1. **Formula Purpose**: Explain what this formula is intended to calculate or accomplish in simple terms.
// 2. **Components and Functions**: Identify and explain each component of the formula, including any functions, cell references, ranges, and operators. Describe what each one does and how they interact with each other.
// 3. **Step-by-Step Execution**: Walk through how Excel processes this formula from start to finish, explaining each calculation or operation that Excel performs.
// 4. **Common Use Cases**: If relevant, mention any common scenarios where this formula might be used, and why it’s useful in those contexts.
// 5. **Additional Tips**: Offer any helpful tips or insights, such as ways to troubleshoot common errors related to this formula or ways to modify it for similar calculations.

// #### Example Input:
// - **Formula**: ${cellFormula}
// - **User Question**: ${userMessage}

// Provide the explanation in simple, conversational language suitable for someone who may be new to Excel formulas. Return normal text format, not markdown.
// `;

//       const response = await axios.post("https://571d-182-156-1-250.ngrok-free.app/api/anthropic", {
//         message: prompt,
//       });
//       console.error(response);
//       // Access the data from the response
//       return response.data.content[0].text || "No explanation found";
//     } catch (error) {
//       console.error("Error calling the local API:", error);
//       console.error(error.message);
//       console.error(JSON.stringify(error.stack));

//       if (axios.isAxiosError(error)) {
//         console.error(error.code);
//         console.error(error.config);
//         if (error.response) {
//           console.error(error.response.data);
//           console.error(error.response.status);
//         }
//       } else {
//         console.error("Non-Axios error:", error);
//       }

//       return "Error fetching explanation from the local API.";
//     }
//   };
return (
  <div className="w-full flex flex-col flex-1 bg-gray-50 rounded-lg border border-gray-300 justify-between">
   <ChatMessages chatMessages={chatMessages} />
    {/* {loading && <Loader />} Display loader if loading is true */}
    {/* <ExplanationBox addMessage={addMessage} /> */}
    <div className="p-2 h-fit">
      <ChatInput onChatSend={handleChatSend} />
    </div>
  </div>
);
};

export default ChatWindow;
