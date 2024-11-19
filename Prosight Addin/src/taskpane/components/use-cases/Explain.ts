import axios from "axios";

export const explainSheet = async (cellFormula: string, userMessage: string) => {

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
  
        const response = await axios.post("https://60cb-49-36-189-20.ngrok-free.app/api/anthropic/upload", {
          message: prompt,
        });
        console.log('content');
        console.log(response.data.content[0].text);
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

}