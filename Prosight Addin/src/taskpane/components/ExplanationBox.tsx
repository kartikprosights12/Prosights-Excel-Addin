import * as React from "react";
import "./App.css";

interface ExplanationBoxProps {
  addMessage: (text: string, type: "question" | "explanation") => void;
}

const ExplanationBox: React.FC<ExplanationBoxProps> = ({ addMessage }) => {
  return (
    <div className="explanation-box">
      <h3>Explanation</h3>
      <p>Sure, revenue figures are located in row 4 given the “Revenue” label in cell B4...</p>
      <div className="button-group">
        <button onClick={() => addMessage("Ask clicked", "question")}>Ask</button>
        <button onClick={() => navigator.clipboard.writeText("Explanation text")}>Copy</button>
        <button onClick={() => addMessage("Apply clicked", "question")}>Apply</button>
      </div>
    </div>
  );
};

export default ExplanationBox;
