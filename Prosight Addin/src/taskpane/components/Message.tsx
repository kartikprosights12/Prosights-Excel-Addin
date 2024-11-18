import * as React from "react";
import "./App.css";

interface MessageProps {
  text: string;
  type: "question" | "explanation";
}

const Message: React.FC<MessageProps> = ({ text, type }) => {
  return (
    <div className={`message ${type === "question" ? "question" : "explanation"}`}>
      <p>{text}</p>
    </div>
  );
};

export default Message;
