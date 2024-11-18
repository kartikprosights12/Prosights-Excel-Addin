import React from "react";
import "./loader.css";

const Loader: React.FC = () => {
  return (
    <div className="loader">
      <span className="dot"></span>
      <span className="dot"></span>
      <span className="dot"></span>
    </div>
  );
};

export default Loader;