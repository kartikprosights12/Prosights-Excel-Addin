import * as React from "react";
import Chat from "./Chat";

interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = () => {
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  return (
    <div className="app-container">
      <header className="header">
        <h2>ProSights Excel Assistant</h2>
      </header>
      <Chat />
    </div>
  );
};

export default App;
