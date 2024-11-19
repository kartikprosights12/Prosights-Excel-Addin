import * as React from "react";
import ChatWindow from "./ChatWindow";

interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = () => {
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.

  return (
    <div className="flex flex-col px-2 py-4 min-h-screen gap-2 bg-gray-100">
      <div className="flex items-center justify-start gap-2">
        <img
          src="https://localhost:3000/assets/prosights-logo.png"
          alt="Prosights Logo"
          className="w-6 h-6 transition-all duration-300 hover:scale-110"
        />
        <h1 className="text-md font-medium">Chat with Prosights AI</h1>
      </div>
      <div className="flex-1 flex">
        <ChatWindow />
      </div>
    </div>
  );
};

export default App;
