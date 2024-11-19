/* global Excel */

import React, { useState, useEffect } from "react";

import { ChatMessageType } from "./ChatMessages";

import {
  AddRegular,
  DocumentRegular,
  AttachRegular,
  ArrowEnterLeftRegular,
  DismissRegular,
} from "@fluentui/react-icons";

export default function ChatInput({ onChatSend }: { onChatSend: (message: ChatMessageType) => void }) {
  // TODO: simplify range values into a single object once multiple cells are supported
  const [selectedAddress, setSelectedAddress] = useState("");
  const [selectedValue, setSelectedValue] = useState("");
  const [selectedFormula, setSelectedFormula] = useState("");

  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  useEffect(() => {
    // Set up the selection changed event handler
    Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(handleSelectionChange);
      await context.sync();
    });

    // Cleanup function
    return () => {
      Excel.run(async (context) => {
        context.workbook.onSelectionChanged.remove(handleSelectionChange);
        await context.sync();
      });
    };
  }, []);

  const handleSelectionChange = async () => {
    // WARNING (Eoin): For Demo purposes this only handles one cell.
    // This can be easily scaled later.
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "formulas"]);
      await context.sync();

      setSelectedAddress(range.address);
      setSelectedValue(range.values[0][0]);
      setSelectedFormula(range.formulas[0][0]);
    });
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setSelectedFiles((prevFiles) => [...prevFiles, ...Array.from(e.target.files)]);
    }
  };

  return (
    <div className="p-2 flex flex-col gap-1 bg-white border border-gray-300 rounded-md">
      {/* Preview Pane for selected cell */}
      {(selectedFormula || selectedValue) && (
        <div className="bg-gray-100 rounded-sm p-2 border border-gray-300">
          <p className="text-xs">
            {selectedFormula ? (
              <ColouredExcelFormula selectedFormula={selectedFormula} />
            ) : (
              <span className="text-xs">{selectedValue}</span>
            )}
          </p>
        </div>
      )}
      {/* Attachments */}
      <div className="w-full flex justify-start gap-1 flex-wrap">
        <div className="flex items-center justify-center h-4 w-4 rounded-sm border border-gray-300">
          <AddRegular className="h-2 w-2" />
        </div>
        {selectedAddress && (
          <div className="flex h-4 gap-2 items-center rounded-sm border border-gray-300 px-1">
            <div className="flex items-center gap-0.5">
              <DocumentRegular className="h-3 w-3" />
              <p className="text-xxs">{selectedAddress.split("!")[1]}</p>
            </div>
            <p className="text-xxs text-gray-400">{selectedAddress.split("!")[0]}</p>
          </div>
        )}
        {selectedFiles.length > 0 &&
          selectedFiles.map((file, index) => (
            <div key={index} className="flex h-4 items-center gap-0.5 rounded-sm border border-gray-300 px-1">
              <DocumentRegular className="h-3 w-3" />
              <p className="text-xxs">{file.name}</p>
              <DismissRegular
                className="h-2 w-2 ml-1 cursor-pointer text-gray-400 hover:text-gray-600"
                onClick={() => setSelectedFiles((files) => files.filter((_, i) => i !== index))}
              />
            </div>
          ))}
      </div>
      {/* Chat Input */}
      <textarea
        placeholder="Ask a question..."
        className="focus:outline-none resize-none overflow-hidden auto-resize text-xs mt-0.5"
        onChange={(e) => {
          e.target.style.height = "auto";
          e.target.style.height = `${e.target.scrollHeight}px`;
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            const value = (e.target as HTMLTextAreaElement).value.trim();
            if (!value) return;
            onChatSend({
              inputMessage: value,
              cellAddress: selectedAddress,
              files: selectedFiles,
            });
            (e.target as HTMLTextAreaElement).value = "";
            setSelectedFiles([]);
            setSelectedAddress("");
            setSelectedFormula("");
            setSelectedValue("");

          }
        }}
      />
      <div className="flex justify-end gap-2">
        <label className="flex text-gray-500 items-center gap-0.5 cursor-pointer hover:text-gray-700 transition-colors duration-150">
          <input type="file" multiple className="hidden" onChange={handleFileSelect} />
          <AttachRegular className="h-3 w-3" />
          <p className="text-xxs">Attach</p>
        </label>
        <div
          className="flex text-gray-500 items-center gap-0.5 cursor-pointer hover:text-gray-700 transition-colors duration-150"
          onClick={() => {}}
        >
          <ArrowEnterLeftRegular className="h-3 w-3" />
          <p className="text-xxs">Chat</p>
        </div>
      </div>
    </div>
  );
}

export function ColouredExcelFormula({ selectedFormula }: { selectedFormula: string }) {
  return (
    <span className="text-xs">
      {typeof selectedFormula === "string" && selectedFormula.startsWith("=") ? (
        <span>
          <span className="text-green-600">=</span>
          {selectedFormula
            .slice(1)
            .split(/([(),+\-*/])/g)
            .map((part, index) => {
              // Functions (uppercase words followed by parentheses)
              if (/^[A-Z]+$/.test(part)) {
                return (
                  <span key={index} className="text-blue-600">
                    {part}
                  </span>
                );
              }
              // Operators
              if (/[(),+\-*/]/.test(part)) {
                return (
                  <span key={index} className="text-gray-500">
                    {part}
                  </span>
                );
              }
              // Cell references
              if (/^[A-Z]+[0-9]+$/i.test(part)) {
                return (
                  <span key={index} className="text-indigo-600">
                    {part}
                  </span>
                );
              }
              return <span key={index}>{part}</span>;
            })}
        </span>
      ) : (
        selectedFormula
      )}
    </span>
  );
}