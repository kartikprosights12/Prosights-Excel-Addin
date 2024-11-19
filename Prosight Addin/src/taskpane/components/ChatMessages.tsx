import { ArrowReplyRegular, DocumentRegular, CopyRegular, PlayRegular } from "@fluentui/react-icons";
import React from "react";
import ReactMarkdown from "react-markdown";

export type ChatMessageType = {
  inputMessage: string;
  cellAddress: string;
  files: File[];
  responseMessage?: string;
};

export default function ChatMessages({ chatMessages }: { chatMessages: ChatMessageType[] }) {
  return (
    <div className="flex-1 flex flex-col gap-4 overflow-y-auto p-2">
      {chatMessages.map((chatMessage, index) => (
        <ChatMessage key={index} chatMessage={chatMessage} />
      ))}
    </div>
  );
}

export function ChatMessage({ chatMessage }: { chatMessage: ChatMessageType }) {
  const cellAddress = chatMessage.cellAddress;
  const files = chatMessage.files;

  return (
    <div className="flex flex-col gap-2">
      <div className="p-2 flex flex-col gap-2 bg-white border border-gray-300 rounded-md">
        {(cellAddress || files?.length > 0) && (
          <div className="flex gap-1 flex-wrap">
            {cellAddress && (
              <div className="flex h-4 gap-2 items-center rounded-sm border border-gray-300 px-1">
                <div className="flex items-center gap-0.5">
                  <DocumentRegular className="h-3 w-3" />
                  <p className="text-xxs">{cellAddress.split("!")[1]}</p>
                </div>
                <p className="text-xxs text-gray-400">{cellAddress.split("!")[0]}</p>
              </div>
            )}
            {files?.length > 0 &&
              files.map((file, index) => (
                <div key={index} className="flex h-4 items-center gap-0.5 rounded-sm border border-gray-300 px-1">
                  <DocumentRegular className="h-3 w-3" />
                  <p className="text-xxs">{file.name}</p>
                </div>
              ))}
          </div>
        )}
        <p className="text-xs"> <ReactMarkdown>{chatMessage.inputMessage}</ReactMarkdown> </p>
      </div>
      {/* Placeholder for AI response */}
      <p className="text-xs p-2">
      {chatMessage.responseMessage || "No response available."}
      </p>
      <div className="flex justify-end gap-2">
        <div className="flex h-4 items-center gap-1 rounded-sm border bg-white border-gray-300 px-1 text-gray-400 hover:text-gray-600 cursor-pointer transition-all duration-150">
          <ArrowReplyRegular className="h-3 w-3" />
          <p className="text-xxs">Ask</p>
        </div>
        <div className="flex h-4 items-center gap-1 rounded-sm border bg-white border-gray-300 px-1 text-gray-400 hover:text-gray-600 cursor-pointer transition-all duration-150">
          <CopyRegular className="h-3 w-3" />
          <p className="text-xxs">Copy</p>
        </div>
        <div className="flex h-4 items-center gap-1 rounded-sm border bg-white border-gray-300 px-1 text-gray-400 hover:text-gray-600 cursor-pointer transition-all duration-150">
          <PlayRegular className="h-3 w-3" />
          <p className="text-xxs">Apply</p>
        </div>
      </div>
    </div>
  );
}