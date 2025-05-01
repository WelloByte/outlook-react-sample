// src/components/OutlookMessages.tsx
import React, { useEffect, useState } from "react";
import { AccountManager } from "../../utils/authConfig";
import { Client } from "@microsoft/microsoft-graph-client";
import { parseReply } from "parse-reply";

interface Message {
  id: string;
  body: string;
  visibleText: string;
  senderName: string;
  senderEmail: string;
  isUser: boolean;
  timestamp: string;
  attachments: { name: string; contentUrl: string }[];
  date: string;
}

const OutlookMessages: React.FC = () => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [chatInput, setChatInput] = useState("");

  useEffect(() => {
    Office.onReady(() => {
      if (Office.context.mailbox) {
        getConversationMessages()
          .then(() => console.log("terpanggil"))
          .catch(() => {
            console.log("tidak");
          });
      }
    });
  }, []);

  const getConversationMessages = async () => {
    try {
      console.log();
      const accountManager = new AccountManager();
      await accountManager.initialize();
      const token = await accountManager.ssoGetAccessToken(["Mail.Read"]);

      const selectedItemsResult = await new Promise((resolve, reject) => {
        Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            return reject(asyncResult.error.message);
          }
          resolve(asyncResult.value);
        });
      });

      const conversationId = selectedItemsResult[0].conversationId;

      const client = Client.init({
        authProvider: (done) => done(null, token),
      });

      const response = await client
        .api("/me/messages")
        .header("Prefer", 'outlook.body-content-type="text"')
        .filter(`conversationId eq '${conversationId}'`)
        .select("subject, body, uniqueBody, sender, createdDateTime, internetMessageHeaders, hasAttachments")
        .top(10)
        .expand("attachments")
        .get();

      const userEmail = Office.context.mailbox.userProfile.emailAddress;

      const getMessages = response.value.map(
        (item: {
          body: any;
          uniqueBody: any;
          sender: {
            emailAddress: {
              name: string;
              address: string;
            };
          };
          createdDateTime: string;
          internetMessageHeaders: any;
          hasAttachments: boolean;
          attachments: any[];
        }) => ({
          body: item.body.content,
          uniqueBody: item.uniqueBody.content,
          sender: {
            emailAddress: {
              name: item.sender.emailAddress.name,
              address: item.sender.emailAddress.address,
            },
          },
          createdDateTime: item.createdDateTime,
          internetMessageHeaders: item.internetMessageHeaders,
          hasAttachments: item.hasAttachments,
          attachments: item.attachments.map((attachment) => ({
            name: attachment.name,
            contentUrl: attachment.contentUrl, // if the attachment object has a contentUrl property
            contentType: attachment.contentType, // if the attachment object has a contentType property
          })),
        })
      );

      console.log(getMessages, "<<<<");
      setMessages(getMessages);
    } catch (error) {
      console.error("Error fetching messages:", error);
    }
  };

  const renderDateSeparators = (messages: Message[]) => {
    let lastDate = "";
    return messages.map((msg) => {
      const showDate = msg.date !== lastDate;
      lastDate = msg.date;
      console.log(msg, "<<<<message");
      return (
        <React.Fragment key={msg.id}>
          {showDate && <div className="date-separator text-center text-sm text-gray-500 my-2">{msg.date}</div>}
          <div
            className={`chat-bubble max-w-[75%] p-3 rounded-lg shadow mb-2 ${
              msg.isUser ? "bg-blue-500 text-white self-end ml-auto" : "bg-gray-100 text-gray-900 self-start mr-auto"
            }`}
          >
            <div className="sender-info" title={msg.senderEmail}>
              {msg.senderName}
            </div>
            <div>{msg.body}</div>
            {/* <div className="message-content" dangerouslySetInnerHTML={{ __html: msg.visibleText }} > */}
            <div className="timestamp">{msg.timestamp}</div>
            {msg.attachments.length > 0 && (
              <div className="attachments-container">
                {msg.attachments.map((att) => (
                  <a
                    key={att.name}
                    className="attachment"
                    href={att.name}
                    // target="_blank"
                    rel="noopener noreferrer"
                  >
                    {att.name}
                  </a>
                ))}
              </div>
            )}
          </div>
        </React.Fragment>
      );
    });
  };

  // const linkify = (text: string) => {
  //   const urlRegex =
  //     /(?:https?:\/\/[^\s<>\[\]]+)|(?:\[[^\]]+https?:\/\/[^\s<>\[\]]+\])|(?:<[^\>]+https?:\/\/[^\s<>\[\]]+>)/g;
  //   return text.replace(urlRegex, (url) => {
  //     const cleanUrl = url.replace(/[\[\]<>]/g, "");
  //     return `<a href="${cleanUrl}" target="_blank" rel="noopener noreferrer">link</a>`;
  //   });
  // };

  const handleSend = () => {
    if (!chatInput.trim()) return;
    const newMsg: Message = {
      id: Date.now().toString(),
      body: chatInput,
      visibleText: chatInput,
      senderName: "You",
      senderEmail: "",
      isUser: true,
      timestamp: new Date().toLocaleTimeString(),
      attachments: [],
      date: new Date().toLocaleDateString(),
    };
    setMessages((prev) => [...prev, newMsg]);
    setChatInput("");

    // Simulate bot response
    setTimeout(() => {
      const botReply: Message = {
        ...newMsg,
        id: Date.now().toString() + "-bot",
        isUser: false,
        senderName: "Bot",
        senderEmail: "bot@example.com",
        visibleText: "This is a bot response.",
      };
      setMessages((prev) => [...prev, botReply]);
    }, 1000);
  };

  return (
    <div id="chat-wrapper">
      <h1>Conversation</h1>
      <div id="chat-container" className="flex flex-col gap-2 overflow-y-auto h-[70vh] p-4">
        {renderDateSeparators(messages)}
      </div>
      <div className="chat-input-area">
        <input
          type="text"
          placeholder="Type a message..."
          value={chatInput}
          onChange={(e) => setChatInput(e.target.value)}
          className="flex-1 border rounded-full px-4 py-2 focus:outline-none"
        />
        <button onClick={handleSend} className="bg-blue-500 text-white px-4 py-2 rounded-full hover:bg-blue-600">
          Send
        </button>
      </div>
    </div>
  );
};

export default OutlookMessages;
