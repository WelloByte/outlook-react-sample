/* global Office console */
import { Client } from "@microsoft/microsoft-graph-client";
// install by executing: npm install @microsoft/microsoft-graph-client

import { AccountManager } from "../authConfig";
//import { makeGraphRequest } from "./msgraph-helper";
import { parseReply } from "parse-reply";

export async function insertText() {
  // Write text to the cursor point in the compose surface.
  // try {
  //   Office.context.mailbox.item?.body.setSelectedDataAsync(
  //     text,
  //     { coercionType: Office.CoercionType.Text },
  //     (asyncResult: Office.AsyncResult<void>) => {
  //       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //         throw asyncResult.error.message;
  //       }
  //     }
  //   );
  // } catch (error) {
  //   console.log("Error: " + error);
  // }

  /* eslint-disable no-useless-escape */
  /*
   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
   * See LICENSE in the project root for license information.
   */

  /* global Office, console, document, HTMLInputElement, setTimeout */

  //import "isomorphic-fetch"; // or import the fetch polyfill you installed

  //var EmailReplyParser = require("email-reply-parser");
  //import EmailReplyParser from "email-reply-parser";
  //import linkify from "linkify-it";

  const accountManager = new AccountManager();

  // This sample shows how to register an event handler in Outlook.
  Office.onReady(() => {
    if (Office.context.mailbox) {
      // Registers an event handler to identify when the user changes the selection in the message list.
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        async () => {
          await getMessages();
        },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
          console.log("Event handler added for the SelectedItemsChanged event.");
        }
      );
    }

    // Initialize MSAL.
    accountManager
      .initialize()
      .then(() => {
        console.log("MSAL initialized.");

        // Identify when the user refreshes the message list.
        // use Office.context.mailbox.item for the current item then get the conversationId and GetConversations
        if (Office.context.mailbox.item) {
          console.log("The pane was refreshed");
          getMessages();
        }
      })
      .catch((error) => {
        console.error("Error initializing MSAL:", error);
      });
  });

  async function getMessages() {
    // Retrieves the selected messages' properties and logs them to the console.
    // Permission to read selected items. need to change permission level on manifest file.
    Office.context.mailbox.getSelectedItemsAsync(async (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      const conversationId = asyncResult.value[0].conversationId;

      asyncResult.value.forEach((message) => {
        console.log(`Item ID: ${message.itemId}`);
        console.log(`Subject: ${message.subject}`);
        console.log(`Item type: ${message.itemType}`);
        console.log(`conversation ID: ${conversationId}`);
        console.dir(message, "<<");
      });

      // Clear existing chat bubbles
      clearChatBubbles();

      const profile = Office.context.mailbox.userProfile;

      // Call fnGetConversations and get all emails with the same conversation ID from Microsoft Graph
      const messages = await fnGetConversations(conversationId);
      console.log(messages, "ini messages");
      let lastDate = null;
      messages.forEach((msg) => {
        const messageDate = new Date(msg.createdDateTime).toLocaleDateString();
        if (messageDate !== lastDate) {
          appendDateSeparator(messageDate);
          lastDate = messageDate;
        }
        if (msg.sender.emailAddress.address === profile.emailAddress) {
          appendChatBubble(
            msg.uniqueBody,
            "user",
            null,
            null,
            new Date(msg.createdDateTime).toLocaleTimeString(),
            msg.attachments
          );
          return;
        }
        appendChatBubble(
          msg.uniqueBody,
          "bot",
          msg.sender.emailAddress.name,
          msg.sender.emailAddress.address,
          new Date(msg.createdDateTime).toLocaleTimeString(),
          msg.attachments
        );

        // Update the H1 element with the sender name
        document.getElementById("msg-sender").textContent = msg.sender.emailAddress.name;
      });
    });
  }

  function clearChatBubbles() {
    const chatContainer = document.getElementById("chat-container");
    while (chatContainer.firstChild) {
      chatContainer.removeChild(chatContainer.firstChild);
    }
  }

  document.getElementById("send-button").addEventListener("click", function () {
    const input = document.getElementById("chat-input") as HTMLInputElement;
    const message = input.value.trim();
    if (message) {
      appendChatBubble(message, "user");
      input.value = "";
      // Simulate bot response
      setTimeout(() => {
        appendChatBubble("This is a bot response.", "bot");
      }, 1000);
    }
  });

  function formatMessage(message) {
    // Remove newlines from the start and end of the message
    message = message.trim();

    // Replace sequences of whitespace (spaces, tabs, zero-width spaces) and newlines with a single space
    // message = message.replace(/[\s\u200B]+/g, " ");

    // Replace more than two consecutive newlines with two newlines
    message = message.replace(/(\n\s*){3,}/g, "\n\n");

    return message;
  }

  function appendChatBubble(
    message,
    senderType,
    senderName = "No Name",
    senderEmail = "No Email Address",
    timestamp = new Date().toLocaleTimeString(),
    attachments = []
  ) {
    const chatContainer = document.getElementById("chat-container");
    const bubble = document.createElement("div");
    bubble.className = `chat-bubble ${senderType}`;

    const header = document.createElement("div");
    header.className = "sender-info";

    if (senderName) {
      const senderNameElem = document.createElement("div");
      senderNameElem.className = "sender-name";
      senderNameElem.textContent = senderName;
      senderNameElem.title = senderEmail; // Tooltip with email address

      header.appendChild(senderNameElem);
    }

    const messageContent = document.createElement("div");
    messageContent.className = "message-content";

    // format message
    const formattedMessage = formatMessage(message);
    // Detect URLs, including those encapsulated in brackets or angle brackets
    const urlRegex =
      /(?:https?:\/\/[^\s<>\[\]]+)|(?:\[[^\]]+https?:\/\/[^\s<>\[\]]+\])|(?:<[^\>]+https?:\/\/[^\s<>\[\]]+>)/g;
    const linkedMessage = formattedMessage.replace(urlRegex, (url) => {
      // Remove brackets or angle brackets if present
      const cleanUrl = url.replace(/[\[\]<>]/g, "");
      return `<a href="${cleanUrl}" target="_blank" rel="noopener noreferrer">link</a>`;
    });

    messageContent.innerHTML = linkedMessage;

    const timestampElem = document.createElement("div");
    timestampElem.className = "timestamp";
    timestampElem.textContent = timestamp;

    bubble.appendChild(header);
    bubble.appendChild(messageContent);
    bubble.appendChild(timestampElem);

    // Append attachments if any
    if (attachments.length > 0) {
      const attachmentsContainer = document.createElement("div");
      attachmentsContainer.className = "attachments-container";

      attachments.forEach((attachment) => {
        const attachmentElem = document.createElement("a");
        attachmentElem.className = "attachment";
        attachmentElem.textContent = attachment.name;
        attachmentElem.href = attachment.contentUrl;
        attachmentElem.target = "_blank"; // Open in a new tab
        attachmentElem.rel = "noopener noreferrer"; // Security measure
        attachmentsContainer.appendChild(attachmentElem);
      });

      bubble.appendChild(attachmentsContainer);
    }

    chatContainer.appendChild(bubble);
  }

  function appendDateSeparator(date) {
    const chatContainer = document.getElementById("chat-container");
    const separator = document.createElement("div");
    separator.className = "date-separator";
    separator.textContent = date;
    chatContainer.appendChild(separator);
  }

  /**
   * Gets all emails based from conversationid
   */
  /**
   * Gets all emails based from conversationId
   */
  async function fnGetConversations(conversationId, count = 10) {
    const accessToken = await accountManager.ssoGetAccessToken(["Mail.Read"]);
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    const response = await client
      .api("/me/messages")
      .header("Prefer", 'outlook.body-content-type="text"')
      .filter(`conversationId eq '${conversationId}'`)
      .select("subject, body, uniqueBody, sender, createdDateTime, internetMessageHeaders, hasAttachments")
      .top(count)
      .expand("attachments")
      .get();

    console.log(response, "ini response");

    const messages = response.value.map((item) => ({
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
        contentUrl: attachment.contentUrl,
        contentType: attachment.contentType,
      })),
    }));

    messages.forEach((msg) => {
      // Optional: log X-Mailer
      if (msg.internetMessageHeaders) {
        const xMailerHeader = msg.internetMessageHeaders.find((header) => header.name === "X-Mailer");
        console.log(xMailerHeader);
      } else {
        console.log("No internetMessageHeaders");
      }

      // ✅ Use parseReply correctly
      const visibleText = parseReply(msg.body);
      console.log("Visible Text:", visibleText);

      // 🟡 Optional: replace `msg.uniqueBody` with visibleText
      appendChatBubble(
        visibleText,
        msg.sender.emailAddress.address === Office.context.mailbox.userProfile.emailAddress ? "user" : "bot",
        msg.sender.emailAddress.name,
        msg.sender.emailAddress.address,
        new Date(msg.createdDateTime).toLocaleTimeString(),
        msg.attachments
      );
    });
    console.log(messages, "<<<");
    return messages;
  }
}
