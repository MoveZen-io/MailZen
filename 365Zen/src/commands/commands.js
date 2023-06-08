function sendTemplate1() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  Office.context.mailbox.item.body.setAsync("This is the body of the email.");
}

function insertSubject() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
}

function insertBody() {
  Office.context.mailbox.item.body.setAsync("This is the body of the email.");
}

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insertTemplate1").onclick = sendTemplate1;
    document.getElementById("insertSubjectButton").onclick = insertSubject;
    document.getElementById("insertBodyButton").onclick = insertBody;
  }
});






// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global global, Office, self, window */

// Office.onReady(() => {
//   // If needed, Office.js is ready to be called
// });

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }

// function getGlobal() {
//   return typeof self !== "undefined"
    // ? self
//     : typeof window !== "undefined"
//     ? window
//     : typeof global !== "undefined"
//     ? global
//     : undefined;
// }

// const g = getGlobal();

// // The add-in command functions need to be available in global scope
// g.action = action;
