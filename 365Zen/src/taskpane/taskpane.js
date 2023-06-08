/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("setTextButton").onclick = setText;
    document.getElementById("insertSubjectButton").onclick = insertSubject;
    document.getElementById("insertBodyButton").onclick = insertBody;
  }
});

function setText() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  Office.context.mailbox.item.body.setAsync("This is the body of the email.");
}

function insertSubject() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
}

function insertBody() {
  Office.context.mailbox.item.body.setAsync("This is the body of the email.");
}

















/* global document, Office */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */
// }


