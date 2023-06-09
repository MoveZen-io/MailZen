
function prependHtmlBody(bodyVar) {
  // Prepare your HTML content
  var htmlContent = bodyVar;

  // Prepend the HTML into the body
  Office.context.mailbox.item.body.prependAsync(htmlContent, { coercionType: Office.CoercionType.Html }, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('HTML body prepend successfully');
      } else {
          console.error(`Failed to prepend HTML body. Error: ${result.error.message}`);
      }
  });
}





function sendSorry() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hi!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}



function sendApp() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>APP!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}



function sendVendor() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>VENDORi!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}




function sendRentalResponse() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Rental Responsei!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}




function sendPayslip() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Still by check?i!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}





function sendVendor() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>VENDORi!</p>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}



Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insertSorry").onclick = sendSorry;
    document.getElementById("insertApp").onclick = sendApp;
    document.getElementById("insertVendor").onclick = sendVendor;

    document.getElementById("insertRentalResponse").onclick = sendRentalResponse;
    document.getElementById("insertPayslip").onclick = sendPayslip;
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
