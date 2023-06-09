
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
  Office.context.mailbox.item.subject.setAsync("Thanks for contacting us");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hi!</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we've' accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p><br>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p><br>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}



function sendApp() {
  Office.context.mailbox.item.subject.setAsync("You've begun the application journey. Next steps");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>How to get your MoveZen application moving forward fast!</p><br>

  <p>Thank you for beggining the application process!<br />
  <br />
  Our application process is broken down to two main parts. Payment, and background information submission. &nbsp;You will need to get us both of those before we can move forward. Here&#39;s how..<br />
  <br />
  If you haven&#39;t submitted your application payment yet, you can do so with the link below:<br />
  <a href="https://movezen.sharepoint.com/:b:/s/acctmanagers/EZgzJDZR3LhMkdV820qNJs4BohN5zw1amoXeB2OuWVlikA?e=yL0V7q">Just to pay the $79 application&nbsp;fee.</a>&nbsp; This is the last step if you&nbsp;haven&#39;t paid, but have submitted your personal information</p>
  
  <p>&nbsp;</p>
  
  <p>If you have paid, but haven&#39;t submitted your legal information for our background check,&nbsp;<a href="https://victoryre.appfolio.com/listings/">you can do that here</a><br />
  &nbsp;<br />
  We&#39;re excited to get started with finalizing! &nbsp;Please, it&#39;s very important that you review this&nbsp;introductory information to make sure we hit the ground running with aligned expectations.<br />
  <br />
  We&#39;ll often email or text important questions and info which is often sent to spam or promotion folders, so whitelist us or regularly check those folders<br />
  <br />
  Below you&#39;ll find a sample lease &amp; the rules of the road which you have hopefully already reviewed, if not you must do so as they will be included with your lease. &nbsp;<br />
  <br />
  You&#39;ll also find crucial information on how to address pre-move in repairs and if you hope to get a pet during your time in the home. &nbsp;<br />
  <br />
  <a href="https://movezen.sharepoint.com/:b:/s/acctmanagers/EZgzJDZR3LhMkdV820qNJs4BohN5zw1amoXeB2OuWVlikA?e=cPpkVQ">Victory Rules &amp; Regulations Link</a><br />
  <br />
  <a href="https://movezen.sharepoint.com/:b:/s/acctmanagers/EUGHc7hjjnlMvpxiqDjwhK4BMbfvsSwl5YqoA7JIHH0xAQ?e=8ELHHz">Sample NC lease</a></p>`);

}



function sendVendor() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>VENDORi!</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p><br>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p><br>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}




function sendRentalResponse() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Rental Responsei!</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p><br>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p><br>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}




function sendPayslip() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Still by check?i!</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p><br>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p><br>

<p><a href="https://twitter.com/victoryrealty" target="_blank">Follow us on Twitter to Receive Automatic Updates on Price Reductions and New Listings</a></p>`);

}





function sendVendor() {
  Office.context.mailbox.item.subject.setAsync("Hello World!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>VENDORi!</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Thank you for your interest in this property.&nbsp; Unfortunately, we have accepted an application on the property and should be collecting the deposit within a day or so.&nbsp; I&#39;ll certainly keep your contact information nearby in the event that something comes up with our current applicant.&nbsp; Thanks again, and sorry for the inconvenience. Please follow us on your favorite social media platform and be the first to hear about new listings as they come available. Thanks! Customer Service Team</p><br>

<p><a href="https://www.facebook.com/VictoryPropertyManagement" target="_blank">Follow us on Facebook to Receive Automatic Updates on Price Reductions and New Listings</a></p><br>

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
