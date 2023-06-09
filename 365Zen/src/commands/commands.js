
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
  Office.context.mailbox.item.subject.setAsync("How to get your MoveZen application moving forward fast!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(` <p>Thank you for beginning the application process!<br />
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
  Office.context.mailbox.item.subject.setAsync("Thanks for your interest in working with us!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>​Hello!<br />
  <br />
  Thanks for your interest in becoming a vendor with Victory Property Management. Attached, you&#39;ll find an information packet and some forms for your reference. In order to complete your vendor onboardings, we will need the following:<br />
  &nbsp;</p>
  
  <ul>
    <li><a href="https://movezen.sharepoint.com/:b:/s/Teams/EfwwnpuUK9FIjAvfRLb6kcwBz8TEiRRt7bNnXbOTFVQJ9w?e=bZxTYG">Learn who we are</a></li>
    <li>Provide your business/individual name</li>
    <li><a href="https://www.irs.gov/pub/irs-pdf/fw9.pdf">Submit a signed W9 form</a></li>
    <li>Provide a certificate of insurance listing Victory Real Estate Inc as an additional insured/interest at the following address: 4002 1/2 Oleander Dr. Suite 1A, Wilmington, NC 28403</li>
    <li>Share your mailing address and contact information</li>
    <li>Describe your field of work and the service area you cover</li>
  </ul>
  
  <p><br />
  Please let us know if you have any questions. We look forward to working with you!<br />
  <br />
  Thanks, Customer Service Team</p>
  `);

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
  prependHtmlBody(`<p>Still Paying Rent With a Check or Money Order?</p><br>

  <p>We have a much better solution using a payment process through your nearby WalMart or CVS!&nbsp; Among others</p>

  <p>&nbsp;</p>
  
  <p>All you do is walk in with a barcode that we would provide to you by text or email, they scan it, you pay, and your rent is instantly&nbsp;funded and will&nbsp;show up on our end that way. It&#39;s the best way to avoid late fees</p>
  
  <p>&nbsp;</p>
  
  <p>More importantly, it&#39;s a lot cheaper than buying multiple money orders</p>
  
  <p>&nbsp;</p>
  
  <p>Finally it&#39;s low risk. With WalMart you know it&#39;s legitimate and safe. Just keep your receipt, and it&#39;s safer&nbsp;than a money order also</p>
  
  <p>&nbsp;</p>
  
  <p>Contact us today and we&#39;ll send yours out!&nbsp; Customer Service Team</p>
  `);

}





function sendApprovedNotice() {
  Office.context.mailbox.item.subject.setAsync("Your MoveZen Application is Approved!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Thank you again for your interest in partnering with us and the homeowner</p><br>

  <p>We are starting the process of getting you set up to move in! &nbsp;From here we need the first month&#39;s rent (nonrefundable) hold payment made within 48 hours. &nbsp;<br />
  <br />
  We&#39;ll collect your security deposit just prior to move in, or you can pay both now. &nbsp;In most cases only the first month&#39;s rent will be non-refundable should you not be able to move forward, so any additional payments are refundable, deposits, pet fees, etc. &nbsp;If we hold a home longer than 45 days the daily prorate hold charges will apply from day 46 on<br />
  <br />
  Your hold payment will ensure that no one else can get the property, a signed lease doesn&#39;t. &nbsp;You have to pay the full hold payment to fully secure the property. You&#39;ll receive a copy of your lease in just a bit, and it must be signed within 3 days or you risk losing the home. &nbsp;No home is completely secured until we have a signed agreement, and consideration (payment)<br />
  <br />
  From this point forward we&#39;ll send several very important emails and texts, so be sure to safelist us or keep an eye on your junk folder. &nbsp;Your tenant portal invite will arrive momentarily, and you are welcome to pay online via your portal IF your move-in is more than 7 days away from your payment date, otherwise, you must per company policy pay by certified check / money order delivered to our office or staff. &nbsp;There are never exceptions to that rule due to significant fraud risks. Please do not miss it. If you pay online, you would have to pay again in certified funds to move in within 7 days, and we&#39;d credit the uncleared funds to the next month<br />
  <br />
  <a href="https://movezen.sharepoint.com/:b:/s/Teams/Ed6cnS8o6Q5ChC7kduknd2oBhPoqj51MZdUXYuipKqtySQ?e=d5nokc">Consult the &quot;rules of the road&quot; if you have additional questions</a><br />
  <br />
  We&#39;re looking forward to having you! &nbsp;Customer Service Team</p>`);

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
