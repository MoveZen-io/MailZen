
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





function sdRefund() {
  Office.context.mailbox.item.subject.setAsync("Security Deposit Refund - [Property Address]");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; We hope you're doing well. Your security deposit disposition is attached to this email. The refund process for your deposit has been initiated, and a check will be sent to the forwarding address you provided, which can be found in the disposition document.</p><br>

<p>If a financially responsible party provided their account information and you're entitled to a full refund, it will be issued via E-Check, and an additional email will be sent to the account holder.</p><br>

<p>If you have any questions or concerns regarding your security deposit refund, please feel free to reach out to us. We also want to take this moment to express our appreciation for your tenancy with MoveZen Property Management. We wish you the very best in all your future endeavors.</p><br>

<p>Thanks! MoveZen</p>

`);

}




function sdBalance() {
  Office.context.mailbox.item.subject.setAsync("Security Deposit Close Out, Balance Due [Property Address]");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(` <p>Hello from MoveZen Accounting,<br />

  <p>You have a balance due for past-due charges and/or move-out charges. Your security deposit has been applied to this balance, but there is still a remaining amount owed. </p><br>

  <p>The balance due is noted on the attached disposition.  Please be advised that this balance must be paid in full within the next 30 days. If the balance is not paid in full by this time, the past-due balance will be turned over to collections.</p><br>

  <p>We understand that unforeseen circumstances may arise and that you may be experiencing financial difficulty. If this is the case, please contact us as soon as possible to discuss potential payment arrangements. It is in everyone's best interest to resolve this matter as soon as possible.</p><br>

  <p>Thank you for your prompt attention to this matter. MoveZen</p><br>

  `);

}




function sdTransfer() {
  Office.context.mailbox.item.subject.setAsync("Security Deposit Transfer [Property Address]");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Security Deposit Transfer [Property Address]</p><br>

<p>We hope you're well. We're reaching out to inform you that your security deposit has been transferred to the property owner. Attached is your disposition for your records.</p><br>

<p>You can now contact the property owner directly for any inquiries or concerns regarding your security deposit. Their contact information is provided below:</p><br>

<p>Owner's Name: [OWNER'S NAME]</p><br>
<p>Owner's Email: [OWNER'S EMAIL]</p><br>
<p>Owner's Phone Number: [OWNER'S PHONE NUMBER]</p><br>
<p>Owner's Mailing Address: [OWNER'S MAILING ADDRESS]</p><br>

<p>We'd also like to take this opportunity to express our appreciation for your tenancy with MoveZen Property Management. We wish you all the best in your future endeavors.</p><br>

<p>Best regards, MoveZen</p><br>


  `);

}





function sdNewMan() {
  Office.context.mailbox.item.subject.setAsync("Security Deposit Transfer [Property Address]");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>

  <p>We hope you're doing well. We wanted to inform you that your security deposit has been transferred to a new management company. They will now handle and manage your security deposit. Here are their details:</p><br>

  <p>[NEW MANAGEMENT COMPANY NAME]</p><br>
  <p>[NEW MANAGEMENT COMPANY EMAIL]</p><br>
  <p>[NEW MANAGEMENT COMPANY PHONE NUMBER]</p><br>
  <p>[NEW MANAGEMENT COMPANY ADDRESS]</p><br>

  <p>For all future inquiries and requests, please reach out to the new management company directly. We suggest contacting them to confirm your deposit's receipt and discuss any potential questions or concerns.</p<br>

  <p>We'd also like to express our appreciation for your tenancy with MoveZen Property Management. We wish you the best in your future endeavors. Thank you for your cooperation during this transition, and we hope you have a positive experience with the new management.</p><br>

  <p>Warm regards, MoveZen</p><br>
  
  `);

}





function filterRemoved() {
  Office.context.mailbox.item.subject.setAsync("We Removed the Filter Charge!");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>

  <p>We hope this message finds you well. We wanted to inform you that we've taken action on your account. Specifically, we've removed one or more HVAC Filter Delivery charges and applied a credit to your account. This credit will automatically offset future rent charges. Please refer to your attached ledger for detailed information.</p><br>

  <p>If you have any questions or require assistance, kindly reach out to your dedicated account manager.</p><br>

  <p>Thanks for choosing MoveZen!</p><br>
  
  `);

}





function payPlan() {
  Office.context.mailbox.item.subject.setAsync("Payment Plan Proposal Request");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>
 
  <p>Thanks for your interest in setting up a payment plan for the balance due. We appreciate your proactive approach and would like to request your payment plan proposal. In order to proceed, we kindly ask you to provide us with the following details:</p><br>

  <p>1. Number of payments</p><br>
  <p>2. Payment amount for each installment</p><br>
  <p>3. Dates on which you intend to make the payments</p><br>

  <p>Once we receive your payment plan proposal, our team will promptly review it and respond with a decision. If you have any questions in the meantime, please let us know. </p><br>

  <p>Regards, MoveZen</p><br>
 
  `);

}





function eCheck() {
  Office.context.mailbox.item.subject.setAsync("E-Check Information - MoveZen Property Management ");
  // Office.context.mailbox.item.body.setAsync("This is the body of the email.");
  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>
  
  <p>We're excited to introduce the option of direct deposit payments for our valued vendors. With E-Check, you'll enjoy faster payments without the wait for paper checks. If this interests you, simply log in to your vendor portal and update your payment method.</p><br>

  <p>If you prefer not to provide your bank information via the portal, please complete the linked form and send it back to us at your convenience.</p><br>

  <p><a href="https://movezen.sharepoint.com/:b:/s/accounting/EZ3b2F0TrYFEp6MVfWadqngB6QM8X_z02f7tQ7dMERXCaw?e=ZatiY7">ACH Form to Submit</a></p><br>

  <p>Should you have any questions or concerns, please feel free to reach out.</p><br>

  <p>Thanks, MoveZen</p><br>
 
  `);

}





function w9Request() {
  Office.context.mailbox.item.subject.setAsync("Information Needed - W9");

  prependHtmlBody(`<p>Hello from MoveZen Accounting,</p><br>
 
  <p>We hope you're doing well. We're reaching out to request a copy of your W9 form for our records. Ensuring our vendor information is accurate and up-to-date is vital for tax compliance and maintaining precise financial records.</p><br>

  <p>Please provide a copy of your W9 form, including your taxpayer identification number (TIN) or social security number (SSN), legal name, and business address. Having this information on hand will facilitate payment processing and help us stay in compliance with tax regulations.</p><br>

  <p>You can access the blank W9 form here: </p><br>

  <p>LINKKKKK</p><br>

  <p>To streamline the process, kindly attach the completed W9 form to your reply. If you have any questions or need assistance, please don't hesitate to email or text us at (910) 795.1668.</p><br>

  <p>Your prompt attention to this matter is appreciated, and we thank you in advance for your cooperation. We value your services and anticipate a continued successful collaboration.</p><br>

  <p>Thanks, MoveZen</p><br>
 
  `);

}






function sendMoveInReminders() {
  Office.context.mailbox.item.subject.setAsync("Move In Info - Thank you for choosing MoveZen!");

  prependHtmlBody(`<table align="center" cellspacing="0" id="m_-5455835037964899298m_-1598750376468129031gmail-m_1119163153852469884bodyTable" style="border-collapse:collapse; height:4165.69px; padding:0px; width:599.965px">
	
			</td>
		</tr>
	</tbody>
</table>`);

}






function sendUtilityNotice() {
  Office.context.mailbox.item.subject.setAsync("It's important you connect utilities before your move, here's a tool to help");

  prependHtmlBody(`<p>Hi!</p><br>

  <p>This is a friendly reminder that it&#39;s time to set up your utilities for your new home.&nbsp;</p><br>
  .theutilityhub.net/partners-page/victory-property-management" target="_blank">Activate Your Utilities Here With Utility Hub</a></p><br><br>
  
  <p><a href="https://movezen360.com/utility-hub-makes-moving-easier/" target="_blank">Read more about Utility Hub</a></p>`);

}






function sendGeneralRentInfo() {
  Office.context.mailbox.item.subject.setAsync("Thanks for your interest. Here's some info to help you get started");

  prependHtmlBody(`<p>Here are some common steps you may find helpful</p><br>

  <p>For questions on the application process, start here.&nbsp; If this doesn&#39;t cover it (it usually does and much more), let us know your specic question and we'll nail it down.&nbsp;<a href="ht
  <br><br>
  
  <p>Please let us know if you have any questions!&nbsp;&nbsp;Customer Service Team</p>`);

}






// function sendTurnoverReserve() {
//   Office.context.mailbox.item.subject.setAsync("Let's nail down the needed turnover funds now to avoid critical delays");

//   prependHtmlBody(`asdf`);

// }





function sendVendorInsur() {
  Office.context.mailbox.item.subject.setAsync("Let's get your insurance updated to avoid critical delays");

  prependHtmlBody(`<p>Hi!</p>

  <p>It looks like yo
  <p>Wilmington, NC 28403&nbsp;<br />
  <br />
  Thanks! MoveZen&nbsp; (we&#39;ll legally change our name in late 2024, this is dba)</p>`);

}





function sendTurnoverReserve() {
  Office.context.mailbox.item.subject.setAsync("Efficiently turning over a rental starts long before your tenant has moved out");

  prependHtmlBody(`<p>Hi!</p><br>
  
  <p>Hope you're doing well.<br><br>
  
  Even for the best 
  <p>&nbsp;</p>
  
  <p>Let us know if you have any questions. Thanks!&nbsp; Customer Service Team</p>`);

}




function sendMoveInspectionRemind() {
  Office.context.mailbox.item.subject.setAsync("The move in inspection is meant for no other reason than to protect you");

  prependHtmlBody(`<p> I hope your move-in went smoothly and you're enjoying the home!</p>

  <p>&nbsp;</p>
  
  <p>We know you&#39;re busy
  
  <p>&nbsp;</p>
  
  <p>Thanks!&nbsp;</p>
  `);

}




function sendComplaintResponse() {
  Office.context.mailbox.item.subject.setAsync("We hear you and we're working on it");

  prependHtmlBody(`<p>Hi</p>

  <p>&nbsp;</p>
  
  <p>Please allow us some time to research things, and we&#39;ll be back in touch soon.</p>
  
  <p>&nbsp;</p>
  
  <p>Thanks! Customer Service Team</p>`);

}




function sendMoveChecklist() {
  Office.context.mailbox.item.subject.setAsync("The most important email we'll send regarding your move out");

  prependHtmlBody(`<p>​Hi!<br />
  <br />
  We made this checkl
  Most importantly, use common sense, don&rsquo;t destroy something in an effort to hide or improve it. &nbsp;If you have questions, ask! &nbsp;We are literally here to help smooth this process out now, rather than fight through it later. &nbsp;We often have great tips for problems that could help you a lot<br />
  <br />
  <br />
  Thanks!</p>`);

}




function sendUnseenDisclaimer() {
  Office.context.mailbox.item.subject.setAsync("Site Unseen Company Warning Disclaimer");

  prependHtmlBody(`<p>​Hi!<br />
  <br />
  Below is our site unseen disclaimer. &nbsp;You&#39;ve probably already been warned, but this is consistently one of the biggest mistakes renters make in our experience. &nbsp;We aren&#39;t going to shoul
  <p><br />
  <br />
  Thanks, Customer Service Team</p>`);

}




function sendReferenceQuestions() {
  Office.context.mailbox.item.subject.setAsync("Rental Reference Questions for a Former Tenant");

  prependHtmlBody(`<p>​Hello, we've' received a rental application for -----------------------------.
  <br />
  <br />
  We wanted to write to ask if
  <br />
  Thanks!</p>`);

}




function sendPlacementInfo() {
  Office.context.mailbox.item.subject.setAsync("Here are the next steps for us to neatly finalize things for you");

  prependHtmlBody(`<table cellspacing="0" style="border-collapse:collapse; font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif; width:100%">
	<tbody>
		<tr>
			<td style="border-bo
				</tbody>
			</table>
			</td>
		</tr>
		<tr>
			<td style="border-bottom:0px; border-top:0px; vertical-align:top">&nbsp;</td>
		</tr>
	</tbody>
</table>`);

}




function sendLeaseBreakNotice() {
  Office.context.mailbox.item.subject.setAsync("Helpful tips and next steps for breaking your lease the safe way");

  prependHtmlBody(`<table align="center" cellspacing="0" style="border-collapse:collapse; max-width:600px; width:600px">
	<tbody>
		<tr>
			<td style="border-bot
			</td>
		</tr>
	</tbody>
</table>`);

}




function sendNoLongHold() {
  Office.context.mailbox.item.subject.setAsync("The home you inquired about isn't available for a pretty good while");

  prependHtmlBody(`<p>Hi!</p>

  <p>&nbsp;</p>
  
  <p><br />
  Thank you&nbsp;very much&nbsp;for your
  <p>We normally list about a month in advance of a home being vacant, however in hot market periods that can shorten up quite a bit</p>
  
  <p>&nbsp;</p>
  
  <p>Hope these tips help a bit in your search.&nbsp; Let us know if we can answer any questions. Thanks!&nbsp;</p>`);

}



function sendPortalResetAll() {
  Office.context.mailbox.item.subject.setAsync("A couple quick portal login steps");

  prependHtmlBody(`<p>​Hi! &nbsp;Sorry you&rsquo;re having trouble getting your portal activated or logged in<br />
  <br />
  The first step is to clear your browser cookies or cache. &nbsp;You can google how to do that relatively easily as it depends on your web browser<br />
  <br />rouble after those two steps, we’ll simply need to delete your login and start from scratch which almost always works, and is fast and simple.  Just reach back out to us and we’ll run that through pretty quickly<br />
  <br />
  <br />
  Thanks!</p>
  `);

}



function sendFieldStaff() {
  Office.context.mailbox.item.subject.setAsync("Thanks for your interest in our field support role!");

  prependHtmlBody(`<p>​Hello!<br />
  <br />
  Thanks for responding to our &nbsp;listing for the field staff 1099 position we have available!<br />
  <br />
  We wanted to ask you a
  We look forward to hearing back from you!<br />
  <br />
  Thanks!</p>`);

}



function sendShowNotice() {
  Office.context.mailbox.item.subject.setAsync("Sorry, it's the dreaded showing notice");

  prependHtmlBody(`<p>​Hi! &nbsp;Yes this is the dreaded showing notice. &nbsp;We hate to bother you, but owners get really hard to deal with if we aren&#39;t making headway to ensure they don&#39;t go a long period of time with a mortgage and no income. That&#39;s the last thing you want before we have to report your move out to them. &nbsp;In fact, the number 1 way to ensure an owner isn&#39;t a pain after move out, is to have a replacement lined up and moving in not too long after you. They are dramatically less concerned about your potential charges 
  We will be in touch and appreciate your patience and understanding<br />
  <br />
  Thanks!</p>`);

}



function sendMoveInInfo() {
  Office.context.mailbox.item.subject.setAsync("Final Move In Instructions");

  prependHtmlBody(`<p>Hi!</p>

  <p>&nbsp;</p>
  
  <p>I hope everything is coming together smoothly for your move!&nbsp; I wanted to send you a note to cover the move-in process to ensure there are no hiccups on move-in day!&nbsp;</p>
  <p>&nbsp;</p>
  
  <p>Thanks!&nbsp;&nbsp;</p>`);

}



function sendRenewalIntro() {
  Office.context.mailbox.item.subject.setAsync("Time is growing short to renew your lease");

  prependHtmlBody(`<p>Hi &nbsp;</p>

  <p>Hope you&rsquo;re doing well today&nbsp;</p>
  
  <p>&nbsp;</p>
  
  <p>We understand renewing or moving decisions are never fun, but those decisions must be made. For a rental owner though, waiting and wondering if your resident will extend can be really difficult, 
  <p>Thanks so much for your understanding. We appreciate you being a part of our MoveZen community!&nbsp;</p>`);
};



Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insertsdRefund").onclick = sdRefund;
    document.getElementById("insertsdBalance").onclick = sdBalance;
    document.getElementById("insertsdTransfer").onclick = sdTransfer;
    document.getElementById("insertsdNewMan").onclick = sdNewMan;
    document.getElementById("insertfilterRemoved").onclick = filterRemoved;
    document.getElementById("insertpayPlan").onclick = payPlan;
    document.getElementById("eCheck").onclick = eCheck;
    document.getElementById("w9Request").onclick = w9Request;
  }
});