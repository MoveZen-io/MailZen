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