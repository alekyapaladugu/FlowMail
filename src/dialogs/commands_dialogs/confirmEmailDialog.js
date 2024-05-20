(function () {
  "use strict";

  Office.onReady(function () {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent, onRegisterMessageComplete);
    $(document).ready(function () {
      registerEvents();
   }) 
  });

  setTimeout(()=>{
    Office.context.ui.messageParent('IAmReady')
  }, 1000);
  

  function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.error.message);
      return;
    } 
  }

  function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    console.log('messaged received', messageFromParent.mail)
    document.getElementById("generated-mail").value  = messageFromParent.mail;
    document.getElementById('suggestions-received').innerHTML = messageFromParent.suggestions;
  }

  function registerEvents() {
    console.log('inside register')
    document.getElementById("confirm-form").onsubmit = confirmForm;
    document.getElementById("generated-mail").onchange = getChangedEmail;
  }

  function confirmForm() {
    let data = document.getElementById("generated-mail").value
    data = data.replace(/\n/g, '<br>')
    Office.context.ui.messageParent(data);
  }

  function getChangedEmail(e) {
    document.getElementById("generated-mail").value = e.target.value
  }

})();