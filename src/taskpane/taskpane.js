/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var templateId = ""

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    registerEvents();
    loadTemplates();
  }
});

function registerEvents() {
  document.getElementById("assignment-temp").onchange = enableInsertBtn
  document.getElementById("insert-button").onclick = insertTemplateToBody
}
//Load Templates
function loadTemplates() {
  document.getElementById("template-list-container").style.display = "flex";
}


function enableInsertBtn() {
  const insertBtn = document.getElementById("insert-button")
  insertBtn.disabled = false
  templateId = this.value
}

function insertTemplateToBody() {
  // let content = document.getElementById("assignment-temp-content").innerHTML
  // Office.context.mailbox.item.body.setSelectedDataAsync(content,
  //   {coercionType: Office.CoercionType.Html}, function(result) {
  //     if (result.status === Office.AsyncResultStatus.Failed) {
  //       showError('Could not insert gist: ' + result.error.message);
  //     }
  // });
  let url = new URI('assignmentQuestion.html').absoluteTo(window.location).toString();
  const dialogOptions = { width: 20, height: 40, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    console.log(settingsDialog)
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
}

function receiveMessage(message) {
  console.log(message)
  settingsDialog.close()
}