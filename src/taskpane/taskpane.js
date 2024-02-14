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
  const radioButtons = document.querySelectorAll('input[name="assignment-temp"]');
  const insertBtn = document.getElementById("insert-button")
  insertBtn.onclick = insertTemplateToBody
  // Add event listener to each radio button
  radioButtons.forEach(function (radioButton) {
    radioButton.addEventListener('change', function () {
      // Enable the inser6 button when a radio button is selected
      templateId = this.value
      enableInsertBtn(insertBtn);
    });
  });
}
//Load Templates
function loadTemplates() {
  document.getElementById("template-list-container").style.display = "flex";
}


function enableInsertBtn(insertBtn) {
  insertBtn.disabled = false
}

function insertTemplateToBody() {
  let url = ""
  if(templateId === "assignment-temp-1") {
    url = new URI('assignmentQuestion.html').absoluteTo(window.location).toString();
  }
  const dialogOptions = { width: 35, height: 50, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    console.log(settingsDialog)
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
}

function receiveMessage(dialogOutput) {
  let msg = JSON.parse(dialogOutput.message)
  if(templateId === "assignment-temp-1") {
    populateBody_Temp1(msg)
  }
  let content = document.getElementById("assignment-temp-content").innerHTML
  Office.context.mailbox.item.body.setSelectedDataAsync(content,
    {coercionType: Office.CoercionType.Html}, function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        showError('Could not insert template: ' + result.error.message);
      }
  });
  
  settingsDialog.close()
}

function populateBody_Temp1(msg) {
  let pname = msg['professorName']
  document.getElementById('professor-name').innerHTML = pname!==undefined ? pname : ''
  let courseNo = msg['courseNumber']
  document.getElementById('course-no').innerHTML = courseNo!==undefined ? courseNo : ''
  let assignmentName = msg['assignmentName']
  document.getElementById('assignment-name').innerHTML = assignmentName!==undefined ? assignmentName : ''
  let assignmentQuestion = msg['assignmentQuestion']
  document.getElementById('assignment-question').innerHTML = assignmentQuestion!==undefined ? assignmentQuestion : ''
  let studentName = msg['studentName']
  document.getElementById('student-name').innerHTML = studentName!==undefined ? studentName : ''
  let studentId = msg['studentId']
  document.getElementById('student-id').innerHTML = studentId!==undefined ? studentId : ''
}