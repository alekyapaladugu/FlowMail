/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var templateId = ""
let settingsDialog;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    registerEvents();
    loadTemplates();
  }
});

function registerEvents() {
  const radioButtons = document.querySelectorAll('input[name="email-template"]');
  const insertBtn = document.getElementById("insert-button")
  insertBtn.onclick = insertTemplateToBody
  // Add event listener to each radio button
  radioButtons.forEach(function (radioButton) {
    radioButton.addEventListener('change', function () {
      // Enable the insert button when a radio button is selected
      templateId = this.value
      loadTemplateContent();
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
  let url = new URI('assignmentQuestion.html').absoluteTo(window.location).toString();
  url = url + '?templateId=' + templateId;
  const dialogOptions = { width: 35, height: 50, displayInIframe: true, promptBeforeOpen: false };
  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    // settingsDialog.messageChild(JSON.stringify({template_id: templateId}))
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
}

function receiveMessage(dialogOutput) {
  let msg = JSON.parse(dialogOutput.message)
  populateTemplateBody(msg)
  let content = document.getElementById("template-content").innerHTML
  Office.context.mailbox.item.body.setSelectedDataAsync(content,
    {coercionType: Office.CoercionType.Html}, function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        showError('Could not insert template: ' + result.error.message);
      }
  });
  
  settingsDialog.close()
}

function loadTemplateContent() {
  let url = '';
  if(templateId === "template-1") {
    url = "assignmentQuestionTemplate.html"
  } else if(templateId === "template-2") {
    url = "gradingQueryTemplate.html"
  } else if(templateId === "template-3") {
    url = "crashCourseTemplate.html"
  } else if(templateId === "template-4") {
    url = "lateSubmissionTemplate.html"
  } else if(templateId === "template-5") {
    url = "topicClarification.html"
  }
  if(url) {
    $("#template-content").load(url, function (response, status, xhr) {
      if (status == "error") {
        var msg = "Sorry but there was an error: ";
        $("#error").html(msg + xhr.status + " " + xhr.statusText);
      }
    });
  }
  
}

function populateTemplateBody(msg) {
  let pname = msg['professorName']
  document.getElementById('professor-name').innerHTML = pname!==undefined ? pname : ''
  let courseNo = msg['courseNumber']
  document.getElementById('course-no').innerHTML = courseNo!==undefined ? courseNo : ''

  if(templateId!=="template-3" && templateId!=="template-5") {
    let assignmentName = msg['assignmentName']
    document.getElementById('assignment-name').innerHTML = assignmentName!==undefined ? assignmentName : ''
    if(templateId!=="template-4") {
      let assignmentQuestion = msg['assignmentQuestion']
      document.getElementById('assignment-question').innerHTML = assignmentQuestion!==undefined ? assignmentQuestion : ''
    }
  }
  
  let studentName = msg['studentName']
  document.getElementById('student-name').innerHTML = studentName!==undefined ? studentName : ''
  let studentId = msg['studentId']
  document.getElementById('student-id').innerHTML = studentId!==undefined ? studentId : ''
  if(templateId === "template-3") {
    let onBaseNumber = msg['onbaseNumber']
    document.getElementById('onbase-no').innerHTML = onBaseNumber!==undefined ? onBaseNumber : ''
  }
  if(templateId === "template-4") {
    let extendDate = msg['extendDate']
    let extendReason = msg['extendReason']
    document.getElementById('extension-date').innerHTML = extendDate!==undefined ? extendDate : ''
    document.getElementById('extension-reason').innerHTML = extendReason!==undefined ? extendReason : ''
  }
  if(templateId === "template-5") {
    let topicName = msg['topicName']
    let topicQuestion = msg['topicQuestion']
    document.getElementById('topic').innerHTML = topicName!==undefined ? topicName : ''
    document.getElementById('topic-question').innerHTML = topicQuestion!==undefined ? topicQuestion : ''
  }
}
