/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
const { default: axios } = require("axios");
let settingsDialog;
let mail;
let suggestions;
let subject;

Office.onReady(() => {s
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  getSelectedText().then(function () {
    event.completed();
  });
}

function getSelectedText() {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
        const text = asyncResult.value;
        // Call the API function to generate mail
        try {
          await generateMail(text);
          openDialog();
        } catch (error) {
          reject(error);
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

async function generateMail(text) {
  try {
    const response = await axios.post('http://127.0.0.1:5000/api/generateMail', 
    {
      "text": text,
    }, 
    {
      headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
      }
    });
    
    console.log(response.data);
    if(response.data[0] !== undefined) {
      mail = response.data[0]["mail"]
      suggestions = response.data[0]["suggestions"]
    } else {
      mail = response.data["mail"]
      suggestions = response.data["suggestions"]
      subject = response.data["subject"]
    }
  
    suggestions = suggestions.replace(/\n/g, '<br>')

  } catch(error) {
    console.log(error)
  }
    
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

function openDialog() {
  let url = new URI('confirmEmailDialog.html').absoluteTo(window.location).toString();
  const dialogOptions = { width: 60, height: 60, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>{
      if (arg.message === "IAmReady"){
        if (Office.context.requirements.isSetSupported('DialogApi', '1.2')) {
          settingsDialog.messageChild(JSON.stringify({
            "mail": mail,
            "suggestions": suggestions
          }), { targetOrigin: url });
        }
      }
      if (arg.message !== "IAmReady"){
        receiveMessage(arg.message)
      }
    });
  });
}

function receiveMessage(message) {
  Office.context.mailbox.item.subject.setAsync(subject,
    (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }
    });
  Office.context.mailbox.item.body.setAsync(message,
    {coercionType: Office.CoercionType.Html}, function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        showError('Could not insert email: ' + result.error.message);
      }
  });
  settingsDialog.close()
}
