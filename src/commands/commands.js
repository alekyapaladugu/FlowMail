/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
const { default: axios } = require("axios");
const {OpenAI} = require("openai");
let settingsDialog;

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  getSelectedText().then(function (selectedText) {
    Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
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
          const generatedMail = await generateMail(text);
          // const confirmedMail = openDialog();
          // resolve(confirmedMail);
          resolve(generatedMail);
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
          "Access-Control-Allow-Origin": "*",
      }
    });
    
    console.log(response.data);
    if(response.data[0] !== undefined) {
      localStorage.setItem('Email', response.data[0]["mail"])
      localStorage.setItem('Suggestions', response.data[0]["suggestions"])
      return response.data[0]["mail"]
    } else {
      localStorage.setItem('Email', response.data["mail"])
      localStorage.setItem('Suggestions', response.data["suggestions"])
      return response.data["mail"]
    }

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
  const dialogOptions = { width: 80, height: 80, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
}

function receiveMessage(message) {
  settingsDialog.close()
  return message
}
