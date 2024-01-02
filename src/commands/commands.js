/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
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

function openDialog(event) {
  let url = new URI('assignmentQuestion.html').absoluteTo(window.location).toString();
  const dialogOptions = { width: 20, height: 40, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
    settingsDialog = result.value;
    settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
}

function receiveMessage(message) {
  console.log(message)
  settingsDialog.close()
}

// Register the function.
Office.actions.associate("openDialog", openDialog);
