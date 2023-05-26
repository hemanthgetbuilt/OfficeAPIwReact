/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let istaskpaneOpen: boolean;

/* global global, Office, self, window */
(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {};
})();

Office.onReady(async () => {
  Office.addin.onVisibilityModeChanged(function (args) {
    istaskpaneOpen = args.visibilityMode === "Taskpane"; // Update count on subsequent opens.
  });
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
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

async function togglepanel(event) {
  if (istaskpaneOpen) {
    Office.addin.hide().then(() => {
      istaskpaneOpen = false;
    });
  } else {
    Office.addin.showAsTaskpane().then(() => {
      istaskpaneOpen = true;
    });
  }
  event.completed();
}

let g = getGlobal() as any;
// The add-in command functions need to be available in global scope
g.action = action;

Office.actions.associate("togglepanel", togglepanel);
