/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let istaskpaneOpen: boolean = true;

/* global global, Office, self, window */
(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {};
})();

Office.onReady(async () => {
  // If needed, Office.js is ready to be called
  // ensureStateInitialized(true);
  await Office.addin.onVisibilityModeChanged(function (args) {
    console.info("onVisibilityModeChanged");
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

async function togglepanel() {
  if (istaskpaneOpen) {
    Office.addin.hide().then(() => {
      istaskpaneOpen = false;
      // g.state.istaskpaneOpen = false;
      console.info("in hide promise");
    });
  } else {
    Office.addin.showAsTaskpane().then(() => {
      istaskpaneOpen = true;
      // g.state.istaskpaneOpen = true;
      console.info("in open promise");
    });
  }
}

let g = getGlobal() as any;
// The add-in command functions need to be available in global scope
g.action = action;

// export async function ensureStateInitialized(isOfficeInitializing) {
//   console.log("ensureInitialize called");
//   let g = getGlobal();
//   let initValue = false;
//   if (isOfficeInitializing) {
//     //we are being called in response to Office Initialize
//     if (g.state !== undefined) {
//       if (g.state.isInitialized === false) {
//         g.state.isInitialized = true;
//       }
//     }
//     if (g.state === undefined) {
//       initValue = true;
//     }
//   }

//   if (g.state === undefined) {
//     g.state = {
//       isStartOnDocOpen: false,
//       isSignedIn: false,
//       isTaskpaneOpen: false,
//       isConnected: false,
//       isSyncEnabled: false,
//       isConnectInProgress: false,
//       isFirstSyncCall: true,
//       isSumEnabled: false,
//       isInitialized: initValue,
//       updateRct: () => {},
//       setTaskpaneStatus: (opened) => {
//         g.state.isTaskpaneOpen = opened;
//       },
//       setConnected: (connected) => {
//         g.state.isConnected = connected;

//         if (connected) {
//           if (g.state.updateRct !== null) {
//             g.state.updateRct("true");
//           }
//         } else {
//           if (g.state.updateRct !== null) {
//             g.state.updateRct("false");
//           }
//         }
//       },
//     };

//     //track startup behavior
//     if (g.state.isInitialized) {
//       let addinState = await Office.addin.getStartupBehavior();
//       console.log("load state is:");
//       console.log("load state" + addinState);
//       if (addinState === Office.StartupBehavior.load) {
//         g.state.isStartOnDocOpen = true;
//       }
//     }

//     //track sign in status
//     // if (localStorage.getItem("loggedIn") === "yes") {
//     //   g.state.isSignedIn = true;
//     // }
//   }
//   // if (g.state.isInitialized) {
//   //   updateRibbon();
//   // }
// }

Office.actions.associate("togglePanel", togglepanel);
