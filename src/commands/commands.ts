//***********************************************************************************************************
//   CONFIDENTIAL
//
//   COPYRIGHT 2018 - 2021
//   Enscape Solutions, LLC DBA BlueCube Energy
//   All Rights Reserved
//
//   NOTICE:  All information contained herein is, and remains the property of Enscape Solutions.
//   The intellectual and technical concepts contained herein are proprietary to Enscape Solutions
//   and are protected by trade secret or copyright law. Dissemination of this information or
//   reproduction of this material is strictly prohibited unless prior written permission is obtained
//   from Enscape Solutions.
//
//   Included third party assets:
//   Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License.
//
//   CONFIDENTIAL
//***********************************************************************************************************

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
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

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action;
