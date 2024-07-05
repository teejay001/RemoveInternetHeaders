/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

let userDialog = null;
let sendEvent = null;

async function onMessageAttachmentsChangedHandler(event) {
  if (event.attachmentStatus === 'added') {
    // const data = `${event.attachmentDetails.name} of size ${event.attachmentDetails.size} has been added ... `;
    setItemInternetHeaders(event.attachmentDetails.name, event.attachmentDetails.size);
    event.completed({ allowEvent: true });
  } else if (event.attachmentStatus === 'removed') {
    // const data = `${event.attachmentDetails.name} of size ${event.attachmentDetails.size} has been removed ... `;
    removeItemInternetHeaders();
    event.completed({ allowEvent: true });
  }
}

function messageHandler(arg) {
  console.log('messageHandler', arg);
  removeItemInternetHeaders();

  userDialog?.close();
  sendEvent.completed({ allowEvent: true });
}

function processOnSend(event) {
  Office.context.mailbox.item?.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error);
      }

      if (asyncResult.value.toLowerCase().includes('prompt')) {
        // Display a dialogue
        sendEvent = event;
        Office.context.ui.displayDialogAsync(`${window.location.origin}/user.html`, {height: 30, width: 20},
          function (asyncResult) {
            dialog = asyncResult.value;
            userDialog = dialog;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
          }
        );
      } else {
        event.completed({ allowEvent: true });
      }
  });
}

function onMessageSendHandler(event) {
  processOnSend(event);
}

function prependToMessageBody(text) {
	Office.context.mailbox.item?.body.prependAsync(text, {}, (asyncResult) => {
		if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
			console.error(`Failed to set body: ${JSON.stringify(asyncResult.error)}`);
		} else {
      console.info(`message body updated successfully`);
    }
	});
}

function setItemInternetHeaders(name, size) {
  Office.context.mailbox.item.internetHeaders.setAsync({ 'header-name': name, 'header-size': size.toString() }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully set headers");
    } else {
      console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
    }
  });
}

function removeItemInternetHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(['header-size'], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully removed selected headers");
    } else {
      console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
    }
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
}