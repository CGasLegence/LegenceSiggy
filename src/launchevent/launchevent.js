/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

// Add start-up logic code here, if any.
Office.onReady();

async function loadSignatureFromFile() {
    // Dynamically build the file path with the user's email
    const filePath = "https://siggy.wearelegence.com/users/corey.gashlin@wearelegence.com.html";
    try {
        const response = await fetch(filePath, { cache: "no-store" }); // no-store ensures no caching
        if (!response.ok) {
            throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
        }
        return response.text();
    } catch (error) {
        console.error('Error fetching HTML file:', error);
        return null;
    }
}

function renderAndUseHtmlAsText(html) {
    // Create a hidden DOM element to render the HTML
    const container = document.createElement("div");
    container.innerHTML = html;

    // Extract rendered text from the HTML
    return container.innerText;
}

async function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();

    if (platform.includes("android")) {
        // Android-specific logic
        console.log("Running Android-specific logic...");
        const plainTextSignature = "Best regards,\nJohn Doe\nCompany Name"; // Plain text fallback
        item.body.setSignatureAsync(plainTextSignature, { coercionType: Office.CoercionType.Text }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        // Android notification
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Legence Corporate Signature Android Test",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("androidNotification", notification);
    } else {
        // Non-Android logic
        console.log("Running non-Android logic...");
        const htmlSignature = await loadSignatureFromFile(); // Load HTML signature
        if (htmlSignature) {
            item.body.setSignatureAsync(htmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                }
                event.completed();
            });
        }

        // Non-Android notification
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Legence Corporate Signature Non-Android Test",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("nonAndroidNotification", notification);
    }
}
