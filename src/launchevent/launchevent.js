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
        return await response.text(); // Return the raw HTML content
    } catch (error) {
        console.error('Error fetching HTML file:', error);
        return null;
    }
}

function renderHtmlToBody(html) {
    // Render the HTML content in a hidden container
    const container = document.createElement("div");
    container.style.display = "none"; // Prevent display
    container.innerHTML = html;

    // Extract innerHTML to simulate rendered HTML output (ready for insertion)
    document.body.appendChild(container);
    const renderedContent = container.innerHTML; // Keep the rendered styles
    document.body.removeChild(container);

    return renderedContent;
}

async function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();

    console.log(`Detected platform: ${platform}`);

    // Load the signature HTML file
    const rawHtmlSignature = await loadSignatureFromFile();
    if (!rawHtmlSignature) {
        console.error("Failed to load the signature.");
        event.completed();
        return;
    }

    if (platform.includes("android")) {
        console.log("Running Android-specific logic...");
        const renderedSignature = renderHtmlToBody(rawHtmlSignature);

        // Set the "rendered" signature as plain text (closest Android support allows)
        item.body.setSignatureAsync(renderedSignature, { coercionType: Office.CoercionType.Text }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        // Notify user
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for Android",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("androidNotification", notification);

    } else {
        console.log("Running non-Android logic...");
        item.body.setSignatureAsync(rawHtmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        // Notify user
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for non-Android platform",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("nonAndroidNotification", notification);
    }
}
