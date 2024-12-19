/*
* Copy-paste rendered HTML into the email body with styling (no raw HTML tags).
*/

Office.onReady();

async function loadSignatureFromFile() {
    const filePath = "https://siggy.wearelegence.com/users/corey.gashlin@wearelegence.com.html";
    try {
        const response = await fetch(filePath, { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
        }
        return await response.text(); // Raw HTML content
    } catch (error) {
        console.error("Error fetching HTML file:", error);
        return null;
    }
}

function extractRenderedContentWithStyles(html) {
    const container = document.createElement("div");
    container.style.visibility = "hidden"; // Ensure it's not visible
    container.innerHTML = html;
    document.body.appendChild(container);

    // Traverse the DOM to extract styled content
    const styledText = extractTextWithInlineStyles(container);
    document.body.removeChild(container);

    return styledText;
}

function extractTextWithInlineStyles(element) {
    const traverseAndExtract = (node) => {
        if (node.nodeType === Node.TEXT_NODE) {
            return node.textContent; // Extract text content
        }

        if (node.nodeType === Node.ELEMENT_NODE) {
            const styles = window.getComputedStyle(node);
            const color = styles.color;
            const fontWeight = styles.fontWeight;
            const fontStyle = styles.fontStyle;
            const fontSize = styles.fontSize;

            // Wrap text with inline styles
            const styledText = Array.from(node.childNodes)
                .map(traverseAndExtract)
                .join("");

            return `<span style="color:${color}; font-weight:${fontWeight}; font-style:${fontStyle}; font-size:${fontSize};">${styledText}</span>`;
        }

        return ""; // Ignore other node types
    };

    return traverseAndExtract(element);
}

async function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();

    const rawHtmlSignature = await loadSignatureFromFile();
    if (!rawHtmlSignature) {
        console.error("Failed to load the signature.");
        event.completed();
        return;
    }

    if (platform.includes("android")) {
        console.log("Running Android-specific logic...");

        const styledContent = extractRenderedContentWithStyles(rawHtmlSignature);
        item.body.setSignatureAsync(styledContent, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for Android",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("androidNotification", notification);

    } else {
        console.log("Running non-Android logic...");

        const styledContent = extractRenderedContentWithStyles(rawHtmlSignature);
        item.body.setSignatureAsync(styledContent, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for non-Android platform",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("nonAndroidNotification", notification);
    }
}
